#!/usr/bin/env python3
"""Build native pivots/charts into a copied course workbook using an explicit analysis plan.

This edits OOXML parts selectively because general spreadsheet exporters do not
preserve this course's DrawingML text boxes and native pivot-chart extensions.
Requires lxml. Native Excel visual and interaction validation remains necessary.
"""
import argparse
from collections import defaultdict
from copy import deepcopy
import json
from math import ceil
from pathlib import Path
import re
import sys
import zipfile
from lxml import etree as X
from workbook_audit import Package, NS, local, rich_text, inspect_workbook
from chart_theme import theme_workbook

S,A,C,R,D = (NS[k] for k in ('s','a','c','r','xdr'))
P='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
C14='http://schemas.microsoft.com/office/drawing/2007/8/2/chart'
FORMS={'line','pie','column','bar','stacked-column','stacked-bar','percent-column','percent-bar','percent-area'}


def el(ns,tag,**attrs):
    return X.Element('{'+ns+'}'+tag,**{k:str(v) for k,v in attrs.items()})


def sub(parent,ns,tag,**attrs):
    child=el(ns,tag,**attrs);parent.append(child);return child


def xml(root):
    return X.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)


def col(index):
    result=''
    while index:
        index,mod=divmod(index-1,26);result=chr(65+mod)+result
    return result


def relpath(part):
    p=Path(part);return str(p.parent/'_rels'/(p.name+'.rels'))


def relations(entries):
    r=X.Element('{'+P+'}Relationships',nsmap={None:P})
    for rid,kind,target in entries:
        sub(r,P,'Relationship',Id=rid,Type=R+'/'+kind,Target=target)
    return r


def cell(parent,address,value,style=0):
    c=sub(parent,S,'c',r=address,s=style)
    if isinstance(value,(int,float)):
        sub(c,S,'v').text=str(value)
    else:
        c.set('t','inlineStr');sub(sub(c,S,'is'),S,'t').text='' if value is None else str(value)
    return c


def load_records(pkg):
    cells=pkg.cells(pkg.sheets['案例数据'])
    headers={re.match('[A-Z]+',r)[0]:v['value'] for r,v in cells.items() if re.search(r'\d+$',r)[0]=='1'}
    if len(set(headers.values()))!=len(headers):raise ValueError('Source headers must be unique')
    records=defaultdict(dict)
    for ref,c in cells.items():
        m=re.fullmatch(r'([A-Z]+)(\d+)',ref)
        if int(m[2])>1 and m[1] in headers:records[int(m[2])][headers[m[1]]]=c['value']
    records=[row for _,row in sorted(records.items()) if any(v is not None for v in row.values())]
    if not records:raise ValueError('No case records found')
    return list(headers.values()),records,max(int(re.search(r'\d+$',r)[0]) for r in cells)


class StyleImporter:
    def __init__(self,parts,reference):
        self.root=X.fromstring(parts['xl/styles.xml']);self.ref=X.fromstring(reference.parts['xl/styles.xml'])
        self.maps={k:{} for k in ['fonts','fills','borders','cellStyleXfs','cellXfs','numFmts']}
    def group(self,name):
        g=self.root.find('{'+S+'}'+name)
        if g is None:
            g=el(S,name,count=0)
            order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
            index=next((i for i,e in enumerate(self.root) if order.index(local(e.tag))>order.index(name)),len(self.root))
            self.root.insert(index,g)
        return g
    def add(self,group,index):
        index=int(index)
        if index in self.maps[group]:return self.maps[group][index]
        node=deepcopy(self.ref.find('{'+S+'}'+group)[index])
        if group in ('cellStyleXfs','cellXfs'):
            for attr,target in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
                if attr in node.attrib:node.set(attr,str(self.add(target,node.get(attr))))
            if 'xfId' in node.attrib:node.set('xfId',str(self.add('cellStyleXfs',node.get('xfId'))))
            num=int(node.get('numFmtId','0'))
            if num>=164:
                old=next((e for e in self.ref.findall('s:numFmts/s:numFmt',NS) if int(e.get('numFmtId'))==num),None)
                if old is not None:
                    if num not in self.maps['numFmts']:
                        new=max([163]+[int(e.get('numFmtId')) for e in self.group('numFmts')])+1
                        copied=deepcopy(old);copied.set('numFmtId',str(new));self.group('numFmts').append(copied)
                        self.group('numFmts').set('count',str(len(self.group('numFmts'))));self.maps['numFmts'][num]=new
                    node.set('numFmtId',str(self.maps['numFmts'][num]))
        g=self.group(group);new=len(g);g.append(node);g.set('count',str(len(g)));self.maps[group][index]=new
        return new
    def from_cell(self,reference,sheet,ref):
        old=reference.cells(reference.sheets[sheet]).get(ref,{}).get('style','0')
        return self.add('cellXfs',old)


def cache_parts(headers,records,lastrow):
    values=[];indexes=[]
    root=X.Element('{'+S+'}pivotCacheDefinition',nsmap={None:S,'r':R},recordCount=str(len(records)),createdVersion='6',refreshedVersion='6',minRefreshableVersion='3',saveData='1',refreshOnLoad='0')
    root.set('{'+R+'}id','rIdRecords')
    sub(sub(root,S,'cacheSource',type='worksheet'),S,'worksheetSource',sheet='案例数据',ref=f'A1:{col(len(headers))}{lastrow}')
    fields=sub(root,S,'cacheFields',count=len(headers))
    for header in headers:
        unique=[];lookup={}
        for row in records:
            value=row.get(header)
            key=(type(value).__name__,value)
            if key not in lookup:lookup[key]=len(unique);unique.append(value)
        values.append(unique);indexes.append(lookup)
        field=sub(fields,S,'cacheField',name=header)
        numeric=all(isinstance(v,(int,float)) for v in unique)
        # Office validates these flags against cached item types, beyond XML schema validity.
        shared=sub(field,S,'sharedItems',count=len(unique),containsString='0' if numeric else '1',containsNumber='1' if numeric else '0',containsSemiMixedTypes='0' if numeric else '1')
        if numeric:
            shared.set('containsInteger','1' if all(float(v).is_integer() for v in unique) else '0')
            shared.set('minValue',str(min(unique)));shared.set('maxValue',str(max(unique)))
        for value in unique:
            if value is None:sub(shared,S,'m')
            else:sub(shared,S,'n' if isinstance(value,(int,float)) else 's',v=value)
    rows=X.Element('{'+S+'}pivotCacheRecords',nsmap={None:S},count=str(len(records)))
    for row in records:
        r=sub(rows,S,'r')
        for i,header in enumerate(headers):sub(r,S,'x',v=indexes[i][(type(row.get(header)).__name__,row.get(header))])
    return root,rows,values


def aggregate(spec,headers,records):
    for name in ['row','value']:
        if spec.get(name) not in headers:raise ValueError(f'Missing {name} field: {spec.get(name)}')
    for name in ['column','filter']:
        if spec.get(name) and spec[name] not in headers:raise ValueError('Missing field: '+spec[name])
    if len({spec[k] for k in ['row','column','filter','value'] if spec.get(k)})!=len([k for k in ['row','column','filter','value'] if spec.get(k)]):raise ValueError('Row, column, filter and value fields must be distinct')
    if any(not isinstance(row.get(spec['value']),(int,float)) for row in records):raise ValueError('Numeric values are unresolved; recalculate source first')
    rowvalues=list(dict.fromkeys(row[spec['row']] for row in records))
    columns=list(dict.fromkeys(row[spec['column']] for row in records)) if spec.get('column') else [spec['value']]
    sums=defaultdict(float)
    for row in records:sums[(row[spec['row']],row[spec['column']] if spec.get('column') else spec['value'])]+=row[spec['value']]
    if spec.get('sort','descending')=='category':rowvalues.sort(key=lambda v:(isinstance(v,str),v))
    elif spec.get('sort','descending')=='descending':rowvalues.sort(key=lambda r:-sum(sums[(r,c)] for c in columns))
    else:raise ValueError('sort must be category or descending')
    matrix=[[sums[(r,c)] for c in columns] for r in rowvalues]
    if spec['chart']=='pie' and (len(columns)!=1 or len(rowvalues)>6):raise ValueError('Use ordinary pie for one dimension and at most six categories; choose another supported form for this plan')
    return rowvalues,columns,matrix


def pivot_table(spec,index,cacheid,headers,shared,rows,columns):
    data_start=5 if spec.get('column') else 4
    totalrow=data_start+len(rows);ncols=len(columns)+(2 if spec.get('column') else 1)
    r=X.Element('{'+S+'}pivotTableDefinition',nsmap={None:S},name=f'CasePT{index:02}',cacheId=str(cacheid),dataCaption='值',updatedVersion='6',minRefreshableVersion='3',createdVersion='6',useAutoFormatting='1',rowGrandTotals='1',colGrandTotals='1')
    attrs={'ref':f'A3:{col(ncols)}{totalrow}','firstHeaderRow':'1','firstDataRow':'2' if spec.get('column') else '1','firstDataCol':'1'}
    if spec.get('filter'):attrs.update(rowPageCount='1',colPageCount='1')
    sub(r,S,'location',**attrs)
    fields=sub(r,S,'pivotFields',count=len(headers))
    for hi,header in enumerate(headers):
        f=sub(fields,S,'pivotField',showAll='0')
        axis='axisRow' if header==spec['row'] else 'axisCol' if header==spec.get('column') else 'axisPage' if header==spec.get('filter') else None
        if header==spec['value']:f.set('dataField','1')
        if axis:
            f.set('axis',axis)
            ordered=rows if axis=='axisRow' else columns if axis=='axisCol' else shared[hi]
            items=sub(f,S,'items',count=len(ordered)+1)
            for value in ordered:sub(items,S,'item',x=shared[hi].index(value))
            sub(items,S,'item',t='default')
    for fieldkey,container,itemcontainer,categories in [('row','rowFields','rowItems',rows),('column','colFields','colItems',columns)]:
        if spec.get(fieldkey):
            sub(sub(r,S,container,count=1),S,'field',x=headers.index(spec[fieldkey]))
            items=sub(r,S,itemcontainer,count=len(categories)+1)
            for i in range(len(categories)):sub(sub(items,S,'i'),S,'x',v=i)
            sub(sub(items,S,'i',t='grand'),S,'x')
    if not spec.get('column'):sub(sub(r,S,'colItems',count=1),S,'i')
    if spec.get('filter'):sub(sub(r,S,'pageFields',count=1),S,'pageField',fld=headers.index(spec['filter']),hier='-1')
    sub(sub(r,S,'dataFields',count=1),S,'dataField',name='求和项:'+spec['value'],fld=headers.index(spec['value']),subtotal='sum',baseField='0',baseItem='0')
    sub(r,S,'pivotTableStyleInfo',name='PivotStyleLight16',showRowHeaders='1',showColHeaders='1',showRowStripes='0',showColStripes='0',showLastColumn='1')
    return r,data_start,totalrow,ncols


def worksheet(spec,rows,columns,matrix,start,totalrow,ncols,styles,index):
    r=X.Element('{'+S+'}worksheet',nsmap={None:S,'r':R})
    sub(r,S,'dimension',ref=f'A1:{col(ncols)}{totalrow}')
    sub(sub(r,S,'sheetViews'),S,'sheetView',workbookViewId='0')
    sub(r,S,'sheetFormatPr',defaultRowHeight='20')
    cs=sub(r,S,'cols');sub(cs,S,'col',min=1,max=1,width='22',customWidth='1');sub(cs,S,'col',min=2,max=max(ncols,2),width='18',customWidth='1')
    data=sub(r,S,'sheetData')
    if spec.get('filter'):
        rr=sub(data,S,'row',r=1);cell(rr,'A1',spec['filter'],styles[0]);cell(rr,'B1','(全部)',styles[1])
    if spec.get('column'):
        rr=sub(data,S,'row',r=3);cell(rr,'A3','求和项:'+spec['value'],styles[0]);cell(rr,'B3','列标签',styles[0])
    header=start-1;rr=sub(data,S,'row',r=header);cell(rr,'A'+str(header),spec['row'],styles[0])
    for i,name in enumerate(columns,2):cell(rr,f'{col(i)}{header}',name,styles[0])
    if spec.get('column'):cell(rr,f'{col(ncols)}{header}','总计',styles[0])
    for i,(name,values) in enumerate(zip(rows,matrix),start):
        rr=sub(data,S,'row',r=i);cell(rr,'A'+str(i),name,styles[1])
        for j,value in enumerate(values,2):cell(rr,f'{col(j)}{i}',value,styles[2])
        if spec.get('column'):cell(rr,f'{col(ncols)}{i}',sum(values),styles[2])
    rr=sub(data,S,'row',r=totalrow);cell(rr,'A'+str(totalrow),'总计',styles[0])
    for i in range(len(columns)):cell(rr,f'{col(i+2)}{totalrow}',sum(v[i] for v in matrix),styles[2])
    if spec.get('column'):cell(rr,f'{col(ncols)}{totalrow}',sum(map(sum,matrix)),styles[2])
    sub(r,S,'pageMargins',left='.7',right='.7',top='.75',bottom='.75',header='.3',footer='.3')
    sub(r,S,'drawing',**{'{'+R+'}id':'rIdDrawing'})
    return r


def chart_cache(parent,tag,formula,values,numeric):
    branch=sub(parent,C,tag);ref=sub(branch,C,'numRef' if numeric else 'strRef');sub(ref,C,'f').text=formula
    cache=sub(ref,C,'numCache' if numeric else 'strCache')
    if numeric:sub(cache,C,'formatCode').text='General'
    sub(cache,C,'ptCount',val=len(values))
    for i,value in enumerate(values):sub(sub(cache,C,'pt',idx=i),C,'v').text=str(value)


def recolor(root,reference):
    theme=X.fromstring(reference.parts['xl/theme/theme1.xml'])
    palette={local(e.tag):next(iter(e)).get('val',next(iter(e)).get('lastClr')) for e in theme.find('.//{'+A+'}clrScheme')}
    alias={'tx1':'dk1','tx2':'dk2','bg1':'lt1','bg2':'lt2'}
    for e in root.findall('.//{'+A+'}schemeClr'):
        color=palette.get(alias.get(e.get('val'),e.get('val')))
        if color and re.fullmatch('[0-9A-Fa-f]{6}',color):e.tag='{'+A+'}srgbClr';e.set('val',color)


def make_chart(reference,spec,index,rows,columns,matrix,start,output_name):
    form=spec['chart']
    base=9 if form=='line' else 16 if form=='percent-area' else 15 if form=='pie' else 10 if 'bar' in form else 11
    root=X.fromstring(reference.parts[f'xl/charts/chart{base}.xml'])
    chart=root.find('{'+C+'}chart');plot=chart.find('{'+C+'}plotArea')
    old=next(e for e in plot if local(e.tag).endswith('Chart'))
    new=el(C,'lineChart' if form=='line' else 'areaChart' if form=='percent-area' else 'pieChart' if form=='pie' else 'barChart')
    if new.tag=='{'+C+'}barChart':sub(new,C,'barDir',val='bar' if 'bar' in form else 'col')
    if form!='pie':sub(new,C,'grouping',val='percentStacked' if form.startswith('percent-') else 'stacked' if form.startswith('stacked-') else 'standard' if form=='line' else 'clustered')
    sub(new,C,'varyColors',val='1' if form=='pie' else '0')
    for i,name in enumerate(columns):
        ser=sub(new,C,'ser');sub(ser,C,'idx',val=i);sub(ser,C,'order',val=i)
        sheet=f'分析内容{index}'
        chart_cache(ser,'tx',f'{sheet}!\u0024{col(i+2)}\u0024{start-1}',[name],False)
        sp=sub(ser,C,'spPr')
        if form=='line':sub(sub(sub(sp,A,'ln',w='28575'),A,'solidFill'),A,'schemeClr',val='accent'+str(i%6+1));sub(sub(ser,C,'marker'),C,'symbol',val='none')
        elif form!='pie':sub(sub(sp,A,'solidFill'),A,'schemeClr',val='accent'+str(i%6+1));sub(sub(sp,A,'ln'),A,'noFill')
        chart_cache(ser,'cat',f'{sheet}!$A${start}:$A${start+len(rows)-1}',rows,False)
        chart_cache(ser,'val',f'{sheet}!\u0024{col(i+2)}\u0024{start}:\u0024{col(i+2)}\u0024{start+len(rows)-1}',[v[i] for v in matrix],True)
    if form=='pie':
        # Latest user preference: direct category + percentage labels inside slices.
        labels=sub(new,C,'dLbls')
        text=sub(labels,C,'txPr');sub(text,A,'bodyPr',wrap='none');sub(text,A,'lstStyle')
        props=sub(sub(sub(text,A,'p'),A,'pPr'),A,'defRPr',sz='1100',b='1')
        sub(sub(props,A,'solidFill'),A,'srgbClr',val='FFFF00')
        sub(labels,C,'dLblPos',val='ctr')
        for name,value in [('showLegendKey',0),('showVal',0),('showCatName',1),('showSerName',0),('showPercent',1)]:sub(labels,C,name,val=value)
        sub(labels,C,'separator').text='\n'
        sub(labels,C,'showLeaderLines',val='0')
        sp=new.find('{'+C+'}ser/{'+C+'}spPr')
        sub(sub(sub(sp,A,'ln',w='19050'),A,'solidFill'),A,'srgbClr',val='FFFFFF')
        sub(new,C,'firstSliceAng',val=270)
        for e in list(plot):
            if local(e.tag) in ('catAx','valAx','dateAx','serAx'):plot.remove(e)
    else:
        if new.tag=='{'+C+'}barChart':
            sub(new,C,'gapWidth',val='80' if len(columns)==1 else '65')
            sub(new,C,'overlap',val='100' if form.startswith(('stacked-','percent-')) else '0')
        for e in old.findall('{'+C+'}axId'):new.append(deepcopy(e))
    plot.replace(old,new)
    title=chart.find('{'+C+'}title')
    ts=title.findall('.//{'+A+'}t') if title is not None else []
    if ts:
        ts[0].text=spec['title']
        for t in ts[1:]:t.text=''
    for legend in chart.findall('{'+C+'}legend'):chart.remove(legend)
    if len(columns)>1 and form!='pie':
        legend=el(C,'legend');sub(legend,C,'legendPos',val='t');sub(legend,C,'overlay',val='0');chart.insert(chart.index(plot)+1,legend)
    # Axis styling is inherited; remove obsolete titles and set units where useful.
    for axis in plot:
        if local(axis.tag) not in ('catAx','valAx'):continue
        for t in axis.findall('{'+C+'}title'):axis.remove(t)
        num=axis.find('{'+C+'}numFmt')
        if num is not None:num.set('formatCode','0%' if form.startswith('percent-') else '#,##0');num.set('sourceLinked','0')
        if local(axis.tag)=='catAx' and 'bar' in form:
            orient=axis.find('{'+C+'}scaling/{'+C+'}orientation')
            if orient is not None:orient.set('val','maxMin')
        if local(axis.tag)=='valAx' and form!='percent-area':
            pos=axis.find('{'+C+'}axPos')
            if pos is not None:pos.set('val','b' if 'bar' in form else 'l')
    pivot=root.find('{'+C+'}pivotSource')
    if pivot is None:pivot=el(C,'pivotSource');root.insert(0,pivot)
    for e in list(pivot):pivot.remove(e)
    sub(pivot,C,'name').text=f'[{output_name}]分析内容{index}!CasePT{index:02}'
    sub(pivot,C,'fmtId',val='0')
    for e in root.findall('{'+C+'}extLst'):root.remove(e)
    ext=sub(sub(root,C,'extLst'),C,'ext',uri='{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}')
    opts=sub(ext,C14,'pivotOptions')
    for name,val in [('dropZoneFilter',int(bool(spec.get('filter')))),('dropZoneCategories',0),('dropZoneData',0),('dropZoneSeries',0),('dropZonesVisible',int(bool(spec.get('filter'))))]:sub(opts,C14,name,val=val)
    recolor(root,reference)
    split_small_pie(root)
    return root


def split_small_pie(root):
    """Use a native secondary pie for crowded small slices, keeping raw points once."""
    plot=root.find('{'+C+'}chart/{'+C+'}plotArea')
    pie=plot.find('{'+C+'}pieChart') if plot is not None else None
    if pie is None:return False
    values=[float(e.text) for e in pie.findall('{'+C+'}ser/{'+C+'}val/{'+C+'}numRef/{'+C+'}numCache/{'+C+'}pt/{'+C+'}v')]
    total=sum(values)
    small=[i for i,v in enumerate(values) if total>0 and v/total<0.08]
    if len(values)<5 or len(small)<2 or len(small)>len(values)-3:return False
    pie.tag='{'+C+'}ofPieChart';pie.insert(0,el(C,'ofPieType',val='pie'))
    for node in pie.findall('{'+C+'}firstSliceAng'):pie.remove(node)
    sub(pie,C,'gapWidth',val='60');sub(pie,C,'splitType',val='cust')
    custom=sub(pie,C,'custSplit')
    for i in small:sub(custom,C,'secondPiePt',val=i)
    sub(pie,C,'secondPieSize',val='75')
    sub(sub(sub(sub(pie,C,'serLines'),C,'spPr'),A,'ln',w='9525'),A,'solidFill').append(el(A,'srgbClr',val='A6A6A6'))
    return True


def chart_background(root,white):
    """Keep the analysis chart opaque and the dashboard chart transparent."""
    sp=root.find('{'+C+'}spPr')
    if sp is None:sp=sub(root,C,'spPr')
    for child in list(sp):
        if local(child.tag) in ('noFill','solidFill','gradFill','blipFill','pattFill','grpFill'):sp.remove(child)
    fill=el(A,'solidFill' if white else 'noFill')
    if white:sub(fill,A,'srgbClr',val='FFFFFF')
    sp.insert(0,fill)


def prepare_dashboard_grid(board):
    """Stable explicit grid: Excel may normalize an implicit 16.8pt default to 14pt."""
    fmt=board.find('{'+S+'}sheetFormatPr')
    fmt.set('defaultRowHeight','14');fmt.set('defaultColWidth','12.5')
    columns=board.find('{'+S+'}cols')
    if columns is None:
        columns=el(S,'cols');board.insert(board.index(fmt)+1,columns)
    retained=[]
    for item in list(columns):
        if int(item.get('max'))>12:
            extra=deepcopy(item);extra.set('min',str(max(13,int(item.get('min')))));retained.append(extra)
        columns.remove(item)
    sub(columns,S,'col',min='1',max='12',width='12.5',customWidth='1')
    columns.extend(retained)
    for row in board.findall('{'+S+'}sheetData/{'+S+'}row'):
        row.set('ht','14');row.set('customHeight','1')


def dashboard_layout(board,analyses):
    """Equal horizontal lengths; content-aware heights, equal within each row."""
    fmt=board.find('{'+S+'}sheetFormatPr');default=float(fmt.get('defaultColWidth','9'))
    widths=[default]*12
    for node in board.findall('{'+S+'}cols/{'+S+'}col'):
        for i in range(max(0,int(node.get('min'))-1),min(12,int(node.get('max')))):
            widths[i]=0 if node.get('hidden')=='1' else float(node.get('width',str(default)))
    pixels=[int(w*7+5) if w else 0 for w in widths]
    outer=12;gap=24;available=sum(pixels)-2*outer-gap
    result={};start=0;length=int(available/2)
    heights={'line':14,'pie':14,'column':17,'bar':15,'percent-bar':15,'stacked-bar':15,'stacked-column':16,'percent-column':16,'percent-area':16}
    while start<len(analyses):
        end=start+1
        while end<len(analyses) and analyses[end]['layer']==analyses[start]['layer']:end+=1
        rowoffset=0
        for i in range(start,end,2):
            pair=analyses[i:min(i+2,end)]
            hs=[float(s.get('height_rows',heights[s['chart']])) for s in pair]
            if any(h<8 or h>40 for h in hs):raise ValueError('height_rows must be between 8 and 40')
            height=ceil(max(hs))
            for j,spec in enumerate(pair):
                px=outer if j==0 else outer+length+gap
                c=0;offset=px
                while c<len(pixels)-1 and offset>=pixels[c]:offset-=pixels[c];c+=1
                result[i+j]=(c,int(offset*9525),length*9525,height,rowoffset)
            rowoffset+=height+2
        start=end
    return result


def anchor(chartid,x,y,w,h,objid,xoff=0):
    a=el(D,'oneCellAnchor');fr=sub(a,D,'from')
    for k,v in [('col',x),('colOff',xoff),('row',y),('rowOff',0)]:sub(fr,D,k).text=str(v)
    sub(a,D,'ext',cx=w,cy=h)
    frame=sub(a,D,'graphicFrame',macro='');nv=sub(frame,D,'nvGraphicFramePr');sub(nv,D,'cNvPr',id=objid,name=f'案例图表 {objid}');sub(nv,D,'cNvGraphicFramePr')
    xf=sub(frame,D,'xfrm');sub(xf,A,'off',x=0,y=0);sub(xf,A,'ext',cx=0,cy=0)
    graphic=sub(frame,A,'graphic');data=sub(graphic,A,'graphicData',uri=C);sub(data,C,'chart',**{'{'+R+'}id':chartid})
    sub(a,D,'clientData');return a


def set_shape_text(shape,text):
    body=shape.find('.//{'+D+'}txBody')
    if body is None:return
    pars=body.findall('{'+A+'}p')
    if not pars:return
    first=pars[0];run=first.find('{'+A+'}r');props=deepcopy(run.find('{'+A+'}rPr')) if run is not None and run.find('{'+A+'}rPr') is not None else None
    for child in list(first):
        if local(child.tag) not in ('pPr',):first.remove(child)
    run=sub(first,A,'r')
    if props is not None:run.append(props)
    sub(run,A,'t').text=text
    for p in pars[1:]:body.remove(p)


def validate_plan(plan,headers,records):
    analyses=plan.get('analyses',[])
    if len(analyses)<6:raise ValueError('Plan must contain at least six analyses')
    if len(set(a.get('chart') for a in analyses))<6:raise ValueError('Plan must contain at least six supported, meaningful chart forms')
    if not any(a.get('filter') for a in analyses):raise ValueError('Plan needs a useful filter field for a dynamic native chart')
    for a in analyses:
        if a.get('chart') not in FORMS:raise ValueError('Unsupported chart form: '+str(a.get('chart')))
        if not a.get('title') or not a.get('layer'):raise ValueError('Each analysis needs a title and layer')
        aggregate(a,headers,records)


def build_case(source,output,plan,reference=None,dashboard='模板设计3',report=None):
    source,output=Path(source),Path(output)
    reference=Path(reference) if reference else Path(__file__).resolve().parents[1]/'assets/completed-reference.xlsx'
    if output.resolve() in (source.resolve(),reference.resolve()) or output.exists():raise ValueError('Use a new output path; do not overwrite source/reference/existing output')
    if report and Path(report).exists():raise ValueError('Report output already exists')
    if ']' in output.name:raise ValueError('Output file name cannot contain ]')
    pkg,ref=Package(source),Package(reference)
    for name in ('案例数据','需求分析',dashboard):
        if name not in pkg.sheets:raise ValueError('Missing source sheet: '+name)
    if any(re.fullmatch('分析内容[1-9]\\d*',n) for n in pkg.sheets):raise ValueError('This input already contains analysis sheets; use the untouched source for rebuilding')
    headers,records,lastrow=load_records(pkg);validate_plan(plan,headers,records)
    parts=dict(pkg.parts)
    def put(path,node):parts[path]=xml(node)
    types=X.fromstring(parts['[Content_Types].xml'])
    def content(path,kind):
        sub(types,CT,'Override',PartName='/'+path,ContentType='application/vnd.openxmlformats-officedocument.'+kind+'+xml')
    styles=StyleImporter(parts,ref)
    styleids=[styles.from_cell(ref,'分析内容1',a) for a in ('A3','A4','B4')]
    book=X.fromstring(parts['xl/workbook.xml']);bookrels=X.fromstring(parts['xl/_rels/workbook.xml.rels'])
    caches=book.find('{'+S+'}pivotCaches')
    if caches is None:
        caches=el(S,'pivotCaches');calc=book.find('{'+S+'}calcPr');ext=book.find('{'+S+'}extLst');book.insert(book.index(calc)+1 if calc is not None else book.index(ext) if ext is not None else len(book),caches)
    cacheid=max([-1]+[int(c.get('cacheId')) for c in caches])+1
    sub(caches,S,'pivotCache',cacheId=cacheid,**{'{'+R+'}id':'rIdCaseCache'})
    sub(bookrels,P,'Relationship',Id='rIdCaseCache',Type=R+'/pivotCacheDefinition',Target='pivotCache/caseCache.xml')
    definition,cachedata,shared=cache_parts(headers,records,lastrow)
    for path,node,kind in [('xl/pivotCache/caseCache.xml',definition,'spreadsheetml.pivotCacheDefinition'),('xl/pivotCache/caseRecords.xml',cachedata,'spreadsheetml.pivotCacheRecords')]:put(path,node);content(path,kind)
    put('xl/pivotCache/_rels/caseCache.xml.rels',relations([('rIdRecords','pivotCacheRecords','caseRecords.xml')]))
    boardpath=pkg.sheets[dashboard];board=X.fromstring(parts[boardpath]);prepare_dashboard_grid(board)
    drawingrel=next(r for r in pkg.rels(boardpath).values() if r['type']=='drawing')
    drawingpath=drawingrel['path'];drawing=X.fromstring(parts[drawingpath])
    # Keep the source header and identity text boxes, updating the original title.
    for a in list(drawing):
        text=rich_text(a)
        if not text:continue
        if '标题' in text and '课程实践' not in text:set_shape_text(a,plan.get('title','销售数据看板'))
        elif '整体情况分析' in text and '课程实践' not in text:drawing.remove(a)
    drawingrels=relations([])
    existingrel=pkg.rels(drawingpath)
    for rid,r in existingrel.items():
        # Source templates normally contain text only; preserve any existing relationship.
        target=__import__('posixpath').relpath(r['path'],str(Path(drawingpath).parent))
        sub(drawingrels,P,'Relationship',Id=rid,Type=R+'/'+r['type'],Target=target)
    heading_reference=X.fromstring(ref.parts['xl/drawings/drawing9.xml'])
    heading=next(a for a in heading_reference if '整体情况分析' in rich_text(a))
    maxobj=max([0]+[int(e.get('id')) for e in drawing.findall('.//{'+D+'}cNvPr')])
    fmt=board.find('{'+S+'}sheetFormatPr')
    rowheight=float(fmt.get('defaultRowHeight','16.8'));columnwidth=float(fmt.get('defaultColWidth','9'))
    width=int((int(columnwidth*7+5)*6-24)*9525);height=int(rowheight*14*12700)
    layout=dashboard_layout(board,plan['analyses'])
    sheets=book.find('{'+S+'}sheets');nextid=max(int(e.get('sheetId')) for e in sheets)+1
    requirement=X.fromstring(parts[pkg.sheets['需求分析']]);reqdata=requirement.find('{'+S+'}sheetData')
    headerrow=deepcopy(reqdata[0]);reqdata.clear();reqdata.append(headerrow)
    manifest={'source':str(source.resolve()),'source_total_amount':sum(row.get('订单金额',0) for row in records),'analyses':[],'native_validation':'pending'}
    layer=None;position=12;layerstart=12;nextrow=12
    for index,spec in enumerate(plan['analyses'],1):
        rows,columns,matrix=aggregate(spec,headers,records)
        pivot,start,totalrow,ncols=pivot_table(spec,index,cacheid,headers,shared,rows,columns)
        sheetpath=f'xl/worksheets/caseAnalysis{index}.xml';pivotpath=f'xl/pivotTables/casePivot{index}.xml';drawpath=f'xl/drawings/caseAnalysis{index}.xml'
        put(sheetpath,worksheet(spec,rows,columns,matrix,start,totalrow,ncols,styleids,index));content(sheetpath,'spreadsheetml.worksheet')
        put(pivotpath,pivot);content(pivotpath,'spreadsheetml.pivotTable')
        put(relpath(pivotpath),relations([('rIdCache','pivotCacheDefinition','../pivotCache/caseCache.xml')]))
        put(relpath(sheetpath),relations([('rIdDrawing','drawing',f'../drawings/caseAnalysis{index}.xml'),('rIdPivot','pivotTable',f'../pivotTables/casePivot{index}.xml')]))
        rid=f'rIdCaseSheet{index}';sub(sheets,S,'sheet',name=f'分析内容{index}',sheetId=nextid+index-1,**{'{'+R+'}id':rid});sub(bookrels,P,'Relationship',Id=rid,Type=R+'/worksheet',Target=f'worksheets/caseAnalysis{index}.xml')
        chart=make_chart(ref,spec,index,rows,columns,matrix,start,output.name)
        chart_name={'line':'折线图','pie':'饼图','column':'柱形图','bar':'条形图','stacked-column':'堆积柱形图','stacked-bar':'堆积条形图','percent-column':'百分比堆积柱形图','percent-bar':'百分比堆积条形图','percent-area':'百分比堆积面积图'}[spec['chart']]
        if chart.find('.//{'+C+'}ofPieChart') is not None:chart_name='复合饼图'
        for suffix in ('A','B'):
            instance=deepcopy(chart);chart_background(instance,suffix=='A')
            cp=f'xl/charts/caseChart{index}{suffix}.xml';put(cp,instance);content(cp,'drawingml.chart')
        analysisdrawing=X.Element('{'+D+'}wsDr',nsmap={'xdr':D,'a':A,'r':R})
        analysiswidth=layout[index-1][2];analysisheight=int(rowheight*layout[index-1][3]*12700)
        analysisdrawing.append(anchor('rIdChart',max(ncols+1,4),2,analysiswidth,analysisheight,2));put(drawpath,analysisdrawing);content(drawpath,'drawing')
        put(relpath(drawpath),relations([('rIdChart','chart',f'../charts/caseChart{index}A.xml')]))
        if layer!=spec['layer']:
            if layer is not None:position=nextrow+1
            layerstart=position;layer=spec['layer'];label=deepcopy(heading);maxobj+=1
            set_shape_text(label,layer+'：')
            for p in label.findall('.//{'+D+'}cNvPr'):p.set('id',str(maxobj));p.set('name',f'分析板块 {index}')
            for key,val in [('from',position-2),('to',position)]:
                node=label.find('{'+D+'}'+key+'/{'+D+'}row')
                if node is not None:node.text=str(val)
                offset=label.find('{'+D+'}'+key+'/{'+D+'}rowOff')
                if offset is not None:offset.text='0'
            drawing.append(label)
        x,xoff,boardwidth,heightrows,rowoffset=layout[index-1]
        y=layerstart+rowoffset;nextrow=max(nextrow,y+heightrows+2);maxobj+=1
        drawing.append(anchor(f'rIdCaseChart{index}',x,y,boardwidth,int(rowheight*heightrows*12700),maxobj,xoff))
        sub(drawingrels,P,'Relationship',Id=f'rIdCaseChart{index}',Type=R+'/chart',Target=f'../charts/caseChart{index}B.xml')
        rr=sub(reqdata,S,'row',r=index+1)
        for c,v in zip('ABCDE',[spec['layer'],spec['title'],'结构分析' if spec['chart']=='pie' else '对比分析+结构分析' if spec['chart'].startswith('percent-') else '趋势分析' if spec['chart']=='line' else '对比分析',chart_name,'是' if spec.get('filter') else '否']):
            original=next((cell for cell in headerrow if cell.get('r','').startswith(c)),None)
            cell(rr,c+str(index+1),v,int(original.get('s','0')) if original is not None else 0)
        manifest['analyses'].append({'index':index,'title':spec['title'],'chart':spec['chart'],'rows':rows,'columns':columns,'values':matrix,'total':sum(map(sum,matrix)),'filter':spec.get('filter')})
    requirement.find('{'+S+'}dimension').set('ref',f'A1:E{len(plan["analyses"])+1}')
    put(pkg.sheets['需求分析'],requirement);put(drawingpath,drawing);put(relpath(drawingpath),drawingrels)
    # Activate only the dashboard; remove accidental all-cell/grouped selections.
    for name,path in pkg.sheets.items():
        node=board if name==dashboard else X.fromstring(parts[path])
        changed=False
        for view in node.findall('{'+S+'}sheetViews/{'+S+'}sheetView'):
            if name==dashboard:
                view.set('tabSelected','1');view.set('topLeftCell','A1')
                view.set('zoomScale','80');view.set('zoomScaleNormal','80')
                for selection in view.findall('{'+S+'}selection'):view.remove(selection)
                sub(view,S,'selection',activeCell='A1',sqref='A1');changed=True
            elif view.get('tabSelected')=='1':view.set('tabSelected','0');changed=True
        if changed:put(path,node)
    views=book.find('{'+S+'}bookViews')
    if views is not None and len(views):views[0].set('activeTab',str(list(pkg.sheets).index(dashboard)))
    put('xl/workbook.xml',book);put('xl/_rels/workbook.xml.rels',bookrels);put('xl/styles.xml',styles.root);put('[Content_Types].xml',types)
    theme_workbook(parts,dashboard)
    output.parent.mkdir(parents=True,exist_ok=True)
    temporary=output.with_suffix('.building.xlsx')
    if temporary.exists():raise ValueError('Temporary output already exists: '+str(temporary))
    try:
        with zipfile.ZipFile(temporary,'w',compression=zipfile.ZIP_DEFLATED) as z:
            for path,payload in parts.items():z.writestr(path,payload)
        # Parse every generated part and check relationship targets before exposing output.
        built=Package(temporary)
        if built.broken_relationships():raise ValueError('Generated workbook has broken relationships')
        for path in parts:
            if path.endswith('.xml'):X.fromstring(parts[path])
        temporary.rename(output)
    finally:
        if temporary.exists():temporary.unlink()
    if report:
        lines=['第三部分：图表信息解读','']
        for spec,item in zip(plan['analyses'],manifest['analyses']):
            totals=[sum(v) for v in item['values']];best=max(range(len(totals)),key=totals.__getitem__);share=totals[best]/sum(totals) if sum(totals) else 0
            finding=spec.get('analysis')
            if not finding:
                metric=spec['value'];name=item['rows'][best]
                finding=f'{name}的{metric}最高，为{totals[best]:,.0f}，占本图合计的{share:.1%}。'
                if spec.get('filter'):finding+='本图默认展示全部，可按'+spec['filter']+'筛选比较。'
            lines.extend([f'{item["index"]}. {spec["title"]}','分析：'+finding,''])
        lines.extend(['总分析',plan.get('summary','本报告从整体走势和构成深入产品、市场及基层表现，各图按实际统计口径比较。可结合动态筛选观察不同类别的差异，再确定具体改进重点。')])
        Path(report).write_text('\n'.join(lines)+'\n',encoding='utf-8')
    return manifest


def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('source',type=Path);p.add_argument('--plan',required=True,type=Path);p.add_argument('--output',required=True,type=Path);p.add_argument('--report',type=Path);p.add_argument('--manifest',type=Path);p.add_argument('--reference',type=Path);p.add_argument('--dashboard',default='模板设计3')
    a=p.parse_args()
    try:
        result=build_case(a.source,a.output,json.loads(a.plan.read_text(encoding='utf-8')),reference=a.reference,dashboard=a.dashboard,report=a.report)
        if a.manifest:a.manifest.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        print(json.dumps({'workbook':str(a.output.resolve()),'report':str(a.report.resolve()) if a.report else None,'analyses':len(result['analyses']),'native_validation':'pending'},ensure_ascii=False));return 0
    except (ValueError,KeyError,OSError,X.XMLSyntaxError) as e:
        print(str(e),file=sys.stderr);return 1


if __name__=='__main__':sys.exit(main())
