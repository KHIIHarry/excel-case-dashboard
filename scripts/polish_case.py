#!/usr/bin/env python3
"""Reflow an existing course dashboard without rebuilding pivots or copying data."""
import argparse
import json
from pathlib import Path
import re
import sys
import zipfile
from lxml import etree as X
from workbook_audit import Package, rich_text
from chart_theme import theme_workbook
from build_case import S,A,C,R,D,el,sub,xml,prepare_dashboard_grid,dashboard_layout,split_small_pie


def chart_form(root):
    plot=root.find('{'+C+'}chart/{'+C+'}plotArea')
    node=next(n for n in plot if X.QName(n).localname.endswith('Chart'))
    kind=X.QName(node).localname
    if kind in ('pieChart','ofPieChart'):return 'pie'
    if kind=='lineChart':return 'line'
    if kind=='areaChart':return 'percent-area'
    grouping=node.find('{'+C+'}grouping')
    prefix={'stacked':'stacked-','percentStacked':'percent-'}.get(grouping.get('val') if grouping is not None else '', '')
    direction=node.find('{'+C+'}barDir')
    return prefix+('bar' if direction is not None and direction.get('val')=='bar' else 'column')


def top(anchor):
    row=anchor.find('{'+D+'}from/{'+D+'}row')
    off=anchor.find('{'+D+'}from/{'+D+'}rowOff')
    return int(row.text)+(int(off.text)/177800 if off is not None else 0)


def move_chart(anchor,x,xoff,y,width,height):
    # Preserve the actual native graphic frame, its IDs and every extension.
    anchor.tag='{'+D+'}oneCellAnchor';anchor.attrib.clear()
    for child in list(anchor):
        if X.QName(child).localname in ('from','to','ext','pos'):anchor.remove(child)
    fr=el(D,'from')
    for key,value in [('col',x),('colOff',xoff),('row',y),('rowOff',0)]:sub(fr,D,key).text=str(value)
    anchor.insert(0,fr);anchor.insert(1,el(D,'ext',cx=width,cy=height))


def polish(source,output,dashboard='模板设计3'):
    source,output=Path(source),Path(output)
    if source.resolve()==output.resolve() or output.exists():raise ValueError('Use a new output path; existing files are never overwritten')
    pkg=Package(source);parts=dict(pkg.parts);board=X.fromstring(parts[pkg.sheets[dashboard]])
    drawingpath=next(r['path'] for r in pkg.rels(pkg.sheets[dashboard]).values() if r['type']=='drawing')
    drawing=X.fromstring(parts[drawingpath]);rels=pkg.rels(drawingpath)
    headings=[a for a in drawing if a.find('.//{'+C+'}chart') is None and '分析' in rich_text(a) and '课程实践' not in rich_text(a)]
    headings.sort(key=top)
    anchors=[a for a in drawing if a.find('.//{'+C+'}chart') is not None]
    anchors.sort(key=lambda a:(top(a),int(a.find('{'+D+'}from/{'+D+'}col').text)))
    specs=[]
    for a in anchors:
        heading=next((h for h in reversed(headings) if top(h)<=top(a)),None)
        layer=rich_text(heading).strip('：: \n') if heading is not None else '情况分析'
        rid=a.find('.//{'+C+'}chart').get('{'+R+'}id');root=X.fromstring(parts[rels[rid]['path']])
        specs.append({'layer':layer,'chart':chart_form(root)})
    prepare_dashboard_grid(board);layout=dashboard_layout(board,specs)
    layer=None;start=12;nextrow=12;heading_index=0;sizes={}
    for i,(a,spec) in enumerate(zip(anchors,specs)):
        if spec['layer']!=layer:
            if layer is not None:start=nextrow+1
            layer=spec['layer']
            if heading_index<len(headings):
                h=headings[heading_index];heading_index+=1
                for key,row in [('from',start-2),('to',start)]:
                    node=h.find('{'+D+'}'+key+'/{'+D+'}row')
                    if node is not None:node.text=str(row)
                    off=h.find('{'+D+'}'+key+'/{'+D+'}rowOff')
                    if off is not None:off.text='0'
        x,xoff,width,rows,offset=layout[i];y=start+offset
        move_chart(a,x,xoff,y,width,int(rows*14*12700));nextrow=max(nextrow,y+rows+2)
        rid=a.find('.//{'+C+'}chart').get('{'+R+'}id');chart=X.fromstring(parts[rels[rid]['path']])
        pivot=chart.find('{'+C+'}pivotSource/{'+C+'}name')
        if pivot is not None:sizes[(pivot.text or '').split(']',1)[-1]]=(width,int(rows*14*12700))
    parts[pkg.sheets[dashboard]]=xml(board);parts[drawingpath]=xml(drawing)
    for name,path in pkg.sheets.items():
        if not re.fullmatch(r'分析内容\d+',name):continue
        for relationship in pkg.rels(path).values():
            if relationship['type']!='drawing':continue
            dp=relationship['path'];dr=X.fromstring(parts[dp]);connections=pkg.rels(dp)
            for item in dr:
                link=item.find('.//{'+C+'}chart')
                if link is None:continue
                cp=connections[link.get('{'+R+'}id')]['path'];chart=X.fromstring(parts[cp])
                pivot=chart.find('{'+C+'}pivotSource/{'+C+'}name')
                dims=sizes.get((pivot.text or '').split(']',1)[-1]) if pivot is not None else None
                if dims:
                    fr=item.find('{'+D+'}from');x=int(fr.find('{'+D+'}col').text);xoff=int(fr.find('{'+D+'}colOff').text);y=int(fr.find('{'+D+'}row').text)
                    move_chart(item,x,xoff,y,*dims)
            parts[dp]=xml(dr)
    converted=[]
    for path in pkg.parts:
        if path.startswith('xl/charts/') and path.endswith('.xml'):
            root=X.fromstring(parts[path]);did_split=split_small_pie(root)
            prefix=root.find('{'+C+'}pivotSource/{'+C+'}name')
            if prefix is not None and prefix.text and prefix.text.startswith('['+source.name+']'):
                prefix.text='['+output.name+']'+prefix.text.split(']',1)[1]
            parts[path]=xml(root)
            if did_split:converted.append(path)
    # Requirements table must name the actual newly split native chart form.
    requirements=parts.get(pkg.sheets.get('需求分析',''))
    if requirements:
        req=X.fromstring(requirements)
        for path in converted:
            root=X.fromstring(parts[path]);name=root.find('{'+C+'}pivotSource/{'+C+'}name')
            match=re.search(r'分析内容(\d+)!',name.text or '') if name is not None else None
            if match:
                target=req.find('.//{'+S+'}c[@r="D'+str(int(match[1])+1)+'"]')
                if target is not None:
                    for child in list(target):target.remove(child)
                    target.set('t','inlineStr');sub(sub(target,S,'is'),S,'t').text='复合饼图'
        parts[pkg.sheets['需求分析']]=xml(req)
    theme_workbook(parts,dashboard)
    output.parent.mkdir(parents=True,exist_ok=True)
    temporary=output.with_suffix('.polishing.xlsx')
    if temporary.exists():raise ValueError('Temporary output exists: '+str(temporary))
    try:
        with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED) as z:
            for path,payload in parts.items():z.writestr(path,payload)
        if Package(temporary).broken_relationships():raise ValueError('Broken workbook relationships')
        temporary.rename(output)
    finally:
        if temporary.exists():temporary.unlink()
    return {'output':str(output.resolve()),'charts':len(anchors),'split_pies':converted,'native_validation':'pending'}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path);parser.add_argument('--output',type=Path,required=True);parser.add_argument('--dashboard',default='模板设计3')
    args=parser.parse_args()
    try:print(json.dumps(polish(args.source,args.output,args.dashboard),ensure_ascii=False));return 0
    except (ValueError,KeyError,OSError,X.XMLSyntaxError) as e:print(str(e),file=sys.stderr);return 1


if __name__=='__main__':sys.exit(main())
