#!/usr/bin/env python3
"""Adapt native chart colors to the selected template, keeping white analysis pages."""
import argparse
from collections import Counter
from copy import deepcopy
import colorsys
import json
from pathlib import Path
import posixpath
import re
import sys
import zipfile
from lxml import etree as X

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
D='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PALETTE=['4472C4','ED7D31','FFC000','70AD47','5B9BD5','A56CB8']


def node(ns,name,**attrs):return X.Element('{'+ns+'}'+name,**{k:str(v) for k,v in attrs.items()})
def add(parent,ns,name,**attrs):
    child=node(ns,name,**attrs);parent.append(child);return child
def local(element):return X.QName(element).localname
def encode(root):return X.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)


def contrast(first,second):
    def luminance(color):
        values=[int(color[i:i+2],16)/255 for i in (0,2,4)]
        values=[v/12.92 if v<=.04045 else ((v+.055)/1.055)**2.4 for v in values]
        return sum(v*w for v,w in zip(values,(.2126,.7152,.0722)))
    low,high=sorted((luminance(first),luminance(second)))
    return (high+.05)/(low+.05)


def blend(first,second,amount):
    return ''.join(f'{round(int(first[i:i+2],16)*(1-amount)+int(second[i:i+2],16)*amount):02X}' for i in (0,2,4))


def readable(color,background,minimum=4.5):
    color=color or '595959'
    if contrast(color,background)>=minimum:return color
    return max(('FFFFFF','17212B','000000'),key=lambda c:contrast(c,background))


def visible_series(color,background,minimum=3):
    if contrast(color,background)>=minimum:return color
    # Retain hue and make only the lightness change needed for the canvas.
    target=max(('FFFFFF','000000'),key=lambda c:contrast(c,background))
    for step in range(1,101):
        candidate=blend(color,target,step/100)
        if contrast(candidate,background)>=minimum:return candidate
    return target


def theme_colors(parts):
    theme=X.fromstring(parts['xl/theme/theme1.xml'])
    scheme=theme.find('.//{'+A+'}clrScheme')
    return {local(e):(next(iter(e)).get('lastClr') if local(next(iter(e)))=='sysClr' else next(iter(e)).get('val')) for e in scheme}


def color_value(element,theme):
    if element is None:return None
    rgb=element.get('rgb') or (element.get('val') if local(element)=='srgbClr' else None)
    if rgb:color=rgb[-6:].upper()
    elif element.get('theme') is not None:
        keys=['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']
        color=theme.get(keys[int(element.get('theme'))])
    elif local(element)=='schemeClr':color=theme.get({'bg1':'lt1','bg2':'lt2','tx1':'dk1','tx2':'dk2'}.get(element.get('val'),element.get('val')))
    elif local(element)=='sysClr':color=element.get('lastClr')
    elif element.get('indexed') is not None:
        color={'0':'000000','1':'FFFFFF','8':'000000','9':'FFFFFF'}.get(element.get('indexed'))
    else:return None
    if not color or not re.fullmatch('[0-9a-fA-F]{6}',color):return None
    tint=float(element.get('tint','0'))
    if tint:
        r,g,b=[int(color[i:i+2],16)/255 for i in (0,2,4)];h,l,s=colorsys.rgb_to_hls(r,g,b)
        l=l*(1+tint) if tint<0 else l*(1-tint)+tint
        color=''.join(f'{round(v*255):02X}' for v in colorsys.hls_to_rgb(h,l,s))
    for change in element:
        amount=float(change.get('val','100000'))/100000
        if local(change)=='lumMod':color=blend('000000',color,amount)
        elif local(change)=='lumOff':color=blend(color,'FFFFFF',amount)
        elif local(change)=='tint':color=blend('FFFFFF',color,amount)
        elif local(change)=='shade':color=blend('000000',color,amount)
    return color


def relationships(parts,path):
    rp=posixpath.join(posixpath.dirname(path),'_rels',posixpath.basename(path)+'.rels')
    if rp not in parts:return {}
    return {e.get('Id'):posixpath.normpath(posixpath.join(posixpath.dirname(path),e.get('Target'))) for e in X.fromstring(parts[rp]) if e.get('TargetMode')!='External'}


def worksheets(parts):
    rels=relationships(parts,'xl/workbook.xml')
    return {e.get('name'):rels[e.get('{'+R+'}id')] for e in X.fromstring(parts['xl/workbook.xml']).find('{'+S+'}sheets')}


def background_color(parts,sheet):
    theme=theme_colors(parts);styles=X.fromstring(parts['xl/styles.xml'])
    fills=styles.find('{'+S+'}fills');xfs=styles.find('{'+S+'}cellXfs');root=X.fromstring(parts[worksheets(parts)[sheet]])
    counts=Counter()
    for cell in root.findall('{'+S+'}sheetData/{'+S+'}row/{'+S+'}c'):
        style=xfs[int(cell.get('s','0'))];fill=fills[int(style.get('fillId','0'))]
        pattern=fill.find('{'+S+'}patternFill')
        if pattern is not None and pattern.get('patternType')=='solid':
            color=color_value(pattern.find('{'+S+'}fgColor'),theme)
            if color:counts[color]+=1
    return counts.most_common(1)[0][0] if counts else 'FFFFFF'


def fill_color(parent,color):
    for child in list(parent):
        if local(child) in ('solidFill','noFill','gradFill','pattFill','blipFill','grpFill'):parent.remove(child)
    fill=node(A,'solidFill');add(fill,A,'srgbClr',val=color)
    preceding={'xfrm','prstGeom','custGeom'}
    if local(parent) in ('rPr','defRPr','endParaRPr'):preceding.add('ln')
    index=max([-1]+[i for i,e in enumerate(parent) if local(e) in preceding])+1
    parent.insert(index,fill)


def insert_before(parent,child,names):
    match=next((e for e in parent if local(e) in names),None)
    parent.insert(parent.index(match) if match is not None else len(parent),child)


def text_color(body,color):
    properties=list(body.findall('.//{'+A+'}rPr'))+list(body.findall('.//{'+A+'}defRPr'))+list(body.findall('.//{'+A+'}endParaRPr'))
    for paragraph in body.findall('.//{'+A+'}p'):
        pp=paragraph.find('{'+A+'}pPr')
        if pp is None:pp=node(A,'pPr');paragraph.insert(0,pp)
        default=pp.find('{'+A+'}defRPr')
        if default is None:default=add(pp,A,'defRPr');properties.append(default)
    for prop in properties:fill_color(prop,color)


def text_properties(parent,color):
    tx=parent.find('{'+C+'}txPr')
    if tx is None:
        tx=node(C,'txPr');add(tx,A,'bodyPr');add(tx,A,'lstStyle');add(tx,A,'p')
        insert_before(parent,tx,{'extLst','externalData','printSettings','userShapes','crossAx','crosses','crossesAt','auto','lblAlgn','lblOffset','tickLblSkip','tickMarkSkip','noMultiLvlLbl','crossBetween','majorUnit','minorUnit','dispUnits','dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','separator','showLeaderLines','leaderLines'})
    text_color(tx,color)
    for rich in parent.findall('{'+C+'}tx/{'+C+'}rich'):text_color(rich,color)


def shape_properties(parent):
    sp=parent.find('{'+C+'}spPr')
    if sp is None:
        sp=node(C,'spPr')
        insert_before(parent,sp,{'txPr','externalData','printSettings','userShapes','crossAx','crosses','crossesAt','auto','lblAlgn','lblOffset','crossBetween','majorUnit','minorUnit','dispUnits','marker','explosion','dPt','dLbls','trendline','errBars','cat','val','xVal','yVal','bubbleSize','smooth','extLst'})
    return sp


def line_color(parent,color,width=None):
    sp=shape_properties(parent);line=sp.find('{'+A+'}ln')
    if line is None:line=add(sp,A,'ln')
    if width is not None:line.set('w',str(width))
    fill_color(line,color)


def adjust_chart(root,background,theme,white=False):
    canvas='FFFFFF' if white else background
    foreground=readable('595959',canvas)
    grid=blend(canvas,foreground,.4)
    if contrast(grid,canvas)<1.5:grid=blend(canvas,foreground,.55)
    area=shape_properties(root)
    if white:fill_color(area,'FFFFFF')
    else:
        for child in list(area):
            if local(child) in ('solidFill','noFill','gradFill','pattFill','blipFill'):area.remove(child)
        area.insert(0,node(A,'noFill'))
    line_color(root,visible_series('6C8EBF',canvas),9525)
    text_properties(root,foreground)
    chart=root.find('{'+C+'}chart');plot=chart.find('{'+C+'}plotArea')
    plot_fill=shape_properties(plot)
    for child in list(plot_fill):
        if local(child) in ('solidFill','noFill','gradFill','pattFill','blipFill'):plot_fill.remove(child)
    plot_fill.insert(0,node(A,'noFill'))
    for target in [chart.find('{'+C+'}title'),chart.find('{'+C+'}legend')]+[e for e in plot if local(e) in ('catAx','valAx','dateAx','serAx')]:
        if target is not None:text_properties(target,foreground)
    for axis in plot:
        if local(axis) not in ('catAx','valAx','dateAx','serAx'):continue
        line_color(axis,grid,9525)
        for line in axis:
            if local(line) in ('majorGridlines','minorGridlines'):line_color(line,grid,6350)
    for form in plot:
        if not local(form).endswith('Chart'):continue
        pie=local(form) in ('pieChart','ofPieChart','pie3DChart','doughnutChart')
        for i,series in enumerate(form.findall('{'+C+'}ser')):
            sp=shape_properties(series)
            old=sp.find('{'+A+'}solidFill/*')
            if old is None:old=sp.find('{'+A+'}ln/{'+A+'}solidFill/*')
            base=color_value(old,theme) or theme.get('accent'+str(i%6+1),PALETTE[i%6])
            data=visible_series(base,canvas)
            if local(form)=='lineChart':line_color(series,visible_series(base,canvas,4.5))
            elif not pie:fill_color(sp,data)
            if not pie:
                for label in series.findall('{'+C+'}dLbls')+form.findall('{'+C+'}dLbls'):
                    position=label.find('{'+C+'}dLblPos')
                    inside=position is not None and position.get('val') in ('ctr','inBase','inEnd')
                    label_color=readable(foreground,data) if inside else foreground
                    text_properties(label,label_color)
                    for point_label in label.findall('{'+C+'}dLbl'):text_properties(point_label,label_color)
                continue
            cache=series.find('{'+C+'}val/{'+C+'}numRef/{'+C+'}numCache')
            count=len(cache.findall('{'+C+'}pt')) if cache is not None else 0
            if not count:continue
            labels=series.find('{'+C+'}dLbls')
            if labels is None:labels=form.find('{'+C+'}dLbls')
            if labels is None:
                labels=node(C,'dLbls');insert_before(form,labels,{'firstSliceAng','holeSize','gapWidth','splitType','splitPos','custSplit','secondPieSize','serLines','extLst'})
                for key,value in [('dLblPos','ctr'),('showLegendKey','0'),('showVal','0'),('showCatName','1'),('showPercent','1')]:add(labels,C,key,val=value)
            text_properties(labels,'17212B')
            # The optional extra dPt styles Excel's generated "Other" only; no new source point.
            for j in range(count+(local(form)=='ofPieChart')):
                point=next((p for p in series.findall('{'+C+'}dPt') if p.find('{'+C+'}idx').get('val')==str(j)),None)
                if point is None:
                    point=node(C,'dPt');add(point,C,'idx',val=j)
                    insert_before(series,point,{'dLbls','trendline','errBars','cat','val','xVal','yVal','extLst'})
                ps=shape_properties(point);existing=ps.find('{'+A+'}solidFill/*')
                fill=visible_series(color_value(existing,theme) or theme.get('accent'+str(j%6+1),PALETTE[j%6]),canvas)
                fill_color(ps,fill);line_color(point,readable('FFFFFF',fill,3),12700)
                label=next((p for p in labels.findall('{'+C+'}dLbl') if p.find('{'+C+'}idx').get('val')==str(j)),None)
                if label is None:
                    label=node(C,'dLbl');add(label,C,'idx',val=j)
                    insert_before(labels,label,{'delete','numFmt','spPr','txPr','dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','separator','showLeaderLines','extLst'})
                if label.find('{'+C+'}txPr') is None and labels.find('{'+C+'}txPr') is not None:
                    insert_before(label,deepcopy(labels.find('{'+C+'}txPr')),{'dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','separator','showLeaderLines','extLst'})
                text_properties(label,readable('FFFF00',fill))
                # Excel does not inherit all global display flags into an explicit dLbl.
                for key,value in [('dLblPos','ctr'),('showLegendKey','0'),('showVal','0'),('showCatName','1'),('showSerName','0'),('showPercent','1'),('showBubbleSize','0')]:
                    item=label.find('{'+C+'}'+key)
                    if item is None:
                        item=node(C,key,val=value);insert_before(label,item,{'extLst'})
                    else:item.set('val',value)
                separator=label.find('{'+C+'}separator')
                if separator is None:
                    separator=node(C,'separator');insert_before(label,separator,{'extLst'})
                separator.text='\n'
            leader=labels.find('{'+C+'}showLeaderLines')
            if leader is None:
                leader=node(C,'showLeaderLines',val='0');insert_before(labels,leader,{'extLst'})
            else:leader.set('val','0')
        for connector in form.findall('{'+C+'}serLines'):line_color(connector,grid,9525)
    return {'background':canvas,'foreground':foreground}


def theme_workbook(parts,dashboard):
    theme=theme_colors(parts);sheets=worksheets(parts);background=background_color(parts,dashboard);changed={};contexts={}
    for name,path in sheets.items():
        if name!=dashboard and not re.fullmatch(r'分析内容\d+',name):continue
        white=name!=dashboard;canvas='FFFFFF' if white else background
        sheet=X.fromstring(parts[path])
        for drawing in sheet.findall('{'+S+'}drawing'):
            dp=relationships(parts,path)[drawing.get('{'+R+'}id')];root=X.fromstring(parts[dp]);rels=relationships(parts,dp)
            for link in root.findall('.//{'+C+'}chart'):
                cp=rels[link.get('{'+R+'}id')]
                if cp in contexts and contexts[cp]!=canvas:raise ValueError('Analysis and dashboard need separate native chart instances: '+cp)
                contexts[cp]=canvas;chart=X.fromstring(parts[cp]);changed[cp]=adjust_chart(chart,background,theme,white);parts[cp]=encode(chart)
            if name==dashboard:
                for body in root.findall('.//{'+D+'}txBody'):
                    shape=body.getparent();sp=shape.find('{'+D+'}spPr')
                    fill=sp.find('{'+A+'}solidFill/*') if sp is not None else None
                    text_color(body,readable('595959',color_value(fill,theme) or canvas))
                parts[dp]=encode(root)
    return {'dashboard':dashboard,'background':background,'charts':changed}


def retheme_file(source,output,dashboard=None):
    from workbook_audit import Package
    source,output=Path(source),Path(output)
    if source.resolve()==output.resolve() or output.exists():raise ValueError('Use a new output; never overwrite the input')
    parts=dict(Package(source).parts)
    if dashboard is None:
        candidates=[name for name in worksheets(parts) if re.fullmatch(r'模[板版]设计\d+',name)]
        if len(candidates)!=1:raise ValueError('Specify --dashboard explicitly when more than one template exists')
        dashboard=candidates[0]
    report=theme_workbook(parts,dashboard)
    for path in report['charts']:
        root=X.fromstring(parts[path]);prefix=root.find('{'+C+'}pivotSource/{'+C+'}name')
        if prefix is not None and prefix.text and prefix.text.startswith('['+source.name+']'):
            prefix.text='['+output.name+']'+prefix.text.split(']',1)[1];parts[path]=encode(root)
    output.parent.mkdir(parents=True,exist_ok=True)
    temporary=output.with_suffix('.coloring.xlsx')
    if temporary.exists():raise ValueError('Temporary output exists: '+str(temporary))
    try:
        with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED) as z:
            for path,data in parts.items():z.writestr(path,data)
        if Package(temporary).broken_relationships():raise ValueError('Broken native workbook relationships')
        temporary.rename(output)
    finally:
        if temporary.exists():temporary.unlink()
    return report


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('source',type=Path);parser.add_argument('--output',type=Path,required=True);parser.add_argument('--dashboard')
    args=parser.parse_args()
    try:
        report=retheme_file(args.source,args.output,args.dashboard);print(json.dumps({'output':str(args.output.resolve()),'dashboard':report['dashboard'],'background':report['background'],'charts':len(report['charts'])},ensure_ascii=False));return 0
    except (ValueError,KeyError,OSError,X.XMLSyntaxError) as e:print(str(e),file=sys.stderr);return 1


if __name__=='__main__':sys.exit(main())
