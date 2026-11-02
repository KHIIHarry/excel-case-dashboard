#!/usr/bin/env python3
"""Read-only OOXML inventory and structural checks. Does not render or certify Excel UI behavior."""
import argparse
from collections import defaultdict
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
}
RID = '{' + NS['r'] + '}id'
ERRORS = {'#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!'}


def local(tag):
    return tag.rsplit('}', 1)[-1]


def number(text):
    try:
        value = Decimal(str(text))
        return int(value) if value == value.to_integral_value() else float(value)
    except (InvalidOperation, ValueError):
        return text


def rich_text(root):
    return ''.join(e.text or '' for e in root.iter() if local(e.tag) == 't')


def text_appearance(root):
    if root is None:
        return {'font_properties': [], 'typefaces': []}
    properties = {}
    typefaces = set()
    for el in root.iter():
        if local(el.tag) in ('rPr', 'defRPr', 'endParaRPr'):
            relevant = {k: v for k, v in el.attrib.items() if k in ('sz', 'b', 'i', 'u', 'strike', 'spc', 'baseline')}
            if relevant:
                properties[tuple(sorted(relevant.items()))] = relevant
        if 'typeface' in el.attrib:
            typefaces.add(el.attrib['typeface'])
    return {'font_properties': list(properties.values()), 'typefaces': sorted(typefaces)}


class Package:
    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(path) as archive:
            bad = archive.testzip()
            if bad:
                raise ValueError('ZIP CRC failure: ' + bad)
            self.parts = {entry.filename: archive.read(entry) for entry in archive.infolist()}
        self.book = self.xml('xl/workbook.xml')
        self.strings = []
        if 'xl/sharedStrings.xml' in self.parts:
            self.strings = [rich_text(e) for e in self.xml('xl/sharedStrings.xml')]
        self.sheets = {}
        rels = self.rels('xl/workbook.xml')
        for sheet in self.book.findall('s:sheets/s:sheet', NS):
            self.sheets[sheet.attrib['name']] = rels[sheet.attrib[RID]]['path']

    def xml(self, path):
        return ET.fromstring(self.parts[path])

    def rels(self, part):
        path = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
        if path not in self.parts:
            return {}
        result = {}
        for rel in self.xml(path):
            a = rel.attrib
            external = a.get('TargetMode') == 'External'
            target = a['Target']
            resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
            result[a['Id']] = {'path': resolved, 'type': a['Type'].rsplit('/', 1)[-1], 'external': external}
        return result

    def cells(self, path):
        result = {}
        for cell in self.xml(path).findall('s:sheetData/s:row/s:c', NS):
            value = cell.find('s:v', NS)
            formula = cell.find('s:f', NS)
            kind = cell.attrib.get('t', 'n')
            text = value.text if value is not None else None
            if kind == 's' and text is not None:
                text = self.strings[int(text)]
            elif kind == 'inlineStr':
                text = rich_text(cell)
            elif kind == 'n' and text is not None:
                text = number(text)
            if text is not None or formula is not None:
                result[cell.attrib['r']] = {'value': text, 'type': kind, 'formula': formula.text if formula is not None else None, 'formula_attributes': dict(formula.attrib) if formula is not None else {}, 'style': cell.attrib.get('s', '0')}
        return result

    def broken_relationships(self):
        broken = []
        for path in self.parts:
            if not path.endswith('.rels'):
                continue
            if path == '_rels/.rels':
                origin = ''
            else:
                folder, filename = path.rsplit('/_rels/', 1)
                origin = posixpath.join(folder, filename[:-5])
            for rel in self.xml(path):
                if rel.attrib.get('TargetMode') == 'External':
                    continue
                target = rel.attrib['Target'].split('#', 1)[0]
                resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(origin), target))
                if resolved not in self.parts:
                    broken.append({'relationship': path, 'target': resolved})
        return broken


def chart_info(pkg, path):
    root = pkg.xml(path)
    plot = root.find('c:chart/c:plotArea', NS)
    forms = []
    series = []
    if plot is not None:
        for obj in plot:
            kind = local(obj.tag)
            if not kind.endswith('Chart'):
                continue
            details = [kind]
            for name in ['barDir', 'grouping', 'ofPieType']:
                el = obj.find('c:' + name, NS)
                if el is not None:
                    details.append(name + '=' + el.attrib.get('val', ''))
            forms.append(':'.join(details))
            for ser in obj.findall('c:ser', NS):
                series.append({
                    'formulas': [e.text for e in ser.findall('.//c:f', NS)],
                    'category_cache': [e.text for branch in ser if local(branch.tag) in ('cat', 'xVal') for e in branch.findall('.//c:pt/c:v', NS)],
                    'value_cache': [number(e.text) for branch in ser if local(branch.tag) in ('val', 'yVal', 'bubbleSize') for e in branch.findall('.//c:pt/c:v', NS)],
                })
    title = root.find('c:chart/c:title', NS)
    title_text = rich_text(title) if title is not None else ''
    if not title_text and title is not None:
        title_text = ''.join(e.text or '' for e in title.findall('.//c:strCache/c:pt/c:v', NS))
    legend = root.find('c:chart/c:legend', NS)
    legend_pos = legend.find('c:legendPos', NS) if legend is not None else None
    pivot = root.find('c:pivotSource/c:name', NS)
    buttons = {local(e.tag): e.attrib.get('val', '') for e in root.iter() if local(e.tag) in ('dropZoneFilter', 'dropZoneCategories', 'dropZoneData', 'dropZoneSeries', 'dropZonesVisible', 'showExpandCollapseFieldButtons')}
    appearance = text_appearance(root)
    appearance['title_font_sizes'] = sorted({int(e.attrib['sz']) for e in title.iter() if 'sz' in e.attrib}) if title is not None else []
    area = root.find('c:spPr', NS)
    appearance['chart_area_properties'] = ET.tostring(area, encoding='unicode') if area is not None else None
    return {
        'part': path, 'title': title_text,
        'title_formula': [e.text for e in title.findall('.//c:f', NS)] if title is not None else [],
        'form': '+'.join(forms) if forms else 'unrecognized',
        'legend': None if legend is None else (legend_pos.attrib.get('val') if legend_pos is not None else 'default'),
        'series': series, 'pivot_source': pivot.text if pivot is not None else None,
        'field_buttons': buttons,
        'appearance': appearance,
        'axis_titles': [rich_text(e) for e in root.findall('.//c:catAx/c:title', NS) + root.findall('.//c:valAx/c:title', NS)],
    }


def drawing_info(pkg, path):
    texts, charts, anchors, other, text_styles = [], [], [], [], []
    rels = pkg.rels(path)
    for anchor in pkg.xml(path):
        label = [e.attrib for e in anchor.findall('.//xdr:cNvPr', NS)]
        for shape in anchor.findall('.//xdr:sp', NS):
            text = rich_text(shape)
            if text:
                texts.append(text)
                properties = shape.find('xdr:spPr', NS)
                text_styles.append({'text': text, **text_appearance(shape), 'shape_properties': ET.tostring(properties, encoding='unicode') if properties is not None else None})
        location = {'kind': local(anchor.tag), 'from': {}, 'to': {}, 'names': label}
        for key in ('from', 'to'):
            branch = anchor.find('xdr:' + key, NS)
            if branch is not None:
                location[key] = {local(e.tag): int(e.text or 0) for e in branch}
        ext = anchor.find('xdr:ext', NS)
        if ext is not None:
            location['extent'] = dict(ext.attrib)
        for el in anchor.iter():
            if local(el.tag) in ('chart', 'chartEx'):
                rid = el.attrib.get(RID)
                relation = rels.get(rid)
                if relation and relation['path'] in pkg.parts:
                    if local(el.tag) == 'chart':
                        chart = chart_info(pkg, relation['path'])
                    else:
                        chart = {'part': relation['path'], 'title': '', 'form': 'chartEx:needs_native_review', 'legend': None, 'series': [], 'pivot_source': None, 'field_buttons': {}, 'axis_titles': []}
                    chart['anchor'] = location
                    charts.append(chart)
        if not any(local(e.tag) in ('chart', 'chartEx') for e in anchor.iter()):
            other.append(location)
        anchors.append(location)
    return texts, charts, anchors, other, text_styles


def sales_metrics(pkg):
    if '案例数据' not in pkg.sheets:
        return {'supported': False, 'reason': 'No 案例数据 sheet'}
    cells = pkg.cells(pkg.sheets['案例数据'])
    headers = {re.sub(r'\d+$', '', ref): cell['value'] for ref, cell in cells.items() if re.search(r'\d+$', ref).group() == '1'}
    required = ['订单金额', '订单数量']
    if not all(k in headers.values() for k in required):
        return {'supported': False, 'headers': headers, 'reason': 'Sales amount and quantity fields are not both present'}
    rows = defaultdict(dict)
    for ref, cell in cells.items():
        match = re.fullmatch(r'([A-Z]+)(\d+)', ref)
        if int(match[2]) > 1 and match[1] in headers:
            rows[int(match[2])][headers[match[1]]] = cell['value']
    total_amount = Decimal(0)
    total_quantity = Decimal(0)
    groups = {key: defaultdict(lambda: [Decimal(0), Decimal(0)]) for key in ['month', 'product', 'market', 'employee']}
    missing = []
    count = 0
    for rownum, row in sorted(rows.items()):
        if not any(v is not None and v != '' for v in row.values()):
            continue
        # Exclude explicit total rows, not legitimate records with zero amount.
        if any(str(v) in ('总计', '合计') for v in row.values()):
            continue
        count += 1
        amount, quantity = row.get('订单金额'), row.get('订单数量')
        if not isinstance(amount, (int, float)) or not isinstance(quantity, (int, float)):
            missing.append(rownum)
            continue
        amount, quantity = Decimal(str(amount)), Decimal(str(quantity))
        total_amount += amount
        total_quantity += quantity
        for key, field in [('product', '产品'), ('market', '所属市场'), ('employee', '员工')]:
            value = row.get(field)
            if value is not None:
                groups[key][str(value)][0] += amount
                groups[key][str(value)][1] += quantity
        # Prefer an actual year-month to avoid pooling different years.
        date = row.get('日期')
        month = row.get('下单月')
        if isinstance(date, (int, float)):
            settings = pkg.book.find('s:workbookPr', NS)
            epoch1904 = settings is not None and settings.attrib.get('date1904') in ('1', 'true')
            date = datetime(1904, 1, 1) + timedelta(days=date) if epoch1904 else datetime(1899, 12, 30) + timedelta(days=date)
            month = date.strftime('%Y-%m')
        elif isinstance(date, str) and re.match(r'^\d{4}[-/]', date):
            try:
                month = datetime.fromisoformat(date.replace('/', '-')).strftime('%Y-%m')
            except ValueError:
                pass
        if month is not None:
            groups['month'][str(month)][0] += amount
            groups['month'][str(month)][1] += quantity
    result = {'supported': True, 'record_count': count, 'total_amount': number(total_amount) if not missing else None, 'total_quantity': number(total_quantity) if not missing else None, 'unresolved_numeric_rows': missing, 'complete': not missing, 'groups': {key: [{'category': cat, 'amount': number(values[0]), 'quantity': number(values[1])} for cat, values in sorted(g.items())] for key, g in groups.items()} if not missing else {}}
    if missing:
        result['known_rows_subtotal'] = {'amount': number(total_amount), 'quantity': number(total_quantity)}
    return result


def inspect_package(pkg):
    result = {'path': str(pkg.path.resolve()), 'sha256': hashlib.sha256(pkg.path.read_bytes()).hexdigest(), 'sheets': [], 'chart_count': 0, 'broken_relationships': pkg.broken_relationships(), 'data': sales_metrics(pkg)}
    style_root = pkg.xml('xl/styles.xml') if 'xl/styles.xml' in pkg.parts else None
    result['workbook_styles'] = {key: [ET.tostring(e, encoding='unicode') for e in style_root.findall('s:' + key + '/*', NS)] if style_root is not None else [] for key in ('fonts', 'fills', 'borders', 'cellXfs')}
    for name, path in pkg.sheets.items():
        root = pkg.xml(path)
        cells = pkg.cells(path)
        texts, charts, anchors, other, text_styles = [], [], [], [], []
        pivots = []
        for rel in pkg.rels(path).values():
            if rel['path'] not in pkg.parts:
                continue
            if rel['type'] == 'drawing':
                t, c, a, o, ts = drawing_info(pkg, rel['path'])
                texts.extend(t); charts.extend(c); anchors.extend(a); other.extend(o); text_styles.extend(ts)
            elif rel['type'] == 'pivotTable':
                pivot = pkg.xml(rel['path'])
                pivots.append({'part': rel['path'], 'name': pivot.attrib.get('name'), 'cache_id': pivot.attrib.get('cacheId'), 'data_fields': [dict(e.attrib) for e in pivot.findall('s:dataFields/s:dataField', NS)], 'page_fields': [dict(e.attrib) for e in pivot.findall('s:pageFields/s:pageField', NS)]})
        dimension = root.find('s:dimension', NS)
        sheet = {'name': name, 'part': path, 'dimension': dimension.attrib.get('ref') if dimension is not None else None, 'nonempty_cells': len(cells), 'preview': [[ref, cell['value'], cell['formula']] for ref, cell in list(cells.items())[:25]], 'texts': texts, 'charts': charts, 'pivot_tables': pivots, 'anchors': anchors, 'other_objects': other, 'formula_errors': [[ref, cell['value']] for ref, cell in cells.items() if cell['type'] == 'e' and cell['value'] in ERRORS], 'row_sizes': [dict(e.attrib) for e in root.findall('s:sheetData/s:row', NS) if 'ht' in e.attrib or 'hidden' in e.attrib], 'column_sizes': [dict(e.attrib) for e in root.findall('s:cols/s:col', NS)]}
        result['sheets'].append(sheet)
        sheet['text_styles'] = text_styles
        result['chart_count'] += len(charts)
    return result


def inspect_workbook(path):
    return inspect_package(Package(path))


def normalize_formula(formula):
    # Quotes around a sheet name are not a different data source.
    return re.sub(r"'([^']+)'!", lambda m: m[1].replace("''", "'") + '!', formula or '')


def signature(chart):
    return (chart['form'], tuple(tuple(normalize_formula(f) for f in s['formulas']) for s in chart['series']), chart['pivot_source'])


def chart_rectangles(pkg, sheet, normalize_implicit=False):
    """Approximate column metrics; exact EMU extents and point-based row heights."""
    root=pkg.xml(sheet['part']);fmt=root.find('s:sheetFormatPr',NS)
    default_width=float(fmt.attrib.get('defaultColWidth','9')) if fmt is not None else 9
    default_height=float(fmt.attrib.get('defaultRowHeight','14')) if fmt is not None else 14
    anchors=[c.get('anchor',{}) for c in sheet['charts']]
    maxcol=max([12]+[a.get(k,{}).get('col',0)+1 for a in anchors for k in ('from','to')])
    maxrow=max([1]+[a.get(k,{}).get('row',0)+1 for a in anchors for k in ('from','to')])
    widths=[int(default_width*7+5)*9525]*maxcol
    for node in root.findall('s:cols/s:col',NS):
        for i in range(int(node.attrib['min'])-1,min(maxcol,int(node.attrib['max']))):
            widths[i]=0 if node.attrib.get('hidden')=='1' else int(float(node.attrib.get('width',default_width))*7+5)*9525
    fallback=14 if normalize_implicit and default_height>14 else default_height
    heights=[fallback*12700]*maxrow
    for node in root.findall('s:sheetData/s:row',NS):
        i=int(node.attrib['r'])-1
        if i>=maxrow:continue
        if node.attrib.get('hidden')=='1':heights[i]=0
        elif 'ht' in node.attrib:heights[i]=float(node.attrib['ht'])*12700
    def point(marker):
        return sum(widths[:marker.get('col',0)])+marker.get('colOff',0),sum(heights[:marker.get('row',0)])+marker.get('rowOff',0)
    result=[]
    for chart in sheet['charts']:
        a=chart.get('anchor',{})
        if not a.get('from'):continue
        x,y=point(a['from'])
        if a.get('extent'):w,h=float(a['extent']['cx']),float(a['extent']['cy'])
        elif a.get('to'):
            right,bottom=point(a['to']);w,h=right-x,bottom-y
        else:continue
        result.append({'part':chart['part'],'x':x,'y':y,'width':w,'height':h})
    return result


def overlapping_charts(rectangles):
    pairs=[]
    for i,a in enumerate(rectangles):
        for b in rectangles[i+1:]:
            if min(a['x']+a['width'],b['x']+b['width'])-max(a['x'],b['x'])>9525 and min(a['y']+a['height'],b['y']+b['height'])-max(a['y'],b['y'])>9525:
                pairs.append([a['part'],b['part']])
    return pairs


def check_workbook(path, source=None, dashboard=None, reference=None):
    pkg = Package(path)
    info = inspect_package(pkg)
    errors, warnings = [], []
    def error(code, **detail):
        errors.append({'code': code, **detail})
    def warn(code, **detail):
        warnings.append({'code': code, **detail})
    sheets = {s['name']: s for s in info['sheets']}
    analyses = [s for s in info['sheets'] if re.fullmatch(r'分析内容[1-9]\d*', s['name'])]
    analyses.sort(key=lambda s: int(re.search(r'\d+$', s['name'])[0]))
    if len(analyses) < 6:
        error('analysis_count', found=len(analyses), minimum=6)
    if analyses and [int(re.search(r'\d+$', s['name'])[0]) for s in analyses] != list(range(1, len(analyses)+1)):
        error('analysis_sequence', sheets=[s['name'] for s in analyses])
    for broken in info['broken_relationships']:
        error('broken_relationship', **broken)
    if '需求分析' not in sheets:
        error('missing_requirements_sheet')
    else:
        values = pkg.cells(pkg.sheets['需求分析'])
        filled = {int(re.search(r'\d+$', ref)[0]) for ref, cell in values.items() if ref.startswith('B') and cell['value'] not in (None, '') and ref != 'B1'}
        if len(filled) != len(analyses):
            error('requirements_count', filled_rows=len(filled), analyses=len(analyses))
    candidates = [s for s in info['sheets'] if re.fullmatch(r'模[板版]设计\d+', s['name']) and s['charts']]
    board = sheets.get(dashboard) if dashboard else (max(candidates, key=lambda s: len(s['charts'])) if candidates else None)
    if board is None:
        error('missing_dashboard', requested=dashboard)
        board_charts = []
    else:
        board_charts = board['charts']
        rectangles=chart_rectangles(pkg,board)
        for rectangle in rectangles:
            if rectangle['height']>0 and rectangle['width']/rectangle['height']<1.4:
                error('chart_proportion',chart=rectangle['part'],width_height_ratio=round(rectangle['width']/rectangle['height'],3))
        for pair in overlapping_charts(rectangles):error('chart_overlap',charts=pair)
        native_pairs=overlapping_charts(chart_rectangles(pkg,board,normalize_implicit=True))
        for pair in native_pairs:
            if pair not in overlapping_charts(rectangles):error('chart_overlap_after_row_normalization',charts=pair)
    if len(board_charts) < 6:
        error('dashboard_chart_count', found=len(board_charts), minimum=6)
    forms = sorted(set(c['form'] for c in board_charts if not c['form'].startswith(('unrecognized', 'chartEx:'))))
    if len(forms) < 6:
        error('chart_form_count', forms=forms, minimum=6)
    if len(board_charts) != len(analyses):
        error('dashboard_analysis_count', charts=len(board_charts), analyses=len(analyses))
    paired = 0
    for sheet in analyses:
        if not sheet['charts']:
            error('missing_analysis_chart', sheet=sheet['name'])
        if not sheet['pivot_tables']:
            error('missing_pivot_table', sheet=sheet['name'])
        for chart in sheet['charts']:
            matches = [c for c in board_charts if signature(c) == signature(chart)]
            if not matches:
                error('unpaired_analysis_chart', sheet=sheet['name'], chart=chart['part'])
            else:
                paired += 1
                if matches[0]['series'] != chart['series']:
                    error('paired_cache_mismatch', sheet=sheet['name'], chart=chart['part'])
    dynamic_candidates = []
    for sheet in info['sheets']:
        for entry in sheet['formula_errors']:
            error('formula_error', sheet=sheet['name'], cell=entry[0], value=entry[1])
        for chart in sheet['charts']:
            if not chart['title'] and not chart.get('title_formula'):
                error('missing_chart_title', sheet=sheet['name'], chart=chart['part'])
            if chart['legend'] not in (None, 't'):
                error('legend_position', sheet=sheet['name'], chart=chart['part'], position=chart['legend'])
            if chart['form'].startswith('chartEx:'):
                warn('unsupported_chart_format', chart=chart['part'])
            for series in chart['series']:
                if any(str(v).strip() in ('总计', '合计', 'Grand Total') for v in series['category_cache']):
                    error('grand_total_plotted', sheet=sheet['name'], chart=chart['part'])
                if any('#REF!' in (f or '') for f in series['formulas']):
                    error('chart_reference_error', chart=chart['part'])
            pivot_source = chart['pivot_source'] or ''
            bracket = re.search(r'\[([^\]]+)\]', pivot_source)
            if bracket and bracket[1] != Path(path).name:
                warn('pivot_workbook_prefix_needs_review', chart=chart['part'], prefix=bracket[1], actual=Path(path).name)
            buttons = chart['field_buttons']
            if sheet is board and buttons.get('dropZoneFilter') == '1' and buttons.get('dropZonesVisible', '1') == '1':
                dynamic_candidates.append(chart['part'])
    slicers = [p for p in pkg.parts if '/slicers/' in p and p.endswith('.xml')]
    if not dynamic_candidates and not slicers:
        warn('no_native_filter_evidence', message='Check a visible useful native control in the actual Excel engine; OOXML flags are not interaction proof.')
    if source:
        src = Package(source)
        source_sheets = {s['name']: s for s in inspect_package(src)['sheets']}
        for name in src.sheets:
            if name not in pkg.sheets:
                error('missing_source_sheet', sheet=name)
        if '案例数据' in src.sheets and '案例数据' in pkg.sheets:
            before, after = src.cells(src.sheets['案例数据']), pkg.cells(pkg.sheets['案例数据'])
            changed = []
            for ref in sorted(set(before) | set(after)):
                a, b = before.get(ref), after.get(ref)
                if a is None or b is None or any(a[k] != b[k] for k in ('value', 'formula', 'formula_attributes')):
                    changed.append(ref)
            if changed:
                error('source_data_changed', cells=changed[:20], total_changes=len(changed))
        # Nonselected sheets are preserved semantically; object styling remains a visual gate.
        for name in src.sheets:
            if name not in sheets or name == '需求分析' or (board is not None and name == board['name']):
                continue
            origin = source_sheets[name]
            if origin['texts'] != sheets[name]['texts']:
                error('source_text_boxes_changed', sheet=name)
    if info['data'].get('unresolved_numeric_rows'):
        warn('unresolved_numeric_values', rows=info['data']['unresolved_numeric_rows'])
    if reference:
        ref = inspect_workbook(reference)
        ref_boards = [s for s in ref['sheets'] if re.fullmatch(r'模[板版]设计\d+', s['name']) and s['charts']]
        warn('reference_style_manual_review', reference=str(Path(reference).resolve()), source_boards=[s['name'] for s in ref_boards], message='Compare inherited fonts, fills, text-box styling and composition visually; apply latest user overrides for dimensions, separate backgrounds and direct pie labels.')
    return {'validation_scope': 'static_package_only', 'ok': not errors, 'path': info['path'], 'dashboard': board['name'] if board else None, 'analysis_count': len(analyses), 'dashboard_chart_count': len(board_charts), 'chart_forms': forms, 'paired_analysis_charts': paired, 'dynamic_candidates': dynamic_candidates, 'slicer_parts': slicers, 'errors': errors, 'warnings': warnings, 'manual_checks': ['Independently reconcile every chosen pivot and chart with current source records and correct units/denominators.', 'Operate a meaningful dashboard control for all and two categories; confirm both chart instances update; restore delivery state.', 'Check native editability and preserve original text-box styling, template fonts/fills and unused sheets.', 'Visually inspect every analysis and the complete dashboard: equal horizontal lengths, type-appropriate vertical heights equal within each row, full inset borders; white analysis chart areas, transparent dashboard chart areas; no pie legend and readable category/percentage labels inside slices; check titles, legends, axes, text boxes and data lines against the actual template background and the separate white analysis canvas.', 'Save and reopen in the intended spreadsheet engine without repair prompts; compare TXT names/order/conclusions with final saved data.']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['inspect', 'check'])
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--source', type=Path)
    parser.add_argument('--dashboard')
    parser.add_argument('--reference', type=Path)
    parser.add_argument('--output', type=Path, help='Write private QA JSON; workbook files are always read-only.')
    args = parser.parse_args()
    try:
        result = inspect_workbook(args.workbook) if args.command == 'inspect' else check_workbook(args.workbook, source=args.source, dashboard=args.dashboard, reference=args.reference)
        payload = json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False)
        if args.output:
            protected = {p.resolve() for p in (args.workbook, args.source, args.reference) if p is not None}
            if args.output.resolve() in protected:
                raise ValueError('QA output must not overwrite a workbook input')
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(payload + '\n', encoding='utf-8')
            print(json.dumps({'report': str(args.output.resolve()), 'ok': result.get('ok'), 'scope': result.get('validation_scope', 'inventory')}, ensure_ascii=False))
        else:
            print(payload)
        return 1 if args.command == 'check' and not result['ok'] else 0
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(json.dumps({'error': str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
