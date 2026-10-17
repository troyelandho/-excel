from pathlib import Path
import json, hashlib, zipfile, xml.etree.ElementTree as ET
import openpyxl

root = Path(__file__).resolve().parent
source = root.parent / 'V2/产业园投资测算_V2_样式调整.xlsx'
wf = openpyxl.load_workbook(source, data_only=False)
wv = openpyxl.load_workbook(source, data_only=True)
result = {'source': str(source), 'sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
          'calculation': str(wf.calculation), 'names': [], 'sheets': []}
for name, obj in wf.defined_names.items():
    result['names'].append({'name':name, 'text':obj.attr_text})
with zipfile.ZipFile(source) as z:
    result['external_parts'] = [n for n in z.namelist() if any(t in n for t in ['externalLink','vbaProject','pivot','connections'])]
    result['chart_parts'] = [n for n in z.namelist() if n.startswith('xl/charts/chart') and n.endswith('.xml')]
lines = []
for ws in wf:
    cached = wv[ws.title]
    cells = {}
    for row in ws:
        for c in row:
            if c.value is None: continue
            cells[c.coordinate] = {'v': cached[c.coordinate].value, 'f': c.value if c.data_type == 'f' else None,
                                    'type': c.data_type, 'style': c.style_id, 'format': c.number_format,
                                    'locked': c.protection.locked,
                                    'font': str(c.font.color), 'fill':c.fill.fgColor.rgb,
                                    'comment':c.comment.text if c.comment else None}
    item = {'name':ws.title, 'state':ws.sheet_state,'dimension':ws.calculate_dimension(),
            'freeze':str(ws.freeze_panes), 'protection':str(ws.protection),
            'validations':[str(d) for d in ws.data_validations.dataValidation],
            'conditional_formats':[str(x) for x in ws.conditional_formatting],
            'merges':[str(r) for r in ws.merged_cells.ranges],
            'hidden_rows':[k for k,v in ws.row_dimensions.items() if v.hidden],
            'hidden_cols':[k for k,v in ws.column_dimensions.items() if v.hidden],
            'charts':len(ws._charts), 'cells':cells}
    result['sheets'].append(item)
    lines.append(f'\n===== {ws.title} [{item["dimension"]}] =====')
    for row in ws:
        filled = [c for c in row if c.value is not None]
        if not filled: continue
        parts = []
        for c in filled:
            v = cached[c.coordinate].value
            if isinstance(v,float): v = round(v,9)
            txt = f'{c.coordinate}: {v}'
            if c.data_type == 'f': txt += f' [={c.value[1:]}]'
            parts.append(txt)
        lines.append(' | '.join(parts))
(root/'workbook.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
(root/'workbook_dump.txt').write_text('\n'.join(lines),encoding='utf-8')
summary = {k: v for k,v in result.items() if k!='sheets'}
summary['sheets']=[{k:v for k,v in s.items() if k not in ['cells','merges','protection','validations','conditional_formats']} |
                   {'nonempty':len(s['cells']),'formulas':sum(c['f'] is not None for c in s['cells'].values()),
                    'errors':[(a,c['v']) for a,c in s['cells'].items() if c['type']=='e'],
                    'validation_count':len(s['validations']),'conditional_format_count':len(s['conditional_formats'])} for s in result['sheets']]
print(json.dumps(summary,ensure_ascii=False,indent=2,default=str))
