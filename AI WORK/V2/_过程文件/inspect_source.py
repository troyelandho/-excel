from pathlib import Path
import json, hashlib, zipfile, posixpath
from collections import Counter
from lxml import etree as E

BASE = Path(__file__).resolve().parents[2]
SRC = BASE / 'V1' / '交付文件' / '产业园投资测算.xlsx'
REF = Path(r'C:\Users\BBEN\Desktop\投资测算AGENT V5-CODEX\REF\城投产业园拿地前测算模板V5.5.xlsx')
NS = {'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def inspect(path):
    with zipfile.ZipFile(path) as z:
        style = E.fromstring(z.read('xl/styles.xml'))
        fonts = style.find('s:fonts', NS)
        fills = style.find('s:fills', NS)
        xfs = style.find('s:cellXfs', NS)
        rels = {n.get('Id'):posixpath.normpath(posixpath.join('xl',n.get('Target')))
                for n in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            strings=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))]
        book=E.fromstring(z.read('xl/workbook.xml'))
        out={'file':str(path),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
             'parts':len(z.namelist()),'fonts':[E.tostring(n).decode() for n in fonts],
             'fills':[E.tostring(n).decode() for n in fills],
             'xfs':[E.tostring(n).decode() for n in xfs],
             'dxfs':[E.tostring(n).decode() for n in style.find('s:dxfs',NS)] if style.find('s:dxfs',NS) is not None else [],
             'sheets':[]}
        for sn in book.find('s:sheets',NS):
            part=rels[sn.get('{'+NS['r']+'}id')]
            if part.startswith('/'):part=part.lstrip('/')
            sh=E.fromstring(z.read(part))
            cells=sh.findall('s:sheetData/s:row/s:c',NS)
            rows=[]
            for row in sh.findall('s:sheetData/s:row',NS):
                labels=[]
                for c in row:
                    v=c.find('s:v',NS); t=c.get('t')
                    txt = strings[int(v.text)] if t=='s' and v is not None else ''.join(c.find('s:is',NS).itertext()) if t=='inlineStr' else v.text if t=='str' and v is not None and c.find('s:f',NS) is None else None
                    if txt:labels.append([c.get('r'),txt,int(c.get('s','0'))])
                if labels:rows.append({'r':row.get('r'),'height':row.get('ht'),'labels':labels})
            out['sheets'].append({'name':sn.get('name'),'part':part,'dimension':sh.find('s:dimension',NS).get('ref'),
                                  'cells':len(cells),'formulas':sum(c.find('s:f',NS) is not None for c in cells),
                                  'style_counts':dict(Counter(c.get('s','0') for c in cells)),
                                  'views':E.tostring(sh.find('s:sheetViews',NS)).decode(),
                                  'cols':E.tostring(sh.find('s:cols',NS)).decode() if sh.find('s:cols',NS) is not None else '',
                                  'merges':[n.get('ref') for n in sh.findall('s:mergeCells/s:mergeCell',NS)],'rows':rows})
        return out

for name,path in [('source',SRC),('reference',REF)]:
    data=inspect(path)
    (Path(__file__).parent / (name+'_inspection.json')).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(name, data['sha256'], 'parts',data['parts'])
    for s in data['sheets']:
        print(s['name'],s['dimension'],s['cells'],s['formulas'],s['style_counts'])
        if name=='source':
            print('COLS',s['cols'])
            for row in s['rows']:
                print(row['r'], row['height'],str(row['labels'])[:550])
    if name=='source':
        for i,x in enumerate(data['xfs']):print('XF',i,x)
