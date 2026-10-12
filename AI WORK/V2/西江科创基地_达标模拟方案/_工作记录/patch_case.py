from pathlib import Path
import json, re, zipfile, copy, hashlib, argparse
from lxml import etree as E

HERE=Path(__file__).resolve().parent
M=json.loads((HERE/'case_manifest.json').read_text(encoding='utf-8'))
N={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
   'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
   'p':'http://schemas.openxmlformats.org/package/2006/relationships',
   'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
S='{'+N['s']+'}'
def parts(p):
    with zipfile.ZipFile(p) as z:return {i.filename:z.read(i) for i in z.infolist()}
def mapping(d):
    rel={x.get('Id'):x.get('Target') for x in E.fromstring(d['xl/_rels/workbook.xml.rels'])}
    return {x.get('name'):'xl/'+rel[x.get('{'+N['r']+'}id')].lstrip('/') if not rel[x.get('{'+N['r']+'}id')].startswith('/') else rel[x.get('{'+N['r']+'}id')].lstrip('/') for x in E.fromstring(d['xl/workbook.xml']).find(S+'sheets')}
def cell_map(xml):return {c.get('r'):c for c in E.fromstring(xml).findall('.//'+S+'sheetData/'+S+'row/'+S+'c')}
def serialize(t):return E.tostring(t,encoding='UTF-8',xml_declaration=True,standalone=True)
def save(d,target,original):
    if target.exists():raise FileExistsError(target)
    with zipfile.ZipFile(original) as src,zipfile.ZipFile(target,'x',zipfile.ZIP_DEFLATED) as dst:
        for i in src.infolist():dst.writestr(i,d[i.filename])
def write_payload(old,new,donor_ss):
    if old.find(S+'f') is not None:raise AssertionError('禁止改写公式'+old.get('r'))
    for c in list(old):old.remove(c)
    t=new.get('t')
    if t=='s':
        text=''.join(donor_ss[int(new.findtext(S+'v'))].itertext())
        old.set('t','inlineStr');o=E.SubElement(old,S+'is');E.SubElement(o,S+'t').text=text
    else:
        if t is None:old.attrib.pop('t',None)
        else:old.set('t',t)
        for c in list(new):old.append(copy.deepcopy(c))
def initial():
    src=Path(M['source']); d=parts(src); mp=mapping(d)
    donor=parts(HERE/'授权输入内容.xlsx'); dm=mapping(donor)
    ss=E.fromstring(donor['xl/sharedStrings.xml']) if 'xl/sharedStrings.xml' in donor else []
    for sheet,changes in M['changes'].items():
        root=E.fromstring(d[mp[sheet]]); sd=root.find(S+'sheetData')
        cm={c.get('r'):c for c in sd.findall('./'+S+'row/'+S+'c')}
        rm={int(r.get('r')):r for r in sd}; dc=cell_map(donor[dm[sheet]])
        for address,value in changes.items():
            row=int(re.search(r'\d+',address).group());col=re.match('[A-Z]+',address).group()
            if row not in rm:
                r=E.Element(S+'row',r=str(row));sd.append(r);rm[row]=r
            if address not in cm:
                style_ref=(col+'5') if row==96 else (col+'7') if col in 'ABCD' else 'D7'
                style=cm.get(style_ref).get('s','0')
                cell=E.Element(S+'c',r=address,s=style);rm[row].append(cell);cm[address]=cell
            write_payload(cm[address],dc[address],ss)
        if sheet=='核心参数输入':
            for address,value in changes.items():
                if address.startswith('D') and int(address[1:])<85 and len(str(value))>55:
                    row=int(address[1:]);rm[row].set('ht','30');rm[row].set('customHeight','1')
            for row in range(97,103):rm[row].set('ht','72');rm[row].set('customHeight','1')
            for row in range(104,108):rm[row].set('ht','46');rm[row].set('customHeight','1')
            rm[95].set('ht','28');rm[95].set('customHeight','1')
            cm['A95'].set('s',cm['A6'].get('s'))
            merges=root.find(S+'mergeCells')
            E.SubElement(merges,S+'mergeCell',ref='A95:K95')
            for row in range(96,108):E.SubElement(merges,S+'mergeCell',ref=f'D{row}:K{row}')
            merges.set('count',str(len(merges)))
            root.find(S+'dimension').set('ref','A1:AH107')
            # 来源名称保持窄列，长信息在D列可读；B列取简短标题。
        def col_no(addr):
            k=0
            for c in re.match('[A-Z]+',addr).group():k=k*26+ord(c)-64
            return k
        for r in rm.values():
            cells=sorted(list(r),key=lambda c:col_no(c.get('r')))
            r[:]=cells
        sd[:]=sorted(list(sd),key=lambda r:int(r.get('r')))
        d[mp[sheet]]=serialize(root)
    target=HERE/'参数已填入_待重算.xlsx';save(d,target,src)
    print(target)
def cached_final():
    src=HERE/'参数已填入_待重算.xlsx';d=parts(src); calc=parts(HERE/'Excel重算副本.xlsx')
    mp=mapping(d);cp=mapping(calc);cache_count=0
    for sheet,member in mp.items():
        root=E.fromstring(d[member]);cc=cell_map(calc[cp[sheet]])
        for cell in root.findall('.//'+S+'sheetData/'+S+'row/'+S+'c'):
            if cell.find(S+'f') is None:continue
            donor=cc[cell.get('r')];v=donor.find(S+'v')
            if v is None:raise AssertionError('重算缓存缺失'+sheet+cell.get('r'))
            for old in cell.findall(S+'v'):cell.remove(old)
            cell.append(copy.deepcopy(v))
            if donor.get('t') is None:cell.attrib.pop('t',None)
            else:cell.set('t',donor.get('t'))
            cache_count+=1
        d[member]=serialize(root)
    for member in d:
        if not re.match(r'xl/charts/chart\d+\.xml$',member):continue
        original=E.fromstring(d[member]);native=E.fromstring(calc[member])
        # 仅刷新图表缓存；原图表类型、坐标、标题、公式引用及视觉设置不动。
        for ref in original.xpath('//c:numRef|//c:strRef',namespaces=N):
            f=ref.find('{'+N['c']+'}f')
            candidates=native.xpath('//c:numRef|//c:strRef',namespaces=N)
            donor=next((c for c in candidates if c.findtext('{'+N['c']+'}f')==f.text),None)
            if donor is None:raise AssertionError('图表缓存引用未找到'+f.text)
            for child in list(ref):
                if E.QName(child).localname.endswith('Cache'):ref.remove(child)
            for child in list(donor):
                if E.QName(child).localname.endswith('Cache'):ref.append(copy.deepcopy(child))
        d[member]=serialize(original)
    target=Path(M['root'])/'西江科创基地_投资测算_达标模拟.xlsx';save(d,target,src)
    print(json.dumps({'final':str(target),'formula_caches_refreshed':cache_count},ensure_ascii=False))
if __name__=='__main__':
    if '--final' in __import__('sys').argv:cached_final()
    else:initial()
