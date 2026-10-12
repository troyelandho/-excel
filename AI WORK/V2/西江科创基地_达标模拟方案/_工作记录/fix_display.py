from pathlib import Path
import json
from patch_case import parts,mapping,save,serialize,S
from lxml import etree as E
h=Path(__file__).resolve().parent
m=json.loads((h/'case_manifest.json').read_text(encoding='utf-8'))
out=Path(m['root'])/'西江科创基地_投资测算_达标模拟.xlsx'
d=parts(out);member=mapping(d)['核心参数输入'];t=E.fromstring(d[member])
r=t.find(S+'sheetData').find(S+'row[@r="7"]');r.set('ht','36');r.set('customHeight','1')
d[member]=serialize(t)
new=out.with_name('西江科创基地_投资测算_显示修正版.xlsx');save(d,new,out)
old=h/'终稿行高调整前.xlsx'
if old.exists():raise FileExistsError(old)
out.rename(old);new.rename(out)
print('仅将项目名称所在第7行加高至36磅，避免换行文字截断。')
