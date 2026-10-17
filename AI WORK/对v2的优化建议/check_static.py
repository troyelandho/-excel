from pathlib import Path
from collections import Counter,defaultdict
import json,re
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries,get_column_letter

root=Path(__file__).parent
d=json.loads((root/'workbook.json').read_text(encoding='utf-8'))
ss={s['name']:s['cells'] for s in d['sheets']}
graph={};missing=[];patterns=[]
for sn,cells in ss.items():
    rows=defaultdict(list)
    for a,c in cells.items():
        f=c['f']
        if not f:continue
        dependencies=set()
        for t in Tokenizer(f).items:
            if t.type!='OPERAND' or t.subtype!='RANGE':continue
            ref=t.value
            if '!' in ref: sh,ref=ref.rsplit('!',1);sh=sh.strip("'")
            else:sh=sn
            if sh not in ss: missing.append((sn,a,sh,ref));continue
            try:
                c1,r1,c2,r2=range_boundaries(ref)
                for r in range(r1,r2+1):
                    for col in range(c1,c2+1):
                        ad=get_column_letter(col)+str(r)
                        if ss[sh].get(ad,{}).get('f'): dependencies.add((sh,ad))
            except:missing.append((sn,a,sh,ref))
        graph[sn,a]=dependencies
        if re.match(r'^[E-Z]\d+$|^A[A-G]\d+$',a) and int(re.search(r'\d+',a)[0]) not in [4,5] and sn not in ['核心参数输入','指标与情景测算','决策看板']:
            row=int(re.search(r'\d+',a)[0])
            try:norm=Translator(f,origin=a).translate_formula('E'+str(row))
            except:norm=f
            rows[row].append((a,norm))
    for row,formulas in rows.items():
        counter=Counter(f for a,f in formulas)
        if len(counter)>1: patterns.append((sn,row,[(n,[(a) for a,f in formulas if f==ff][:4],ff) for ff,n in counter.items()]))
state={};cycles=[]
def visit(n,path):
    if state.get(n)==1:cycles.append(path[path.index(n):]+[n]);return
    if state.get(n)==2:return
    state[n]=1
    for dep in graph.get(n,[]):visit(dep,path+[n])
    state[n]=2
for n in graph:visit(n,[])
report={'formula_count':len(graph),'cycles':cycles,'unparsed_or_missing_refs':missing,'inconsistent_period_patterns':patterns}
(root/'static_check.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
