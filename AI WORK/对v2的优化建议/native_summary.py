import json,math,sys
from pathlib import Path
root=Path(__file__).parent
orig=json.loads((root/'workbook.json').read_text(encoding='utf-8'))
def cv(rows,cell):
    import re
    from openpyxl.utils.cell import column_index_from_string
    col,row=re.match(r'([A-Z]+)(\d+)',cell).groups()
    ri=int(row)-(1 if rows[0][0] is None else 2); ci=column_index_from_string(col)-1
    if ri<0 or ri>=len(rows) or ci>=len(rows[ri]):return None
    return rows[ri][ci]
for p in sorted(root.glob('native_*.json')):
    if p.name=='native_validation.json':continue
    d=json.loads(p.read_text(encoding='utf-8-sig')); s=d['sheets']; m=s['指标与情景测算']; b=s['决策看板']
    print(d['case'], 'errors',len(d['errors']), 'state',cv(b,'A5'),'checks_failed',[(r,cv(b,f'A{r}'),cv(b,f'D{r}')) for r in range(35,65) if cv(b,f'D{r}')!='通过'])
    print('metrics', {cell:cv(m,cell) for cell in ['B6','B7','B8','B9','B10','B11','B12','B13','B14','B15','B16','D6','D7','D8','D13','D14']})
    if d['case']=='base':
        diff=[]
        for sh in orig['sheets']:
            for cell,c in sh['cells'].items():
                v=c['v']; n=cv(s[sh['name']],cell)
                if isinstance(v,(float,int)) and isinstance(n,(float,int)):
                    equal=math.isclose(v,n,abs_tol=1e-7,rel_tol=1e-9)
                else: equal=v==n or (v is None and n=='')
                if not equal:diff.append((sh['name'],cell,v,n))
        print('CACHE_DIFFERENCES',diff[:25], 'count',len(diff))
