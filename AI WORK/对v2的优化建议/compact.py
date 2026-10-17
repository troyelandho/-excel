import json, re
from pathlib import Path
data=json.loads((Path(__file__).parent/'workbook.json').read_text(encoding='utf-8'))
out=[]
for s in data['sheets']:
    rows={}
    for a,c in s['cells'].items(): rows.setdefault(int(re.search(r'\d+',a)[0]),[]).append((a,c))
    out.append('\n===== '+s['name']+' =====')
    compact=s['name'] not in ['使用说明','核心参数输入','指标与情景测算','决策看板']
    for row,cs in rows.items():
        if compact and row in [4,5]: continue
        vals=[]
        for a,c in cs:
            col=re.match(r'[A-Z]+',a)[0]
            if compact and c['f'] and col not in ['B','C','D','E','H','K','X','AH']: continue
            v=c['v']
            if isinstance(v,float): v=round(v,6)
            vals.append(a+'='+str(v)+(' {'+c['f']+'}' if c['f'] else ''))
        out.append(' | '.join(vals))
(Path(__file__).parent/'compact.txt').write_text('\n'.join(out),encoding='utf-8')
