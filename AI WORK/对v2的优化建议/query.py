import sys,json,re
from pathlib import Path
d=json.loads((Path(__file__).parent/'workbook.json').read_text(encoding='utf-8'))
sheet=sys.argv[1]; lo=int(sys.argv[2]) if len(sys.argv)>2 else 1; hi=int(sys.argv[3]) if len(sys.argv)>3 else 999
cols=sys.argv[4].split(',') if len(sys.argv)>4 else ['A','B','C','D','E','H','K','X']
fc=sys.argv[5].split(',') if len(sys.argv)>5 else ['B','H']
s=next(s for s in d['sheets'] if s['name']==sheet)
for r in range(lo,hi+1):
    parts=[]
    for col in cols:
        a=col+str(r); c=s['cells'].get(a)
        if not c:continue
        v=c['v']; v=round(v,8) if isinstance(v,float) else v
        f=c['f'] if col in fc else None
        if f: f=re.sub(r'^=IF\([A-Z]+\$4<=核心参数输入!\$B\$9,(.*),0\)$',r'=\1',f)
        parts.append(a+': '+str(v)+(' {'+f+'}' if f else ''))
    if parts:print(' | '.join(parts))
