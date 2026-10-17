import json
from pathlib import Path
n='#,##0.00;[Red](#,##0.00);"—"'
p='0.00%;[Red](0.00%);"—"'
xs=[{'sheet':s,'cell':c,'numberFormat':n} for s,c in [('指标与情景','B6:B54'),('决策看板','B9:D13'),('决策看板','B20:B27'),('决策看板','B32:D32'),('决策看板','B36:B37')]]
xs += [{'sheet':'指标与情景','cell':'B'+str(i),'numberFormat':p} for i in [6,7,19,22,23,24,38,39,42,43]]
xs += [{'sheet':'核心参数输入','cell':'B'+str(i),'numberFormat':'0'} for i in [8,9,16,17,18,53,59,60,64,73,74,80,82,95,96,97,98,99,100,101,103,104,105,106,107,111,112,116,124,126,127,136]]
xs += [{'sheet':s,'cell':c,'numberFormat':'0'} for s,c in [('季度投融资','B28:B31'),('季度投融资','E28:E31'),('决策看板','F8'),('决策看板','B35')]]
(Path(__file__).parent/'formula_repairs.json').write_text(json.dumps(xs,ensure_ascii=False),encoding='utf-8')
