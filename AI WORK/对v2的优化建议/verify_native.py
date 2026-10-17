from pathlib import Path
import json,math,re
from openpyxl.utils.cell import column_index_from_string
root=Path(__file__).parent
def value(case,sheet,address):
 s=data[case]['sheets'][sheet];col,row=re.fullmatch(r'([A-Z]+)(\d+)',address).groups();off=1 if s[0][0] is None else 2
 return s[int(row)-off][column_index_from_string(col)-1]
data={p.stem[7:]:json.loads(p.read_text(encoding='utf-8-sig')) for p in root.glob('native_*.json') if p.name!='native_validation.json'}
results=[]
def check(name,test):
 results.append({'check':name,'pass':bool(test)})
def close(a,b):return isinstance(a,(int,float)) and isinstance(b,(int,float)) and math.isclose(a,b,abs_tol=1e-7,rel_tol=1e-9)
check('all completed cases contain no Excel formula errors',all(not d['errors'] for d in data.values()))
if 'pressure_zero' in data:
 check('zero stress reproduces baseline for every metric',all(close(value('pressure_zero','指标与情景测算',f'B{r}'),value('pressure_zero','指标与情景测算',f'D{r}')) or value('pressure_zero','指标与情景测算',f'B{r}')==value('pressure_zero','指标与情景测算',f'D{r}') for r in range(6,29)))
if 'no_loan' in data:
 check('no loan gives equal project/equity IRR and DSCR not applicable',close(value('no_loan','指标与情景测算','B6'),value('no_loan','指标与情景测算','B7')) and value('no_loan','指标与情景测算','B12')=='不适用')
if 'equity_10000' in data:
 check('capital shortage 22810 without changing theoretical IRR',close(value('equity_10000','指标与情景测算','B14'),22810) and close(value('equity_10000','指标与情景测算','B7'),value('base','指标与情景测算','B7')))
 check('capital shortage result labelled theoretical',value('equity_10000','指标与情景测算','B27')=='理论结果，资金尚未落实')
if 'disposal' in data:
 check('disposal only year 20, net proceeds 29100',all(close(value('disposal','销售与出租收入',f'{col}26'),29100 if t==20 else 0) for t,col in enumerate(['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH'])))
if 'horizon_10_disposal' in data:
 check('horizon 10 moves disposal and working capital recovery',close(value('horizon_10_disposal','销售与出租收入','N26'),29100) and close(value('horizon_10_disposal','投资与资金安排','N32'),500) and close(value('horizon_10_disposal','利润与现金流','X26'),0))
if 'horizon_30' in data:
 check('horizon 30 moves working capital recovery to year 30',close(value('horizon_30','投资与资金安排','AH32'),500) and close(value('horizon_30','投资与资金安排','X32'),0))
if 'first_draw_3' in data:
 check('first draw 3 means maturity 12',close(value('first_draw_3','融资与还本付息','B6'),3) and close(value('first_draw_3','融资与还本付息','B7'),12))
for case,cell in [('horizon_2_invalid','D36'),('pure_hold','D37'),('pure_sale','D37'),('blank_rent','D35'),('negative_cost','D46'),('zero_collection','D45')]:
 if case in data: check(case+' flagged',value(case,'决策看板',cell)=='未通过' and value(case,'决策看板','A5')=='待判断：请检查参数')
if 'high_sale_price' in data:
 for metric,cfrow in [('B10',26),('B11',50)]:
  cum=0;pb=None
  for t in range(31):
   col=__import__('openpyxl').utils.cell.get_column_letter(t+4);flow=value('high_sale_price','利润与现金流',f'{col}{cfrow}');old=cum;cum+=flow
   if t>0 and old<0<=cum:pb=t-1-old/flow;break
  actual=value('high_sale_price','指标与情景测算',metric)
  check('independent payback '+metric,close(pb,actual) if pb is not None else actual=='测算期内未回收')
output={'case_count':len(data),'checks':results,'all_pass':all(c['pass'] for c in results)}
(root/'native_validation.json').write_text(json.dumps(output,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(output,ensure_ascii=False,indent=2))
