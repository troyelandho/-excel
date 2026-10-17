import json,math
from pathlib import Path
root=Path(__file__).parent
data=json.loads((root/'scenario_results.json').read_text(encoding='utf-8-sig'))
by={x['case']:x for x in data};checks=[]
def ck(name,ok):checks.append({'check':name,'pass':bool(ok)})
def eq(a,b):return isinstance(a,(int,float)) and isinstance(b,(int,float)) and math.isclose(a,b,abs_tol=1e-6,rel_tol=1e-8)
base=by['base'];b=base['metrics']
ck('19 scenarios completed',len(by)==19)
ck('no Excel formula errors in any tested scenario',all(not x['errors'] for x in data))
expected_bad={'blank_rent','negative_cost','explicit_zero_override','shifted_terminal'}
ck('valid scenarios reconcile',all(not x['failedChecks'] for x in data if x['case'] not in expected_bad))
for name in expected_bad:ck(name+' explicitly flagged',len(by[name]['failedChecks'])>0)
ck('zero stress equals base',all(eq(by['zero_stress']['metrics']['B'+str(r)],by['zero_stress']['metrics']['E'+str(r)]) for r in [6,7,8,13,14]))
ck('no loan unlevered/equity IRR consistent',eq(by['no_loan']['metrics']['B6'],by['no_loan']['metrics']['B7']) and by['no_loan']['metrics']['B12']=='不适用')
ck('equity shortage is explicit and theory unchanged',eq(by['equity_10000']['metrics']['B14'],22810) and eq(by['equity_10000']['metrics']['B7'],b['B7']) and eq(by['equity_10000']['capitalTotal'],10000))
ck('pure sale has no held area/rent/operating cost',eq(by['pure_sale']['metrics']['B21'],0) and eq(by['pure_sale']['rentCash'],0) and eq(by['pure_sale']['opCost'],0))
ck('pure hold has no sale collections',eq(by['pure_hold']['saleCash'],0))
ck('loan draw lag shifts first draw and maturity',eq(by['draw_lag_4']['first'],5) and eq(by['draw_lag_4']['maturity'],44))
ck('financing fee excludes project return',eq(by['finance_fee']['metrics']['B8'],b['B8']))
ck('explicit zero does not use default',eq(by['explicit_zero_override']['firstFormation'],0) and by['explicit_zero_override']['planValid']==0)
tables=json.loads((root/'native_table_checks.json').read_text(encoding='utf-8-sig'))
for x in tables:
 if 'table' in x:ck(x['test']+' native data table matches model',eq(x['table'],x['model']))
 else:ck('monthly lag exposes deficit while quarter totals tie',eq(x['baseGap'],0) and x['lateGap']>0 and abs(x['quarterReconciliation'])<1e-6 and x['valid']==1)
report={'case_count':len(by),'checks':checks,'all_pass':all(x['pass'] for x in checks),'scenario_metrics':{k:v['metrics'] for k,v in by.items()}}
(root/'scenario_validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in report.items() if k!='scenario_metrics'},ensure_ascii=False,indent=2))
