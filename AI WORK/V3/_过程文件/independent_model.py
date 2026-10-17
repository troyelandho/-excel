"""Independent numeric model used for QA; never writes an Excel workbook."""
from pathlib import Path
import json,math

ROOT=Path(__file__).parent
MANIFEST=json.loads((ROOT/'manifest.json').read_text(encoding='utf-8')) if (ROOT/'manifest.json').exists() else None
def calculate(params,build,sale,stress=False,overrides=None):
 p=lambda r:params[r]
 ov=overrides or {}
 N=int(p(9));end=N*4;K=int(p(16));T=range(end+1)
 share=1 if p(95)==2 else 0 if p(95)==3 else p(14)
 hold=1-share;area=p(10)*p(11)*p(12);ha=area*hold;sa=area*share
 landprice=ov.get(6,p(21));price=ov.get(7,p(31))*(1-p(78) if stress else 1);monthly=ov.get(8,p(36))*(1-p(79) if stress else 1)
 costfactor=ov.get(9,1)*(1+p(81) if stress else 1)
 delay=int(ov.get(10,p(107))+(p(124) if stress else 0));completion=K*4+delay;start=(int(p(18))-1)*4+delay+1
 def shift(series,lag):return [series[t-lag] if t>=lag else 0 for t in T]
 formation=[];signed=[]
 for t in T:
  b=t-delay
  formation.append(build[(b-1)//4+1]*p(128+(b-1)%4) if 1<=b<=K*4 else 0)
  signed.append(sale[(b-1)//4+1]*p(128+(b-1)%4) if share>0 and 1<=b<=120 else 0)
 payment=shift(formation,int(p(105)));weights=shift(payment,int(p(106)));delivered=shift(signed,int(p(126)))
 draw=[p(57)*x for x in weights];first=next((t for t,x in enumerate(draw) if x>1e-6),0);maturity=first+int(p(59))*4-1 if p(57)>0 else 0
 repaystart=int(p(60))*4+delay;interval=int(p(127));repaycount=max(1,(maturity-repaystart)//interval+1+(1 if (maturity-repaystart)%interval else 0)) if p(57)>0 else 0
 opening=[];principal=[];interest=[];closing=[];ci=[];oi=[];fee=[]
 for t in T:
  op=closing[-1] if t else 0
  pay=0
  if p(57)>0 and repaystart<=t<=maturity and ((t-repaystart)%interval==0 or t==maturity):pay=op+draw[t] if t==maturity else min(op+draw[t],p(57)/repaycount)
  it=max(0,op+draw[t]/2-pay/2)*p(58)/4
  opening.append(op);principal.append(pay);interest.append(it);closing.append(op+draw[t]-pay);ci.append(it if t<=completion else 0);oi.append(it if t>completion else 0);fee.append(p(57)*p(113) if p(57)>0 and t==first else 0)
 land=p(10)*landprice/10000+p(114)
 shared=(p(10)*p(11)*p(22)/10000+p(13)*p(23)/10000+p(24)+p(25)+p(26))*(1+p(27))*costfactor
 fit=p(115)*costfactor if ha>0 else 0;nonland=shared+fit;wc=p(28) if ha>0 else 0;net=1-(p(120) if p(96)==2 else 0)
 saleP=(land+shared*net)*share;saleE=saleP+(sum(ci)+sum(fee))*share;heldP=(shared*hold+fit)*net;heldE=heldP+(sum(ci)+sum(fee))*hold;heldland=land*hold
 invest=[(land if t==0 else 0)+nonland*payment[t]+(wc if t==completion+1 else 0) for t in T]
 q={r:[0]*(end+1) for r in range(10,77)}
 for t in T:
  q[10][t]=sa*delivered[t];q[11][t]=q[10][t]*price/10000
  q[12][t]=q[11][t]*p(102)+(q[11][t-int(p(103))] if t>=p(103) else 0)*(1-p(102))
  q[13][t]=sum(q[11][:t+1])-sum(q[12][:t+1])
  if t>=start:
   age=(t-start)//4;target=p(38+age) if age<3 else p(41);previous=0 if age==0 else p(37+age) if age<=3 else p(41)
   if stress:target=max(0,target-p(80)/100);previous=max(0,previous-p(80)/100)
   q[14][t]=target;q[15][t]=previous+(target-previous)*((t-start)%4+1)/4
   q[20][t]=monthly*(1+p(37))**age
  q[16][t]=q[17][t-1] if t else 0;q[17][t]=ha*q[15][t];q[18][t]=max(0,q[17][t]-q[16][t]);q[19][t]=q[16][t]+q[18][t]*(1-p(125))
  q[21][t]=q[19][t]*q[20][t]*3/10000
  q[22][t]=sum(q[18][t-j]*max(0,min(3,p(42)+3*p(125)-j*3)-max(0,3*p(125)-j*3)) for j in range(5) if t-j>=0)*q[20][t]/10000
  q[23][t]=max(0,q[21][t]-q[22][t]);q[24][t]=q[23][t]*(1-p(43))/(1+(p(119) if p(96)==2 else 0));q[25][t]=q[23][t]*p(43)
  q[26][t]=q[25][t-int(p(104))] if t>=p(104) else 0;q[27][t]=q[25][t]*p(44);q[28][t]=q[27][t-int(p(104))] if t>=p(104) else 0
  q[29][t]=q[11][t]-q[11][t]/(1+p(118)) if p(96)==2 else 0;q[30][t]=q[23][t]-q[23][t]/(1+p(119)) if p(96)==2 else 0;q[31][t]=q[27][t]-q[27][t]/(1+p(119)) if p(96)==2 else 0
  q[32][t]=q[11][t]-q[29][t];q[33][t]=q[23][t]+q[27][t]-q[30][t]-q[31][t];q[34][t]=q[32][t]+q[33][t]
  q[35][t]=q[11][t]*p(32);q[36][t]=q[11][t]*(p(33) if p(96)==1 else p(121))
  q[37][t]=p(47)*ha/max(1,p(108))*(1+p(48))**((t-completion-1)//4)/4 if ha>0 and t>completion else 0
  q[38][t]=(q[33][t]-q[24][t])*p(49)
  q[39][t]=p(50)*ha/max(1,p(108))/4+(p(117) if t==p(116)*4 else 0) if ha>0 and t>completion else 0
  q[40][t]=(q[23][t]+q[27][t])*(p(51) if p(96)==1 else p(122))
  q[41][t]=heldP*(1-p(54))/p(53)/4 if completion<t<=completion+p(53)*4 else 0;q[42][t]=heldE*(1-p(54))/p(53)/4 if completion<t<=completion+p(53)*4 else 0
  q[43][t]=heldland/p(99)/4 if p(98)==1 and completion<t<=completion+p(99)*4 else 0
  q[44][t]=saleP*delivered[t];q[45][t]=saleE*delivered[t]
  q[46][t]=q[34][t]-sum(q[r][t] for r in range(35,42))-q[43][t]-q[44][t]-q[24][t]
  q[47][t]=q[34][t]-sum(q[r][t] for r in range(35,41))-q[42][t]-q[43][t]-q[45][t]-q[24][t]-oi[t]
  q[49][t]=nonland*payment[t]*p(120) if p(96)==2 else 0;q[50][t]=q[52][t-1] if t else 0
  q[51][t]=max(0,sum(q[r][t] for r in [29,30,31])-q[49][t]-q[50][t]);q[52][t]=max(0,q[49][t]+q[50][t]-sum(q[r][t] for r in [29,30,31]))
  q[55][t]=heldland+heldP*sum(formation[:t+1])-sum(q[41][:t+1])-sum(q[43][:t+1])
  q[56][t]=q[55][t]+sum(ci[:t+1])*hold+sum(fee[:t+1])*hold-sum(q[42][:t+1])+sum(q[41][:t+1])
  q[57][t]=(land+shared*net*sum(formation[:t+1]))*share+(sum(ci[:t+1])+sum(fee[:t+1]))*share-sum(q[45][:t+1])
  q[58][t]=p(65)*(p(139) if stress else 1) if ha>0 and t==end else 0;q[59][t]=q[58][t]*(1-p(66))
  q[60][t]=q[59][t]-q[55][t] if ha>0 and t==end else 0;q[61][t]=q[59][t]-q[56][t] if ha>0 and t==end else 0
 # Annual loss stock is kept by vintage and used FIFO; the terminal variant only changes N.
 annual={}
 for levered in [False,True]:
  opRow=47 if levered else 46;gainRow=61 if levered else 60;stock={};tax=[];taxexit=[];details=[]
  for y in range(N+1):
   profit=sum(q[opRow][t] for t in T if (0 if t==0 else (t-1)//4+1)==y)
   gain=sum(q[gainRow][t] for t in T if (0 if t==0 else (t-1)//4+1)==y)
   stock={v:x for v,x in stock.items() if p(101)==0 or y-v<=p(101)} if p(100)==1 else {}
   available=sum(stock.values());used=min(max(0,profit),available)
   tax.append(max(0,profit-used)*p(52));taxexit.append(max(0,profit+gain-min(max(0,profit+gain),available))*p(52))
   remaining=used
   for v in sorted(stock):x=min(remaining,stock[v]);stock[v]-=x;remaining-=x
   if profit<0 and p(100)==1:stock[y]=-profit
   details.append((profit,available,used,sum(stock.values())))
  annual[levered]={'tax':tax,'exit':taxexit,'details':details}
 lag=int(p(112))
 for t in T:
  def due(r):return (q[r][t-lag] if t>=lag else 0)+(sum(q[r][j] for j in T if j+lag>end) if t==end else 0)
  q[53][t]=due(51);q[54][t]=due(36)+due(40)
  for levered,ex,row in [(False,False,62),(False,True,63),(True,False,64),(True,True,65)]:
   taxes=annual[levered]['exit' if ex else 'tax'];q[row][t]=sum(taxes[y] for y in range(N+1) if y*4+p(111)==t)+ (sum(taxes[y] for y in range(N+1) if y*4+p(111)>end) if t==end else 0)
  q[66][t]=q[47][t]-(annual[True]['tax'][t//4] if t>0 and t%4==0 else 0)
  q[67][t]=q[47][t]+q[61][t]-(annual[True]['exit'][t//4] if t>0 and t%4==0 else 0)
  q[68][t]=q[12][t]+q[26][t]+q[28][t];q[69][t]=q[35][t]+q[37][t]+q[38][t]+q[39][t]+q[53][t]+q[54][t]
  q[70][t]=q[33][t]-q[24][t]-q[37][t]-q[38][t]-q[39][t]-q[40][t]
 paths={}
 for ex in [False,True]:
  cf=[];pre=[]
  for t in T:
   recovery=wc*p(110 if ex else 109) if t==end else 0;disposal=q[59][t] if ex else 0
   cf.append(q[68][t]-q[69][t]-q[63 if ex else 62][t]-invest[t]+recovery+disposal)
   pre.append(q[68][t]-q[69][t]-q[65 if ex else 64][t]-invest[t]+recovery+disposal+draw[t]-interest[t]-principal[t]-fee[t])
  reserve=[0]*(end+1)
  for t in range(end-1,-1,-1):reserve[t]=max(0,reserve[t+1]-pre[t+1])
  capital=[];dist=[];cash=[]
  for t in T:
   opening=cash[-1] if t else 0;minimum=0 if t==0 or (ex and t==end) else p(123)
   eq=max(0,minimum-opening-pre[t]);distribution=max(0,opening+pre[t]+eq-reserve[t]-(0 if ex and t==end else p(123)))
   capital.append(eq);dist.append(distribution);cash.append(opening+pre[t]+eq-distribution)
  paths[ex]={'cf':cf,'pre':pre,'reserve':reserve,'capital':capital,'dist':dist,'cash':cash,'eqcf':[a-b for a,b in zip(dist,capital)]}
 path=paths[bool(p(64))];gap=[];capcash=[];capeq=[];capdist=[];cum=0;priorgap=0;hist=0
 for t in T:
  op=capcash[-1] if t else 0;minimum=p(123) if t>0 and (p(64)==0 or t<end) else 0;need=max(0,-(op+path['pre'][t]-priorgap)+minimum)
  eq=min(need,max(0,p(61)-cum));cum+=eq;netcash=op+path['pre'][t]-priorgap+eq;g=max(0,-netcash);hist=max(hist,g)
  dist=0 if hist>.01 else max(0,netcash-path['reserve'][t]-(p(123) if p(64)==0 or t<end else 0))
  capcash.append(max(0,netcash)-dist);gap.append(g);capeq.append(eq);capdist.append(dist);priorgap=g
 cfads=[q[68][t]-q[69][t]-q[64][t] for t in T]
 dscr=[]
 for y in range(1,N+1):
  tids=range((y-1)*4+1,y*4+1);service=sum(principal[t]+interest[t] for t in tids)
  if y*4>completion and service>1e-6:dscr.append((y,sum(cfads[t] for t in tids)/service))
 return {'op':q,'paths':paths,'capital':path['capital'],'eqcf':path['eqcf'],'cf':path['cf'],'gap':gap,'capcash':capcash,'capeq':capeq,'capdist':capdist,'draw':draw,'principal':principal,'interest':interest,'closing':closing,'ci':ci,'annualtax':annual,'minimum_dscr':min((x[1] for x in dscr),default=None),'investment':land+nonland+sum(ci)+sum(fee)+wc,'firstdraw':first,'maturity':maturity}

def compare_native(path):
 m=json.loads((ROOT/'manifest.json').read_text(encoding='utf-8'));native=json.loads(Path(path).read_text(encoding='utf-8-sig'))
 params={x['row']:x['value'] for x in m['source']['parameters'] if not x['formula']}
 params.update({x[0]:x[2] for x in m['added']})
 def actual(sheet,row,t):
  d=native['sheets'][sheet];return d['values'][row-d['row']][t+4-d['col']]
 differences=[];count=0;metrics=[]
 for stress in [False,True]:
  r=calculate(params,m['source']['build'],m['source']['sale'],stress);offset=110 if stress else 0
  for row,vals in r['op'].items():
   if row not in [10,11,12,13,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70]:continue
   for t,val in enumerate(vals):
    got=actual('季度经营税费',row+offset,t);count+=1
    if not isinstance(got,(int,float)) or not math.isclose(val,got,abs_tol=1e-6,rel_tol=1e-9):differences.append(('季度经营税费',row+offset,t,val,got))
  for row,key in [(45,'cf'),(47,'eqcf'),(49,'capital'),(58,'gap'),(57,'capcash'),(55,'capeq'),(56,'capdist')]:
   for t,val in enumerate(r[key]):
    got=actual('季度现金流',row+offset,t);count+=1
    if not isinstance(got,(int,float)) or not math.isclose(val,got,abs_tol=1e-6,rel_tol=1e-9):differences.append(('季度现金流',row+offset,t,val,got))
  metrics.append({'stress':stress,'required_equity':sum(r['capital']),'gap':max(r['gap']),'min_dscr':r['minimum_dscr'],'investment':r['investment']})
 report={'checked_cells':count,'difference_count':len(differences),'differences':differences[:30],'metrics':metrics}
 (ROOT/'independent_validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':
 import sys
 compare_native(sys.argv[1])
