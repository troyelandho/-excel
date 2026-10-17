import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
process.on('uncaughtException',e=>{console.error('BUILD ERROR: '+e.message);process.exit(1);});

// One formula builder owns both of the visible scenarios inherited from V2.
const here=new URL('.',import.meta.url), out=new URL('../',here);
const source=JSON.parse(await fs.readFile(new URL('source.json',here),'utf8'));
if(process.argv.includes('--fix')){
 const file=decodeURIComponent(new URL('候选_V3.xlsx',here).pathname.slice(1));
 const w=await SpreadsheetFile.importXlsx(await FileBlob.load(file));
 const fixes=JSON.parse(await fs.readFile(new URL('formula_repairs.json',here),'utf8'));
 for(const x of fixes){const rg=w.worksheets.getItem(x.sheet).getRange(x.cell);if(x.formula)rg.formulas=[[x.formula]];if(x.numberFormat)rg.setNumberFormat(x.numberFormat);if(x.format)rg.format=x.format;}
 w.recalculate();const blob=await SpreadsheetFile.exportXlsx(w);await blob.save(file);console.log('targeted fixes exported');process.exit(0);
}
if(process.argv.includes('--patch')){
 const file=decodeURIComponent(new URL('候选_V3.xlsx',here).pathname.slice(1));
 const w=await SpreadsheetFile.importXlsx(await FileBlob.load(file));
 const meta=JSON.parse(await fs.readFile(new URL('manifest.json',here),'utf8'));
 const inp=w.worksheets.getItem('核心参数输入');
 const old=inp.getRange('B150').formulas[0][0].slice(1);
 const rates=[12,14,27,32,33,38,39,40,41,43,44,49,51,52,54,58,66,69,70,71,72,78,79,102,109,110,113,118,119,120,121,122,125,128,129,130,131,137];
 const lags=[103,104,105,106,107,111,112,124,126,148];
 const guards=[...rates.map(r=>`AND(B${r}>=0,B${r}<=1)`),...lags.map(r=>`AND(B${r}>=0,B${r}=INT(B${r}))`),'B8=INT(B8)','B8>=1900','B8<=2200','B37>=-0.99','B37<=1','B48>=-0.99','B48<=1','B81>=0','B81<=10','B144>=0','B145>=0','B146>=0','B147>=0'];
 inp.getRange('B150').formulas=[[`=IF(AND(${old}=1,${guards.join(',')}),1,0)`]];
 const ss=w.worksheets.getItem('敏感性与边界');
 const sig=ss.getRange('B68').formulas[0][0];ss.getRange('B68').formulas=[[sig+'&"|"&'+[6,7,8,9,10].map(r=>`IF(B${r}="","_",B${r})`).join('&"|"&')]];
 // A compact monthly rephasing check for a selected critical quarter.
 const s=w.worksheets.add('关键月份核对');meta.sheets.push('关键月份核对');meta.monthlyInputs=['B5:B6','B11:D16'];
 s.showGridLines=false;s.getRange('A1:G44').format={font:{name:'Microsoft YaHei',size:10,color:'#334155'},rowHeight:23,verticalAlignment:'center'};
 s.getRange('A1:A44').format.columnWidth=43;s.getRange('B1:D44').format.columnWidth=20;s.getRange('E1:G44').format.columnWidth=20;
 const V=(ad,x)=>s.getRange(ad).values=[[x]],F=(ad,x)=>s.getRange(ad).formulas=[['='+x]],P=r=>`'核心参数输入'!$B$${r}`;
 V('A2','关键月份核对 · 所选季度资金先后');s.getRange('A2:G2').format={font:{size:16,bold:true},rowHeight:30};
 V('A5','项目季度覆盖（留空自动取资金高峰）');V('A6','核对情景：1基准、2组合压力');V('B6',1);
 V('A7','当前核对项目季度');F('B7',`IF(ISNUMBER(B5),B5,MATCH(MAX(INDEX('季度现金流'!$D$1:$DT$200,77+110*(B6-1),0)),INDEX('季度现金流'!$D$1:$DT$200,77+110*(B6-1),0),0)-1)`);
 V('A8','季度名称');F('B8',"INDEX('季度计划'!$D$5:$DT$5,1,B7+1)");
 s.getRange('A10:G10').values=[['时点权重／现金项目','第1月','第2月','第3月','季度合计／末值','季度模型值','勾稽差额']];s.getRange('A10:G10').format={fill:'#D9E2F3',font:{bold:true},wrapText:true,rowHeight:32};
 for(const [r,label,weights]of [[11,'经营及退出现金流入',[1/3,1/3,1/3]],[12,'经营成本、税及投资支付',[1/3,1/3,1/3]],[13,'贷款到账',[1/3,1/3,1/3]],[14,'股东出资',[1,0,0]],[15,'还本付息及融资费用',[0,0,1]],[16,'股东现金分配',[0,0,1]]]){V(`A${r}`,label);s.getRange(`B${r}:D${r}`).values=[weights];F(`E${r}`,`SUM(B${r}:D${r})`);}
 for(const rg of ['B5:B6','B11:D16'])s.getRange(rg).format={fill:'#FFF9C4',font:{color:'#1256A0'}};
 s.getRange('B11:E16').setNumberFormat('0.00%');s.getRange('B20:G34').setNumberFormat('#,##0.00;[Red](#,##0.00);"—"');
 s.dataValidations.add({range:'B11:D16',rule:{type:'decimal',operator:'between',formula1:0,formula2:1}});s.dataValidations.add({range:'B5',rule:{type:'whole',operator:'between',formula1:1,formula2:120}});s.dataValidations.add({range:'B6',rule:{type:'whole',operator:'between',formula1:1,formula2:2}});
 const get=(sheet,row,prior=false)=>`INDEX('${sheet}'!$D$1:$DT$200,${row}+110*($B$6-1),MAX(1,MIN(121,$B$7+${prior?0:1})))`;
 const endcash=get('季度现金流',51),begin=get('季度现金流',51,true);
 const monthly={21:[`${get('季度经营税费',68)}+IF(${P(64)}=1,${get('季度经营税费',59)}+${get('季度现金流',17)},${get('季度现金流',16)})`,11],22:[`${get('季度经营税费',69)}+IF(${P(64)}=1,${get('季度经营税费',65)},${get('季度经营税费',64)})`,12],23:[get('季度投融资',51),12],24:[get('季度投融资',39),13],25:[get('季度现金流',49),14],26:[`${get('季度投融资',46)}+${get('季度投融资',49)}`,15],27:[get('季度现金流',50),16]};
 for(const [r,label]of [[20,'期初可用留存现金'],[21,'经营、退出及周转现金流入'],[22,'经营成本及税费现金'],[23,'土地、建设及周转现金投入'],[24,'合同贷款提款到账'],[25,'季度计划股东出资'],[26,'合同还本付息及融资费用'],[27,'季度计划股东分配'],[28,'月末现金余额（未补口）'],[29,'月内未覆盖资金缺口']])V(`A${r}`,label);
 for(const [c,prev]of [['B',null],['C','B'],['D','C']]){
  F(`${c}20`,prev?`${prev}28`:`IF($B$7=0,0,${begin})`);
  for(const [r,[formula,weight]]of Object.entries(monthly))F(`${c}${r}`,`(${formula})*${c}$${weight}`);
  F(`${c}28`,`${c}20+${c}21-${c}22-${c}23+${c}24+${c}25-${c}26-${c}27`);F(`${c}29`,`MAX(0,-${c}28)`);
 }
 for(const [r,[formula]]of Object.entries(monthly)){F(`E${r}`,`SUM(B${r}:D${r})`);F(`F${r}`,formula);F(`G${r}`,`E${r}-F${r}`);}
 F('E28','D28');F('F28',endcash);F('G28','E28-F28');V('A33','所选季度月内最大缺口');F('B33','MAX(B29:D29)');V('A34','月度权重及季度勾稽有效');F('B34',`IF(AND(B7>=1,B7<=${P(9)}*4,MIN(B11:D16)>=0,MAX(B11:D16)<=1,ABS(E11-1)<0.000001,ABS(E12-1)<0.000001,ABS(E13-1)<0.000001,ABS(E14-1)<0.000001,ABS(E15-1)<0.000001,ABS(E16-1)<0.000001,ABS(G28)<0.01),1,0)`);V('A35','所选季度月度资金安排通过');F('B35','IF(AND(B34=1,B33<=0.01),1,0)');
 s.getRange('A38:G40').merge();V('A38','本表仅重排所选季度的到账、支付和出资先后，不自动增加资本金或改变合同计划。利息沿用季度已测金额；月内缺口需通过更早到账或提前出资落实。只检验所选季度，不代表全期月度最大缺口；精确边界仍按季度约束求解。');s.getRange('A38:G40').format={wrapText:true,fill:'#F7F9FC'};s.freezePanes.freezeRows(10);
 s.getRange('B29:D29').conditionalFormats.add('cellIs',{operator:'greaterThan',formula:0.01,format:{fill:'#FBE0E0',font:{color:'#A72C2C'}}});
 const d=w.worksheets.getItem('决策看板');d.getRange('A41:G42').merge();d.getRange('A41').formulas=[['="所选季度月内最大缺口（万元）："&TEXT(\'关键月份核对\'!B33,"#,##0.00")&IF(\'关键月份核对\'!B35=1,"；该季度月度安排通过","；月度安排需调整")']];d.getRange('A41:G42').format={wrapText:true,fill:'#F7F9FC'};
 // Targeted repairs discovered in QA are applied through the same authoring API.
 if(await fs.stat(new URL('formula_repairs.json',here)).then(()=>true).catch(()=>false)){
  const fixes=JSON.parse(await fs.readFile(new URL('formula_repairs.json',here),'utf8'));
  for(const x of fixes){const rg=w.worksheets.getItem(x.sheet).getRange(x.cell);if(x.formula)rg.formulas=[[x.formula]];if(x.numberFormat)rg.setNumberFormat(x.numberFormat);if(x.format)rg.format=x.format;}
 }
 w.recalculate();const blob=await SpreadsheetFile.exportXlsx(w);await blob.save(file);await fs.writeFile(new URL('manifest.json',here),JSON.stringify(meta,null,2));console.log('targeted validation and monthly check added');process.exit(0);
}
const wb=Workbook.create();
const names=['决策看板','核心参数输入','指标与情景','年度汇总','季度投融资','季度经营税费','季度现金流','年度所得税','季度计划','敏感性与边界','模型校验','使用说明'];
const sh=Object.fromEntries(names.map(n=>[n,wb.worksheets.add(n)]));
const C={ink:'#334155',title:'#1F2937',blue:'#4472C4',navy:'#1F4E78',head:'#D9E2F3',pale:'#F7F9FC',key:'#EAF2FF',input:'#FFF9C4',inputInk:'#1256A0',link:'#237047',red:'#A72C2C',bad:'#FBE0E0',rule:'#D9DEE7'};
const NUM='#,##0.00;[Red](#,##0.00);"—"', PCT='0.00%;[Red](0.00%);"—"';
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const END=col(124), O=110, Q=120;
const ref=(s,a)=>`'${s}'!${a}`, p=r=>ref('核心参数输入',`$B$${r}`);
const a=(s,r,t,o=0)=>ref(s,`${col(4+t)}${r+o}`);
const rr=(s,r,o=0)=>ref(s,`$D$${r+o}:$${END}$${r+o}`);
const active=(s,r,o=0,start=0)=>`${ref(s,`$${col(4+start)}$${r+o}`)}:INDEX(${rr(s,r,o)},1,${p(9)}*4+1)`;
const inv=(r,k=0)=>ref('季度投融资',`$${k?'E':'B'}$${r}`);
const v=(s,ad,x)=>sh[s].getRange(ad).values=[[x]];
const f=(s,ad,x)=>sh[s].getRange(ad).formulas=[[x.startsWith('=')?x:'='+x]];
const rows=(s,r,fn,n=121)=>sh[s].getRange(`D${r}:${col(n+3)}${r}`).formulas=[Array.from({length:n},(_,i)=>'='+fn(i,col(4+i)))];
const valueRow=(s,r,x)=>sh[s].getRange(`D${r}:${col(x.length+3)}${r}`).values=[x];
const label=(s,r,text,unit='万元')=>{v(s,`A${r}`,text);v(s,`B${r}`,unit);};
function section(s,r,text,last='L'){
 const rg=sh[s].getRange(`A${r}:${last}${r}`);rg.format={fill:C.blue,font:{bold:true,color:'#FFFFFF'},rowHeight:24};v(s,`A${r}`,text);
}
function setup(s,last='L',height=180){
 const w=sh[s];w.showGridLines=false;
 w.getRange(`A1:${last}${height}`).format={font:{name:'Microsoft YaHei',size:10,color:C.ink},rowHeight:21,verticalAlignment:'center'};
 w.getRange(`A1:A${height}`).format.columnWidth=43;
 w.getRange(`B1:B${height}`).format.columnWidth=18;
 w.getRange(`C1:C${height}`).format.columnWidth=13;
 w.getRange(`D1:${last}${height}`).format.columnWidth=16;
 w.getRange(`D1:${last}${height}`).setNumberFormat(NUM);
 v(s,'A2',`${s} · V3模拟测算`);w.getRange('A2:L2').format={font:{bold:true,size:16,color:C.title},rowHeight:30};
 w.freezePanes.freezeRows(6);w.freezePanes.freezeColumns(3);
}
function headerQ(s){
 rows(s,4,t=>s==='季度计划'?String(t):a('季度计划',4,t));
 rows(s,5,(t,c)=>s==='季度计划'?`IF(${c}4=0,TEXT(${p(8)},"0")&"年末 / 第0年",TEXT(${p(8)}+INT((${c}4-1)/4)+1,"0")&"Q"&TEXT(MOD(${c}4-1,4)+1,"0"))`:a('季度计划',5,t));
 rows(s,6,(t,c)=>s==='季度计划'?`IF(${c}4=0,0,INT((${c}4-1)/4)+1)`:a('季度计划',6,t));
 sh[s].getRange(`D4:${END}6`).format={fill:C.head,font:{bold:true},horizontalAlignment:'center'};
 sh[s].getRange(`D4:${END}4`).setNumberFormat('0');sh[s].getRange(`D6:${END}6`).setNumberFormat('0');
 v(s,'A4','季度索引');v(s,'A5','日历期间');v(s,'A6','项目年度');
}
function headerY(s){valueRow(s,4,Array.from({length:31},(_,t)=>t));rows(s,5,(t,c)=>`${p(8)}+${c}$4`,31);sh[s].getRange('D4:AH5').format={fill:C.head,font:{bold:true},horizontalAlignment:'center'};sh[s].getRange('D4:AH5').setNumberFormat('0');v(s,'A4','项目年度');v(s,'A5','日历年份');}
function input(s,ad,unit=''){sh[s].getRange(ad).format={fill:C.input,font:{color:C.inputInk},numberFormat:unit==='%'?PCT:unit==='文本'?'@':NUM};}
const controls=[];
function validate(s,range,type,min,max){sh[s].dataValidations.add({range,rule:{type,operator:'between',formula1:min,formula2:max}});}
for(const s of names)setup(s,['季度投融资','季度经营税费','季度现金流','季度计划'].includes(s)?END:['年度汇总','年度所得税'].includes(s)?'AH':'L',s==='年度所得税'?300:['季度投融资','季度经营税费','季度现金流'].includes(s)?205:180);
console.log('sheets and styles ready');
sh['决策看板'].tabColor=C.navy;sh['核心参数输入'].tabColor=C.blue;sh['使用说明'].tabColor='#94A3B8';

// Preserve the actual saved V2 inputs, not generator defaults.
v('核心参数输入','A5','参数');v('核心参数输入','B5','输入值');v('核心参数输入','C5','单位');v('核心参数输入','D5','口径、来源与状态');
sh['核心参数输入'].getRange('D1:K160').format.columnWidth=13;
for(const x of source.parameters){
 v('核心参数输入',`A${x.row}`,x.name);v('核心参数输入',`C${x.row}`,x.unit);
 if(!x.formula){v('核心参数输入',`B${x.row}`,x.value);input('核心参数输入',`B${x.row}`,x.unit);controls.push(`B${x.row}`);}
 sh['核心参数输入'].mergeCells(`D${x.row}:K${x.row}`);v('核心参数输入',`D${x.row}`,x.note);sh['核心参数输入'].getRange(`D${x.row}:K${x.row}`).format={font:{color:'#64748B'},wrapText:true};
 v('核心参数输入',`L${x.row}`,'V2保存参数');
}
for(const [r,n] of [[6,'基础与建筑参数'],[20,'投资参数'],[30,'销售参数'],[35,'出租参数'],[46,'运营与税费'],[56,'融资与资本金'],[63,'期末处置'],[68,'模拟判断标准'],[77,'组合压力']])section('核心参数输入',r,n);
valueRow('核心参数输入',87,Array.from({length:31},(_,i)=>i));valueRow('核心参数输入',88,source.build);valueRow('核心参数输入',89,source.sale);v('核心参数输入','A88','年度建设分配比例');v('核心参数输入','A89','年度签约销售比例');input('核心参数输入','D88:AH89','%');
controls.push('D88:AH89');validate('核心参数输入','D88:AH89','decimal',0,1);
section('核心参数输入',93,'V3补充假设（黄色输入；默认均为示意、待确认）');
const added=[
 [95,'经营模式',1,'编号','1租售结合，2纯出售，3纯持有；混合模式读取原B14'],
 [96,'流转税计算口径',1,'编号','1沿用V2综合率；2分离增值税，必须填B118:B122财务参数'],
 [97,'新增假设已完成项目核对',0,'0/1','示意默认0；只有逐项核对后改1，仍不代表项目值得投资'],
 [98,'持有土地是否摊销',1,'0/1','V3示意开启；仅财务假设，非税务适用结论'],
 [99,'持有土地摊销年限',50,'年','示意新增；从建成后开始，零残值；须核对剩余年限'],
 [100,'是否结转所得税亏损',1,'0/1','按年度、分年份、先发生先使用；不跨项目共享'],
 [101,'亏损可结转年限',5,'年','示意新增，非实际适用规则；0表示无到期，最大30年'],
 [102,'销售交付当季收款比例',0.7,'%','其余按B103延迟；收入按交付确认，未收保留应收'],
 [103,'销售尾款延迟季度数',1,'季度','示意新增；尾款超出测算期不自动收回，并提示应收余额'],
 [104,'租金收款延迟季度数',0,'季度','收缴率视为最终可收回比例，其余记坏账；收款另按此延迟'],
 [105,'建设付款相对工程形成滞后',0,'季度','可在季度计划逐季覆盖；形成成本与付款现金分开'],
 [106,'贷款提款相对建设付款滞后',0,'季度','既定额度；滞后会增加股东垫资，不自动新增贷款'],
 [107,'整体开发延期季度数',0,'季度','土地已于第0年投入；非土地开工及后续租售整体后移，建设跨度不变'],
 [108,'固定及维护成本对应持有面积',60750,'平方米','V2持有面积基准；持有规模变化时按面积比例调整成本'],
 [109,'持续持有期末周转回收比例',0,'%','默认不因观察期结束回收周转资金'],
 [110,'出售退出期末周转回收比例',1,'%','示意全额；仅持有物业退出，不将未售存货或应收自动变现'],
 [111,'所得税现金支付滞后季度数',0,'季度','年度计算、年末支付；期末未支付税款按终值负债结算扣除'],
 [112,'流转税现金支付滞后季度数',0,'季度','含增值税及其他流转税；期末未付款作为负债结算扣除'],
 [113,'贷款一次性融资费用率',0,'%','首次提款支付并资本化，作为融资现金流出，不进入项目投资现金流'],
 [114,'拿地附加现金成本',0,'万元','示意为零，须核对包含范围；归入土地成本且第0年支付'],
 [115,'持有部分专项装修配套投资',0,'万元','示意为零，全部归持有部分；按建设比例形成及支付，不与共同成本重复'],
 [116,'持有重大更新发生项目年',10,'年','仅当B117非零生效；年末一次发生，现金费用化，财务假设待确认'],
 [117,'持有重大更新支出',0,'万元','不随面积自动缩放，已含常年维护的部分不得重复填写'],
 [118,'销售收入增值税率',0,'%','仅口径2生效；零为待核对示意参数，不代表实际税率'],
 [119,'出租及其他收入增值税率',0,'%','仅口径2生效；按含税收入分离销项税'],
 [120,'建设支付可抵扣进项税额比例',0,'%','仅口径2生效；以含税支付额为分母，同时扣减资本成本；余额不自动退款'],
 [121,'销售其他流转税费率',0,'%','仅口径2生效，按含税销售金额；口径1仍用原B33综合率'],
 [122,'运营其他流转税费率',0,'%','仅口径2生效，按含税租赁及其他收入；口径1仍用原B51'],
 [123,'最低安全现金余额',0,'万元','需要时填写；自首个正季度维持，期末出售释放'],
 [124,'压力情景额外开发延期',0,'季度','除原四项压力外可再增加延期；默认0保留原压力设定'],
 [125,'新增租户季内入驻位置',0.5,'比例','0季初、0.5季中、1季末；出租率输入改为各出租年末目标'],
 [126,'销售签约至交付延迟',0,'季度','交付不得早于建成；收入、销售成本及首款以交付为起点'],
 [127,'合同还本间隔季度数',4,'季度','4每年、1每季；从原开始还本年度的年末开始；到期结清计划本金'],
 [128,'年度计划第1季度权重',0.25,'%','四季权重合计100%；用于建设、销售默认分解，可逐季覆盖'],
 [129,'年度计划第2季度权重',0.25,'%','V3示意均匀分配，非实际合同付款安排'],
 [130,'年度计划第3季度权重',0.25,'%','用于原年度计划分解'],
 [131,'年度计划第4季度权重',0.25,'%','用于原年度计划分解'],
 [132,'地价边界搜索上限',3200,'元/平方米','财务测试范围，非市场区间；超范围只报告边界未找到'],
 [133,'售价边界搜索上限',26000,'元/平方米','同一基准模型反复求解；修改参数后运行刷新工具'],
 [134,'月租金边界搜索上限',208,'元/平方米/月','财务测试范围，非市场区间'],
 [135,'建设成本倍数测试上限',2,'倍','以V2非土地共同建设及专项装修为基数'],
 [136,'开发延期测试上限',8,'季度','用于延期单因素边界，搜索不超过此范围'],
 [137,'出租率测试上限',1,'%','用于稳定出租率财务边界，不可超过100%'],
 [139,'压力情景期末售价系数',1,'倍','默认1沿用V2处置价；需要时单独压低退出价格']
];
for(const [r,n,x,u,note] of added){v('核心参数输入',`A${r}`,n);v('核心参数输入',`B${r}`,x);v('核心参数输入',`C${r}`,u);input('核心参数输入',`B${r}`,u);sh['核心参数输入'].mergeCells(`D${r}:K${r}`);v('核心参数输入',`D${r}`,note);sh['核心参数输入'].getRange(`D${r}:K${r}`).format={wrapText:true,font:{color:'#64748B'}};v('核心参数输入',`L${r}`,'V3示意新增，待核对');controls.push(`B${r}`);}
section('核心参数输入',141,'当前生效口径（公式）');
const activeRows={142:['实际销售比例',`IF(${p(95)}=2,1,IF(${p(95)}=3,0,${p(14)}))`],143:['实际持有比例',`1-${p(142)}`],144:['生效土地单价',`IF(ISNUMBER('敏感性与边界'!$B$6),'敏感性与边界'!$B$6,${p(21)})`],145:['生效销售均价',`IF(ISNUMBER('敏感性与边界'!$B$7),'敏感性与边界'!$B$7,${p(31)})`],146:['生效初始月租金',`IF(ISNUMBER('敏感性与边界'!$B$8),'敏感性与边界'!$B$8,${p(36)})`],147:['生效非土地成本倍数',`IF(ISNUMBER('敏感性与边界'!$B$9),'敏感性与边界'!$B$9,1)`],148:['生效整体延期季度',`IF(ISNUMBER('敏感性与边界'!$B$10),'敏感性与边界'!$B$10,${p(107)})`],149:['经营模式名称',`CHOOSE(${p(95)},"租售结合","纯出售","纯持有")`],152:['税费口径',`IF(${p(96)}=1,"综合税费率（沿用V2）","分离增值税（财务参数）")`]};
for(const [r,[n,x]]of Object.entries(activeRows)){v('核心参数输入',`A${r}`,n);f('核心参数输入',`B${r}`,x);}
v('核心参数输入','A15','实际出租占比（模式后）');f('核心参数输入','B15',p(143));
for(const r of [14,142,143])sh['核心参数输入'].getRange(`B${r}`).setNumberFormat(PCT);
const noteOverrides={14:'混合模式使用；纯出售／纯持有由B95切换，原55%示意值保留',38:'第1个出租年度年末目标；V3按季内入驻位置形成平均面积',39:'第2个出租年度年末目标',40:'第3个出租年度年末目标',41:'第4出租年末及以后目标；不模拟租户更替',50:'常年维护按基准持有面积缩放；重大更新另列B117',52:'年度计税，按B100:B101结转亏损；与现金支付时间分开',57:'既定贷款额度；季度提款可覆盖，提款滞后不自动增贷',60:'原开始年度年末首笔还本；整体延期时配套后移',64:'当前展示口径开关；持有与出售结果始终并列计算',65:'仅持有物业的第N年含税售价；变更期限应重估',66:'综合退出费用已包含交易税费，退出不再重复计销项增值税'};
for(const [r,n]of Object.entries(noteOverrides))v('核心参数输入',`D${r}`,n);
const coreNumeric=source.parameters.filter(x=>typeof x.value==='number'&&!x.formula).map(x=>`B${x.row}`);
const rateRows=[12,14,27,32,33,38,39,40,41,43,44,49,51,52,54,58,66,69,70,71,72,78,79,102,109,110,113,118,119,120,121,122,125,128,129,130,131,137];
for(const r of rateRows)validate('核心参数输入',`B${r}`,'decimal',0,1);
for(const r of [9,16,17,18,53,59,60,99])validate('核心参数输入',`B${r}`,'whole',1,r===53||r===99?100:30);
for(const r of [103,104,105,106,107,111,112,124,126,136])validate('核心参数输入',`B${r}`,'whole',0,120);
for(const r of [64,82,97,98,100])validate('核心参数输入',`B${r}`,'whole',0,1);
validate('核心参数输入','B95','whole',1,3);validate('核心参数输入','B96','whole',1,2);validate('核心参数输入','B101','whole',0,30);validate('核心参数输入','B127','whole',1,4);
for(const r of [37,48])validate('核心参数输入',`B${r}`,'decimal',-0.99,1);
validate('核心参数输入','B42','decimal',0,12);validate('核心参数输入','B80','decimal',0,100);validate('核心参数输入','B81','decimal',0,10);
const requiredAdded=added.map(x=>`B${x[0]}`);
f('核心参数输入','B150',`IF(AND(COUNT(${coreNumeric.join(',')})=${coreNumeric.length},COUNT(${requiredAdded.join(',')})=${added.length},B7<>"",B9=INT(B9),B9>B16,B9<=30,B16>=1,B16=INT(B16),B10>0,B11>0,B12>0,B12<=1,B13>=0,B14>=0,B14<=1,B95>=1,B95<=3,B95=INT(B95),OR(B96=1,B96=2),MIN(B21:B26,B28,B31,B36,B47,B50,B57,B61,B65,B108,B114:B115,B117,B123)>=0,MIN(B38:B41)>=0,MAX(B38:B41)<=1,B38<=B39,B39<=B40,B40<=B41,B42>=0,B42<=12,B43>=0,B43<=1,B37>-1,B48>-1,B53>=1,B53=INT(B53),B99>=1,B99=INT(B99),B101>=0,B101=INT(B101),B69>=0,B70>=0,B71>=0,B72>=0,B73>0,B74>0,B75>0,SUM(B128:B131)>0,ABS(SUM(B128:B131)-1)<0.000001,MIN(B128:B131)>=0,MIN(D88:AH89)>=0,MAX(D88:AH89)<=1,COUNT(D88:AH88)=31,COUNT(D89:AH89)=31,ABS(SUM(D88:AH88)-1)<0.000001,SUM(D89:AH89)<=1.000001,OR(B142=0,AND(B17>B16,B17<=B9,B17=INT(B17))),OR(B143=0,AND(B18>B16,B18<=B9,B18=INT(B18),B108>0)),B102>=0,B102<=1,B105>=0,B105=INT(B105),B106>=0,B106=INT(B106),B148>=0,B148=INT(B148),B127>=1,B127=INT(B127),MIN(B118:B122)>=0,MAX(B118:B122)<=1,OR(B98=0,B98=1),OR(B100=0,B100=1),OR(B64=0,B64=1),OR(B82=0,B82=1)),1,0)`);
v('核心参数输入','A150','输入基础有效标志');
console.log('source and added assumptions written');

// Quarter plans: blank overrides retain defaults; an explicit zero is respected.
for(const s of ['季度计划','季度投融资','季度经营税费','季度现金流'])headerQ(s);
for(const [r,n]of [[20,'工程形成比例覆盖'],[21,'现金付款比例覆盖'],[22,'贷款提款权重覆盖'],[23,'销售签约比例覆盖']]){label('季度计划',r,n,'%');input('季度计划',`D${r}:${END}${r}`,'%');validate('季度计划',`D${r}:${END}${r}`,'decimal',0,1);}
const delayed=(r,t,lag,o=0,s='季度计划')=>`IF(AND(${t}-${lag}>=0,${t}-${lag}<=120),INDEX(${rr(s,r,o)},1,MAX(1,MIN(121,${t}-${lag}+1))),0)`;
for(const k of [0,1]){
 const o=k*O, delay=`(${p(148)}${k?'+'+p(124):''})`;
 section('季度计划',8+o,k?'组合压力 · 季度计划':'基准 · 季度计划',END);
 for(const [r,n]of [[10,'默认建设形成比例'],[11,'生效工程形成比例'],[12,'生效现金付款比例'],[13,'生效贷款提款权重'],[14,'销售签约比例'],[15,'销售交付比例'],[16,'出租经营季度年龄'],[17,'建设完成季度']])label('季度计划',r+o,n,r<16?'%':'季度');
 for(const [r,fn]of [
 [10,(t,c)=>t===0?'0':`IF(AND(${t}-${delay}>0,${t}-${delay}<=${p(16)}*4),INDEX('核心参数输入'!$D$88:$AH$88,1,INT((${t}-${delay}-1)/4)+2)*INDEX('核心参数输入'!$B$128:$B$131,MOD(${t}-${delay}-1,4)+1),0)`],
 [11,(t,c)=>t===0?'0':`IF(AND(${t}-${delay}>=0,${t}-${delay}<=120),IF(COUNT(INDEX($D$20:$${END}$20,1,MAX(1,MIN(121,${t}-${delay}+1))))=1,INDEX($D$20:$${END}$20,1,MAX(1,MIN(121,${t}-${delay}+1))),${c}${10+o}),0)`],
 [12,(t,c)=>t===0?'0':`IF(AND(${t}-${delay}>=0,${t}-${delay}<=120),IF(COUNT(INDEX($D$21:$${END}$21,1,MAX(1,MIN(121,${t}-${delay}+1))))=1,INDEX($D$21:$${END}$21,1,MAX(1,MIN(121,${t}-${delay}+1))),${delayed(11,t,p(105),o)}),0)`],
 [13,(t,c)=>t===0?'0':`IF(AND(${t}-${delay}>=0,${t}-${delay}<=120),IF(COUNT(INDEX($D$22:$${END}$22,1,MAX(1,MIN(121,${t}-${delay}+1))))=1,INDEX($D$22:$${END}$22,1,MAX(1,MIN(121,${t}-${delay}+1))),${delayed(12,t,p(106),o)}),0)`],
 [14,(t,c)=>t===0?'0':`IF(${p(142)}=0,0,IF(AND(${t}-${delay}>0,${t}-${delay}<=120),IF(COUNT(INDEX($D$23:$${END}$23,1,MAX(1,MIN(121,${t}-${delay}+1))))=1,INDEX($D$23:$${END}$23,1,MAX(1,MIN(121,${t}-${delay}+1))),INDEX('核心参数输入'!$D$89:$AH$89,1,INT((${t}-${delay}-1)/4)+2)*INDEX('核心参数输入'!$B$128:$B$131,MOD(${t}-${delay}-1,4)+1)),0))`],
 [15,t=>delayed(14,t,p(126),o)],
 [16,t=>`MAX(0,${t}-((${p(18)}-1)*4+${delay}))`],
 [17,t=>`${p(16)}*4+${delay}`]
 ])rows('季度计划',r+o,(t,c)=>`IF(${t}<=${p(9)}*4,${fn(t,c)},0)`);
 for(const r of [10,11,12,13,14,15])sh['季度计划'].getRange(`D${r+o}:${END}${r+o}`).setNumberFormat(PCT);
}
v('季度计划','A26','覆盖说明');v('季度计划','D26','黄色覆盖区为原计划时点；整体延期后统一后移。空白用默认，0表示明确不发生。');
f('季度计划','B30',`IF(AND(${[0,O].map(o=>`ABS(SUM(${active('季度计划',11,o)})-1)<0.000001,ABS(SUM(${active('季度计划',12,o)})-1)<0.000001,OR(${p(57)}=0,ABS(SUM(${active('季度计划',13,o)})-1)<0.000001),SUM(${active('季度计划',14,o)})<=1.000001,SUMPRODUCT(${rr('季度计划',15,o)},--('季度计划'!$D$4:$${END}$4<=${p(16)}*4+${p(148)}${o?'+'+p(124):''}))=0`).join(',')}),1,0)`);v('季度计划','A30','计划完整有效标志');
console.log('quarterly plans written');

// Investment and loan calculations. Capitalised interest never feeds debt principal.
section('季度投融资',7,'投资构成与面积');v('季度投融资','B6','基准');v('季度投融资','E6','组合压力');
const investmentNames={8:'地上建安',9:'地下建安',10:'共同配套',11:'前期费用',12:'建设管理及其他',13:'非土地共同成本小计',14:'预备费',15:'共同非土地成本含预备费',16:'持有专项装修配套',17:'非土地现金投资总额',18:'建设期利息',19:'融资手续费',20:'初始运营周转资金',21:'项目总投资（含融资费用）',22:'销售成本基数（项目）',23:'持有建筑成本（项目）',24:'持有土地成本',25:'销售成本基数（融资）',26:'持有建筑成本（融资）',27:'持有分摊投资含周转',28:'首次提款季度',29:'合同到期季度',30:'首笔还本季度',31:'计划还本次数',32:'地上可经营面积'};
v('季度投融资','A7','土地及拿地附加成本');for(const [r,n]of Object.entries(investmentNames))v('季度投融资',`A${r}`,n);
const heldArea=`(${p(10)}*${p(11)}*${p(12)}*${p(143)})`, saleArea=`(${p(10)}*${p(11)}*${p(12)}*${p(142)})`;
for(const k of [0,1]){
 const o=k*O, cc=k?'E':'B', factor=`${p(147)}${k?'*(1+'+p(81)+')':''}`, net=`(1-IF(${p(96)}=2,${p(120)},0))`, x=r=>`$${cc}$${r}`;
 const scalars={7:`${p(10)}*${p(144)}/10000+${p(114)}`,8:`${p(10)}*${p(11)}*${p(22)}/10000*${factor}`,9:`${p(13)}*${p(23)}/10000*${factor}`,10:`${p(24)}*${factor}`,11:`${p(25)}*${factor}`,12:`${p(26)}*${factor}`,13:`SUM(${cc}8:${cc}12)`,14:`${x(13)}*${p(27)}`,15:`SUM(${cc}13:${cc}14)`,16:`IF(${p(143)}>0,${p(115)}*${factor},0)`,17:`SUM(${cc}15:${cc}16)`,18:`SUM(${rr('季度投融资',44,o)})`,19:`${p(57)}*${p(113)}`,20:`IF(${p(143)}>0,${p(28)},0)`,21:`SUM(${cc}7,${cc}17:${cc}20)`,22:`(${x(7)}+${x(15)}*${net})*${p(142)}`,23:`(${x(15)}*${p(143)}+${x(16)})*${net}`,24:`${x(7)}*${p(143)}`,25:`${x(22)}+(${x(18)}+${x(19)})*${p(142)}`,26:`${x(23)}+(${x(18)}+${x(19)})*${p(143)}`,27:`SUM(${cc}24,${cc}26,${cc}20)`,28:`IF(${p(57)}=0,0,IFERROR(MATCH(1,${rr('季度投融资',60,o)},0)-1,0))`,29:`IF(${p(57)}=0,0,${x(28)}+${p(59)}*4-1)`,30:`${p(60)}*4+${p(148)}${k?'+'+p(124):''}`,31:`IF(${p(57)}=0,0,MAX(1,INT((${x(29)}-${x(30)})/MAX(1,${p(127)}))+1+IF(MOD(${x(29)}-${x(30)},MAX(1,${p(127)}))<>0,1,0)))`,32:`${p(10)}*${p(11)}*${p(12)}`};
 for(const [r,x]of Object.entries(scalars))f('季度投融资',`${cc}${r}`,x);
 const ls={35:'本季形成非土地投资',36:'土地现金支付',37:'建设现金付款',38:'周转资金投入',39:'本季贷款提款',40:'期初贷款本金',41:'合同计划还本',42:'本季现金利息',43:'期末计划贷款本金',44:'其中建设期利息',45:'其中经营期利息',46:'计划本息合计',47:'累计建设付款含土地',48:'投资现金含资本化融资费用',49:'本季融资手续费',50:'到期计划本金余额',51:'项目投资现金（不含融资费）',52:'累计工程形成成本含土地',60:'首次提款辅助标志'};
 section('季度投融资',34+o,k?'组合压力 · 季度投资和合同还款':'基准 · 季度投资和合同还款',END);for(const [r,n]of Object.entries(ls))label('季度投融资',+r+o,n);
 for(const [r,fn]of [
 [35,t=>`${inv(17,k)}*${a('季度计划',11,t,o)}`],[36,t=>t===0?inv(7,k):'0'],[37,t=>`${inv(17,k)}*${a('季度计划',12,t,o)}`],
 [38,t=>`IF(${t}=${p(16)}*4+${p(148)}${k?'+'+p(124):''}+1,${inv(20,k)},0)`],
 [39,t=>`${p(57)}*${a('季度计划',13,t,o)}`],[40,(t,c)=>t?a('季度投融资',43,t-1,o):'0'],
 [41,(t,c)=>`IF(AND(${p(57)}>0,${t}>=${inv(30,k)},${t}<=${inv(29,k)},OR(MOD(${t}-${inv(30,k)},MAX(1,${p(127)}))=0,${t}=${inv(29,k)})),IF(${t}=${inv(29,k)},${c}${40+o}+${c}${39+o},MIN(${c}${40+o}+${c}${39+o},${p(57)}/MAX(1,${inv(31,k)}))),0)`],
 [42,(t,c)=>`MAX(0,${c}${40+o}+${c}${39+o}/2-${c}${41+o}/2)*${p(58)}/4`],
 [43,(t,c)=>`${c}${40+o}+${c}${39+o}-${c}${41+o}`],
 [44,(t,c)=>`IF(${t}<=${p(16)}*4+${p(148)}${k?'+'+p(124):''},${c}${42+o},0)`],
 [45,(t,c)=>`${c}${42+o}-${c}${44+o}`],[46,(t,c)=>`${c}${41+o}+${c}${42+o}`],
 [47,(t,c)=>`SUM($D${36+o}:${c}${36+o})+SUM($D${37+o}:${c}${37+o})`],
 [48,(t,c)=>`SUM(${c}${36+o}:${c}${38+o})+${c}${44+o}+${c}${49+o}`],
 [49,t=>`IF(AND(${p(57)}>0,${t}=${inv(28,k)}),${inv(19,k)},0)`],
 [50,(t,c)=>`IF(AND(${p(57)}>0,${t}=${inv(29,k)}),${c}${43+o},0)`],
 [51,(t,c)=>`SUM(${c}${36+o}:${c}${38+o})`],
 [52,(t,c)=>`SUM($D${36+o}:${c}${36+o})+SUM($D${35+o}:${c}${35+o})`],
 [60,(t,c)=>`IF(${c}${39+o}>0.000001,1,0)`]
 ])rows('季度投融资',r+o,(t,c)=>`IF(${t}<=${p(9)}*4,${fn(t,c)},0)`);
}
f('季度投融资','B33',`IF(AND(${[0,1].map(k=>`OR(${p(57)}=0,AND(${inv(28,k)}>0,${inv(29,k)}<=${p(9)}*4,${inv(30,k)}>${p(16)}*4+${p(148)}${k?'+'+p(124):''},${inv(30,k)}<=${inv(29,k)},${p(57)}<=${inv(17,k)}+0.01)),ABS(SUM(${rr('季度投融资',39,k*O)})-${p(57)})<=0.01,MIN(${rr('季度投融资',43,k*O)})>=-0.01`).join(',')}),1,0)`);v('季度投融资','A33','融资方案基础有效标志');
console.log('investment and debt written');

// Operating build. Amounts accrued, amounts collected and cash taxes are distinct.
const opNames={10:'销售交付面积（平方米）',11:'含税销售收入',12:'销售实收现金',13:'期末销售应收款',14:'出租年度目标比例',15:'季度末出租率',16:'期初出租面积（平方米）',17:'期末出租面积（平方米）',18:'新增出租面积（平方米）',19:'季度平均出租面积（平方米）',20:'月租金（元/平方米/月）',21:'免租前合同租金',22:'新增租户免租减收',23:'合同租金收入（含税）',24:'预计租金坏账费用',25:'租金预计可收回额',26:'租金实收现金',27:'其他经营收入（含税）',28:'其他经营实收现金',29:'销售销项增值税',30:'出租销项增值税',31:'其他销项增值税',32:'销售收入（利润口径）',33:'持有收入（利润口径）',34:'营业收入合计（利润口径）',35:'销售费用',36:'销售其他流转税费计提',37:'持有固定运营成本',38:'持有变动成本',39:'常年维护及重大更新',40:'持有其他流转税费计提',41:'持有建筑折旧（项目）',42:'持有建筑折旧（融资）',43:'持有土地摊销',44:'销售成本结转（项目）',45:'销售成本结转（融资）',46:'不含融资经营税前利润',47:'融资经营税前利润',49:'本季可抵扣建设进项税',50:'期初未抵扣进项余额',51:'本季应付增值税',52:'期末未抵扣进项余额',53:'增值税现金支付及期末结算',54:'其他流转税现金支付及期末结算',55:'持有账面净值（项目）',56:'持有账面净值（融资）',57:'未售存货（融资）',58:'退出方案持有物业售价',59:'退出方案净处置现金',60:'退出方案处置损益（项目）',61:'退出方案处置损益（融资）',62:'项目所得税现金（持续持有）',63:'项目所得税现金（期末出售）',64:'融资所得税现金（持续持有）',65:'融资所得税现金（期末出售）',66:'净利润（持续持有）',67:'净利润（期末出售）',68:'经营及销售现金流入合计',69:'现金成本税费（不含所得税）',70:'持有NOI（含维护，不含所得税）',71:'销售其他流转税现金',72:'持有其他流转税现金',73:'期末税款结算（增值税）',74:'期末税款结算（其他流转税）',75:'期末税款结算（项目所得税持有）',76:'期末税款结算（融资所得税持有）'};
const taxStart=(k,levered)=>10+k*150+(levered?75:0);
const yTax=(k,levered,row)=>ref('年度所得税',`$D$${taxStart(k,levered)+row}:$AH$${taxStart(k,levered)+row}`);
const taxCash=(t,k,levered,exit)=>`SUMIFS(${yTax(k,levered,exit?9:5)},'年度所得税'!$D$4:$AH$4,(${t}-${p(111)})/4)+IF(${t}=${p(9)}*4,SUMIFS(${yTax(k,levered,exit?9:5)},'年度所得税'!$D$4:$AH$4,">"&(${p(9)}*4-${p(111)})/4),0)`;
for(const k of [0,1]){
 const o=k*O, delay=`(${p(148)}${k?'+'+p(124):''})`, start=`((${p(18)}-1)*4+${delay}+1)`, complete=`(${p(16)}*4+${delay})`;
 section('季度经营税费',8+o,k?'组合压力 · 交付、收款、成本与税费':'基准 · 交付、收款、成本与税费',END);
 for(const [r,n]of Object.entries(opNames))label('季度经营税费',+r+o,n);
 const op=(r,t)=>a('季度经营税费',r,t,o);
 const decDep=(basis,t)=>`IF(AND(${t}>${complete},${t}<=${complete}+${p(53)}*4),${basis}*(1-${p(54)})/${p(53)}/4,0)`;
 const sumRow=(r,t,c)=>`SUM($D${r+o}:${c}${r+o})`;
 const dueTax=(r,t)=>`${delayed(r,t,p(112),o,'季度经营税费')}+IF(${t}=${p(9)}*4,SUMPRODUCT(${rr('季度经营税费',r,o)},--('季度经营税费'!$D$4:$${END}$4>${p(9)}*4-${p(112)})),0)`;
 for(const [r,fn]of [
 [10,t=>`${saleArea}*${a('季度计划',15,t,o)}`],[11,(t,c)=>`${c}${10+o}*${p(145)}${k?'*(1-'+p(78)+')':''}/10000`],
 [12,(t,c)=>`${c}${11+o}*${p(102)}+(${delayed(11,t,p(103),o,'季度经营税费')})*(1-${p(102)})`],
 [13,(t,c)=>`${sumRow(11,t,c)}-${sumRow(12,t,c)}`],
 [14,t=>`IF(${t}<${start},0,IF(INT((${t}-${start})/4)=0,${p(38)},IF(INT((${t}-${start})/4)=1,${p(39)},IF(INT((${t}-${start})/4)=2,${p(40)},${p(41)}))))`],
 [15,(t,c)=>`IF(${t}<${start},0,MAX(0,IF(INT((${t}-${start})/4)=0,0,IF(INT((${t}-${start})/4)=1,${p(38)},IF(INT((${t}-${start})/4)=2,${p(39)},IF(INT((${t}-${start})/4)=3,${p(40)},${p(41)}))))-${k?p(80)+'/100':'0'})+(MAX(0,${c}${14+o}-${k?p(80)+'/100':'0'})-MAX(0,IF(INT((${t}-${start})/4)=0,0,IF(INT((${t}-${start})/4)=1,${p(38)},IF(INT((${t}-${start})/4)=2,${p(39)},IF(INT((${t}-${start})/4)=3,${p(40)},${p(41)}))))-${k?p(80)+'/100':'0'}))*(MOD(${t}-${start},4)+1)/4)`],
 [16,t=>t?op(17,t-1):'0'],[17,(t,c)=>`${heldArea}*${c}${15+o}`],[18,(t,c)=>`MAX(0,${c}${17+o}-${c}${16+o})`],[19,(t,c)=>`${c}${16+o}+${c}${18+o}*(1-${p(125)})`],
 [20,t=>`IF(${t}<${start},0,${p(146)}*(1+${p(37)})^INT((${t}-${start})/4)${k?'*(1-'+p(79)+')':''})`],
 [21,(t,c)=>`${c}${19+o}*${c}${20+o}*3/10000`],
 [22,(t,c)=>`(${Array.from({length:5},(_,j)=>t-j<0?'0':`${op(18,t-j)}*MAX(0,MIN(3,${p(42)}+3*${p(125)}-${j}*3)-MAX(0,3*${p(125)}-${j}*3))`).join('+')})*${c}${20+o}/10000`],
 [23,(t,c)=>`MAX(0,${c}${21+o}-${c}${22+o})`],
 [24,(t,c)=>`${c}${23+o}*(1-${p(43)})/(1+IF(${p(96)}=2,${p(119)},0))`],
 [25,(t,c)=>`${c}${23+o}*${p(43)}`],[26,t=>delayed(25,t,p(104),o,'季度经营税费')],
 [27,(t,c)=>`${c}${25+o}*${p(44)}`],[28,t=>delayed(27,t,p(104),o,'季度经营税费')],
 [29,(t,c)=>`IF(${p(96)}=2,${c}${11+o}-${c}${11+o}/(1+${p(118)}),0)`],
 [30,(t,c)=>`IF(${p(96)}=2,${c}${23+o}-${c}${23+o}/(1+${p(119)}),0)`],
 [31,(t,c)=>`IF(${p(96)}=2,${c}${27+o}-${c}${27+o}/(1+${p(119)}),0)`],
 [32,(t,c)=>`${c}${11+o}-${c}${29+o}`],[33,(t,c)=>`${c}${23+o}+${c}${27+o}-${c}${30+o}-${c}${31+o}`],[34,(t,c)=>`${c}${32+o}+${c}${33+o}`],
 [35,(t,c)=>`${c}${11+o}*${p(32)}`],[36,(t,c)=>`${c}${11+o}*IF(${p(96)}=1,${p(33)},${p(121)})`],
 [37,t=>`IF(AND(${heldArea}>0,${t}>${complete}),${p(47)}*${heldArea}/MAX(1,${p(108)})*(1+${p(48)})^INT((${t}-${complete}-1)/4)/4,0)`],
 [38,(t,c)=>`(${c}${33+o}-${c}${24+o})*${p(49)}`],
 [39,t=>`IF(AND(${heldArea}>0,${t}>${complete}),${p(50)}*${heldArea}/MAX(1,${p(108)})/4+IF(${t}=${p(116)}*4,${p(117)},0),0)`],
 [40,(t,c)=>`(${c}${23+o}+${c}${27+o})*IF(${p(96)}=1,${p(51)},${p(122)})`],
 [41,t=>decDep(inv(23,k),t)],[42,t=>decDep(inv(26,k),t)],
 [43,t=>`IF(AND(${p(98)}=1,${t}>${complete},${t}<=${complete}+${p(99)}*4),${inv(24,k)}/MAX(1,${p(99)})/4,0)`],
 [44,t=>`${inv(22,k)}*${a('季度计划',15,t,o)}`],[45,t=>`${inv(25,k)}*${a('季度计划',15,t,o)}`],
 [46,(t,c)=>`${c}${34+o}-SUM(${c}${35+o}:${c}${41+o})-${c}${43+o}-${c}${44+o}-${c}${24+o}`],
 [47,(t,c)=>`${c}${34+o}-SUM(${c}${35+o}:${c}${40+o})-${c}${42+o}-${c}${43+o}-${c}${45+o}-${c}${24+o}-${a('季度投融资',45,t,o)}`],
 [49,t=>`IF(${p(96)}=2,${a('季度投融资',37,t,o)}*${p(120)},0)`],[50,t=>t?op(52,t-1):'0'],
 [51,(t,c)=>`MAX(0,SUM(${c}${29+o}:${c}${31+o})-${c}${49+o}-${c}${50+o})`],
 [52,(t,c)=>`MAX(0,${c}${49+o}+${c}${50+o}-SUM(${c}${29+o}:${c}${31+o}))`],
 [53,t=>dueTax(51,t)],
 [54,t=>`${dueTax(36,t)}+${dueTax(40,t)}`],
 [55,(t,c)=>`${inv(24,k)}+${inv(23,k)}*SUM(${ref('季度计划',`$D$${11+o}:${c}${11+o}`)})-${sumRow(41,t,c)}-${sumRow(43,t,c)}`],
 [56,(t,c)=>`${c}${55+o}+SUM(${ref('季度投融资',`$D$${44+o}:${c}${44+o}`)})*${p(143)}+SUM(${ref('季度投融资',`$D$${49+o}:${c}${49+o}`)})*${p(143)}-${sumRow(42,t,c)}+${sumRow(41,t,c)}`],
 [57,(t,c)=>`(${inv(7,k)}+${inv(15,k)}*(1-IF(${p(96)}=2,${p(120)},0))*SUM(${ref('季度计划',`$D$${11+o}:${c}${11+o}`)}))*${p(142)}+(SUM(${ref('季度投融资',`$D$${44+o}:${c}${44+o}`)})+SUM(${ref('季度投融资',`$D$${49+o}:${c}${49+o}`)}))*${p(142)}-${sumRow(45,t,c)}`],
 [58,t=>`IF(AND(${heldArea}>0,${t}=${p(9)}*4),${p(65)}${k?'*'+p(139):''},0)`],[59,(t,c)=>`${c}${58+o}*(1-${p(66)})`],
 [60,(t,c)=>`IF(AND(${heldArea}>0,${t}=${p(9)}*4),${c}${59+o}-${c}${55+o},0)`],[61,(t,c)=>`IF(AND(${heldArea}>0,${t}=${p(9)}*4),${c}${59+o}-${c}${56+o},0)`],
 [62,t=>taxCash(t,k,false,false)],[63,t=>taxCash(t,k,false,true)],[64,t=>taxCash(t,k,true,false)],[65,t=>taxCash(t,k,true,true)],
 [66,(t,c)=>`${c}${47+o}-IF(AND(${t}>0,MOD(${t},4)=0),INDEX(${yTax(k,true,5)},1,${t}/4+1),0)`],
 [67,(t,c)=>`${c}${47+o}+${c}${61+o}-IF(AND(${t}>0,MOD(${t},4)=0),INDEX(${yTax(k,true,9)},1,${t}/4+1),0)`],
 [68,(t,c)=>`${c}${12+o}+${c}${26+o}+${c}${28+o}`],
 [69,(t,c)=>`${c}${35+o}+SUM(${c}${37+o}:${c}${39+o})+${c}${53+o}+${c}${54+o}`],
 [70,(t,c)=>`${c}${33+o}-${c}${24+o}-SUM(${c}${37+o}:${c}${40+o})`],
 [71,t=>dueTax(36,t)],[72,t=>dueTax(40,t)],
 [73,t=>`IF(${t}=${p(9)}*4,SUMPRODUCT(${rr('季度经营税费',51,o)},--('季度经营税费'!$D$4:$${END}$4>${p(9)}*4-${p(112)})),0)`],
 [74,t=>`IF(${t}=${p(9)}*4,SUMPRODUCT(${rr('季度经营税费',36,o)}+${rr('季度经营税费',40,o)},--('季度经营税费'!$D$4:$${END}$4>${p(9)}*4-${p(112)})),0)`],
 [75,t=>`IF(${t}=${p(9)}*4,SUMIFS(${yTax(k,false,5)},'年度所得税'!$D$4:$AH$4,">"&(${p(9)}*4-${p(111)})/4),0)`],
 [76,t=>`IF(${t}=${p(9)}*4,SUMIFS(${yTax(k,true,5)},'年度所得税'!$D$4:$AH$4,">"&(${p(9)}*4-${p(111)})/4),0)`]
 ])rows('季度经营税费',r+o,(t,c)=>`IF(${t}<=${p(9)}*4,${fn(t,c)},0)`);
 for(const r of [14,15])sh['季度经营税费'].getRange(`D${r+o}:${END}${r+o}`).setNumberFormat(PCT);
}

console.log('operating and tax cash schedules written');
// Annual income tax, loss vintages and FIFO utilisation, independently for project/equity.
headerY('年度所得税');
for(const k of [0,1])for(const levered of [false,true]){
 const start=taxStart(k,levered),o=k*O, mat=start+25, incomeRow=levered?47:46,gainRow=levered?61:60;
 const matrix=Array.from({length:56},()=>Array(31).fill(''));
 section('年度所得税',start-2,`${k?'组合压力':'基准'} · ${levered?'融资':'项目'}所得税及亏损年份`, 'AH');
 const ns=['经营税前利润','期末处置损益（出售方案）','本年可用以前年度亏损','本年弥补亏损','计税利润（持续持有）','所得税（持续持有）','本年新增可结转亏损','期末亏损余额','本年到期作废亏损','所得税（期末出售）','含出售税前利润','当前口径所得税'];
 ns.forEach((n,i)=>label('年度所得税',start+i,n));
 for(let y=0;y<=30;y++){
  const c=col(4+y),prev=col(3+y), row=i=>`${c}${start+i}`;
  const before= y?`SUMIFS(${ref('年度所得税',`$${prev}$${mat}:$${prev}$${mat+30}`)},$B$${mat}:$B$${mat+30},">="&IF(${p(101)}=0,0,${y}-${p(101)}))`:'0';
  const forms={0:`SUMIFS(${rr('季度经营税费',incomeRow,o)},'季度经营税费'!$D$6:$${END}$6,${y})`,1:`SUMIFS(${rr('季度经营税费',gainRow,o)},'季度经营税费'!$D$6:$${END}$6,${y})`,2:`IF(${p(100)}=1,${before},0)`,3:`MIN(MAX(0,${row(0)}),${row(2)})`,4:`MAX(0,${row(0)}-${row(3)})`,5:`${row(4)}*${p(52)}`,6:`IF(${p(100)}=1,MAX(0,-${row(0)}),0)`,7:`SUM(${c}${mat}:${c}${mat+30})`,8:y?`MAX(0,${prev}${start+7}-${row(2)})`:'0',9:`MAX(0,${row(0)}+${row(1)}-MIN(MAX(0,${row(0)}+${row(1)}),${row(2)}))*${p(52)}`,10:`${row(0)}+${row(1)}`,11:`IF(${p(64)}=1,${row(9)},${row(5)})`};
  for(const [r,x]of Object.entries(forms))matrix[+r][y]=`=IF(${y}<=${p(9)},${x},0)`;
  for(let vintage=0;vintage<=30;vintage++){
   const r=mat+vintage;if(y===0){v('年度所得税',`A${r}`,`第${vintage}年亏损结余`);v('年度所得税',`B${r}`,vintage);}
   let form='0';
   if(y===vintage)form=row(6);
   else if(y>vintage){
    const prefix=vintage===0?'0':`SUMIFS($${prev}$${mat}:$${prev}$${r-1},$B$${mat}:$B$${r-1},">="&IF(${p(101)}=0,0,${y}-${p(101)}))`;
    form=`IF(AND(${p(100)}=1,OR(${p(101)}=0,${y}-${vintage}<=${p(101)})),MAX(0,${prev}${r}-MAX(0,${row(3)}-${prefix})),0)`;
   }
   matrix[25+vintage][y]=`=IF(${y}<=${p(9)},${form},0)`;
  }
 }
 sh['年度所得税'].getRange(`D${start}:AH${start+55}`).formulas=matrix;
}

console.log('annual income tax and loss vintages written');
// Cashflow, reverse reserve, theoretical full equity and capped executable funding.
for(const k of [0,1]){
 const o=k*O, cf=(r,t)=>a('季度现金流',r,t,o), op=(r,t)=>a('季度经营税费',r,t,o), fi=(r,t)=>a('季度投融资',r,t,o);
 section('季度现金流',8+o,k?'组合压力 · 现金流与资金约束':'基准 · 现金流与资金约束',END);
 const labels={10:'持续持有项目经营现金',11:'项目投资现金支出',12:'贷款提款',13:'全部利息现金支付',14:'合同计划还本',15:'融资手续费现金支付',16:'持续持有周转回收',17:'出售退出周转回收',18:'项目现金流（持续持有）',19:'项目现金流（期末出售）',20:'股东投入前净现金（持有）',21:'股东投入前净现金（出售）',25:'持有后续资金留存需求',26:'持有理论期初现金',27:'持有完成计划所需出资',28:'持有可分配现金',29:'持有理论期末现金',30:'持有累计所需资本金',31:'资本金现金流（持续持有）',32:'持有理论资金平衡差额',35:'出售后续资金留存需求',36:'出售理论期初现金',37:'出售完成计划所需出资',38:'出售可分配现金',39:'出售理论期末现金',40:'出售累计所需资本金',41:'资本金现金流（期末出售）',42:'出售理论资金平衡差额',45:'当前项目现金流',46:'当前股东投入前净现金',47:'当前理论资本金现金流',48:'当前计划留存需求',49:'当前所需股东投入',50:'当前理论股东分配',51:'当前理论期末现金',52:'额度内期初现金',53:'期初未落实缺口',54:'额度内当季资金需要',55:'额度内实际可投入资本金',56:'额度内允许分配现金',57:'额度内期末可用现金',58:'期末未落实资金缺口',59:'历史最大未落实缺口',60:'累计额度内可投入资本金',61:'额度内资金平衡差额',62:'项目现金流符号',63:'项目最近非零符号',64:'项目现金流变号',65:'经营偿债现金CFADS（不含退出）',66:'当季经营偿债覆盖倍数',67:'需动用留存或股东的偿债金额',68:'累计项目现金流',69:'当季项目折现现金流',70:'累计折现项目现金流',71:'静态首次跨零辅助',72:'动态首次跨零辅助',73:'资本金现金流符号',74:'资本金最近非零符号',75:'资本金现金流变号',76:'累计融资前现金（含融资成本）',77:'融资前累计资金需要',78:'累计已支付投资含融资费用'};
 for(const [r,n]of Object.entries(labels))label('季度现金流',+r+o,n);
 const forms=[
 [10,t=>`${op(68,t)}-${op(69,t)}-${op(62,t)}`],[11,t=>fi(51,t)],[12,t=>fi(39,t)],[13,t=>fi(42,t)],[14,t=>fi(41,t)],[15,t=>fi(49,t)],
 [16,t=>`IF(${t}=${p(9)}*4,${inv(20,k)}*${p(109)},0)`],[17,t=>`IF(${t}=${p(9)}*4,${inv(20,k)}*${p(110)},0)`],
 [18,(t,c)=>`${c}${10+o}-${c}${11+o}+${c}${16+o}`],[19,(t,c)=>`${op(68,t)}-${op(69,t)}-${op(63,t)}-${c}${11+o}+${c}${17+o}+${op(59,t)}`],
 [20,(t,c)=>`${op(68,t)}-${op(69,t)}-${op(64,t)}-${c}${11+o}+${c}${16+o}+${c}${12+o}-${c}${13+o}-${c}${14+o}-${c}${15+o}`],
 [21,(t,c)=>`${op(68,t)}-${op(69,t)}-${op(65,t)}-${c}${11+o}+${c}${17+o}+${op(59,t)}+${c}${12+o}-${c}${13+o}-${c}${14+o}-${c}${15+o}`]
 ];
 for(const exit of [false,true]){
  const b=exit?35:25, pre=exit?21:20;
  forms.push([b,t=>t===Q?'0':`IF(${t}<${p(9)}*4,MAX(0,${cf(b,t+1)}-${cf(pre,t+1)}),0)`]);
  forms.push([b+1,t=>t?cf(b+4,t-1):'0']);
  forms.push([b+2,(t,c)=>`MAX(0,${t===0?'0':exit?`IF(${t}=${p(9)}*4,0,${p(123)})`:p(123)}-(${c}${b+1+o}+${c}${pre+o}))`]);
  forms.push([b+3,(t,c)=>`MAX(0,${c}${b+1+o}+${c}${pre+o}+${c}${b+2+o}-${c}${b+o}-${exit?`IF(${t}=${p(9)}*4,0,${p(123)})`:p(123)})`]);
  forms.push([b+4,(t,c)=>`${c}${b+1+o}+${c}${pre+o}+${c}${b+2+o}-${c}${b+3+o}`]);
  forms.push([b+5,(t,c)=>`SUM($D${b+2+o}:${c}${b+2+o})`]);
  forms.push([b+6,(t,c)=>`${c}${b+3+o}-${c}${b+2+o}`]);
  forms.push([b+7,(t,c)=>`${c}${b+1+o}+${c}${pre+o}+${c}${b+2+o}-${c}${b+3+o}-${c}${b+4+o}`]);
 }
 forms.push(
 [45,(t,c)=>`IF(${p(64)}=1,${c}${19+o},${c}${18+o})`],[46,(t,c)=>`IF(${p(64)}=1,${c}${21+o},${c}${20+o})`],[47,(t,c)=>`IF(${p(64)}=1,${c}${41+o},${c}${31+o})`],[48,(t,c)=>`IF(${p(64)}=1,${c}${35+o},${c}${25+o})`],[49,(t,c)=>`IF(${p(64)}=1,${c}${37+o},${c}${27+o})`],[50,(t,c)=>`IF(${p(64)}=1,${c}${38+o},${c}${28+o})`],[51,(t,c)=>`IF(${p(64)}=1,${c}${39+o},${c}${29+o})`],
 [52,t=>t?cf(57,t-1):'0'],[53,t=>t?cf(58,t-1):'0'],
 [54,(t,c)=>`MAX(0,-(${c}${52+o}+${c}${46+o}-${c}${53+o})+IF(AND(${t}>0,OR(${p(64)}=0,${t}<${p(9)}*4)),${p(123)},0))`],
 [55,(t,c)=>`MIN(${c}${54+o},MAX(0,${p(61)}-${t?cf(60,t-1):'0'}))`],
 [56,(t,c)=>`IF(${c}${59+o}>0.01,0,MAX(0,${c}${52+o}+${c}${46+o}-${c}${53+o}+${c}${55+o}-${c}${48+o}-IF(OR(${p(64)}=0,${t}<${p(9)}*4),${p(123)},0)))`],
 [57,(t,c)=>`MAX(0,${c}${52+o}+${c}${46+o}-${c}${53+o}+${c}${55+o})-${c}${56+o}`],
 [58,(t,c)=>`MAX(0,-(${c}${52+o}+${c}${46+o}-${c}${53+o}+${c}${55+o}))`],
 [59,(t,c)=>`MAX(${t?cf(59,t-1):'0'},${c}${58+o})`],[60,(t,c)=>`${t?cf(60,t-1):'0'}+${c}${55+o}`],
 [61,(t,c)=>`${c}${52+o}-${c}${53+o}+${c}${46+o}+${c}${55+o}-${c}${56+o}-${c}${57+o}+${c}${58+o}`],
 [62,(t,c)=>`IF(ABS(${c}${45+o})<=0.000001,0,SIGN(${c}${45+o}))`],[63,(t,c)=>`IF(${c}${62+o}=0,${t?cf(63,t-1):'0'},${c}${62+o})`],[64,(t,c)=>t?`IF(AND(${c}${62+o}<>0,${cf(63,t-1)}<>0,${c}${62+o}<>${cf(63,t-1)}),1,0)`:'0'],
 [65,t=>`${op(68,t)}-${op(69,t)}-${op(64,t)}`],
 [66,(t,c)=>`IF(AND(${t}>${p(16)}*4+${p(148)}${k?'+'+p(124):''},${fi(46,t)}>0.000001),${c}${65+o}/${fi(46,t)},"")`],
 [67,(t,c)=>`MAX(0,${fi(46,t)}-${c}${65+o})`],[68,(t,c)=>`${t?cf(68,t-1):'0'}+${c}${45+o}`],[69,(t,c)=>`${c}${45+o}/(1+${p(71)})^(${t}/4)`],[70,(t,c)=>`${t?cf(70,t-1):'0'}+${c}${69+o}`],
 [71,(t,c)=>t?`IF(AND(${cf(68,t-1)}<0,${c}${68+o}>=0),1,0)`:'0'],[72,(t,c)=>t?`IF(AND(${cf(70,t-1)}<0,${c}${70+o}>=0),1,0)`:'0'],
 [73,(t,c)=>`IF(ABS(${c}${47+o})<=0.000001,0,SIGN(${c}${47+o}))`],[74,(t,c)=>`IF(${c}${73+o}=0,${t?cf(74,t-1):'0'},${c}${73+o})`],[75,(t,c)=>t?`IF(AND(${c}${73+o}<>0,${cf(74,t-1)}<>0,${c}${73+o}<>${cf(74,t-1)}),1,0)`:'0'],
 [76,(t,c)=>`${t?cf(76,t-1):'0'}+${c}${46+o}-${c}${12+o}+${c}${14+o}`],[77,(t,c)=>`MAX(0,-${c}${76+o})`],[78,(t,c)=>`SUM(${ref('季度投融资',`$D$${48+o}:${c}${48+o}`)})`]);
 for(const [r,fn]of forms)rows('季度现金流',r+o,(t,c)=>`IF(${t}<=${p(9)}*4,${fn(t,c)},0)`);
 label('季度现金流',79+o,'当前累计所需股东出资');rows('季度现金流',79+o,(t,c)=>`IF(${t}<=${p(9)}*4,SUM($D${49+o}:${c}${49+o}),0)`);
}
console.log('cashflow and reserve written');

// Annual summaries keep stocks at year-end and sum only additive flows.
headerY('年度汇总');
const annualNames={10:'销售收入（利润口径）',11:'持有收入（利润口径）',12:'销售实收现金',13:'租金实收现金',14:'其他经营实收现金',15:'现金成本及流转税',16:'融资所得税现金',17:'净利润',18:'投资支付含资本化融资费用',19:'贷款提款',20:'计划还本',21:'全部利息现金',22:'经营偿债现金CFADS',23:'合同本息',24:'年度经营DSCR',25:'项目投资现金流',26:'资本金现金流（完整出资）',27:'所需股东投入',28:'理论股东分配',29:'年末留存现金',30:'年内最大未落实缺口',31:'偿债需要其他资金支持',32:'年末计划贷款余额',33:'累计项目现金流',34:'本年项目折现现金流',35:'累计项目折现现金流',36:'持有期末账面净值',37:'年末销售应收',38:'年末进项税留抵',39:'周转资金回收',40:'持有物业出售净现金',41:'期末税款及负债结算',42:'最低DSCR辅助筛选'};
for(const k of [0,1]){
 const o=k*O,ay=k*60;
 section('年度汇总',8+ay,k?'组合压力 · 年度汇总':'基准 · 年度汇总','AH');
 for(const [r,n]of Object.entries(annualNames))label('年度汇总',+r+ay,n,r==24||r==42?'倍':'万元');
 const summ=(s,r,y)=>`SUMIFS(${rr(s,r,o)},${ref(s,`$D$6:$${END}$6`)},${y})`;
 const ending=(s,r,y)=>`INDEX(${rr(s,r,o)},1,${y}*4+1)`;
 for(let y=0;y<=30;y++){
  const c=col(4+y),ar=r=>`${c}${r+ay}`;
  const forms={10:summ('季度经营税费',32,y),11:summ('季度经营税费',33,y),12:summ('季度经营税费',12,y),13:summ('季度经营税费',26,y),14:summ('季度经营税费',28,y),15:summ('季度经营税费',69,y),16:`IF(${p(64)}=1,${summ('季度经营税费',65,y)},${summ('季度经营税费',64,y)})`,17:`IF(${p(64)}=1,${summ('季度经营税费',67,y)},${summ('季度经营税费',66,y)})`,18:summ('季度投融资',48,y),19:summ('季度投融资',39,y),20:summ('季度投融资',41,y),21:summ('季度投融资',42,y),22:summ('季度现金流',65,y),23:summ('季度投融资',46,y),24:`IF(AND(${y}*4>${p(16)}*4+${p(148)}${k?'+'+p(124):''},${ar(23)}>0.000001),${ar(22)}/${ar(23)},"")`,25:summ('季度现金流',45,y),26:summ('季度现金流',47,y),27:summ('季度现金流',49,y),28:summ('季度现金流',50,y),29:ending('季度现金流',51,y),30:y===0?a('季度现金流',58,0,o):`MAX(${ref('季度现金流',`$${col(4+(y-1)*4+1)}$${58+o}:$${col(4+y*4)}$${58+o}`)})`,31:summ('季度现金流',67,y),32:ending('季度投融资',43,y),33:ending('季度现金流',68,y),34:summ('季度现金流',69,y),35:ending('季度现金流',70,y),36:ending('季度经营税费',56,y),37:ending('季度经营税费',13,y),38:ending('季度经营税费',52,y),39:`IF(${p(64)}=1,${summ('季度现金流',17,y)},${summ('季度现金流',16,y)})`,40:`IF(${p(64)}=1,${summ('季度经营税费',59,y)},0)`,41:`${summ('季度经营税费',73,y)}+${summ('季度经营税费',74,y)}+IF(${p(64)}=1,0,${summ('季度经营税费',76,y)})`,42:`IF(ISNUMBER(${ar(24)}),${ar(24)},1E+99)`};
  for(const [r,x]of Object.entries(forms))f('年度汇总',`${c}${+r+ay}`,`IF(${y}<=${p(9)},${x},${r==42?'1E+99':r==24?'""':'0'})`);
 }
 sh['年度汇总'].getRange(`D${42+ay}:AH${42+ay}`).setNumberFormat('0.00;[Red](0.00);"—"');
}

// Metrics use quarterly equally spaced periods, annual effective IRR and quarterly discounting.
const metricNames={6:'项目税后IRR（年化）',7:'资本金IRR（年化）',8:'项目NPV',9:'资本金NPV',10:'静态投资回收期',11:'动态投资回收期',12:'最低年度经营DSCR',13:'所需累计资本金',14:'最大未落实资金缺口',15:'到期计划本金余额',16:'总投资含建设期融资费用',17:'既定贷款额度',18:'可投入资本金额度',19:'测算期销售交付比例',20:'销售交付完成时间',21:'持有出租面积',22:'稳定运营现金保本出租率',23:'持有稳定NOI收益率',24:'期末净处置占现金流入',25:'项目IRR变号提示',26:'资本金IRR变号提示',27:'资本金收益可执行性',28:'稳定观察项目年度',29:'融资前峰值资金需要',30:'融资前峰值发生季度',31:'未落实缺口高峰季度',32:'累计资本金投入高峰季度',33:'最低额度内期末现金',34:'期末销售应收余额',35:'累计净利润',36:'流转税口径',37:'贷款合同到期时间',38:'持续持有项目IRR',39:'期末出售项目IRR',40:'持续持有项目NPV',41:'期末出售项目NPV',42:'持续持有资本金IRR',43:'期末出售资本金IRR',44:'当前口径周转回收',45:'持有账面净值（出售前）',46:'期末未抵扣进项税',47:'最大理论留存现金',48:'最低DSCR发生年度',49:'模型基础有效标志',50:'已建模全部财务条件达标标志',51:'资金约束达标标志',52:'项目NPV盈亏差额',53:'稳定经营现金NOI',54:'已建模参数确认状态'};
for(const [r,n]of Object.entries(metricNames))v('指标与情景',`A${r}`,n);
v('指标与情景','A4','指标');v('指标与情景','B4','基准');v('指标与情景','C4','单位');v('指标与情景','D4','阈值');v('指标与情景','E4','组合压力');v('指标与情景','F4','基准状态');v('指标与情景','G4','压力状态');
const percentageMetrics=new Set([6,7,19,22,23,24,38,39,42,43]);
const irr=(r,o)=>`IF(AND(COUNTIFS(${active('季度现金流',r,o)},"<0")>0,COUNTIFS(${active('季度现金流',r,o)},">0")>0),IFERROR((1+IRR(${active('季度现金流',r,o)}))^4-1,"无法计算"),"无法计算")`;
const npv=(r,o,rate)=>`NPV((1+${p(rate)})^(1/4)-1,${active('季度现金流',r,o,1)})+${a('季度现金流',r,0,o)}`;
const payback=(o,dynamic)=>{
 const mark=active('季度现金流',dynamic?72:71,o),cum=rr('季度现金流',dynamic?70:68,o),flow=rr('季度现金流',dynamic?69:45,o);
 return `IF(COUNTIFS(${active('季度现金流',dynamic?69:45,o)},"<0")=0,"不适用",IF(COUNTIFS(${mark},1)=0,"测算期内未回收",(MATCH(1,${mark},0)-2-INDEX(${cum},1,MAX(1,MATCH(1,${mark},0)-1))/INDEX(${flow},1,MATCH(1,${mark},0)))/4))`;
};
for(const k of [0,1]){
 const o=k*O,c=k?'E':'B',ay=k*60,m=r=>`${c}${r}`;
 const when=(s,r,kind='MAX')=>`INDEX('季度计划'!$D$5:$${END}$5,1,MATCH(${kind}(${active(s,r,o)}),${active(s,r,o)},0))`;
 const stableY=`${p(18)}+3+INT((${p(148)}${k?'+'+p(124):''})/4)`,stableQ=`MIN(${p(9)}*4,(${p(18)}+3)*4+${p(148)}${k?'+'+p(124):''})`;
 const stable=(r)=>`INDEX(${rr('季度经营税费',r,o)},1,${stableQ}+1)`;
 const forms={6:irr(45,o),7:irr(47,o),8:npv(45,o,71),9:npv(47,o,72),10:payback(o,false),11:payback(o,true),12:`IF(${p(57)}=0,"不适用",IF(COUNT('年度汇总'!$D$${24+ay}:$AH$${24+ay})=0,"无法计算",MIN('年度汇总'!$D$${42+ay}:$AH$${42+ay})))`,13:`SUM(${active('季度现金流',49,o)})`,14:`MAX(${active('季度现金流',58,o)})`,15:`IF(${p(57)}=0,0,IF(${inv(29,k)}>${p(9)}*4,"无法计算",INDEX(${rr('季度投融资',43,o)},1,MIN(121,${inv(29,k)}+1))))`,16:inv(21,k),17:p(57),18:p(61),19:`SUM(${active('季度计划',15,o)})`,20:`IF(${p(142)}=0,"不适用",IF(${m(19)}<0.999999,"期内未完成",INDEX('季度计划'!$D$5:$${END}$5,1,MATCH(1,${ref('指标与情景',`$D$${90+o}:$${END}$${90+o}`)},0))))`,21:heldArea,22:`IF(${heldArea}=0,"不适用",IF(${p(9)}<${stableY},"不适用",IF(${heldArea}*${stable(20)}*12/10000*${p(43)}*(1+${p(44)})*(1-${p(49)}-IF(${p(96)}=1,${p(51)},${p(122)}))<=0,"无法计算",(${stable(37)}*4+${p(50)}*${heldArea}/MAX(1,${p(108)}))/(${heldArea}*${stable(20)}*12/10000*${p(43)}*(1+${p(44)})*(1-${p(49)}-IF(${p(96)}=1,${p(51)},${p(122)}))))`,23:`IF(OR(${heldArea}=0,${inv(27,k)}=0,${p(9)}<${stableY}),"不适用",${stable(70)}*4/${inv(27,k)})`,24:`IF(SUM(${active('季度经营税费',68,o)})+SUM(${active('季度经营税费',59,o)})=0,"不适用",IF(${p(64)}=1,SUM(${active('季度经营税费',59,o)})/(SUM(${active('季度经营税费',68,o)})+SUM(${active('季度经营税费',59,o)})),0))`,25:`IF(SUM(${active('季度现金流',64,o)})>1,"多次变号，IRR需结合NPV审视","单次变号或无变号")`,26:`IF(SUM(${active('季度现金流',75,o)})>1,"多次变号，IRR需结合NPV审视","单次变号或无变号")`,27:`IF(${m(14)}>0.01,"理论收益，资金尚未落实","完整出资、留存后分配口径")`,28:`IF(${p(9)}>=${stableY},${stableY},"未覆盖稳定观察年")`,29:`MAX(${active('季度现金流',77,o)})`,30:when('季度现金流',77),31:`IF(${m(14)}<=0.01,"无未落实缺口",${when('季度现金流',58)})`,32:when('季度现金流',k?40:30),33:`MIN(${active('季度现金流',57,o)})`,34:`INDEX(${rr('季度经营税费',13,o)},1,${p(9)}*4+1)`,35:`SUM('年度汇总'!$D$${17+ay}:$AH$${17+ay})`,36:p(152),37:`IF(${p(57)}=0,"无贷款",INDEX('季度计划'!$D$5:$${END}$5,1,MAX(1,MIN(121,${inv(29,k)}+1))))`,38:irr(18,o),39:irr(19,o),40:npv(18,o,71),41:npv(19,o,71),42:irr(31,o),43:irr(41,o),44:`SUM(${active('季度现金流',16,o)})*(1-${p(64)})+SUM(${active('季度现金流',17,o)})*${p(64)}`,45:`INDEX(${rr('季度经营税费',56,o)},1,${p(9)}*4+1)`,46:`INDEX(${rr('季度经营税费',52,o)},1,${p(9)}*4+1)`,47:`MAX(${active('季度现金流',51,o)})`,48:`IF(${p(57)}=0,"不适用",IF(ISNUMBER(${m(12)}),INDEX('年度汇总'!$D$5:$AH$5,1,MATCH(${m(12)},'年度汇总'!$D$${24+ay}:$AH$${24+ay},0)),"无法计算"))`,49:`IF(AND(${p(150)}=1,'季度计划'!$B$30=1,'季度投融资'!$B$33=1,${inv(21,k)}>0),1,0)`,50:`IF(AND(${m(49)}=1,COUNTIFS(${k?'G':'F'}6:${k?'G':'F'}15,"未通过")=0,COUNTIFS(${k?'G':'F'}6:${k?'G':'F'}15,"待判断")=0),1,0)`,51:`IF(AND(ISNUMBER(${m(13)}),${m(13)}<=${p(61)}+0.01,${m(14)}<=0.01,ISNUMBER(${m(15)}),ABS(${m(15)})<=0.01),1,0)`,52:m(8),53:`IF(OR(${heldArea}=0,${p(9)}<${stableY}),"不适用",${stable(70)}*4)`,54:`IF(${p(97)}=1,"新增假设已核对","新增假设待项目核对")`};
 const fullRent=`(${heldArea}*${stable(20)}*12/10000)`;
 const margin=`(${p(43)}*(1+${p(44)})*(1-${p(49)})/(1+IF(${p(96)}=2,${p(119)},0))-(1+${p(43)}*${p(44)})*IF(${p(96)}=1,${p(51)},${p(122)}))`;
 forms[22]=`IF(OR(${heldArea}=0,${p(9)}<${stableY}),"不适用",IF(${fullRent}*${margin}<=0,"无法计算",(${stable(37)}*4+${p(50)}*${heldArea}/MAX(1,${p(108)}))/(${fullRent}*${margin})))`;
 forms[32]=when('季度现金流',79);
 for(const [r,x]of Object.entries(forms))f('指标与情景',`${c}${r}`,x);
 rows('指标与情景',90+o,(t,cq)=>`IF(SUM(${ref('季度计划',`$D$${15+o}:${cq}${15+o}`)})>=0.999999,1,0)`);
 for(const r of percentageMetrics)sh['指标与情景'].getRange(`${c}${r}`).setNumberFormat(PCT);
}
const threshold={6:69,7:70,8:null,9:null,10:73,11:74,12:75,13:61,14:null,15:null};
for(const [r,pr]of Object.entries(threshold)){
 f('指标与情景',`D${r}`,pr?p(pr):'0');
 for(const [data,state]of [['B','F'],['E','G']]){
  const x=`${data}${r}`,n=+r;
  let test=n===12?`${x}>=D${r}`:n===10||n===11||n===13?`${x}<=D${r}${n===13?'+0.01':''}`:n===14||n===15?`ABS(${x})<=0.01`:`${x}>=D${r}`;
  f('指标与情景',`${state}${r}`,`IF(${x}="不适用",${n===12?'"通过"':'"未通过"'},IF(NOT(ISNUMBER(${x})),IF(${x}="无法计算","待判断","未通过"),IF(${test},"通过","未通过")))`);
 }
}
for(const r of [6,7])sh['指标与情景'].getRange(`D${r}`).setNumberFormat(PCT);
for(const [r,n]of Object.entries(metricNames))v('指标与情景',`C${r}`,percentageMetrics.has(+r)?'%':[10,11,28].includes(+r)?'年':+r===12?'倍':[20,25,26,27,30,31,32,36,37,48,54].includes(+r)?'说明':[49,50,51].includes(+r)?'0/1':+r===21?'平方米':'万元');
const metricNotes={6:'季度IRR按(1+r)^4－1年化；只取第0季至第N年末',7:'完整出资并预留后续资金后的分配；不代表未落实资金已到位',8:'按年折现率换算季度有效利率，第0季不折现',10:'首次跨零、季内线性插值；若以后再次转负需另审视',12:'按年度CFADS／年度合同本息；季度覆盖另见现金流表',15:'合同计划还款后的余额，不等于实际已经偿还',22:'仅稳定运营现金保本，不含所得税、折旧、本息和投资回收；无新增免租',23:'稳定观察季度NOI×4／持有分摊总投资；非IRR',29:'排除贷款和股东流入及本金偿还，包含融资利息和手续费',38:'持续持有不将物业、存货或应收自动变现；周转按B109',39:'仅出售持有物业，扣费用与所得税，周转按B110',49:'输入、计划、融资的本地必要条件；独立校验在模型校验表',50:'只代表当前已建模财务条件；不包含未确认假设和市场论证'};
for(const [r,n]of Object.entries(metricNotes)){v('指标与情景',`H${r}`,n);sh['指标与情景'].getRange(`H${r}:L${r}`).merge();sh['指标与情景'].getRange(`H${r}:L${r}`).format={wrapText:true,font:{color:'#64748B'}};}

// Independent terminal review. Outputs and builds do not depend on this audit sheet.
v('模型校验','A5','校验事项');v('模型校验','B5','基准差额/标志');v('模型校验','C5','压力差额/标志');v('模型校验','D5','判断');v('模型校验','E5','说明');
const checks=[];
function ck(name,x,y,kind='zero',note=''){checks.push({name,x,y,kind,note});}
ck('输入完整及基础范围',p(150),p(150),'one','必要输入完整、单位及范围正确');
ck('季度建设、付款及销售时点','\'季度计划\'!$B$30','\'季度计划\'!$B$30','one','形成与支付分别检验，交付不早于建成');
ck('贷款方案基础有效','\'季度投融资\'!$B$33','\'季度投融资\'!$B$33','one','首笔、到期、合同还款及额度范围');
ck('经营面积分解',`${saleArea}+${heldArea}-${p(10)}*${p(11)}*${p(12)}`,`${saleArea}+${heldArea}-${p(10)}*${p(11)}*${p(12)}`);
ck('非土地工程形成总额',`SUM(${rr('季度投融资',35)})-${inv(17)}`,`SUM(${rr('季度投融资',35,O)})-${inv(17,1)}`);
ck('非土地实际付款总额',`SUM(${rr('季度投融资',37)})-${inv(17)}`,`SUM(${rr('季度投融资',37,O)})-${inv(17,1)}`);
ck('投资现金与总投资',`SUM(${rr('季度投融资',48)})-${inv(21)}`,`SUM(${rr('季度投融资',48,O)})-${inv(21,1)}`);
ck('提款额度一致',`SUM(${rr('季度投融资',39)})-${p(57)}`,`SUM(${rr('季度投融资',39,O)})-${p(57)}`);
ck('贷款滚动平衡',`SUM(${rr('季度投融资',39)})-SUM(${rr('季度投融资',41)})-INDEX(${rr('季度投融资',43)},1,${p(9)}*4+1)`,`SUM(${rr('季度投融资',39,O)})-SUM(${rr('季度投融资',41,O)})-INDEX(${rr('季度投融资',43,O)},1,${p(9)}*4+1)`);
ck('持有理论现金平衡',`SUMPRODUCT(ABS(${rr('季度现金流',32)}))`,`SUMPRODUCT(ABS(${rr('季度现金流',32,O)}))`);
ck('出售理论现金平衡',`SUMPRODUCT(ABS(${rr('季度现金流',42)}))`,`SUMPRODUCT(ABS(${rr('季度现金流',42,O)}))`);
ck('额度内现金及缺口平衡',`SUMPRODUCT(ABS(${rr('季度现金流',61)}))`,`SUMPRODUCT(ABS(${rr('季度现金流',61,O)}))`);
ck('资本金未突破额度',`MAX(0,SUM(${rr('季度现金流',55)})-${p(61)})`,`MAX(0,SUM(${rr('季度现金流',55,O)})-${p(61)})`);
for(const [r,ratio,n]of [[16,109,'持续持有周转'],[17,110,'出售退出周转']]){
 ck(`${n}回收总额`,`SUM(${rr('季度现金流',r)})-${inv(20)}*${p(ratio)}`,`SUM(${rr('季度现金流',r,O)})-${inv(20,1)}*${p(ratio)}`);
 ck(`${n}仅在第N年末`,`SUMPRODUCT(ABS(${rr('季度现金流',r)}),--('季度现金流'!$D$4:$${END}$4<>${p(9)}*4))`,`SUMPRODUCT(ABS(${rr('季度现金流',r,O)}),--('季度现金流'!$D$4:$${END}$4<>${p(9)}*4))`);
 ck(`${n}第N年末金额`,`INDEX(${rr('季度现金流',r)},1,${p(9)}*4+1)-${inv(20)}*${p(ratio)}`,`INDEX(${rr('季度现金流',r,O)},1,${p(9)}*4+1)-${inv(20,1)}*${p(ratio)}`);
}
ck('持有物业处置仅在第N年末',`SUMPRODUCT(ABS(${rr('季度经营税费',58)}),--('季度经营税费'!$D$4:$${END}$4<>${p(9)}*4))`,`SUMPRODUCT(ABS(${rr('季度经营税费',58,O)}),--('季度经营税费'!$D$4:$${END}$4<>${p(9)}*4))`);
ck('持有物业处置金额',`SUM(${rr('季度经营税费',58)})-IF(${heldArea}>0,${p(65)},0)`,`SUM(${rr('季度经营税费',58,O)})-IF(${heldArea}>0,${p(65)}*${p(139)},0)`);
ck('销售应收滚动',`SUM(${rr('季度经营税费',11)})-SUM(${rr('季度经营税费',12)})-INDEX(${rr('季度经营税费',13)},1,${p(9)}*4+1)`,`SUM(${rr('季度经营税费',11,O)})-SUM(${rr('季度经营税费',12,O)})-INDEX(${rr('季度经营税费',13,O)},1,${p(9)}*4+1)`);
ck('销售与持有资本成本分摊',`${inv(25)}+${inv(26)}+${inv(24)}-(${inv(7)}+${inv(17)}*(1-IF(${p(96)}=2,${p(120)},0))+${inv(18)}+${inv(19)})`,`${inv(25,1)}+${inv(26,1)}+${inv(24,1)}-(${inv(7,1)}+${inv(17,1)}*(1-IF(${p(96)}=2,${p(120)},0))+${inv(18,1)}+${inv(19,1)})`);
ck('纯出售无持有运营支出',`IF(${p(143)}=0,SUM(${rr('季度经营税费',37)})+SUM(${rr('季度经营税费',39)})+SUM(${rr('季度投融资',38)}),0)`,`IF(${p(143)}=0,SUM(${rr('季度经营税费',37,O)})+SUM(${rr('季度经营税费',39,O)})+SUM(${rr('季度投融资',38,O)}),0)`);
ck('纯持有无销售收入',`IF(${p(142)}=0,SUM(${rr('季度经营税费',11)}),0)`,`IF(${p(142)}=0,SUM(${rr('季度经营税费',11,O)}),0)`);
checks.forEach((x,i)=>{const r=7+i;v('模型校验',`A${r}`,x.name);f('模型校验',`B${r}`,x.x);f('模型校验',`C${r}`,x.y);f('模型校验',`D${r}`,x.kind==='one'?`IF(AND(B${r}=1,C${r}=1),"通过","未通过")`:`IF(AND(ISNUMBER(B${r}),ISNUMBER(C${r}),ABS(B${r})<=0.01,ABS(C${r})<=0.01),"通过","未通过")`);v('模型校验',`E${r}`,x.note||'金额差额容差0.01万元；零值为勾稽一致');});
sh['模型校验'].getRange(`B7:C${checks.length+6}`).setNumberFormat('0.00;[Red](0.00);0.00');
v('模型校验','A38','使用说明');v('模型校验','E38','本表为独立复核终点，不向测算和决策公式反馈。通过表示本表已覆盖关系一致，不代表项目可拿地。');

// Native sensitivities: economic inputs on this sheet genuinely drive the model.
for(const [r,n]of [[6,'临时土地单价覆盖'],[7,'临时销售均价覆盖'],[8,'临时月租金覆盖'],[9,'临时建设成本倍数覆盖'],[10,'临时整体延期季度覆盖']]){v('敏感性与边界',`A${r}`,n);input('敏感性与边界',`B${r}`);}
v('敏感性与边界','D6','黄色覆盖值留空使用核心参数；0是有效测试值。原生数据表仅临时替换这些真实驱动。');
v('敏感性与边界','D8','双变量矩阵为实时原生数据表。下方精确边界是刷新时快照，参数变化会提示需刷新。');
const dataTables=[];
function sensitivity(start,label,output,rowInput,columnInput,top,left,format=NUM){
 section('敏感性与边界',start,label,'J');const anchor= start+2;
 f('敏感性与边界',`D${anchor}`,output);
 sh['敏感性与边界'].getRange(`E${anchor}:${col(4+top.length)}${anchor}`).formulas=[top.map(x=>'='+x)];
 sh['敏感性与边界'].getRange(`D${anchor+1}:D${anchor+left.length}`).formulas=left.map(x=>['='+x]);
 const range=`D${anchor}:${col(4+top.length)}${anchor+left.length}`;
 dataTables.push({range,rowInput,columnInput});
 // Native tables are finalised in Excel; complex artifact construction exceeded several minutes.
 sh['敏感性与边界'].getRange(`E${anchor+1}:${col(4+top.length)}${anchor+left.length}`).setNumberFormat(format);
 sh['敏感性与边界'].getRange(`D${anchor}:${col(4+top.length)}${anchor}`).format={fill:C.head,font:{bold:true}};
 sh['敏感性与边界'].getRange(`D${anchor+1}:D${anchor+left.length}`).format={fill:C.head,font:{bold:true}};
}
const priceSteps=[0.8,0.9,1,1.1,1.2].map(x=>`${p(31)}*${x}`),rentSteps=[0.8,0.9,1,1.1,1.2].map(x=>`${p(36)}*${x}`),landSteps=[0,0.5,1,1.5,2].map(x=>`${p(21)}*${x}`),costSteps=[0.8,0.9,1,1.1,1.2].map(String);
sensitivity(13,'售价（横轴）／月租金（纵轴）对项目NPV的影响',"'指标与情景'!B8",'B7','B8',priceSteps,rentSteps);
sensitivity(23,'售价／月租金：已建模全部条件达标标志（1达标、0未达标）',"'指标与情景'!B50",'B7','B8',priceSteps,rentSteps,'0');
sensitivity(33,'建设成本倍数（横轴）／土地单价（纵轴）对项目NPV的影响',"'指标与情景'!B8",'B9','B6',costSteps,landSteps);
sensitivity(43,'建设成本倍数／土地单价对峰值资本金的影响',"'指标与情景'!B13",'B9','B6',costSteps,landSteps);
sensitivity(53,'售价（横轴）／整体延期季度（纵轴）对项目NPV的影响',"'指标与情景'!B8",'B7','B10',[`${p(31)}*0.9`,p(31),`${p(31)}*1.1`],['0','1','2','4']);
section('敏感性与边界',64,'精确单因素边界（同一基准模型反复求解的刷新快照）');
v('敏感性与边界','A66','边界刷新状态');v('敏感性与边界','A67','刷新时间');v('敏感性与边界','A68','当前输入签名');v('敏感性与边界','A69','上次刷新签名');
f('敏感性与边界','B66','IF(B68=B69,"与当前参数一致","参数已变化，需运行刷新工具")');v('敏感性与边界','B69','尚未刷新');
const allSignatureRefs=[...source.parameters.filter(x=>!x.formula).map(x=>p(x.row)),...added.map(x=>p(x[0])),...Array.from({length:31},(_,i)=>ref('核心参数输入',`${col(i+4)}88`)),...Array.from({length:31},(_,i)=>ref('核心参数输入',`${col(i+4)}89`))];
// Excel string concatenation is explicit and includes every original/new input,
// plus every optional quarterly override. No volatile dates in the signature.
const chunks=[];
const sigRefs=[...allSignatureRefs,...[20,21,22,23].flatMap(r=>Array.from({length:121},(_,i)=>ref('季度计划',`${col(i+4)}${r}`)))];
for(let i=0;i<sigRefs.length;i+=70){const row=105+chunks.length;f('敏感性与边界',`B${row}`,sigRefs.slice(i,i+70).map(x=>`IF(${x}="","_",${x})`).join('&"|"&'));chunks.push(`B${row}`);}
f('敏感性与边界','B68',chunks.join('&"|"&'));
const bounds=[['项目NPV=0的最高土地单价','元/平方米'],['全部财务条件下最高土地单价','元/平方米'],['项目NPV=0的最低销售均价','元/平方米'],['全部财务条件下最低销售均价','元/平方米'],['项目NPV=0的最低月租金','元/平方米/月'],['全部财务条件下最低月租金','元/平方米/月'],['全部财务条件下最低稳定出租率','%'],['全部财务条件下最高建设成本倍数','倍'],['全部财务条件下最大整体延期','季度']];
v('敏感性与边界','A72','边界事项');v('敏感性与边界','B72','结果');v('敏感性与边界','C72','单位');v('敏感性与边界','D72','解释');
bounds.forEach(([n,u],i)=>{v('敏感性与边界',`A${73+i}`,n);v('敏感性与边界',`B${73+i}`,'待刷新');v('敏感性与边界',`C${73+i}`,u);v('敏感性与边界',`D${73+i}`,'同一模型求解；超出搜索范围或无可行值明确提示');});
v('敏感性与边界','A84','边界口径');v('敏感性与边界','D84','其他条件保持基准；全部条件含收益、回收期、年度DSCR、资本金额度及缺口。示意假设确认开关不参与数值求解。');

// Dashboard: parallel return and funding information, without claiming a land decision.
const dash=sh['决策看板'];dash.freezePanes.unfreeze();
dash.getRange('A1:A120').format.columnWidth=37;dash.getRange('B1:B120').format.columnWidth=20;dash.getRange('C1:C120').format.columnWidth=13;dash.getRange('D1:D120').format.columnWidth=17;dash.getRange('E1:E120').format.columnWidth=20;dash.getRange('F1:F120').format.columnWidth=15;dash.getRange('G1:G120').format.columnWidth=24;
dash.getRange('I1:P120').format.columnWidth=12;
dash.getRange('A4:G5').merge();f('决策看板','A4',`IF(OR('指标与情景'!B49=0,AND(${p(82)}=1,'指标与情景'!E49=0)),"待判断：输入或计划条件无效",IF(AND('指标与情景'!B50=1,OR(${p(82)}=0,'指标与情景'!E50=1)),IF(${p(97)}=1,"已纳入模型的财务条件达标","财务条件达标；新增假设待项目核对"),"已纳入模型的财务条件未全部达标"))`);dash.getRange('A4:G5').format={fill:C.key,font:{bold:true,size:14},wrapText:true};
const info=[['A7','项目名称','B7',p(7)],['E7','经营模式','F7',p(149)],['A8','当前期末口径','B8',`IF(${p(64)}=1,"期末出售持有物业","持续持有观察期")`],['E8','测算期限（年）','F8',p(9)],['A9','总投资（万元）','B9',"'指标与情景'!B16"],['E9','既定贷款（万元）','F9',p(57)],['A10','所需累计资本金（万元）','B10',"'指标与情景'!B13"],['E10','资本金额度（万元）','F10',p(61)],['A11','最大未落实缺口（万元）','B11',"'指标与情景'!B14"],['E11','缺口发生季度','F11',"'指标与情景'!B31"],['A12','融资前峰值需要（万元）','B12',"'指标与情景'!B29"],['E12','峰值发生季度','F12',"'指标与情景'!B30"],['A13','理论最大留存现金（万元）','B13',"'指标与情景'!B47"],['E13','收益口径','F13',"'指标与情景'!B27"]];
for(const [ad,n,target,formula]of info){v('决策看板',ad,n);f('决策看板',target,formula);}
for(const ad of ['B7:D7','B8:D8','B9:D9','B10:D10','B11:D11','B12:D12','B13:D13','F7:G7','F8:G8','F9:G9','F10:G10','F11:G11','F12:G12','F13:G13']){dash.getRange(ad).merge();dash.getRange(ad).format={fill:C.key,font:{bold:true},wrapText:true};}
dash.getRange('A15:G15').merge();f('决策看板','A15',`${p(152)}&"；季度计算、年度汇总；"&'指标与情景'!B54`);dash.getRange('A15:G15').format={font:{color:'#64748B'},rowHeight:27};
v('决策看板','A17','指标／资金约束');v('决策看板','B17','基准');v('决策看板','C17','单位');v('决策看板','D17','阈值');v('决策看板','E17','压力');v('决策看板','F17','基准状态');v('决策看板','G17','压力状态');
for(let r=6;r<=15;r++){const d=r+12;v('决策看板',`A${d}`,metricNames[r]);for(const [dst,src]of [['B','B'],['C','C'],['D','D'],['E','E'],['F','F'],['G','G']])f('决策看板',`${dst}${d}`,ref('指标与情景',`${src}${r}`));if([6,7].includes(r))dash.getRange(`B${d}:E${d}`).setNumberFormat(PCT);}
section('决策看板',29,'持续持有与期末出售（同一经营、融资假设）','G');
v('决策看板','A30','指标');v('决策看板','B30','持续持有');v('决策看板','D30','期末出售');v('决策看板','F30','说明');
for(const [dr,n,hold,exit,u]of [[31,'项目IRR',38,39,'%'],[32,'项目NPV',40,41,'万元'],[33,'资本金IRR',42,43,'%']]){v('决策看板',`A${dr}`,n);f('决策看板',`B${dr}`,ref('指标与情景',`B${hold}`));f('决策看板',`D${dr}`,ref('指标与情景',`B${exit}`));v('决策看板',`F${dr}`,u);if(u==='%')dash.getRange(`B${dr}:D${dr}`).setNumberFormat(PCT);}
v('决策看板','A35','最低经营DSCR发生年度');f('决策看板','B35',"'指标与情景'!B48");v('决策看板','E35','稳定经营现金保本出租率');f('决策看板','F35',"'指标与情景'!B22");dash.getRange('F35').setNumberFormat(PCT);
v('决策看板','A36','期末销售应收（万元）');f('决策看板','B36',"'指标与情景'!B34");v('决策看板','E36','持续持有账面净值（万元）');f('决策看板','F36',"'指标与情景'!B45");
v('决策看板','A37','累计净利润（万元）');f('决策看板','B37',"'指标与情景'!B35");v('决策看板','E37','精确边界刷新状态');f('决策看板','F37',"'敏感性与边界'!B66");
dash.getRange('A39:G40').merge();f('决策看板','A39',`"经营DSCR衡量当年经营现金；前期销售留存另行支持偿债。"&IF('指标与情景'!B14>0.01," 存在未落实缺口，当前计划不能视为已执行。","")&" "&'指标与情景'!B25`);dash.getRange('A39:G40').format={fill:C.pale,wrapText:true};
const chartHeaders=['项目年','经营CFADS','合同本息','年末留存','未落实缺口','所需股东出资','股东分配','累计项目现金'];
dash.getRange('A80:H80').values=[chartHeaders];
for(let y=0;y<=30;y++){const r=81+y,c=col(4+y);f('决策看板',`A${r}`,`IF(${y}<=${p(9)},${y},"")`);for(const [dest,row]of [['B',22],['C',23],['D',29],['E',30],['F',27],['G',28],['H',33]])f('决策看板',`${dest}${r}`,`IF(${y}<=${p(9)},'年度汇总'!${c}${row},"")`);}
for(const [ranges,title,pos,colors]of [
 [['A80:C111'],'年度经营现金与合同本息（万元）',['I7','Q22'],['#4472C4','#A85F50']],
 [['A80:A111','D80:D111','E80:E111'],'留存现金与未落实缺口（万元）',['I25','Q40'],['#237047','#B34B4B']],
 [['A80:A111','F80:F111','G80:G111'],'股东投入与分配（万元）',['A43','H61'],['#4472C4','#6B8E77']]
]){
 const ch=dash.charts.add('line',ranges.length===1?dash.getRange(ranges[0]):ranges.map(r=>dash.getRange(r)));ch.title=title;ch.setPosition(...pos);ch.titleTextStyle.fontSize=12;ch.titleTextStyle.typeface='Microsoft YaHei';ch.legend={position:'bottom'};ch.yAxis={numberFormatCode:'#,##0',numberFormatSourceLinked:false};ch.xAxis={axisType:'textAxis'};ch.series.items.forEach((s,i)=>{s.line={fill:colors[i%colors.length],style:'solid',width:2};});
}

// Explanation records financial assumptions rather than legal or market claims.
const notes=[
 ['用途','拿地前方案测算；原示意参数及阈值保留。财务条件达标只对应已纳入模型的口径，不能替代业务假设核对。'],
 ['生成来源','V2保存内容为参数来源，V1生成代码用于追溯关系；V3没有调用旧生成器初始化或覆盖V2。'],
 ['使用步骤','先改核心参数黄色区及年度计划，再按合同需要改季度覆盖；查看指标与情景、季度缺口和模型校验；参数改动后运行同目录刷新_V3.ps1更新精确边界。'],
 ['季度与年度','底层全期按季度等距计算，年报展示稳定期；第0年为输入年份年末，第一正季度为次年一季度。季度内仍使用平均时点，关键月份应另作支付核对。'],
 ['经营模式','B95切换租售结合、纯出售、纯持有；原B14仅混合模式生效；纯出售不承担持有固定成本、维护或周转资金。'],
 ['投资范围','共同建设费用按经营面积分摊；持有专项装修全部归持有；拿地附加成本归土地。新增零值不证明项目不存在该费用，须核对包含范围。'],
 ['计划与延期','工程形成、付款和提款分别排布，黄色覆盖留空用年度分解、0表示明确零；整体延期表示土地已投后开工和租售后移，非建设跨度延长。融资首提款及合同还本配套后移。'],
 ['销售','签约分解为季度，可延迟交付；收入与成本按交付确认；首款和尾款分季收回；期末未收不自动变现，不自动作为现金偿债。'],
 ['出租','原首3年和稳定出租率在V3解释为出租年度年末目标；年内线性增长，季内按B125入驻。免租按每季新增租户队列跨季扣减，最长12个月。'],
 ['收缴与坏账','收缴率为最终可回收比例，未收部分列预计坏账；可回收部分按收款滞后到现金。未建立租户更替，不支持出租率下降路径。'],
 ['流转税','口径1保留销售8%及运营6%综合率；口径2按填写的销项率及进项比例分离，另设其他流转税率。新增零税率是示意待确认，未套用实际税务规则。'],
 ['所得税','按年度利润计税；项目口径剔除融资，股东口径包含融资。亏损按发生年份先进先用，按输入年限到期；处置仅影响末年，分别计算持有与出售所得税。'],
 ['折旧摊销','建筑直线折旧至残值，持有土地按输入开关和年限摊销；建设期利息及手续费仅融资口径资本化；未将销售成本和折旧重复扣现金。'],
 ['现金税与期末负债','税费按计提和支付滞后分开；超出观察期的已计提税款在末期作为负债结算扣除，避免延期缴付虚增回报。留抵进项不自动退款。'],
 ['融资','一笔固定额度贷款，季内平均提款和还本计息；建设利息和手续费资本金支付、不滚入本金；合同还款计划不随资金不足自动减少。'],
 ['资本金与留存','理论路径按需补资并反向预留后续已知资金需求；额度路径只投入额度内资本金，缺口单列，发生过缺口后停止分配。缺口路径只定位无法执行的计划。'],
 ['持有和退出','持续持有不将物业净值、存货、应收或周转自动变现；出售仅处置持有物业并扣费用和税，周转按各路径独立比例回收。退出费用为含综合交易税费口径，不重复征收销项增值税。'],
 ['收益指标','季度IRR年化为(1+r)^4－1；折现率按有效季度换算；第0期不折现。回收期取首次跨零并在季度内插值；多次变号需结合NPV与后续资金审视。'],
 ['偿债与运营保本','年度DSCR不包含退出、周转回收、贷款及股东流入；前期销售现金留存另表展示。稳定运营保本只覆盖运营支出及流转税，不代表投资回本或本息覆盖。'],
 ['敏感性与边界','原生双变量数据表随Excel重算。精确边界通过同一模型扫描及二分求解，记录参数签名和刷新时间；无可行值、超搜索上限和不适用分别说明。边界数值不参与财务计算。'],
 ['适用范围','同质业态、一次开发，最多30年；未建立多业态差异成本、多分期、租户更替、预售资金流和多笔债务。涉及这些情况须扩展对应驱动。'],
 ['已确认修正','周转回收和物业处置分别校验总额、末期金额和其他期间为零，修正V2仅核总额却标示时点通过的问题；冻结表头及标签列。'],
 ['维护','黄色单元格为输入，蓝字；公式受保护，文件无宏、无外部链接。保护仅防误改，可在Excel取消保护；新增年度超30年须同步扩展所有范围。']
];
sh['使用说明'].freezePanes.unfreeze();sh['使用说明'].getRange('A1:A60').format.columnWidth=24;sh['使用说明'].getRange('B1:L60').format.columnWidth=12;
notes.forEach(([n,txt],i)=>{const r=5+i*2;v('使用说明',`A${r}`,n);sh['使用说明'].getRange(`B${r}:L${r+1}`).merge();v('使用说明',`B${r}`,txt);sh['使用说明'].getRange(`B${r}:L${r+1}`).format={wrapText:true,verticalAlignment:'center',fill:i%2?C.pale:'#FFFFFF'};sh['使用说明'].getRange(`A${r}:A${r+1}`).merge();sh['使用说明'].getRange(`A${r}:A${r+1}`).format={font:{bold:true},fill:C.head,wrapText:true};});

// Display formatting: neutral passing checks, restrained red exceptions.
for(const s of names){
 const w=sh[s];
 for(const [range,formula]of s==='模型校验'?[[`D7:D${checks.length+6}`,'D7="未通过"']]:s==='指标与情景'?[['F6:G15','F6="未通过"']]:s==='决策看板'?[['F18:G27','F18="未通过"']]:[]){w.getRange(range).conditionalFormats.addCustom(formula,{fill:C.bad,font:{color:C.red,bold:true}});}
 if(['季度投融资','季度经营税费','季度现金流','季度计划','年度汇总','年度所得税'].includes(s)){
  const used=w.getUsedRange();
  used.format.font.name='Microsoft YaHei';
  w.getRange('C1:C300').format.columnWidth=3;
 }
}
sh['指标与情景'].getRange('A1:A205').format.columnWidth=40;sh['指标与情景'].getRange('G1:G205').format.columnWidth=18;
sh['模型校验'].getRange('A1:A50').format.columnWidth=39;sh['模型校验'].getRange('E1:E50').format.columnWidth=70;
sh['敏感性与边界'].getRange('A1:A120').format.columnWidth=45;sh['敏感性与边界'].getRange('B1:B120').format.columnWidth=28;sh['敏感性与边界'].getRange('C1:C120').format.columnWidth=17;
sh['敏感性与边界'].getRange('B68:B69').setNumberFormat(';;;');sh['敏感性与边界'].getRange('B105:B115').setNumberFormat(';;;');
sh['核心参数输入'].getRange('L1:L160').format.columnWidth=25;
sh['季度投融资'].getRange('B7:B27').setNumberFormat(NUM);sh['季度投融资'].getRange('E7:E27').setNumberFormat(NUM);
sh['模型校验'].getRange('A5:E5').format={fill:C.head,font:{bold:true},horizontalAlignment:'center'};
sh['指标与情景'].getRange('A4:G4').format={fill:C.head,font:{bold:true},horizontalAlignment:'center'};
dash.getRange('A17:G17').format={fill:C.head,font:{bold:true},horizontalAlignment:'center'};
// Native table definitions are applied and verified in the intended Excel engine.
const manifest={source,output:'产业园投资测算_V3_决策测算.xlsx',sheets:names,sourceInputs:source.parameters.filter(x=>!x.formula).map(x=>`B${x.row}`),controls,quarterInputs:[20,21,22,23].map(r=>`D${r}:${END}${r}`),dataTables,boundRows:bounds.map((b,i)=>({row:73+i,name:b[0],unit:b[1]})),O,END,Q,checks:checks.map((x,i)=>({row:7+i,name:x.name})),added};
await fs.writeFile(new URL('manifest.json',here),JSON.stringify(manifest,null,2));
console.log('formula build complete; recalculating artifact engine');
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'季度投融资',range:'A7:F33',maxChars:1600,tableMaxRows:4,tableMaxCols:6})).ndjson);
const blob=await SpreadsheetFile.exportXlsx(wb);await blob.save(decodeURIComponent(new URL('候选_V3.xlsx',here).pathname.slice(1)));
console.log('exported candidate');
for(const [sheet,range,name]of [['决策看板','A1:G40','决策看板'],['核心参数输入','A93:L107','参数'],['年度汇总','A8:K29','年度'],['模型校验','A5:E19','校验']]){
 try{const im=await wb.render({sheetName:sheet,range,scale:1.3,format:'png'});await fs.writeFile(new URL(`预览_${name}.png`,here),new Uint8Array(await im.arrayBuffer()));console.log('rendered '+sheet);}catch(e){console.log('render error '+sheet+': '+String(e));}
}
