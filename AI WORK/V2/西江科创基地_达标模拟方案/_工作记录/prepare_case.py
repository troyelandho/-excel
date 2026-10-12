from pathlib import Path
import sys, json, hashlib, math, openpyxl

ROOT = Path(__file__).resolve().parent.parent
V2 = ROOT.parent / '产业园投资测算_V2_样式调整.xlsx'
sys.path.insert(0, str(ROOT.parent.parent / 'V1' / '交付文件'))
import generate_excel as definitions  # 仅读取参数地址/行号；不调用任何生成函数
import validate_model as independent

book = openpyxl.load_workbook(V2, data_only=True)
formulas = openpyxl.load_workbook(V2, data_only=False)
sheet = book['核心参数输入']
params = {k: sheet.cell(r, 2).value for k, r in definitions.INPUT_ROWS.items() if k != 'rent_share'}
params.update(name='高明西江科创基地（模拟）', start_year=2023, horizon=20,
              land_area=45249.99, far=2.05, operable_share=.92,
              under_area=95000-45249.99*2.05, sale_share=.85,
              build_years=2, sale_start=3, rent_start=3,
              land_price=350, above_cost=1700, under_cost=2800,
              infra=1200, pre=500, other_build=550, contingency=.06,
              working_capital=200, sale_price=5300, rent=18,
              occ1=.60, occ2=.75, occ3=.85, occ_stable=.90,
              free_months=3, fixed_cost=70, maintenance=60,
              loan=9000, interest=.038, loan_term=5, repay_start=3,
              equity_cap=17000, terminal_on=0, terminal_value=0,
              stress_required=1)
params['build_plan']=[0,.45,.55]+[0]*28
params['sale_plan']=[0,0,0,.5,.3,.2]+[0]*25

sources=[
 dict(id='S1',title='佛山新闻网／佛山日报：科创产业基地交付',date='2026-01-06',
      url='https://www.foshannews.net/h/191/20260106/843690.html',
      fact='项目名称、总建筑面积近9.5万平方米、产业方向；已交付。',quality='地方主流媒体，最新规模口径'),
 dict(id='S2',title='中建四局：项目2号楼封顶',date='2025-02-01',
      url='https://4cd.cscec.com/xwzx/yxdt/202502/3853859.html',
      fact='恒良路以北、恒荷路以东；科创大楼、生产车间、人才公寓；早期总建面约9.3万㎡。',quality='施工企业一手发布，早期约数'),
 dict(id='S3',title='南方Plus：高明首个科创基地启用（转载）',date='2025-12-29',
      url='https://www.sohu.com/a/970617989_100116740',
      fact='2025-12-29正式交付；占地约67.87亩；规划10栋，2号和7号楼为反向飞地载体。',quality='主流媒体署名转载，非招商兑现证明'),
 dict(id='S4',title='施工总承包招标公告（佛山资讯网转载）',date='2024-05-06',
      url='https://www.fs0757.com/home/article/view/id/656511.html',
      fact='招标公告检索摘要列用地面积45249.99㎡，工业用地、50年。',quality='转载摘要，未取得原公告附件／测绘成果'),
 dict(id='S5',title='久久厂房网：同名项目招商信息',date='检索2026-10-12',
      url='https://fs.99cfw.com/t/rQqy.htm',
      fact='同名项目介绍列总建面约9.47万㎡、容积率约2.05；无可核验成交均价。',quality='平台挂牌检索摘要，仅弱参考，容积率仍作为模拟口径'),
 dict(id='S6',title='广东宏正：园建绿化工程采购邀请',date='2025-08-05',
      url='https://www.gdhzec.com/mobile/info_2507001.php?flv1id=85&id=3203&lvz=2',
      fact='单项园建绿化采购预算约110万元；采购人佛山市高信建设发展有限公司。',quality='采购代理一手公告；不是全项目预算或结算。后续采购终止，不作实际支出')]

notes={
 'name':'公开名称据S1；本表为开发模拟方案，非项目实际经营结果。区位：高明荷城恒良路北、恒荷路东（S2）。',
 'start_year':'模拟第0年2023，建设2024—2025，运营2026起；对应2025年底交付背景，不是实际财务历史。',
 'horizon':'20年，项目第0—20年，日历2023—2043；保持原动态范围规则。',
 'land_area':'公开信息S4：45,249.99㎡，约67.87亩；招标转载摘要，未取得实测成果。',
 'far':'模拟采用2.05（S5平台约数弱参考）；用于推算地上92,762.48㎡，非确认计容指标。',
 'operable_share':'模拟92%；扣除8%公共及非经营面积，所有共同成本计入。',
 'under_area':'模拟2,237.52㎡＝95,000－土地面积×2.05；使总建面贴近S1约数，未核实地下实测面积。',
 'sale_share':'方案目标出售85%、持有15%；未核实真实租售比例。销售物业出售前后均不计租金。',
 'build_years':'模拟2年（2024—2025）；非土地支出45%/55%，2025年底完成后次年经营。',
 'sale_start':'模拟第3年＝2026起销售；3年去化50%/30%/20%，当年全额回款，需客户及付款安排支持。',
 'rent_start':'模拟第3年＝2026出租；60%/75%/85%/90%逐年爬坡，非实际签约出租率。',
 'land_price':'模拟350元/㎡，约23.33万元/亩；未找到该地块可核验成交金额，仅土地取得价款。',
 'above_cost':'目标成本1700元/㎡：普通工业／研发主体、标准机电及基本装修；实验室、生产设备及租户专装另由租户承担。',
 'under_cost':'模拟2800元/㎡；地下土建、基础机电，不与地上或园区配套重复。',
 'infra':'模拟1200万元：道路、管网、园建、供配电等公共配套；S6单项110万元仅作量级参考，不能视为结算。',
 'pre':'模拟500万元：勘察、设计、咨询；不含建安或建设管理。',
 'other_build':'模拟550万元：建设管理、监理等；不含运营固定成本和建设利息。',
 'contingency':'模拟6%；全部非土地建设费用为基数，含建安、配套、前期、管理，保持原公式。',
 'working_capital':'模拟200万元；第3年投入，第20年回收，非经营费用和折旧。',
 'sale_price':'达标目标5300元/㎡（含税回款口径）；未发现可靠成交依据，偏积极，须高品质研发／定制客户支撑，非市场预测。',
 'sale_expense':'沿用模拟2.5%销售回款费用；含营销、佣金等，需在目标售价之外承担。',
 'sale_tax':'沿用模拟综合8%销售税费；不建立复杂税务或抵扣模型。',
 'rent':'模拟18元/㎡/月；持有部分工业／研发空间混合租金，未核实本项目实际合同。',
 'rent_growth':'沿用模拟1.5%年增长；长期平滑增长，不代表每份合同的递增条款。',
 'occ1':'模拟首个出租年60%；新增出租面积当年一次免租3个月。',
 'occ2':'模拟第2个出租年75%；不重复给原已出租面积免租。',
 'occ3':'模拟第3个出租年85%；不是已签约或已投产面积。',
 'occ_stable':'模拟第4个出租年起90%；压力下降10个百分点至80%，不模拟租户更替。',
 'free_months':'沿用模拟新租平均3个月免租；同一面积仅扣一次。',
 'collection':'沿用模拟98%收缴率；坏账、收款折扣通过此参数简化反映。',
 'other_income':'沿用模拟实收租金的5%；配套服务净收入口径，不计租户用电等全额代收。',
 'fixed_cost':'模拟70万元/年，仅持有约1.28万㎡：人员、管理、保险等；空置仍承担，售出物业运营费用由业主承担。',
 'fixed_growth':'沿用模拟2%年增长；自竣工次年起，不与出租率挂钩。',
 'variable_cost':'沿用模拟12%持有经营收入；含日常保洁、公共能耗及小修，不含年度维护更新。',
 'maintenance':'模拟60万元/年：持有及分摊公共设施的周期更新；当期费用化、税前扣除，不再资本化或折旧。',
 'operating_tax':'沿用模拟6%持有经营收入综合税费；不模拟抵扣或单独税种。',
 'income_tax':'沿用模拟25%；正利润缴税，负利润不缴，不结转亏损。',
 'dep_years':'沿用30年直线折旧；土地不折旧，项目口径排除融资利息。',
 'residual':'沿用模拟5%残值；不会自动形成处置现金收入。',
 'loan':'模拟9000万元，按45%/55%建设计划提款；贷款额度在压力情景保持不变。',
 'interest':'模拟3.8%年利率，非融资报价；利息按（期初＋提款/2－还本/2）计算。',
 'loan_term':'模拟5年，自2024首次提款起，到期2028；销售期内清偿以避免长期租金承担大额本金。',
 'repay_start':'模拟第3年2026起等额本金，每年3000万元，2028还清；不额外提前还本。',
 'equity_cap':'模拟可落实17000万元上限；基准和组合压力按需投入，未落实资金不会自动补齐。',
 'terminal_on':'关闭；达标不依赖期末出售持有物业或未售存货。',
 'terminal_value':'0；处置关闭，此值不参与现金流。',
 'terminal_fee':'沿用3%处置交易费假设；开关关闭时不产生费用。',
 'stress_required':'开启；售价－10%、租金－10%、出租率－10个百分点、非土地成本＋10%同时重算，均须通过。'}
changes={ '核心参数输入':{},'使用说明':{},'决策看板':{} }
for key,row in definitions.INPUT_ROWS.items():
    if key=='rent_share':continue
    changes['核心参数输入'][f'B{row}']=params[key]
    if key in notes:changes['核心参数输入'][f'D{row}']=notes[key]
for year in range(31):
    col=openpyxl.utils.get_column_letter(year+4)
    changes['核心参数输入'][f'{col}88']=params['build_plan'][year]
    changes['核心参数输入'][f'{col}89']=params['sale_plan'][year]
changes['核心参数输入']['A2']='核心参数输入 · 项目模拟方案'
changes['核心参数输入']['A92']='模拟计划：2024—2025建设45%/55%；2026—2028销售50%/30%/20%。出售85%，持有15%；公开资料与目标假设已在D列区分。'
changes['使用说明'].update(A2='使用说明 · 西江科创基地模拟方案',
 B5='以佛山高明西江产业新城科创产业基地的公开规模为参考，模拟2023拿地、2024—2025建设、2026起租售。实际项目已于2025年底交付。土地与总建面有公开资料；地下拆分、成本、售价、租金及融资均含模拟假设，不能视为实际财务结果。',
 B69='保持原压力幅度：售价－10%、租金－10%、出租率－10个百分点、非土地建设成本＋10%；本方案开启压力须通过。所有判断门槛沿用V2，期末处置关闭。基准和压力均按原公式重算。',
 B81='本副本直接更新V2的黄色参数及相关说明，保留原公式和样式。Excel实际重算后保存缓存，日常修改参数仍自动联动。目标售价5300元/㎡须有交易支持，未核实已实现。资料编号及链接见“核心参数输入”第95行起；详细假设和验证见附带方案说明。')
changes['决策看板']['A2']='决策看板 · 西江科创基地模拟方案'
changes['核心参数输入']['A95']='公开资料来源（检索日期：2026-10-12）'
for col,val in zip('ABCD',['编号','来源名称','发布日期','链接及用途']):changes['核心参数输入'][f'{col}96']=val
for row,s in enumerate(sources,97):
    short=['佛山新闻网','中建四局','南方Plus转载','招标公告转载','久久厂房网','广东宏正'][row-97]
    for col,val in zip('ABCD',[s['id'],short,s['date'],s['url']+'\n'+s['fact']+' '+s['quality']]):
        changes['核心参数输入'][f'{col}{row}']=val
changes['核心参数输入'].update(A104='资料口径与达标前提',D104='S1总建面近9.5万㎡、S2约9.3万㎡、S5约9.47万㎡均为公开约数；本方案采用95,000㎡量级，地上／地下拆分未经证实。')
changes['核心参数输入'].update(A105='目标成交条件',D105='5300元/㎡为偏积极的达标目标，三年全部回款为计划，不是成交事实。需高品质研发／定制客户支撑；若做普通厂房且价差不足，项目可能不通过。')
changes['核心参数输入'].update(A106='成本控制条件',D106='1700元/㎡地上建安为控制目标，项目总投资约2.19亿元；不含租户实验室、工艺设备和专用精装修。实际若由园区承担，应提高建安及配套参数。')
changes['核心参数输入'].update(A107='持有与偿债条件',D107='持有部分占15%，租金仅辅助整体回报；贷款依靠2026—2028销售现金流还清。稳定期NOI收益率较低，不应将整体达标理解为出租业务高收益。')

cases=[dict(name='交付方案',overrides={}),
 dict(name='售价下降10%',overrides={'sale_price':4770}),
 dict(name='租金下降10%',overrides={'rent':16.2}),
 dict(name='出租率下降10个百分点',overrides={'occ1':.50,'occ2':.65,'occ3':.75,'occ_stable':.80}),
 dict(name='非土地成本增加10%',overrides={'above_cost':1870,'under_cost':3080,'infra':1320,'pre':550,'other_build':605}),
 dict(name='利率提高2个百分点',overrides={'interest':.058}),
 dict(name='资本金额度不足',overrides={'equity_cap':10000}),
 dict(name='售价跌至4000元',overrides={'sale_price':4000}),
 dict(name='期限12年',overrides={'horizon':12}),
 dict(name='期限30年',overrides={'horizon':30})]
def native(x):
    if hasattr(x,'tolist'):return native(x.tolist())
    if isinstance(x,float) and not math.isfinite(x):return None
    if isinstance(x,dict):return {k:native(v) for k,v in x.items()}
    if isinstance(x,(tuple,list)):return [native(v) for v in x]
    return x
predictions=[dict(name=c['name'],base=native(independent.oracle({**params,**c['overrides']})),
                  stress=native(independent.oracle({**params,**c['overrides']},stress=True))) for c in cases]
manifest=dict(source=str(V2),source_sha256=hashlib.sha256(V2.read_bytes()).hexdigest(),
              root=str(ROOT),changes=changes,params=params,sources=sources,cases=cases,
              input_rows=definitions.INPUT_ROWS,
              maps=dict(plan=definitions.PLAN,investment_annual={k:definitions.INV[k] for k in ['land_cash','nonland_cash','build_cash','wc_in','wc_out','project_invest_cash','finance_cost_cash','total_cash']},
                        income=definitions.INC,costs=definitions.COST,financing=definitions.FIN,cashflow=definitions.CF,
                        metrics=definitions.METRIC_ROWS,investment=definitions.INV),
              predictions=predictions)
(ROOT/'_工作记录'/'case_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(dict(source_sha256=manifest['source_sha256'],base=predictions[0]['base']['metrics'],stress=predictions[0]['stress']['metrics']),ensure_ascii=False,indent=2))
