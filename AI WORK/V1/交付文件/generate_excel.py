#!/usr/bin/env python3
"""Portable annual industrial-park model. All editable assumptions live in Excel.

Python 3.12+, XlsxWriter. No proprietary runtime, macros, external links or data.
Use --recalculate for an isolated Microsoft Excel COM recalculation on Windows.
"""
from __future__ import annotations
import argparse
import json
import subprocess
from datetime import datetime
from pathlib import Path
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

SHEETS = ['使用说明','核心参数输入','建筑面积与开发计划','投资与资金安排',
          '销售与出租收入','运营成本与税费','融资与还本付息','利润与现金流',
          '指标与情景测算','决策看板']
MAX_YEAR = 30
OFFSET = 80  # Two visible calculation blocks: base and simultaneous downside.
COLORS = dict(ink='#243746', navy='#203E56', blue='#1256A0', yellow='#FFF2CC',
              pale='#EAF0F5', green='#DDEEDC', red='#FBE0E0', gray='#E7E7E7',
              teal='#2B7C82', orange='#CC8740')
PARAMETERS = [
 ('基础与建筑参数', [
 ('name','项目名称','模拟产业园 · 租售结合','文本','全部数据为虚构模拟数据'),
 ('start_year','测算起始年份',2026,'年','第0年的日历年份；年度支出采用年末时点'),
 ('horizon','测算期限',20,'年','实际范围第0年至第N年，预留至30年；须覆盖贷款到期'),
 ('land_area','土地面积',100000,'平方米','地上面积＝土地面积×容积率'),
 ('far','容积率',1.5,'倍','地上总建筑面积口径'),
 ('operable_share','地上可经营面积占比',0.90,'%','其余为公共及非经营空间；共同成本仍计入'),
 ('under_area','地下建筑面积',15000,'平方米','地下空间不独立产生收入'),
 ('sale_share','可经营面积销售占比',0.55,'%','严格大于0且小于100%；销售部分出售前也不出租'),
 ('rent_share','出租占比（自动）',None,'%','自动＝1－销售占比'),
 ('build_years','建设年限',3,'年','第1年至第K年；土地在第0年支付'),
 ('sale_start','销售开始年份',4,'项目第几年','必须晚于建设完成年；年度销售比例在下方输入'),
 ('rent_start','出租开始年份',4,'项目第几年','必须晚于建设完成年；运营固定成本从K+1年开始')]),
 ('投资参数', [
 ('land_price','土地单价',800,'元/平方米','仅土地取得价款；不计入建筑折旧'),
 ('above_cost','地上单方建安成本',2400,'元/平方米','不含下列配套、前期、管理费用'),
 ('under_cost','地下单方建安成本',3000,'元/平方米','地下工程；不与地上重复'),
 ('infra','园区配套工程费用',4500,'万元','园区道路、管网、公共配套；不含建安主体'),
 ('pre','勘察设计咨询前期费用',1800,'万元','不含建安及管理费用'),
 ('other_build','管理及其他建设费用',2200,'万元','建设管理等；不含融资利息和运营成本'),
 ('contingency','预备费比例',0.08,'%','基数＝全部非土地建设费用（建安、配套、前期、管理）'),
 ('working_capital','初始运营周转资金',500,'万元','K+1年投入，期末全额回收；不计入费用及折旧')]),
 ('销售参数', [
 ('sale_price','销售均价',6500,'元/平方米','简化含税成交/实收口径；销售当年全额收款'),
 ('sale_expense','销售费用率',0.025,'%','销售回款×费用率'),
 ('sale_tax','简化销售相关税费率',0.08,'%','含税回款×综合率；不另行建立增值税抵扣')]),
 ('出租参数', [
 ('rent','初始月租金',26,'元/平方米/月','出租首年价格；按12个月换算后除10000为万元'),
 ('rent_growth','年租金增长率',0.015,'%','从出租开始年递增'),
 ('occ1','出租第1年出租率',0.50,'%','全年占用面积比例；新增面积仅当年一次免租'),
 ('occ2','出租第2年出租率',0.70,'%','前期爬坡参数集中输入'),
 ('occ3','出租第3年出租率',0.85,'%','出租率须不下降；不模拟租户更替'),
 ('occ_stable','第4年起稳定出租率',0.90,'%','稳定出租率直接驱动第4个出租年度以后'),
 ('free_months','新增出租面积平均免租月数',3,'月','0—12月；同一面积不重复免租'),
 ('collection','租金收缴率',0.98,'%','实收租金，同时作为简化收入确认口径'),
 ('other_income','其他经营收入比例',0.05,'%','实收租金×比例；属于持有经营收入')]),
 ('运营与税费参数', [
 ('fixed_cost','运营首年固定成本',300,'万元/年','仅持有部分：人员、管理、保险等；K+1年即承担'),
 ('fixed_growth','固定成本年度增长率',0.02,'%','自K+1年起递增，与是否出租无关'),
 ('variable_cost','变动运营成本率',0.12,'%','实收租金及其他经营收入×比例；不含维护更新'),
 ('maintenance','年度维护更新支出',180,'万元/年','K+1年起每年发生；费用化并当期税前扣除，不再折旧'),
 ('operating_tax','简化运营相关税费率',0.06,'%','租金及其他经营收入×综合率；不另计复杂税种'),
 ('income_tax','企业所得税率',0.25,'%','年度正利润缴税；负利润不缴，不模拟亏损结转'),
 ('dep_years','持有建筑物折旧年限',30,'年','K+1年起直线法；项目口径不含融资利息，资本金口径含资本化利息'),
 ('residual','建筑物残值率',0.05,'%','仅影响账面价值，不能自动变成现金收入')]),
 ('融资与资本金参数', [
 ('loan','贷款金额',30000,'万元','一笔贷款；按年度建设投资比例提款，不能超过非土地建设支出'),
 ('interest','贷款年利率',0.042,'%','利息＝（期初本金＋当年提款/2－当年还本/2）×利率'),
 ('loan_term','贷款期限',10,'年','从首次提款年算起：到期年＝首次提款年＋期限－1'),
 ('repay_start','开始偿还本金年份',4,'项目第几年','晚于建设完成年；到期前等额本金；不主动额外提前还款'),
 ('equity_cap','可投入资本金总额度',33000,'万元','按需投入且不突破额度；缺口单列，不作为真实资金')]),
 ('期末处置参数', [
 ('terminal_on','是否计入期末处置',0,'0关闭/1开启','默认关闭；仅处置持有物业，未售存货不自动变现'),
 ('terminal_value','持有物业期末处置价值',30000,'万元','直接给定第N年的持有物业含税总售价，非账面值'),
 ('terminal_fee','期末处置费用率',0.03,'%','处置价×综合交易费用率（含简化交易税费）；处置净收益另纳入所得税')]),
 ('决策阈值（模拟判断标准，可自行修改）', [
 ('project_irr_min','项目税后IRR最低值',0.08,'%','模拟阈值，非实际投资标准'),
 ('equity_irr_min','资本金IRR最低值',0.10,'%','资金不足时仅为完整出资的理论结果'),
 ('project_discount','项目折现率',0.08,'%','项目NPV及动态回收期折现率'),
 ('equity_discount','资本金折现率',0.10,'%','资本金NPV折现率'),
 ('static_limit','静态投资回收期上限',15,'年','含第0年及建设期；项目投资现金流口径'),
 ('dynamic_limit','动态投资回收期上限',20,'年','按项目折现率；测算期内未回收视为不达标'),
 ('dscr_min','最低偿债覆盖倍数',1.20,'倍','仅统计建成以后存在本息偿付的年度')]),
 ('组合压力情景参数', [
 ('stress_sale','销售价格下降比例',0.10,'%','基准售价×(1－下降比例)'),
 ('stress_rent','租金下降比例',0.10,'%','基准各年租金×(1－下降比例)'),
 ('stress_occ','出租率下降百分点',10,'百分点','10表示下降10个百分点，按10/100扣减，不是乘90%'),
 ('stress_cost','非土地建设成本增加比例',0.10,'%','作用于建安、配套、前期、管理及对应预备费；土地不变'),
 ('stress_required','是否要求压力情景也通过',0,'0否/1是','默认仅观察；开启后两种情景都必须通过')])]
DEFAULTS = {k: v for _,items in PARAMETERS for k,_,v,_,_ in items if v is not None}
INPUT_ROWS = {}
_row = 6
for _group,_items in PARAMETERS:
    _row += 1
    for _k,*_rest in _items:
        INPUT_ROWS[_k] = _row
        _row += 1
    _row += 1
PLAN_HEADER = _row + 3
BUILD_PLAN_ROW = PLAN_HEADER + 1
SALE_PLAN_ROW = PLAN_HEADER + 2
DEFAULT_BUILD_PLAN = [0,0.30,0.40,0.30] + [0]*27
DEFAULT_SALE_PLAN = [0,0,0,0,0.30,0.40,0.30] + [0]*24

INV = dict(land=6,above=7,under=8,infra=9,pre=10,other=11,pre_cont=12,
           contingency=13,nonland=14,cap_interest=15,development=16,wc=17,total=18,
           sale_basis_project=19,dep_basis_project=20,held_land=21,
           sale_basis_equity=22,dep_basis_equity=23,held_invest=24,
           land_cash=28,nonland_cash=29,build_cash=30,wc_in=31,wc_out=32,
           project_invest_cash=33,finance_cost_cash=34,total_cash=35)
AREA = dict(land=6,above=7,operable=8,sale=9,held=10,under=11)
PLAN = dict(build=20,sale=21,cum_sale=22,sold=23,unsold=24,rent_age=25,
            occ=26,occupied=27,new=28,complete=29)
INC = dict(sale_ratio=10,sale_area=11,sale_price=12,sales=13,occ=14,occupied=15,
           new=16,rent_age=17,rent_price=18,free_loss=19,rent_cash=20,other=21,
           held_income=22,revenue=23,terminal_gross=24,terminal_fee=25,terminal_net=26)
COST = dict(sale_expense=10,sale_tax=11,fixed=12,variable=13,maintenance=14,
            operating_tax=15,cash_total=16,sale_cost_project=17,sale_cost_equity=18,
            dep_project=19,dep_equity=20,book_project=21,book_equity=22,
            inventory=23,gain_project=24,gain_equity=25)
FIN = dict(first_draw=6,maturity=7,repay_count=8,opening=10,draw=11,draw_flag=12,
           principal=13,interest=14,closing=15,cap_interest=16,op_interest=17,
           service=18,maturity_balance=19)
CF = dict(revenue=10,cash_cost=11,ebit_project=12,ebit_equity=13,op_interest=14,
          pretax=15,tax_project=16,tax_equity=17,profit=18,cfads=19,op_project=20,
          land=21,build=22,wc_in=23,wc_out=24,terminal=25,project=26,draw=27,
          interest=28,principal=29,tax_bridge=30,before_equity=31,reserve=32,
          theory_open=33,required_equity=34,theory_distribution=35,theory_close=36,
          cum_required=37,equity=38,actual_open=39,gap_open=40,actual_need=41,
          actual_equity=42,actual_distribution=43,actual_close=44,gap=45,
          max_gap=46,cum_actual=47,actual_equity_cf=48,cum_project=49,
          discounted=50,cum_discounted=51,dscr=52,dscr_candidate=53,
          static_cross=54,dynamic_cross=55,sign_project=56,last_sign_project=57,
          changes_project=58,theory_balance=59,actual_balance=60,reserve_check=61,
          sign_equity=62,last_sign_equity=63,changes_equity=64,
          sale_cash_contribution=65,held_cash_contribution=66)
METRIC_ROWS = dict(project_irr=6,equity_irr=7,project_npv=8,equity_npv=9,
                   static_payback=10,dynamic_payback=11,dscr=12,required_equity=13,
                   max_gap=14,maturity_balance=15,total_investment=16,loan=17,
                   equity_cap=18,sale_through=19,sale_complete=20,stable_occ=21,
                   breakeven=22,noi_yield=23,terminal_share=24,
                   project_sign_warning=25,equity_sign_warning=26,equity_label=27,
                   stable_year=28)
CONDITIONS = [
 ('project_irr','项目税后IRR','project_irr_min','ge'),
 ('equity_irr','资本金IRR','equity_irr_min','ge'),
 ('project_npv','项目NPV',None,'nonnegative'),
 ('equity_npv','资本金NPV',None,'nonnegative'),
 ('static_payback','静态回收期','static_limit','le'),
 ('dynamic_payback','动态回收期','dynamic_limit','le'),
 ('dscr','最低偿债覆盖倍数','dscr_min','dscr'),
 ('required_equity','所需累计资本金','equity_cap','le_money'),
 ('max_gap','未落实资金缺口最大值',None,'zero'),
 ('maturity_balance','贷款到期未偿本金',None,'zero')]

def col(year): return xl_col_to_name(3 + year)
def q(sheet, addr): return "'%s'!%s" % (sheet, addr)
def p(key): return q(SHEETS[1], '$B$%d' % INPUT_ROWS[key])
def at(sheet, row, year, off=0): return q(SHEETS[sheet], f'{col(year)}{row+off}')
def scalar(sheet, row, off=0): return q(SHEETS[sheet],f'$B${row+off}')
def series(sheet,row,off=0): return q(SHEETS[sheet],f'$D${row+off}:$AH${row+off}')
def active_series(sheet,row,off=0,start=0):
    return q(SHEETS[sheet],f'${col(start)}${row+off}')+':INDEX('+series(sheet,row,off)+','+p('horizon')+'+1)'

class Builder:
    def __init__(self, path):
        self.path = path
        self.wb = xlsxwriter.Workbook(str(path))
        self.wb.set_properties(dict(title='产业园投资测算（模拟数据）',author='模拟学习模型',
                                    comments='Annual mixed-sale-and-rental financial model'))
        self.wb.set_calc_mode('auto')
        self.ws = [self.wb.add_worksheet(n) for n in SHEETS]
        common = dict(font_name='Microsoft YaHei',font_size=10,font_color=COLORS['ink'],valign='vcenter')
        self.fm = {}
        def fmt(key, **kwargs): self.fm[key] = self.wb.add_format({**common,**kwargs})
        fmt('text'); fmt('wrap',text_wrap=True); fmt('title',bold=True,font_size=16)
        fmt('section',bold=True,bg_color=COLORS['pale'])
        fmt('header',bold=True,bg_color=COLORS['navy'],font_color='white',align='center')
        fmt('number',num_format='#,##0.00;[Red](#,##0.00);"—"')
        fmt('percent',num_format='0.00%',italic=True)
        fmt('integer',num_format='0'); fmt('input_number',bg_color=COLORS['yellow'],font_color=COLORS['blue'],num_format='#,##0.00')
        fmt('input_percent',bg_color=COLORS['yellow'],font_color=COLORS['blue'],num_format='0.00%')
        fmt('input_integer',bg_color=COLORS['yellow'],font_color=COLORS['blue'],num_format='0')
        fmt('input_text',bg_color=COLORS['yellow'],font_color=COLORS['blue'])
        fmt('result',bold=True,bg_color=COLORS['pale'],num_format='#,##0.00;[Red](#,##0.00);"—"')
        fmt('good',bg_color=COLORS['green'],font_color='#236332')
        fmt('bad',bg_color=COLORS['red'],font_color='#A72C2C')
        fmt('na',bg_color=COLORS['gray'],font_color='#6F6F6F')
        for i,s in enumerate(self.ws):
            s.hide_gridlines(2); s.set_default_row(21)
            s.set_column('A:A',48,self.fm['text']); s.set_column('B:B',17,self.fm['number'])
            s.set_column('C:C',12,self.fm['text']); s.set_column('D:AH',15,self.fm['number'])
            s.set_landscape(); s.set_paper(9); s.fit_to_pages(1,0)
            s.set_margins(0.25,0.25,0.4,0.4)
            s.set_footer('&L模拟数据 · 学习用途&C&P / &N')
            if i>0:s.freeze_panes(5,3)
            s.write('A2',SHEETS[i]+' · 模拟数据',self.fm['title'])
            s.set_row(1,29)
        self.ws[0].set_tab_color('#82929E'); self.ws[1].set_tab_color('#D3A53E')
        self.ws[-1].set_tab_color(COLORS['navy'])

    def f(self,si,address,formula,kind='number'):
        self.ws[si].write_formula(address,formula if formula.startswith('=') else '='+formula,self.fm[kind])
    def label(self,si,row,label,unit='',note=''):
        s=self.ws[si];s.write(row-1,0,label,self.fm['text']);s.write(row-1,1,unit,self.fm['text'])
        if note:s.write(row-1,2,note,self.fm['text'])
    def yearly(self,si,row,y,formula,off=0,kind='number',mask=True):
        if mask: formula=f'IF({col(y)}$4<={p("horizon")},{formula},0)'
        self.f(si,f'{col(y)}{row+off}',formula,kind)
    def headers(self,si):
        s=self.ws[si]
        for y in range(31):
            self.f(si,f'{col(y)}4',q(SHEETS[1],f'{col(y)}{PLAN_HEADER}'),'header')
            self.f(si,f'{col(y)}5',f'{p("start_year")}+{col(y)}4','integer')
        s.write('A4','项目年度 / 日历年份',self.fm['header'])
        s.repeat_rows(3,4)
    def band(self,si,row,text):
        s=self.ws[si]; s.merge_range(row-1,0,row-1,33,text,self.fm['section']);s.set_row(row-1,25)
    def status_cf(self,si,rng):
        for term,fmt in [('通过','good'),('未通过','bad'),('不适用','na'),('待判断','bad')]:
            self.ws[si].conditional_format(rng,dict(type='text',criteria='begins with',value=term,format=self.fm[fmt]))

    def inputs(self):
        s=self.ws[1]; s.set_column('A:A',46);s.set_column('B:B',22);s.set_column('C:C',20);s.set_column('D:AH',15)
        for j,text in enumerate(['参数名称','输入值','单位']):s.write(4,j,text,self.fm['header'])
        s.merge_range('D5:K5','说明',self.fm['header'])
        row=6
        for group,items in PARAMETERS:
            s.merge_range(row-1,0,row-1,10,group,self.fm['section']); row+=1
            for key,name,value,unit,note in items:
                s.write(row-1,0,name);s.write(row-1,2,unit);s.merge_range(row-1,3,row-1,10,note,self.fm['wrap']);s.set_row(row-1,30)
                if value is None:self.f(1,f'B{row}',f'1-{p("sale_share")}','percent')
                else:
                    kind='text' if isinstance(value,str) else 'percent' if unit=='%' else 'integer' if key in ['start_year','horizon','build_years','sale_start','rent_start','dep_years','loan_term','repay_start','terminal_on','stress_required','free_months'] else 'number'
                    s.write(row-1,1,value,self.fm['input_'+kind])
                    if not isinstance(value,str):
                        lo,hi=0,1e9
                        dtype='decimal'
                        if unit=='%':hi=1
                        if key in ['rent_growth','fixed_growth']:lo=-0.5
                        if key=='stress_occ':hi=100
                        if key in ['horizon','build_years','sale_start','rent_start','loan_term','repay_start','dep_years','start_year','terminal_on','stress_required','free_months']:
                            dtype='integer';lo=1;hi=30
                            if key=='dep_years':hi=100
                            if key=='start_year':lo=1900;hi=2200
                            if key in ['terminal_on','stress_required']:lo=0;hi=1
                            if key=='free_months':lo=0;hi=12
                        if key=='sale_share':lo=0.001;hi=0.999
                        if key=='far':lo=0.01;hi=20
                        s.data_validation(row-1,1,row-1,1,dict(validate=dtype,criteria='between',minimum=lo,maximum=hi,
                           ignore_blank=False,error_type='stop',error_title='请输入有效模拟参数',error_message='请遵守数值范围；粘贴后仍需检查看板校验。'))
                row+=1
            row+=1
        s.merge_range(PLAN_HEADER-3,0,PLAN_HEADER-3,7,'年度计划：建设及销售比例均以项目总量为基数；年度索引为第0年至第30年',self.fm['section'])
        for y in range(31):s.write(PLAN_HEADER-1,3+y,y,self.fm['header'])
        for r,label,values in [(BUILD_PLAN_ROW,'年度非土地建设投资比例',DEFAULT_BUILD_PLAN),(SALE_PLAN_ROW,'年度销售比例',DEFAULT_SALE_PLAN)]:
            s.write(r-1,0,label);s.write(r-1,1,'合计见看板');s.write(r-1,2,'%')
            for y,v in enumerate(values):s.write(r-1,3+y,v,self.fm['input_percent'])
            s.data_validation(r-1,3,r-1,33,dict(validate='decimal',criteria='between',minimum=0,maximum=1,ignore_blank=False))
        s.merge_range(SALE_PLAN_ROW+2,0,SALE_PLAN_ROW+2,9,'提示：调整建设/销售时点时，请同步调整黄色年度计划；出租率由出租模块的首3年及稳定期参数自动形成。',self.fm['wrap'])
        s.set_row(SALE_PLAN_ROW+2,35);s.freeze_panes(5,3);s.print_area(0,0,SALE_PLAN_ROW+3,33)

    def area_plan(self):
        self.headers(2)
        labels={'land':'土地面积','above':'地上建筑面积','operable':'地上可经营面积','sale':'总可售面积','held':'总持有出租面积','under':'地下建筑面积'}
        forms={'land':p('land_area'),'above':f'{p("land_area")}*{p("far")}',
               'operable':f'$B$7*{p("operable_share")}','sale':f'$B$8*{p("sale_share")}',
               'held':f'$B$8*{p("rent_share")}','under':p('under_area')}
        for k,r in AREA.items():self.ws[2].write(f'A{r}',labels[k]);self.f(2,f'B{r}',forms[k]);self.ws[2].write(f'C{r}','平方米')
        for off in (0,OFFSET):
            self.band(2,18+off,'基准情景 · 年度开发与租售计划' if not off else '组合压力情景 · 年度开发与租售计划（共用面积及销售比例）')
            for k,name,unit in [('build','建设分配比例','%'),('sale','销售比例','%'),('cum_sale','累计销售去化率','%'),('sold','当年出售面积','平方米'),('unsold','期末未售面积','平方米'),('rent_age','出租第几个年度','年'),('occ','实际出租率','%'),('occupied','已出租面积','平方米'),('new','新增出租面积','平方米'),('complete','累计销售完成标记','0/1')]:self.label(2,PLAN[k]+off,name,unit)
            for y in range(31):
                c=col(y);prev=col(y-1)
                forms={
                  'build':f'IF(AND({c}$4>=1,{c}$4<={p("build_years")}),{q(SHEETS[1],c+str(BUILD_PLAN_ROW))},0)',
                  'sale':f'IF({c}$4>={p("sale_start")},{q(SHEETS[1],c+str(SALE_PLAN_ROW))},0)',
                  'cum_sale':f'SUM($D{PLAN["sale"]+off}:{c}{PLAN["sale"]+off})',
                  'sold':f'{c}{PLAN["sale"]+off}*$B$9',
                  'unsold':f'$B$9*(1-{c}{PLAN["cum_sale"]+off})',
                  'rent_age':f'IF({c}$4>={p("rent_start")},{c}$4-{p("rent_start")}+1,0)',
                  'occ':f'IF({c}{PLAN["rent_age"]+off}=0,0,MAX(0,IF({c}{PLAN["rent_age"]+off}=1,{p("occ1")},IF({c}{PLAN["rent_age"]+off}=2,{p("occ2")},IF({c}{PLAN["rent_age"]+off}=3,{p("occ3")},{p("occ_stable")})))-'+(f'{p("stress_occ")}/100' if off else '0')+'))',
                  'occupied':f'$B$10*{c}{PLAN["occ"]+off}',
                  'new':f'MAX(0,{c}{PLAN["occupied"]+off}-'+(f'{prev}{PLAN["occupied"]+off}' if y else '0')+')',
                  'complete':f'IF({c}{PLAN["cum_sale"]+off}>=0.999999,1,0)'}
                for k,r in PLAN.items():self.yearly(2,r,y,forms[k],off,'percent' if k in ['build','sale','cum_sale','occ'] else 'integer' if k=='rent_age' else 'number')
        self.ws[2].print_area(0,0,OFFSET+30,33)

    def investment(self):
        self.headers(3)
        names=['土地取得费用','地上建安费用','地下建安费用','配套工程费用','前期费用','管理及其他建设费用','非土地建设费用小计','预备费','非土地建设费用含预备费','建设期融资费用（资本金支付）','开发总成本含融资费用','初始运营周转资金','项目总投资含建设期融资费用','销售成本基数（项目口径）','持有折旧基数（项目口径）','持有分摊土地成本','销售成本基数（融资口径）','持有折旧基数（融资口径）','持有分摊总投资含周转资金']
        for off in (0,OFFSET):
            self.ws[3].write(f'A{3+off}','基准情景' if not off else '组合压力情景',self.fm['section'])
            factor=f'(1+{p("stress_cost")})' if off else '1'
            forms={
             'land':f'{p("land_area")}*{p("land_price")}/10000',
             'above':f'{scalar(2,AREA["above"])}*{p("above_cost")}/10000*{factor}',
             'under':f'{p("under_area")}*{p("under_cost")}/10000*{factor}',
             'infra':f'{p("infra")}*{factor}','pre':f'{p("pre")}*{factor}','other':f'{p("other_build")}*{factor}',
             'pre_cont':f'SUM(B{7+off}:B{11+off})','contingency':f'B{12+off}*{p("contingency")}',
             'nonland':f'SUM(B{12+off}:B{13+off})','cap_interest':f'SUM({series(6,FIN["cap_interest"],off)})',
             'development':f'B{6+off}+B{14+off}+B{15+off}','wc':p('working_capital'),
             'total':f'SUM(B{16+off}:B{17+off})','sale_basis_project':f'(B{6+off}+B{14+off})*{p("sale_share")}',
             'dep_basis_project':f'B{14+off}*{p("rent_share")}',
             'held_land':f'B{6+off}*{p("rent_share")}',
             'sale_basis_equity':f'B{16+off}*{p("sale_share")}',
             'dep_basis_equity':f'(B{14+off}+B{15+off})*{p("rent_share")}',
             'held_invest':f'B{16+off}*{p("rent_share")}+B{17+off}'}
            for (k,r),name in zip(list(INV.items())[:19],names):
                self.ws[3].write(f'A{r+off}',name);self.f(3,f'B{r+off}',forms[k],'result' if k in ['total','nonland','development'] else 'number');self.ws[3].write(f'C{r+off}','万元')
            self.band(3,26+off,'年度现金支出（成本结转与折旧不在此再次支付）')
            for k,name in [('land_cash','土地款'),('nonland_cash','非土地建设现金支出'),('build_cash','年度建设支出合计'),('wc_in','运营周转资金投入'),('wc_out','运营周转资金收回'),('project_invest_cash','项目口径年度投资支出'),('finance_cost_cash','建设期利息现金支付'),('total_cash','含融资费用年度投资支出')]:self.label(3,INV[k]+off,name,'万元')
            for y in range(31):
                c=col(y)
                forms={'land_cash':f'IF({c}$4=0,$B${6+off},0)',
                 'nonland_cash':f'$B${14+off}*{at(2,PLAN["build"],y,off)}',
                 'build_cash':f'SUM({c}{28+off}:{c}{29+off})',
                 'wc_in':f'IF({c}$4={p("build_years")}+1,{p("working_capital")},0)',
                 'wc_out':f'IF(AND({c}$4={p("horizon")},{p("horizon")}>{p("build_years")}),{p("working_capital")},0)',
                 'project_invest_cash':f'{c}{30+off}+{c}{31+off}',
                 'finance_cost_cash':at(6,FIN['cap_interest'],y,off),
                 'total_cash':f'{c}{33+off}+{c}{34+off}'}
                for k in list(INV)[19:]:self.yearly(3,INV[k],y,forms[k],off)
        self.ws[3].print_area(0,0,OFFSET+36,33)

    def financing(self):
        self.headers(6)
        for off in (0,OFFSET):
            self.ws[6].write(f'A{3+off}','基准情景 · 单笔贷款' if not off else '组合压力情景 · 贷款额度不增加',self.fm['section'])
            for r,name in [(6,'首次提款年'),(7,'贷款到期年'),(8,'等额本金还款年度数')]:self.ws[6].write(f'A{r+off}',name)
            self.f(6,f'B{6+off}',f'IF({p("loan")}=0,0,IFERROR(MATCH(1,{series(6,12,off)},0)-1,0))','integer')
            self.f(6,f'B{7+off}',f'IF({p("loan")}=0,0,B{6+off}+{p("loan_term")}-1)','integer')
            self.f(6,f'B{8+off}',f'IF({p("loan")}=0,0,MAX(1,B{7+off}-{p("repay_start")}+1))','integer')
            for k,name in [('opening','期初贷款本金'),('draw','当年贷款提款'),('draw_flag','提款年度标记'),('principal','当年还本'),('interest','当年利息'),('closing','期末贷款本金'),('cap_interest','其中：建设期利息'),('op_interest','其中：建成后费用化利息'),('service','当年本息合计'),('maturity_balance','到期未还本金（到期年标记）')]:self.label(6,FIN[k]+off,name,'万元' if k!='draw_flag' else '0/1')
            for y in range(31):
                c=col(y);prev=col(y-1)
                forms={
                 'opening':f'{prev}{15+off}' if y else '0',
                 'draw':f'{p("loan")}*{at(2,PLAN["build"],y,off)}',
                 'draw_flag':f'IF({c}{11+off}>0.000001,1,0)',
                 'principal':f'IF(AND({p("loan")}>0,{c}$4>={p("repay_start")},{c}$4<=$B${7+off}),MIN({c}{10+off}+{c}{11+off},{p("loan")}/MAX(1,$B${8+off})),0)',
                 'interest':f'MAX(0,{c}{10+off}+{c}{11+off}/2-{c}{13+off}/2)*{p("interest")}',
                 'closing':f'MAX(0,{c}{10+off}+{c}{11+off}-{c}{13+off})',
                 'cap_interest':f'IF({c}$4<={p("build_years")},{c}{14+off},0)',
                 'op_interest':f'{c}{14+off}-{c}{16+off}',
                 'service':f'{c}{13+off}+{c}{14+off}',
                 'maturity_balance':f'IF({c}$4=$B${7+off},{c}{15+off},0)'}
                for k in list(FIN)[3:]:self.yearly(6,FIN[k],y,forms[k],off)
        self.ws[6].print_area(0,0,OFFSET+20,33)

    def income(self):
        self.headers(4)
        labels=['年度销售比例','当年销售面积','销售价格','销售回款','实际出租率','已出租面积','新增出租面积','出租第几个年度','月租金','新增面积一次免租损失','实收租金','其他经营收入','持有经营收入合计','营业收入合计','期末持有物业处置价','处置交易费用含简税','期末处置净收入（所得税前）']
        for off in (0,OFFSET):
            self.band(4,7+off,'基准情景 · 销售与持有分别核算' if not off else '组合压力情景 · 同时降低售价、租金和出租率')
            for (k,r),name in zip(INC.items(),labels):self.label(4,r+off,name,'%' if k in ['sale_ratio','occ'] else '平方米' if k in ['sale_area','occupied','new'] else '元/平方米' if k=='sale_price' else '元/平方米/月' if k=='rent_price' else '年' if k=='rent_age' else '万元')
            for y in range(31):
                c=col(y)
                forms={
                 'sale_ratio':at(2,PLAN['sale'],y,off),'sale_area':at(2,PLAN['sold'],y,off),
                 'sale_price':p('sale_price')+(f'*(1-{p("stress_sale")})' if off else ''),
                 'sales':f'{c}{11+off}*{c}{12+off}/10000',
                 'occ':at(2,PLAN['occ'],y,off),'occupied':at(2,PLAN['occupied'],y,off),
                 'new':at(2,PLAN['new'],y,off),'rent_age':at(2,PLAN['rent_age'],y,off),
                 'rent_price':f'IF({c}{17+off}>0,{p("rent")}*(1+{p("rent_growth")})^MAX(0,{c}{17+off}-1)'+(f'*(1-{p("stress_rent")})' if off else '')+',0)',
                 'free_loss':f'{c}{16+off}*{c}{18+off}*{p("free_months")}/10000',
                 'rent_cash':f'({c}{15+off}*{c}{18+off}*12/10000-{c}{19+off})*{p("collection")}',
                 'other':f'{c}{20+off}*{p("other_income")}','held_income':f'SUM({c}{20+off}:{c}{21+off})',
                 'revenue':f'{c}{13+off}+{c}{22+off}',
                 'terminal_gross':f'IF(AND({p("terminal_on")}=1,{c}$4={p("horizon")}),{p("terminal_value")},0)',
                 'terminal_fee':f'{c}{24+off}*{p("terminal_fee")}',
                 'terminal_net':f'{c}{24+off}-{c}{25+off}'}
                for k,r in INC.items():self.yearly(4,r,y,forms[k],off,'percent' if k in ['sale_ratio','occ'] else 'number')
        self.ws[4].print_area(0,0,OFFSET+27,33)

    def costs(self):
        self.headers(5)
        labels=['销售费用现金支出','销售相关税费现金支出','持有固定运营成本','持有变动运营成本','维护更新支出（当期费用化）','运营相关税费','现金成本及经营税费合计','销售成本结转（项目口径）','销售成本结转（融资口径）','持有建筑折旧（项目口径）','持有建筑折旧（融资口径）','持有处置前账面净值（项目）','持有处置前账面净值（融资）','期末未售开发存货（融资）','期末处置损益（项目口径）','期末处置损益（融资口径）']
        for off in (0,OFFSET):
            self.band(5,7+off,'基准情景 · 现金成本与非现金成本分开' if not off else '组合压力情景 · 建设成本上升并重算结转及折旧')
            for (k,r),name in zip(COST.items(),labels):self.label(5,r+off,name,'万元')
            for y in range(31):
                c=col(y);age=f'{c}$4-{p("build_years")}'
                forms={
                 'sale_expense':f'{at(4,INC["sales"],y,off)}*{p("sale_expense")}',
                 'sale_tax':f'{at(4,INC["sales"],y,off)}*{p("sale_tax")}',
                 'fixed':f'IF({age}>0,{p("fixed_cost")}*(1+{p("fixed_growth")})^MAX(0,{age}-1),0)',
                 'variable':f'{at(4,INC["held_income"],y,off)}*{p("variable_cost")}',
                 'maintenance':f'IF({age}>0,{p("maintenance")},0)',
                 'operating_tax':f'{at(4,INC["held_income"],y,off)}*{p("operating_tax")}',
                 'cash_total':f'SUM({c}{10+off}:{c}{15+off})',
                 'sale_cost_project':f'{scalar(3,INV["sale_basis_project"],off)}*{at(2,PLAN["sale"],y,off)}',
                 'sale_cost_equity':f'{scalar(3,INV["sale_basis_equity"],off)}*{at(2,PLAN["sale"],y,off)}',
                 'dep_project':f'IF(AND({age}>0,{age}<={p("dep_years")}),{scalar(3,INV["dep_basis_project"],off)}*(1-{p("residual")})/MAX(1,{p("dep_years")}),0)',
                 'dep_equity':f'IF(AND({age}>0,{age}<={p("dep_years")}),{scalar(3,INV["dep_basis_equity"],off)}*(1-{p("residual")})/MAX(1,{p("dep_years")}),0)',
                 'book_project':f'(SUM({q(SHEETS[3],f"$D${INV['land_cash']+off}:{c}{INV['land_cash']+off}")})+SUM({q(SHEETS[3],f"$D${INV['nonland_cash']+off}:{c}{INV['nonland_cash']+off}")}))*{p("rent_share")}-SUM($D{19+off}:{c}{19+off})',
                 'book_equity':f'(SUM({q(SHEETS[3],f"$D${INV['land_cash']+off}:{c}{INV['land_cash']+off}")})+SUM({q(SHEETS[3],f"$D${INV['nonland_cash']+off}:{c}{INV['nonland_cash']+off}")})+SUM({q(SHEETS[6],f"$D${FIN['cap_interest']+off}:{c}{FIN['cap_interest']+off}")}))*{p("rent_share")}-SUM($D{20+off}:{c}{20+off})',
                 'inventory':f'(SUM({q(SHEETS[3],f"$D${INV['land_cash']+off}:{c}{INV['land_cash']+off}")})+SUM({q(SHEETS[3],f"$D${INV['nonland_cash']+off}:{c}{INV['nonland_cash']+off}")})+SUM({q(SHEETS[6],f"$D${FIN['cap_interest']+off}:{c}{FIN['cap_interest']+off}")}))*{p("sale_share")}-SUM($D{18+off}:{c}{18+off})',
                 'gain_project':f'IF(AND({p("terminal_on")}=1,{c}$4={p("horizon")}),{at(4,INC["terminal_net"],y,off)}-{c}{21+off},0)',
                 'gain_equity':f'IF(AND({p("terminal_on")}=1,{c}$4={p("horizon")}),{at(4,INC["terminal_net"],y,off)}-{c}{22+off},0)'}
                for k,r in COST.items():self.yearly(5,r,y,forms[k],off)
        self.ws[5].print_area(0,0,OFFSET+26,33)

    def cashflow(self):
        self.headers(7)
        names=['营业收入（销售＋持有）','现金成本及经营税费','无融资经营利润EBIT','融资口径经营利润EBIT','建成后利息费用','融资口径处置前税前利润','项目口径所得税含处置损益','融资口径所得税含处置损益','融资口径净利润含处置损益','可用于偿债现金CFADS（不含处置）','项目经营现金流（不含处置）','土地现金支出','非土地建设现金支出','周转资金投入','周转资金收回','持有处置净现金（所得税前）','项目投资税后现金流','贷款提款（非营业收入）','全部现金利息支付','本金偿还','所得税口径调整（项目税－融资税）','资本金投入前资金净流量','期末后续支出及偿债留存需求','理论期初资金余额','完成计划当年所需资本金','完整出资可分配现金','理论期末资金余额','累计所需资本金','资本金现金流（完整出资理论）','额度内期初现金余额','期初未落实资金缺口','额度内当年资金需要','额度内当年实际可投资本金','额度约束下可分配现金','额度内期末可用现金余额','期末未落实资金缺口','至当年最大未落实缺口','累计额度内可投资本金','额度内资本金现金流','累计项目现金流','当年项目折现现金流','累计项目折现现金流','当年运营偿债覆盖倍数','最低DSCR筛选值（非偿债年忽略）','静态回收首次跨零标记','动态回收首次跨零标记','项目现金流符号','项目最近非零符号','项目现金流变号标记','理论资金平衡差额','额度内资金平衡差额','分配后未来追加资本金校验','资本金现金流符号','资本金最近非零符号','资本金现金流变号标记']
        for off in (0,OFFSET):
            self.band(7,7+off,'基准情景 · 利润、现金流与资金约束' if not off else '组合压力情景 · 全链条重算')
            names_all=names+['销售净回款贡献（所得税前）','持有净经营现金贡献（所得税前）']
            for (k,r),name in zip(CF.items(),names_all):self.label(7,r+off,name,'倍' if k in ['dscr','dscr_candidate'] else '标记' if k.startswith('sign') or 'cross' in k or 'changes' in k or 'last_sign' in k else '万元')
            # All reserve cells depend only on ex-equity cash flows, never on distributions.
            for y in range(31):
                c=col(y);prev=col(y-1);nxt=col(y+1)
                local=lambda k:f'{c}{CF[k]+off}'
                old=lambda k:f'{prev}{CF[k]+off}' if y else '0'
                inc=lambda k:at(4,INC[k],y,off)
                cost=lambda k:at(5,COST[k],y,off)
                fin=lambda k:at(6,FIN[k],y,off)
                inv=lambda k:at(3,INV[k],y,off)
                forms={
                 'revenue':inc('revenue'),'cash_cost':cost('cash_total'),
                 'ebit_project':f'{local("revenue")}-{local("cash_cost")}-{cost("sale_cost_project")}-{cost("dep_project")}',
                 'ebit_equity':f'{local("revenue")}-{local("cash_cost")}-{cost("sale_cost_equity")}-{cost("dep_equity")}',
                 'op_interest':fin('op_interest'),'pretax':f'{local("ebit_equity")}-{local("op_interest")}',
                 'tax_project':f'MAX(0,{local("ebit_project")}+{cost("gain_project")})*{p("income_tax")}',
                 'tax_equity':f'MAX(0,{local("pretax")}+{cost("gain_equity")})*{p("income_tax")}',
                 'profit':f'{local("pretax")}+{cost("gain_equity")}-{local("tax_equity")}',
                 'cfads':f'{local("revenue")}-{local("cash_cost")}-MAX(0,{local("pretax")})*{p("income_tax")}',
                 'op_project':f'{local("revenue")}-{local("cash_cost")}-MAX(0,{local("ebit_project")})*{p("income_tax")}',
                 'land':inv('land_cash'),'build':inv('nonland_cash'),'wc_in':inv('wc_in'),'wc_out':inv('wc_out'),'terminal':inc('terminal_net'),
                 'project':f'{local("revenue")}-{local("cash_cost")}-{local("tax_project")}-{local("land")}-{local("build")}-{local("wc_in")}+{local("wc_out")}+{local("terminal")}',
                 'draw':fin('draw'),'interest':fin('interest'),'principal':fin('principal'),
                 'tax_bridge':f'{local("tax_project")}-{local("tax_equity")}',
                 'before_equity':f'{local("project")}+{local("tax_bridge")}+{local("draw")}-{local("interest")}-{local("principal")}',
                 'reserve':f'MAX(0,{nxt}{32+off}-{nxt}{31+off})' if y<30 else '0',
                 'theory_open':old('theory_close'),
                 'required_equity':f'MAX(0,-({local("theory_open")}+{local("before_equity")}))',
                 'theory_distribution':f'MAX(0,{local("theory_open")}+{local("before_equity")}+{local("required_equity")}-{local("reserve")})',
                 'theory_close':f'{local("theory_open")}+{local("before_equity")}+{local("required_equity")}-{local("theory_distribution")}',
                 'cum_required':f'{old("cum_required")}+{local("required_equity")}',
                 'equity':f'{local("theory_distribution")}-{local("required_equity")}',
                 'actual_open':old('actual_close'),'gap_open':old('gap'),
                 'actual_need':f'MAX(0,-({local("actual_open")}+{local("before_equity")}-{local("gap_open")}))',
                 'actual_equity':f'MIN({local("actual_need")},MAX(0,{p("equity_cap")}-{old("cum_actual")}))',
                 'gap':f'MAX(0,-({local("actual_open")}+{local("before_equity")}-{local("gap_open")}+{local("actual_equity")}))',
                 'max_gap':f'MAX({old("max_gap")},{local("gap")})',
                 'actual_distribution':f'IF({local("max_gap")}>0.01,0,MAX(0,{local("actual_open")}+{local("before_equity")}-{local("gap_open")}+{local("actual_equity")}-{local("reserve")}))',
                 'actual_close':f'MAX(0,{local("actual_open")}+{local("before_equity")}-{local("gap_open")}+{local("actual_equity")})-{local("actual_distribution")}',
                 'cum_actual':f'{old("cum_actual")}+{local("actual_equity")}',
                 'actual_equity_cf':f'{local("actual_distribution")}-{local("actual_equity")}',
                 'cum_project':f'{old("cum_project")}+{local("project")}',
                 'discounted':f'{local("project")}/(1+{p("project_discount")})^{c}$4',
                 'cum_discounted':f'{old("cum_discounted")}+{local("discounted")}',
                 'dscr':f'IF(AND({c}$4>{p("build_years")},{fin("service")}>0.000001),{local("cfads")}/{fin("service")},"")',
                 'dscr_candidate':f'IF(ISNUMBER({local("dscr")}),{local("dscr")},1E99)',
                 'static_cross':f'IF(AND({c}$4>0,{old("cum_project")}<0,{local("cum_project")}>='+'0),1,0)',
                 'dynamic_cross':f'IF(AND({c}$4>0,{old("cum_discounted")}<0,{local("cum_discounted")}>='+'0),1,0)',
                 'sign_project':f'IF(ABS({local("project")})<=0.000001,0,SIGN({local("project")}))',
                 'last_sign_project':f'IF({local("sign_project")}=0,{old("last_sign_project")},{local("sign_project")})',
                 'changes_project':f'IF(AND({local("sign_project")}<>0,{old("last_sign_project")}<>0,{local("sign_project")}<>{old("last_sign_project")}),1,0)',
                 'theory_balance':f'{local("theory_open")}+{local("before_equity")}+{local("required_equity")}-{local("theory_distribution")}-{local("theory_close")}',
                 'actual_balance':f'{local("actual_open")}-{local("gap_open")}+{local("before_equity")}+{local("actual_equity")}-{local("actual_distribution")}-{local("actual_close")}+{local("gap")}',
                 'reserve_check':f'IF({local("theory_distribution")}>0.01,SUM({nxt}{34+off}:$AH{34+off}),0)' if y<30 else '0',
                 'sign_equity':f'IF(ABS({local("equity")})<=0.000001,0,SIGN({local("equity")}))',
                 'last_sign_equity':f'IF({local("sign_equity")}=0,{old("last_sign_equity")},{local("sign_equity")})',
                 'changes_equity':f'IF(AND({local("sign_equity")}<>0,{old("last_sign_equity")}<>0,{local("sign_equity")}<>{old("last_sign_equity")}),1,0)',
                 'sale_cash_contribution':f'{inc("sales")}-{cost("sale_expense")}-{cost("sale_tax")}',
                 'held_cash_contribution':f'{inc("held_income")}-{cost("fixed")}-{cost("variable")}-{cost("maintenance")}-{cost("operating_tax")}'}
                for k,r in CF.items():
                    # DSCR placeholder remains text, the screening helper remains 1E99 outside horizon.
                    mask=k not in ['dscr','dscr_candidate']
                    formula=forms[k]
                    if k=='dscr':formula=f'IF({c}$4<={p("horizon")},{formula},"")'
                    if k=='dscr_candidate':formula=f'IF({c}$4<={p("horizon")},{formula},1E99)'
                    self.yearly(7,r,y,formula,off,mask=mask)
        self.ws[7].print_area(0,0,OFFSET+67,33)

    def metrics(self):
        s=self.ws[8];s.set_column('A:A',48);s.set_column('B:B',24);s.set_column('C:C',19);s.set_column('D:D',24);s.set_column('E:E',24);s.set_column('F:F',18);s.set_column('G:G',75)
        for a,text in [('A4','指标 / 观察项'),('B4','基准情景'),('C4','单位'),('D4','组合压力情景'),('G4','口径 / 提示')]:s.write(a,text,self.fm['header'])
        definitions={
         'project_irr':('项目税后IRR','%','不含融资，所得税按无融资口径；IRR只取第0年至第N年'),
         'equity_irr':('资本金IRR','%','完整出资理论：实际所需股东投入与留存后分配现金'),
         'project_npv':('项目NPV','万元','NPV(第1年至第N年)＋第0年；按项目折现率'),
         'equity_npv':('资本金NPV','万元','按资本金折现率，完整出资理论'),
         'static_payback':('项目静态回收期','年','含建设期；跨年线性插值'),
         'dynamic_payback':('项目动态回收期','年','使用折现项目现金流；跨年线性插值'),
         'dscr':('运营期最低偿债覆盖倍数','倍','CFADS/本息；无贷款为不适用；排除处置、周转回收及融资流入'),
         'required_equity':('所需累计资本金','万元','按完整计划及留存规则计算所需投入'),
         'max_gap':('未落实资金缺口最大值','万元','额度内资金平衡的历史最大未落实缺口'),
         'maturity_balance':('贷款到期未偿本金','万元','只取到期年；期限未覆盖时参数校验失败'),
         'total_investment':('项目总投资','万元','土地＋非土地建设＋建设期利息＋周转资金'),
         'loan':('贷款金额','万元','压力情景保持额度不变'),
         'equity_cap':('可投入资本金额度','万元','唯一额度输入，无额外资本金上限'),
         'sale_through':('测算期累计销售去化率','%','年度销售占总可售面积的比例累计'),
         'sale_complete':('完成销售年份','项目年','首次达到100%；未完成销售不自动产生处置回款'),
         'stable_occ':('稳定期实际出租率','%','取第4个出租年度；未覆盖显示不适用'),
         'breakeven':('持有经营保本出租率','%','收入覆盖固定、变动、维护及经营税；不含所得税、折旧、本息'),
         'noi_yield':('持有稳定期净经营收益率','%','第4个出租年度NOI/持有分摊总投资，含维护；不等于IRR'),
         'terminal_share':('期末处置净收入占现金流入','%','所得税前处置净收入/销售＋租金＋其他＋处置净收入＋周转回收'),
         'project_sign_warning':('项目IRR变号提示','文本','忽略中间零现金流；多次变号需谨慎理解IRR'),
         'equity_sign_warning':('资本金IRR变号提示','文本','同样检查完整出资资本金现金流'),
         'equity_label':('资本金收益结果标识','文本','缺口大于0时最终决策仍未通过'),
         'stable_year':('稳定观察年份','项目年','出租开始年＋3，第4出租年度；仍有少量新增免租')}
        for k,r in METRIC_ROWS.items():
            name,unit,note=definitions[k];s.write(f'A{r}',name);s.write(f'C{r}',unit);s.write(f'G{r}',note,self.fm['wrap']);s.set_row(r-1,30)
        for off,out in [(0,'B'),(OFFSET,'D')]:
            cr=lambda k:series(7,CF[k],off)
            ar=lambda k:active_series(7,CF[k],off)
            index=lambda sheet,row,y:f'INDEX({series(sheet,row,off)},{y}+1)'
            project_irr=f'IF(AND(COUNTIF({ar("project")},"<0")>0,COUNTIF({ar("project")},">0")>0),IFERROR(IRR({ar("project")}),"无法计算"),"无法计算")'
            eq_irr=f'IF(AND(COUNTIF({ar("equity")},"<0")>0,COUNTIF({ar("equity")},">0")>0),IFERROR(IRR({ar("equity")}),"无法计算"),"无法计算")'
            forms={
             'project_irr':project_irr,'equity_irr':eq_irr,
             'project_npv':f'NPV({p("project_discount")},{active_series(7,CF["project"],off,1)})+{at(7,CF["project"],0,off)}',
             'equity_npv':f'NPV({p("equity_discount")},{active_series(7,CF["equity"],off,1)})+{at(7,CF["equity"],0,off)}',
             'dscr':f'IF({p("loan")}=0,"不适用",IF(COUNT({ar("dscr")})=0,"无法计算",MIN({ar("dscr_candidate")})))',
             'required_equity':f'SUM({ar("required_equity")})','max_gap':f'MAX({ar("gap")})',
             'maturity_balance':f'IF({p("loan")}=0,0,IF(OR({scalar(6,FIN["maturity"],off)}>30,{scalar(6,FIN["maturity"],off)}>{p("horizon")}),"无法计算",{index(6,FIN["closing"],scalar(6,FIN["maturity"],off))}))',
             'total_investment':scalar(3,INV['total'],off),'loan':p('loan'),'equity_cap':p('equity_cap'),
             'sale_through':f'SUM({active_series(2,PLAN["sale"],off)})',
             'sale_complete':f'IF({out}19<0.999999,"测算期内未完成",IFERROR(MATCH(1,{series(2,PLAN["complete"],off)},0)-1,"测算期内未完成"))',
             'stable_year':f'IF({p("horizon")}>={p("rent_start")}+3,{p("rent_start")}+3,"不适用")',
             'stable_occ':f'IF(ISNUMBER({out}28),{index(2,PLAN["occ"],out+"28")},"不适用")',
             'project_sign_warning':f'IF(SUM({ar("changes_project")})>1,"多次变号：IRR可能有歧义","单次变号或无变号")',
             'equity_sign_warning':f'IF(SUM({ar("changes_equity")})>1,"多次变号：IRR可能有歧义","单次变号或无变号")',
             'equity_label':f'IF({out}14>0.01,"理论结果，资金尚未落实","按完整出资计划测算")',
             'terminal_share':f'IF(SUM({series(4,INC["revenue"],off)})+SUM({series(4,INC["terminal_net"],off)})+SUM({series(3,INV["wc_out"],off)})=0,"不适用",SUM({series(4,INC["terminal_net"],off)})/(SUM({series(4,INC["revenue"],off)})+SUM({series(4,INC["terminal_net"],off)})+SUM({series(3,INV["wc_out"],off)})))'}
            for which,flag,cf,cum in [('static_payback','static_cross','project','cum_project'),('dynamic_payback','dynamic_cross','discounted','cum_discounted')]:
                m=f'MATCH(1,{ar(flag)},0)'
                forms[which]=f'IF(COUNTIF({ar(cf)},"<0")=0,"不适用",IF(COUNTIF({ar(flag)},1)=0,"测算期内未回收",{m}-2-INDEX({cr(cum)},MAX(1,{m}-1))/INDEX({cr(cf)},{m})))'
            stable=out+'28'
            rev100=f'{scalar(2,AREA["held"])}*{index(4,INC["rent_price"],stable)}*12/10000*{p("collection")}*(1+{p("other_income")})'
            fixed=f'{index(5,COST["fixed"],stable)}+{p("maintenance")}'
            forms['breakeven']=f'IF(ISNUMBER({stable}),IF({rev100}*(1-{p("variable_cost")}-{p("operating_tax")})<=0,"无法计算",({fixed})/({rev100}*(1-{p("variable_cost")}-{p("operating_tax")}))),"不适用")'
            noi=f'{index(4,INC["held_income"],stable)}-{index(5,COST["fixed"],stable)}-{index(5,COST["variable"],stable)}-{index(5,COST["maintenance"],stable)}-{index(5,COST["operating_tax"],stable)}'
            forms['noi_yield']=f'IF(ISNUMBER({stable}),IF({scalar(3,INV["held_invest"],off)}>0,({noi})/{scalar(3,INV["held_invest"],off)},"无法计算"),"不适用")'
            for k,r in METRIC_ROWS.items():self.f(8,f'{out}{r}',forms[k],'percent' if definitions[k][1]=='%' else 'wrap' if definitions[k][1]=='文本' else 'number')
        self.f(8,'G22','IF(OR(AND(ISNUMBER(B22),B22>1),AND(ISNUMBER(D22),D22>1)),"保本出租率超过100%：满租也无法覆盖经营现金支出","保本率假设稳定全年租赁、无新增免租；不含折旧、本息及所得税")','wrap')
        s.merge_range('A31:G31','逐项判断（金额容差0.01万元；回收期未回收即未通过；阈值均为模拟值）',self.fm['section'])
        for a,t in [('A33','指标 / 安全条件'),('B33','基准测算值'),('C33','标准'),('D33','基准状态'),('E33','压力测算值'),('F33','压力状态')]:s.write(a,t,self.fm['header'])
        for i,(key,label,threshold,kind) in enumerate(CONDITIONS,34):
            s.write(f'A{i}',label)
            for out,val,status in [('B','B','D'),('D','E','F')]:
                self.f(8,f'{val}{i}',f'{out}{METRIC_ROWS[key]}','percent' if 'irr' in key else 'number')
                std=p(threshold) if threshold else '0'
                self.f(8,f'C{i}',std,'percent' if 'irr' in key else 'number')
                v=f'{val}{i}';t=f'C{i}'
                expr=f'{v}>={t}-0.00000001' if kind in ['ge','dscr'] else f'{v}<={t}+0.00000001' if kind=='le' else f'{v}<={t}+0.01' if kind=='le_money' else f'{v}>=-0.01' if kind=='nonnegative' else f'ABS({v})<=0.01'
                self.f(8,f'{status}{i}',f'IF({v}="不适用",'+('"不适用"' if kind=='dscr' else '"未通过"')+f',IF(NOT(ISNUMBER({v})),IF({v}="无法计算","待判断","未通过"),IF({expr},"通过","未通过")))','text')
        checks=q(SHEETS[9],'$D$35:$D$64')
        for out,statuscol,head,reason in [('B','D','B46','B48'),('D','F','E46','B50')]:
            self.f(8,head,f'IF(OR(COUNTIF({checks},"未通过")>0,COUNTIF({statuscol}34:{statuscol}43,"待判断")>0),"待判断：请检查参数",IF(COUNTIF({statuscol}34:{statuscol}43,"未通过")>0,"未通过","通过"))','wrap')
            terms=[f'IF({statuscol}{i}="未通过","{label}未达标；",IF({statuscol}{i}="待判断","{label}无法计算；",""))' for i,(_,label,_,_) in enumerate(CONDITIONS,34)]
            self.ws[8].merge_range(reason+':G'+reason[1:],'',self.fm['wrap'])
            self.f(8,reason,f'IF(COUNTIF({checks},"未通过")>0,"参数校验未通过；","")&'+ '&'.join(terms),'wrap')
            self.ws[8].set_row(int(reason[1:])-1,58)
        s.write('A46','情景总体状态');s.write('A48','基准不通过原因');s.write('A50','压力不通过原因')
        self.status_cf(8,'D34:D43');self.status_cf(8,'F34:F43');self.status_cf(8,'B46:G46')
        s.freeze_panes(5,1);s.print_area('A1:G51')

    def dashboard(self):
        s=self.ws[9];s.set_column('A:A',48);s.set_column('B:B',22);s.set_column('C:C',22);s.set_column('D:D',18);s.set_column('E:E',24);s.set_column('F:F',20);s.set_column('G:G',26);s.set_column('H:L',12)
        s.merge_range('A5:G6','',self.fm['title'])
        overall=q(SHEETS[8],'B46');down=q(SHEETS[8],'E46')
        self.f(9,'A5',f'IF(OR(LEFT({overall},3)="待判断",AND({p("stress_required")}=1,LEFT({down},3)="待判断")),"待判断：请检查参数",IF(AND({overall}="通过",OR({p("stress_required")}=0,{down}="通过")),"通过：当前模拟参数下具备投资条件","未通过"))','title')
        self.status_cf(9,'A5:G6')
        cards=[('项目名称',p('name'),'可经营面积（平方米）',scalar(2,AREA['operable'])),
               ('测算期限（年）',p('horizon'),'销售 / 出租比例',f'TEXT({p("sale_share")},"0%")&" / "&TEXT({p("rent_share")},"0%")'),
               ('地上面积（平方米）',scalar(2,AREA['above']),'持有面积（平方米）',scalar(2,AREA['held'])),
               ('项目总投资（万元）',q(SHEETS[8],'B16'),'贷款（万元）',p('loan')),
               ('所需资本金（万元）',q(SHEETS[8],'B13'),'资本金额度（万元）',p('equity_cap')),
               ('最大未落实缺口（万元）',q(SHEETS[8],'B14'),'压力情景是否须通过',f'IF({p("stress_required")}=1,"是","否，仅供观察")')]
        for r,(l,v,l2,v2) in enumerate(cards,8):
            s.write(f'A{r}',l);s.merge_range(f'B{r}:C{r}','',self.fm['result']);self.f(9,f'B{r}',v,'wrap' if r==8 else 'result')
            s.merge_range(f'E{r}:F{r}',l2,self.fm['text']);self.f(9,f'G{r}',v2,'wrap' if r in [9,13] else 'result')
        s.merge_range('A15:G16','',self.fm['wrap'])
        self.f(9,'A15',f'{q(SHEETS[8],"B48")}&IF({p("stress_required")}=1," 压力情景："&{q(SHEETS[8],"B50")},"")','wrap');s.set_row(14,35);s.set_row(15,30)
        for a,t in [('A18','指标 / 安全条件'),('B18','基准值'),('C18','通过标准'),('D18','状态'),('E18','压力值'),('F18','压力状态')]:s.write(a,t,self.fm['header'])
        for r,(_,label,_,kind) in enumerate(CONDITIONS,19):
            s.write(f'A{r}',label)
            for c,src in [('B','B'),('C','C'),('D','D'),('E','E'),('F','F')]:self.f(9,f'{c}{r}',q(SHEETS[8],f'{src}{r+15}'),'percent' if r in [19,20] and c in ['B','C','E'] else 'text' if c in ['D','F'] else 'number')
        s.merge_range('A30:G31','',self.fm['wrap']);self.f(9,'A30',q(SHEETS[8],'B27')+'&"；"&'+q(SHEETS[8],'B25'),'wrap')
        s.merge_range('A33:G33','参数与计算校验（两种情景均核对；全部通过才可作出投资判断）',self.fm['section'])
        for a,t in [('A34','校验项目'),('B34','说明 / 校验口径'),('D34','结果')]:s.write(a,t,self.fm['header'])
        build_rng=q(SHEETS[1],f'$D${BUILD_PLAN_ROW}:$AH${BUILD_PLAN_ROW}');sale_rng=q(SHEETS[1],f'$D${SALE_PLAN_ROW}:$AH${SALE_PLAN_ROW}')
        yr=q(SHEETS[1],f'$D${PLAN_HEADER}:$AH${PLAN_HEADER}')
        numeric=[p(k) for k in DEFAULTS if k!='name']
        checks=[
         ('必填输入完整','数值必填，年度计划不得留空',f'AND(COUNT({",".join(numeric)})={len(numeric)},COUNT({build_rng})=31,COUNT({sale_rng})=31,{p("name")}<>"")'),
         ('测算期限与建设年限','整数，1≤建设年限<期限≤30',f'AND({p("horizon")}=INT({p("horizon")}),{p("horizon")}<=30,{p("build_years")}=INT({p("build_years")}),{p("build_years")}>=1,{p("horizon")}>{p("build_years")})'),
         ('面积与租售比例','面积>0，经营比例及销售出租比例有效',f'AND({p("land_area")}>0,{p("far")}>0,{p("under_area")}>=0,{p("operable_share")}>0,{p("operable_share")}<=1,{p("sale_share")}>0,{p("sale_share")}<1)'),
         ('面积分解平衡','可售＋持有＝可经营',f'ABS({scalar(2,9)}+{scalar(2,10)}-{scalar(2,8)})<=0.01'),
         ('建设计划合计','非土地建设比例累计100%，非负',f'AND(ABS(SUM({build_rng})-1)<=0.000001,MIN({build_rng})>=0,MAX({build_rng})<=1)'),
         ('建设计划时点','第0年及建设期外不得安排建设',f'SUMPRODUCT({build_rng},--(({yr}<1)+({yr}>{p("build_years")})))=0'),
         ('销售计划合计','累计不超过100%，允许未售存货',f'AND(SUM({sale_rng})<=1.000001,MIN({sale_rng})>=0,MAX({sale_rng})<=1)'),
         ('销售计划时点','早于销售开始或超出期限不得销售',f'SUMPRODUCT({sale_rng},--(({yr}<{p("sale_start")})+({yr}>{p("horizon")})))=0'),
         ('开发时点','销售出租都在建成后且在期内开始',f'AND({p("sale_start")}>{p("build_years")},{p("rent_start")}>{p("build_years")},{p("sale_start")}<={p("horizon")},{p("rent_start")}<={p("horizon")},{p("sale_start")}=INT({p("sale_start")}),{p("rent_start")}=INT({p("rent_start")}))'),
         ('出租率与爬坡','0≤首年≤次年≤第三年≤稳定率≤1',f'AND({p("occ1")}>=0,{p("occ1")}<={p("occ2")},{p("occ2")}<={p("occ3")},{p("occ3")}<={p("occ_stable")},{p("occ_stable")}<=1)'),
         ('免租与收缴','0—12月，收缴0—100%',f'AND({p("free_months")}>=0,{p("free_months")}<=12,{p("collection")}>0,{p("collection")}<=1)'),
         ('价格成本及额度非负','全部金额/成本/贷款/额度非负', 'MIN('+','.join(p(k) for k in ['land_price','above_cost','under_cost','infra','pre','other_build','working_capital','sale_price','rent','fixed_cost','maintenance','loan','equity_cap','terminal_value'])+')>=0'),
         ('费用税率','所有费用、税率、残值及预备费在0—1', 'AND(MIN('+','.join(p(k) for k in ['contingency','sale_expense','sale_tax','variable_cost','operating_tax','income_tax','residual','terminal_fee','other_income'])+')>=0,MAX('+','.join(p(k) for k in ['contingency','sale_expense','sale_tax','variable_cost','operating_tax','income_tax','residual','terminal_fee','other_income'])+')<=1)'),
         ('增长率折现率及利率','增长>-100%；折现率非负，贷款利率0—100%',f'AND({p("rent_growth")}>-1,{p("fixed_growth")}>-1,{p("interest")}>=0,{p("interest")}<=1,{p("project_discount")}>=0,{p("equity_discount")}>=0)'),
         ('折旧年限','折旧年限为1—100整数',f'AND({p("dep_years")}>=1,{p("dep_years")}<=100,{p("dep_years")}=INT({p("dep_years")}))'),
         ('贷款金额与建设需求','贷款不能超过基准非土地建设成本',f'{p("loan")}<={scalar(3,INV["nonland"])}+0.01'),
         ('贷款期限及偿还时点','建设后开始还本，期限从首次提款年计',f'OR({p("loan")}=0,AND({p("loan_term")}>=1,{p("loan_term")}=INT({p("loan_term")}),{p("repay_start")}=INT({p("repay_start")}),{p("repay_start")}>{p("build_years")},{p("repay_start")}<={scalar(6,FIN["maturity"])}))'),
         ('期限覆盖贷款到期','N≥首次提款年＋期限－1',f'OR({p("loan")}=0,{p("horizon")}>={scalar(6,FIN["maturity"])})'),
         ('贷款提款总量','两情景均按计划提款，不自动增贷',f'AND(ABS(SUM({series(6,FIN["draw"])})-{p("loan")})<=0.01,ABS(SUM({series(6,FIN["draw"],OFFSET)})-{p("loan")})<=0.01)'),
         ('贷款余额与还本平衡','提款－还本＝期末本金，余额非负', 'AND('+','.join(f'ABS(SUM({series(6,11,o)})-SUM({series(6,13,o)})-INDEX({series(6,15,o)},{p("horizon")}+1))<=0.01' for o in [0,OFFSET])+')'),
         ('理论资金平衡','期初＋净流＋出资－分配＝期末', 'AND('+','.join(f'SUMPRODUCT(ABS({series(7,59,o)}))<=0.01' for o in [0,OFFSET])+')'),
         ('额度内资金平衡','缺口作为待落实支出，不当成现金', 'AND('+','.join(f'SUMPRODUCT(ABS({series(7,60,o)}))<=0.01' for o in [0,OFFSET])+')'),
         ('资本金不突破额度','累计实际可投不超过输入额度',f'AND(SUM({series(7,42)})<={p("equity_cap")}+0.01,SUM({series(7,42,OFFSET)})<={p("equity_cap")}+0.01)'),
         ('留存后不立即补资','分配后不再出现计划内追加资本金',f'AND(MAX({series(7,61)})<=0.01,MAX({series(7,61,OFFSET)})<=0.01)'),
         ('压力参数范围','下降比例0—100%，出租率按百分点',f'AND({p("stress_sale")}>=0,{p("stress_sale")}<=1,{p("stress_rent")}>=0,{p("stress_rent")}<=1,{p("stress_occ")}>=0,{p("stress_occ")}<=100,{p("stress_cost")}>=0)'),
         ('开关参数','两个开关均为0或1',f'AND(OR({p("terminal_on")}=0,{p("terminal_on")}=1),OR({p("stress_required")}=0,{p("stress_required")}=1))'),
         ('总投资现金核对','支付投资合计＝总投资（不扣周转回收）','AND('+','.join(f'ABS(SUM({series(3,35,o)})-{scalar(3,18,o)})<=0.01' for o in [0,OFFSET])+')'),
         ('第N年回收与处置时点','周转仅期末回收，处置仅持有物业','AND('+','.join(f'ABS(SUM({series(3,32,o)})-{p("working_capital")})<=0.01' for o in [0,OFFSET])+')'),
         ('判断标准有效','IRR阈值非负、回收期上限及DSCR>0',f'AND({p("project_irr_min")}>=0,{p("equity_irr_min")}>=0,{p("static_limit")}>0,{p("dynamic_limit")}>0,{p("dscr_min")}>0)'),
         ('项目有投资且持有分摊合理','总投资>0；销售＋持有分摊开发成本一致',f'AND({scalar(3,18)}>0,ABS({scalar(3,22)}+{scalar(3,23)}+{scalar(3,21)}-{scalar(3,16)})<=0.01)')]
        self.check_definitions=checks
        for r,(name,note,expr) in enumerate(checks,35):
            s.write(f'A{r}',name);s.merge_range(f'B{r}:C{r}',note,self.fm['wrap']);s.set_row(r-1,31)
            self.f(9,f'D{r}',f'IF({expr},"通过","未通过")','text')
        self.status_cf(9,'D19:D28');self.status_cf(9,'F19:F28');self.status_cf(9,'D35:D64')
        s.merge_range('A66:G66','基准情景图表 · 修改输入及期限后自动更新',self.fm['section'])
        for j,t in enumerate(['项目年度','销售收入','实收租金','其他收入','可用于偿债现金','应还本息','累计项目现金流','累计折现现金流']):s.write(119,j,t,self.fm['header'])
        for y in range(31):
            r=121+y;self.f(9,f'A{r}',q(SHEETS[1],f'{col(y)}{PLAN_HEADER}'),'integer')
            for out,si,row in [('B',4,13),('C',4,20),('D',4,21),('E',7,19),('F',6,18),('G',7,49),('H',7,51)]:self.f(9,f'{out}{r}',at(si,row,y))
        for name,c in [('ChartYear','A'),('ChartSales','B'),('ChartRent','C'),('ChartOther','D'),('ChartCFADS','E'),('ChartDebt','F'),('ChartCum','G'),('ChartDiscounted','H')]:
            self.wb.define_name(name,q(SHEETS[9],f'${c}$121')+':INDEX('+q(SHEETS[9],f'${c}$121:${c}$151')+','+p('horizon')+'+1)')
        charts=[('column','stacked','年度销售与出租收入构成',[('销售回款','ChartSales',COLORS['navy']),('实收租金','ChartRent',COLORS['teal']),('其他经营收入','ChartOther',COLORS['orange'])],67),
                ('column',None,'可用于偿债的现金与本息',[('可用于偿债现金','ChartCFADS',COLORS['teal']),('应还本金及利息','ChartDebt',COLORS['orange'])],84),
                ('line',None,'累计项目现金流与累计折现现金流',[('累计项目现金流','ChartCum',COLORS['navy']),('累计折现现金流','ChartDiscounted',COLORS['orange'])],101)]
        for typ,sub,title,items,r in charts:
            chart=self.wb.add_chart(dict(type=typ,**({'subtype':sub} if sub else {})))
            for label,n,color in items:
                d=dict(name=label,categories="='决策看板'!ChartYear",values=f"='决策看板'!{n}",
                       categories_data=list(range(int(DEFAULTS['horizon'])+1)),values_data=[0]*(int(DEFAULTS['horizon'])+1))
                d['line' if typ=='line' else 'fill']=dict(color=color)
                if typ!='line':d['border']=dict(none=True)
                chart.add_series(d)
            chart.set_title(dict(name=title,name_font=dict(name='Microsoft YaHei',size=12)))
            chart.set_x_axis(dict(name='项目年度',label_position='low',num_font=dict(name='Microsoft YaHei',size=9)))
            chart.set_y_axis(dict(name='万元',num_format='#,##0',num_font=dict(name='Microsoft YaHei',size=9),major_gridlines=dict(visible=True,line=dict(color='#DFE5E9'))))
            chart.set_legend(dict(position='bottom',font=dict(name='Microsoft YaHei',size=9)))
            chart.set_chartarea(dict(border=dict(none=True)));chart.set_size(dict(width=1060,height=320))
            s.insert_chart(r-1,0,chart)
        s.freeze_panes(7,0);s.set_zoom(85);s.print_area('A1:H117');s.set_h_pagebreaks([33,66,83,100]);s.set_v_pagebreaks([])

    def instructions(self):
        s=self.ws[0];s.set_column('A:A',28);s.set_column('B:G',18)
        lines=[
         ('模型用途','拿地前的租售结合产业园学习测算。所有默认数据及投资阈值均为模拟值。年度模型覆盖第0年、第1—N年，最多30年；后续年末现金流，建设投资及运营作年度简化。'),
         ('使用步骤','1 在“核心参数输入”修改黄色单元格；2 同步调整建设和销售年度计划；3 在Excel中自动重算；4 查看“决策看板”的校验、指标及原因；5 在计算表核对各年变化。'),
         ('工作表结构','使用说明 → 核心参数输入 → 建筑面积与开发计划 → 投资与资金安排 → 销售与出租收入 → 运营成本与税费 → 融资与还本付息 → 利润与现金流 → 指标与情景测算 → 决策看板。'),
         ('颜色与编辑','浅黄色底蓝色文字＝手动输入；中性色＝公式及结果；绿色＝通过，红色＝未通过或待判断，灰色＝不适用。所有公式可查看编辑，无保护、宏和外部链接。'),
         ('单位','金额万元；面积平方米；地价、建安及售价元/平方米；租金元/平方米/月。面积×单价÷10000；年租金＝面积×月租金×12÷10000。百分比输入10%=0.10；出租率压力10个百分点输入10。'),
         ('收益与资金条件','IRR越高越好，NPV≥0，回收期越短越好，DSCR越高越好。收益达标不代表资金充足：资本金总额度、资金缺口和到期本金也必须满足。NPV第0年不折现。'),
         ('两类现金流','项目投资现金流排除贷款、股东出资及全部本息、建设期融资费用，所得税按无融资利润；资本金现金流＝完整出资下留存后股东分配－按需股东投入。项目总投资另包含建设期融资费用。'),
         ('状态说明','“无法计算”＝IRR无解或必要条件不足，不用0冒充；“测算期内未回收”＝对应累计现金流在期内没有由负转非负，视为未达标；“不适用”＝无贷款或未覆盖辅助稳定观察年等。必要收益指标无解时待判断。'),
         ('建设与分摊','土地第0年支付；其余非土地支出第1—K年按输入比例分配。建安、配套、前期、管理各范围不重叠；预备费以非土地上述成本为基数。共同开发成本按销售/持有经营面积比例分摊，非经营面积及地下成本一起分摊。'),
         ('收入与免租','销售部分出售前也不出租，出售后不再贡献租金；持有部分出租。年度销售占总可售面积比例，当年收款。出租年内面积视为全年占用，只对本年增加的出租面积扣一次免租，不模拟租户更替，故爬坡率不得下降。租金按收缴后现金口径确认收入，其他收入按实收租金比例。'),
         ('税务与成本简化','收入均为简化含税实收口径，销售及运营综合税费也按该口径；不建立增值税抵扣、土地增值税及复杂税制。所得税＝max(年度利润,0)×税率，不结转亏损。固定成本由持有部分承担，自K+1年即支付，空置不豁免；变动成本仅按持有收入；维护更新当年费用化并税前扣除，不重复资本化或折旧。'),
         ('非现金成本','销售成本按年度销售比例结转，未售保留存货；建设期间在建及存货按累计已支付成本显示。持有建筑非土地成本直线折旧，土地不折旧。融资利润将建设利息分摊至销售成本及持有折旧；项目口径剔除融资利息。建设成本只在支付年扣现金，结转与折旧不再次扣现金。'),
         ('贷款规则','一笔贷款按建设比例提款；首次提款年算贷款期限，到期年＝首次提款年＋期限－1。建设完成后到期前等额本金；利息＝(期初本金＋本年提款/2－本年还本/2)×利率。建设期利息资本金承担且不滚入本金。销售回款可还计划内本息，不额外提前还款。'),
         ('留存与分配','从期末向前计算留存需求R(t)=max(0,R(t+1)－下一年资本金前净流量)。因此先保留已知后续建设和计划偿债所需现金，剩余才分配；资本金只在现金不足当年补入，不为凑留存而提前投入。该确定性规则不单列额外安全现金。'),
         ('额度与缺口','先按完整计划计算理论所需资本金；再按唯一额度投入。不足部分作为未落实计划支出缺口滚动，后续净回款先覆盖缺口；发生过缺口的额度内路径停止分配。缺口从不作为现金/贷款/真实出资。该路径仅定位无法执行的计划，不代表项目能在无资金情况下施工运营；理论资本金收益标明资金尚未落实，最大历史缺口仍导致未通过。'),
         ('期末处理','默认处置关闭，持有净值及未售存货均不自动转现金。开启后第N年仅收持有物业输入处置价，扣综合交易费用，净额减处置前账面净值作为所得税损益，与本年利润合并，不跨年结转。初始运营周转资金第N年全额收回；不回收其他未单列项目。'),
         ('压力情景','各计算表第二个可见区同时降低销售价、租金、出租率并增加非土地建设成本及对应预备费；土地、贷款额度不变。出租率按百分点扣减并最低为0。默认只供观察，开启开关后基准及压力均须通过。'),
         ('辅助指标','稳定观察取第4出租年度，包含当年少量新增免租；保本出租率另按稳定全年无新增免租收入测算，只覆盖运营支出与经营税费，超过100%会提示。NOI收益率按含维护的稳定年NOI/持有分摊投资（含周转及分摊融资费用），不是IRR。处置依赖比例采用所得税前净处置收入。'),
         ('期限与IRR','调整N时现金流有效范围、IRR/NPV/回收期及处置/周转回收时点一起变化；预留年度不纳入IRR。N不得早于贷款到期。多次现金流变号会提示IRR可能有歧义；NPV及资金约束仍须审视。'),
         ('生成与重算','代码仅初始化默认示例；重新运行不读取或保留旧Excel中手改参数，同名文件自动加时间标识。Excel设置打开自动重算。可在Windows安装Excel后运行--recalculate；未实际重算的新文件有完整公式但缓存初始值不代表结果。验证详情见同目录“验证说明.md”。')]
        r=5
        for title,body in lines:
            s.write(r-1,0,title,self.fm['section']);s.merge_range(r-1,1,r+1,6,body,self.fm['wrap'])
            for rr in range(r-1,r+2):s.set_row(rr,25)
            r+=4
        s.print_area(0,0,r,6);s.set_portrait();s.fit_to_pages(1,0)

    def build(self):
        self.inputs();self.area_plan();self.investment();self.financing();self.income();self.costs();self.cashflow();self.metrics();self.dashboard();self.instructions()
        self.ws[0].activate();self.wb.close()

def unique_path(directory:Path,filename='产业园投资测算.xlsx'):
    path=directory/filename
    if path.exists():
        stem=Path(filename).stem;ext=Path(filename).suffix
        stamp=datetime.now().strftime('%Y%m%d_%H%M%S_%f');path=directory/f'{stem}_{stamp}{ext}'
    return path

def recalculate_excel(path:Path):
    """PowerShell COM only operates an isolated new Excel instance, never a user's session."""
    script=r'''param([string]$File)
$ErrorActionPreference='Stop'
$excel=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
 $book=$excel.Workbooks.Open($File,0,$false)
 $excel.Calculation=-4105
 $excel.CalculateFullRebuild()
 if ($excel.CalculationState -ne 0) { throw 'Excel calculation did not finish' }
 $book.Save()
 Write-Output ('Excel '+$excel.Version+': recalculated and saved')
} finally {
 if ($null -ne $book) { $book.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
 if ($null -ne $excel) { $excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
'''
    import tempfile
    with tempfile.TemporaryDirectory(prefix='park_recalc_') as td:
        ps=Path(td)/'recalc.ps1';ps.write_text(script,encoding='utf-8-sig')
        subprocess.run(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(ps),'-File',str(path.resolve())],check=True)

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir',type=Path,default=Path(__file__).resolve().parent)
    parser.add_argument('--filename',default='产业园投资测算.xlsx')
    parser.add_argument('--recalculate',action='store_true',help='Recalculate using installed Windows Microsoft Excel')
    args=parser.parse_args();args.output_dir.mkdir(parents=True,exist_ok=True)
    path=unique_path(args.output_dir,args.filename)
    Builder(path).build()
    if args.recalculate:recalculate_excel(path)
    print(path.resolve())
if __name__=='__main__':main()
