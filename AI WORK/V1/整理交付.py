from pathlib import Path
import json
import shutil
import zipfile
import hashlib
import sys
base_dir = Path(__file__).resolve().parent
sys.path.insert(0,str(base_dir / '交付文件'))
import generate_excel as g
import validate_model as v

out=base_dir / '交付文件'
summary=json.loads((base_dir / '验证最终/validation_summary.json').read_text(encoding='utf-8'))
shutil.copy2(base_dir / '验证最终/交付验证完成.xlsx',out/'产业园投资测算.xlsx')
shutil.copy2(base_dir / '验证最终/validation_summary.json',out/'验证摘要.json')
baseline=summary['baseline'];base=baseline['scenarios'][0]['metrics'];stress=baseline['scenarios'][1]['metrics']

def fmt(value,percent=False):
    if isinstance(value,str):return value
    return f'{value*100:.2f}%' if percent else f'{value:,.2f}'

rows=[('项目税后IRR',6,True),('资本金IRR',7,True),('项目NPV（万元）',8,False),('资本金NPV（万元）',9,False),
      ('静态回收期（年）',10,False),('动态回收期（年）',11,False),('最低偿债覆盖倍数',12,False),
      ('所需累计资本金（万元）',13,False),('最大未落实缺口（万元）',14,False),('到期未还本金（万元）',15,False),
      ('项目总投资（万元）',16,False),('贷款（万元）',17,False),('可投入资本金（万元）',18,False),
      ('累计销售去化率',19,True),('完成销售项目年',20,False),('稳定期实际出租率',21,True),
      ('持有经营保本出租率',22,True),('持有稳定期NOI收益率',23,True),('期末处置净收入占流入',24,True)]
lines=['# 验证说明','',
       '验证日期：2026-10-11（用户时区 Asia/Shanghai）。模型仅用于个人学习，默认参数及阈值均为模拟值。',
       '', '## 默认结果与判断','',
       '基准判断：**未通过**。30项参数与计算校验全部通过，资金总额度足以完成默认计划，但收益和年度偿债覆盖指标未达标。组合压力情景也未通过，且出现资本金缺口。默认关闭期末物业处置，未自动将账面净值或未售库存转为回款。','',
       '默认地上面积150,000平方米、可经营面积135,000平方米、可售74,250平方米、持有出租60,750平方米。建设第1—3年，销售及出租第4年起，销售第6年完成；贷款第1年首次提款，第10年到期，第4—10年等额本金偿还。','',
       '| 核心/观察指标 | 基准 | 组合压力 |','|---|---:|---:|']
for label,row,pct in rows:lines.append(f'| {label} | {fmt(base[str(row)],pct)} | {fmt(stress[str(row)],pct)} |')
lines += ['', '不通过原因：'+baseline['scenarios'][0]['reasons'], '',
          '压力原因：'+baseline['scenarios'][1]['reasons'], '',
          '默认项目无融资开发投资60,920万元，周转资金500万元，建设期利息1,890万元，因此含融资费用总投资63,310万元。资本金32,810万元在建设期按需投入；初始运营周转投入由后续回款支持，不必机械地将总投资减贷款作为峰值资本金需求。', '',
          '最低DSCR只看当年经营可偿债现金，包含当年销售，不包含之前年度已留存的销售回款。默认销售回款留存可覆盖计划本息，贷款期末余额为0，但销售结束后的年度租金偿债覆盖仍低于1.20的模拟门槛。','',
          '## 已完成的验证','',
          '| 验证层次 | 是否完成 | 实际方法与证据 |',
          '|---|---|---|',
          '| 生成代码实际运行 | 是 | 完整生成器及 --recalculate 选项均实际运行成功；最终工作簿由该代码生成 |',
          '| 结构、公式引用及校验 | 是 | 10张表顺序正确、8,183个公式、3张图、数值输入及年度计划的数据验证覆盖；检查公式依赖图无循环、无外部链接 |',
          '| 独立数值核对 | 是 | 单独Python数值模型核对面积、逐年建设、销售、免租租金、运营及税、贷款、成本结转与折旧、两类现金流、留存/资本金/缺口及核心指标 |',
          f'| Excel实际重算 | 是 | Microsoft Excel {summary["excel_version"]}，独立隐藏实例；各次修改后 CalculateFullRebuild 并读取结果，最终保存重算缓存；非“只检查公式存在” |',
          f'| 修改参数联动 | 是 | 19组情景＋默认恢复，两个情景均核对，共{summary["independent_value_comparisons"]:,}个数值/结果比较；期限变化后图表点数随N变为N+1 |',
          '| 自动不达标原因 | 是 | 逐项状态与独立规则比较，并检查每个不通过或无法计算指标在原因文本中出现；补验基准通过但压力失败时的开关 |',
          '| Excel公式错误扫描 | 是 | 实际重算后的全部已用单元格扫描，Excel错误单元格0个；IRR无解使用明确文字，不替换成0 |',
          '| 视觉检查 | 是（范围见下） | Excel原生PDF渲染并查看各表开头主要区及3张图；另复看建设期存货/持有账面值与最终图表排版 |',
          '| 默认参数恢复 | 是 | 所有标量输入及31年建设/销售输入逐项恢复；最终交付仅保存默认示例 |', '',
          '独立IRR使用现金流折现因子多项式求根，NPV逐年折现加总，回收期逐年累计并插值，留存需求另用未来累计现金流最低值计算；没有使用Excel缓存反推预期值。金额数值比较采用约1e-6万元绝对容差及相对容差，决策采用0.01万元容差。','',
          '## 情景测试结果','',
          '以下均是实际改变Excel输入后的结果；每行同时重算内置组合压力情景，表中列出修改后基准区。额外乐观情景只检验“通过”逻辑，**没有成为默认示例**。','',
          '| 测试 | 项目IRR | 总投资（万元） | 所需资本金（万元） | 最大缺口（万元） | 总判断 |',
          '|---|---:|---:|---:|---:|---|']
for r in summary['cases']:
    if r['name']=='恢复默认':continue
    m=r['metrics']
    lines.append(f'| {r["name"]} | {fmt(m["6"],True)} | {fmt(m["16"])} | {fmt(m["13"])} | {fmt(m["14"])} | {r["decision"]} |')
lines += ['',
          '特别确认：售价下降10%与出租率下降10个百分点均降低回款；成本增加10%使非土地投资及分摊成本增加，贷款额度不变；利率提高至6.5%提高建设融资费用和资本金需要，而项目无融资IRR保持不变；资本金额度2亿元时识别12,810万元最大未落实缺口，不视为真实资金。','',
          '期限12年/30年：IRR、NPV、回收期只读取实际期限；周转资金在相应期末回收，图表分别13/31个年度点。期限8年及到期超出30年会待判断，且到期未还本金显示无法计算，避免越界公式错误。期末处置开启后按第N年持有资产净处置和所得税重新计算。无贷款时DSCR不适用。销售仅完成70%时余下30%继续为存货，不出租或自动变现。','',
          '出租推迟到第6年时，第4—5年仍支付固定成本和维护。新增面积一次免租，旧已出租面积不重复扣免租。持有面积上限不含任何出售面积。建设支出独立进入现金流，销售成本结转和折旧只影响利润及税基，不重复扣现金。分配前留存未来计划净支出，分配后没有再次追加理论资本金。','',
          '高固定成本情景的保本出租率超过100%，提示文字实际显示；无正现金流情景的两种IRR显示无法计算并待判断。基准通过、压力失败时，压力要求开关将总判断变为未通过。','',
          '## 范围与未完成的检查','',
          '- 完成的是Windows桌面Excel 16.0与本包明确列出的19组情景；未在LibreOffice、WPS、网页Excel、macOS或其他Excel版本重算，不能宣称跨软件已验证。',
          '- 视觉检查覆盖每个工作表的开头主要区和全部3张图，并复看变化区；没有逐页检查全部30年列的每一张纸面打印页，也没有穷举所有可能输入组合。',
          '- 本次Excel直接PNG图表导出返回空文件，因此改用Excel原生PDF渲染完成图表视觉检查；空PNG不作为验证证据或交付内容。',
          '- 不验证法律、审批、真实市场、复杂税种、税收优惠、详细工程预算、租户替换及实际融资审批；这些均不在用户要求的学习简化模型范围内。',
          '- 任意覆盖公式、删除行列、粘贴绕过数据验证或引入不合理参数，会超出本次验证范围；看板检查用于发现常见问题，不能代替所有边界验证。','',
          '完整的验证源代码和机器可读摘要随包提供。默认数据恢复、公式错误扫描、独立计算及参数联动已实际执行，未完成项目如上明确列出。','']
(out/'验证说明.md').write_text('\n'.join(lines),encoding='utf-8')
files=['产业园投资测算.xlsx','generate_excel.py','validate_model.py','requirements.txt','README.md','验证说明.md','验证摘要.json']
zip_path=out/'产业园投资测算_完整交付包.zip'
assert not zip_path.exists(),'Avoid overwriting an existing package'
with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
    for name in files:z.write(out/name,'产业园投资测算交付/'+name)
with zipfile.ZipFile(zip_path) as z:assert z.testzip() is None
for name in files+[zip_path.name]:
    path=out/name
    print(name,path.stat().st_size,hashlib.sha256(path.read_bytes()).hexdigest()[:16])
print(v.structural_check(out/'产业园投资测算.xlsx'))
