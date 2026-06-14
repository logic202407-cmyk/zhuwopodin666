from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = Path('校园电动车充电安全监测与自动断电设计_课程设计报告_真实截图终稿.docx')
OUT = Path('校园电动车充电安全监测与自动断电设计_课程设计报告_团队分工与参考文献已更新.docx')

doc = Document(SRC)

# References are taken from the user-provided “参考文献内容汇编-高切合度版.docx”.
NEW_REFS = [
    '[1] 王钢，邵其专，李晓华. 电路容差对继电保护装置可靠性影响的分析方法[J]. 华南理工大学学报(自然科学版)，2007，35(3)：61-65+76.',
    '[2] 王淑娟，沙友涛，张辉，翟国富. 基于EDA技术的可靠性容差设计方法[J]. Journal of Zhejiang University SCIENCE A，2007，8(3)：426-432.',
    '[3] 周治良，郭成统. 电动汽车模式二充电剩余电流检测设计[J]. 电器与能效管理技术，2025(8)：52-57+84. DOI:10.16628/j.cnki.2095-8188.2025.08.008.',
    '[4] WANG Y, HAN Z, SHI J. 锂离子电池保护IC中的比较器电路设计[C]//2014 IEEE ICIEA. Hangzhou: IEEE，2014：1427-1432. DOI:10.1109/ICIEA.2014.6931392.',
    '[5] GHOSH S, BAKSHI S, GHOSH S K, et al. 基于LM324的低成本欠压过流保护装置[J]. 2016.',
    '[6] 雷蕾，吕亚兵. 电动自行车充电保护系统设计研究[J]. 消防科学与技术，2020，39(9)：1202-1205.',
    '[7] 集成运放温度报警电路[J]. 电子世界，2014(16).',
    '[8] LDO过流与温度保护电路的分析与设计[J]. 电子器件，2006，29(1)：127-130.',
    '[9] 山东华宇工学院. 基于智能检测的电动车充电保护装置设计[J]. 现代信息科技，2023，7(19)：195-198.',
    '[10] ANALOG DEVICES. Overload circuitry protects batteries and power supplies[EB/OL]. https://www.analog.com/en/resources/design-notes/overload-circuitry-protects-batteries-and-power-supplies.html.',
    '[11] CADENCE. PSpice User Guide[EB/OL]. San Jose，2024.',
    '[12] 贾新章，武岳山. 电子电路CAD技术——基于OrCAD 17.2[M]. 西安：西安电子科技大学出版社，2018.',
    '[13] 刘明山，周原. OrCAD和PSpice电路设计与仿真[M]. 北京：机械工业出版社，2019.'
]

# Body paragraphs with citation positions matched to the user's reference-compilation suggestions.
REPLACEMENTS = {
    '随着校园电动车数量': '随着校园电动车数量不断增加，宿舍区、教学区和生活区周边的集中充电需求逐渐增大。电动车充电装置通常需要长时间连续工作，若在充电过程中出现充电电压过高、充电电流过大、电池温度异常升高或散热条件变差等情况，可能导致电池性能下降、器件过应力，甚至引发安全隐患。已有电动自行车充电保护系统研究表明，电压、电流和温度是充电保护中需要重点监测的关键状态量[6,9]；分立元件检测方案在降低软件失效风险和缩短响应时间方面也具有一定优势[3]。因此，充电系统不仅需要完成基本能量传输，还应具备对关键状态量的检测、故障判断和自动切断能力。',
    '电子电路 CAD 技术能够': '电子电路 CAD 技术能够在实际制作电路之前完成原理图绘制、元件参数设置、网络连接检查以及仿真验证。利用 OrCAD Capture 和 PSpice 可以对电路进行瞬态分析、温度分析、参数扫描、容差分析和电应力分析，从而在设计阶段判断电路功能是否正确、参数设置是否合理以及器件是否具有足够安全裕量[11-13]。本设计以校园电动车充电安全为应用背景，采用比较器、二极管逻辑和晶体管开关等常见模拟电子电路模块，设计一种具有过压、过流、过温检测和自动断电功能的充电安全监测电路。',
    '本设计不是对真实高功率充电主回路': '本设计不是对真实高功率充电主回路的直接缩比，而是对安全监测和控制逻辑的等效验证。采用 VPULSE 模拟传感器输出，可将电压、电流和温度统一表示为比较器可处理的电压量，使研究重点集中在故障判定、逻辑合成和断电执行。与单片机控制相比，纯模拟比较器和分立元件方案不依赖程序运行状态，响应路径更直接[3]；以 LM324 和晶体管构建保护装置也具有结构简单、成本低和便于仿真的特点[5]。三参数联合保护相对于单一保护具有更完整的故障覆盖范围；加入高温、容差、最坏条件和电应力分析，可使结论从标称功能正确扩展到参数变化时仍具可靠性和改进依据。',
    '检测环节可采用单片机': '检测环节可采用单片机采集后软件判断，也可采用模拟比较器直接判定。单片机方案便于显示和通信，但存在程序异常、上电自检和响应延迟等问题；采用分立元件检测的保护方案能够降低软件失效风险并提高响应速度[3]。同时，锂离子电池保护 IC 中普遍采用比较器完成过充、过放和过流检测，说明模拟比较器方案具有成熟应用基础[4]。同类电动自行车充电保护系统也通常围绕电压、电流和温度等参数建立保护逻辑[6]，因此本课题选择 LM324 比较器和二极管逻辑实现硬件判定。',
    '系统由前端检测': '系统由前端检测、故障合成、充电允许控制、状态等效支路和负载断电模块组成。相关锂离子保护 IC 采用比较器完成电压和电流阈值判断[4]，而基于 LM324 的低成本保护装置也验证了“四运放比较器+晶体管开关”架构的可行性[5]。本课题将 LM324 的四路运放分别用于三通道故障检测和 EN 生成，将 Q2N2222 用作负载开关，使整体电路具备清晰的模块化结构。传感器等效信号接 LM324 同相端，参考电压接反相端，输入超过阈值时相应故障输出变高；D1、D2、D3 将三路故障汇接至 FAULT_RAW，R1=10 kΩ 在无故障时下拉该节点。',
    '过压通道 V2': '过压通道 V2 由 3.5 V 跳变至 4.5 V，TD=2 ms、PW=2 ms，VOV=4 V；过流通道 V4 由 1.5 V 跳变至 3.5 V，TD=6 ms、PW=2 ms，VOC=2.5 V；过温通道 V6 由 2.0 V 跳变至 4.0 V，TD=10 ms、PW=2 ms，VOT=3 V。锂离子电池保护 IC 中常通过开环比较器将电池状态量与参考阈值比较[4]；温度报警电路也常将热敏元件输出转化为电压后送入运放比较器进行阈值判断[7]。本课题采用相同思想，用三路 LM324 完成 OV、OC 和 OT 检测。',
    'D1～D3 均采用': 'D1～D3 均采用 1N4148，阳极分别连接 OV、OC、OT，阴极汇接 FAULT_RAW。U4A 同相端接 2.5 V，反相端接 FAULT_RAW，正常时输出 EN 高电平，故障时输出低电平。Q1 基极经 R3=10 kΩ 接 FAULT_RAW；Q2、Q3 基极分别经 R6=10 kΩ、RB=10 kΩ 接 EN。过流与温度保护电路通常利用比较器输出控制后级开关或保护支路[8]，过载保护设计中也常采用比较器与锁存/关断结构实现快速切断[10]。本课题用二极管或门和 Q2N2222 开关管完成多故障合成与等效断电。R4、R5 均为 330 Ω；D4、D5 为普通 1N4148，不是 LED。最终电路没有继电器和蜂鸣器。',
    '将 VOV 设为全局参数': '将 VOV 设为全局参数。参数扫描和容差设计研究表明，通过改变关键参数并观察输出是否满足设计规范，可以判断保护电路对参数偏移的敏感程度[1,2]。本课题首先按 3.8～4.2 V、步长 0.1 V 扫描，因各阈值均低于 4.5 V 故障输入，全部曲线均可触发保护；随后扩大到 4.3～4.7 V 的临界范围，以观察阈值接近故障输入时的漏保护风险。',
    '最坏条件设置为': '最坏条件设置为 85 ℃、VCC12=10.8 V、VCC5=4.75 V、VOV=4.2 V、VOC=2.7 V、VOT=3.2 V、RLOAD=242 Ω，三路故障高电平分别为 4.3 V、2.8 V 和 3.3 V，仅比阈值高 0.1 V。可靠性容差分析通常需要在不利参数组合下检验输出是否仍满足设计规范[1,2]，因此该组设置同时包含高温、低供电、阈值偏高和负载偏差等因素。正常理论电流为 4.75/242=19.63 mA。',
    '对关键电阻设置': '对关键电阻设置 5%～10% 容差并进行 100 次随机运行，测量 Max/Min(V(EN))、Max/Min(V(FAULT_RAW)) 和 Max/Min(I(RLOAD))。已有研究将 PSpice 电路仿真与 Monte Carlo 方法结合，用输出是否满足容差或设计规范作为可靠性判据[1]；基于 EDA 的容差设计方法也强调通过统计均值、标准差和 Yield 评价电路可靠性[2]。本课题以 FAULT_RAW、EN 和负载电流的正确动作为判据进行统计验证。',
    '“Smoke Analysis succeeded”': '“Smoke Analysis succeeded”仅表示分析运行成功，必须结合颜色和 % Max 判断器件安全。过载保护设计笔记强调保护电路应关注关断路径和器件安全裕量[10]，而 PSpice 高级分析资料可用于建立电应力、降额与 Smoke 检查流程[11]。因此，本课题在瞬态仿真基础上进一步检查 RLOAD、R4 和 R5 的功耗利用率。初次结果中 RLOAD、R4 和 R5 的 PDM 项出现红色过应力。',
    '本课程设计完成了校园电动车': '本课程设计完成了校园电动车充电安全监测与自动断电电路的方案比较、原理图设计、参数计算和多层次仿真。真实 OrCAD 原理图与 PSpice 波形表明，三类故障均能被正确识别，任一故障均可使 FAULT_RAW 变高、EN 变低并切断 RLOAD。温度扫描中负载电流变化约 0.062%，最坏条件下未出现漏动作，100 次 Monte Carlo 结果保持稳定。该多层次验证过程与保护电路可靠性评估和 EDA 容差设计方法相一致[1,2]。Smoke 初次分析发现功耗降额问题，提高额定功率后红色过应力项消失。由此可认为该设计达到了多参数监测和自动断电的预期目标。'
}

def set_run_font(run, size=12, bold=False, cn='宋体'):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def replace_paragraph_text(paragraph, new_text, size=12, bold=False, cn='宋体'):
    for r in paragraph.runs:
        r.text = ''
    r = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    r.text = new_text
    set_run_font(r, size=size, bold=bold, cn=cn)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.first_line_indent = Pt(24)
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.space_after = Pt(0)

# Restore team cover information.
for table in doc.tables:
    if len(table.rows) >= 2 and len(table.columns) >= 2:
        labels = [row.cells[0].text.strip().replace('：','') for row in table.rows]
        if any('姓名' in x for x in labels) and any('学号' in x for x in labels):
            for row in table.rows:
                label = row.cells[0].text.strip()
                if '姓名' in label:
                    replace_paragraph_text(row.cells[1].paragraphs[0], '郑翔元，李凯', 14)
                    row.cells[1].paragraphs[0].paragraph_format.first_line_indent = Pt(0)
                    row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                elif '学号' in label:
                    replace_paragraph_text(row.cells[1].paragraphs[0], '2410160076\n2410110188', 14)
                    row.cells[1].paragraphs[0].paragraph_format.first_line_indent = Pt(0)
                    row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            break

# Replace selected body paragraphs and insert exact citation positions from the provided reference compilation.
for p in doc.paragraphs:
    txt = p.text.strip()
    for prefix, replacement in REPLACEMENTS.items():
        if txt.startswith(prefix):
            replace_paragraph_text(p, replacement, 12)
            break

# Insert one team-division page after catalog and before Chapter 1.
def insert_para_before(anchor_para, text='', align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, size=12, bold=False, cn='宋体', before=0, after=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    if first_indent:
        p.paragraph_format.first_line_indent = Pt(24)
    if text:
        set_run_font(p.add_run(text), size=size, bold=bold, cn=cn)
    anchor_para._p.addprevious(p._p)
    return p

def set_cell(cell, text, bold=False):
    cell.text = ''
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(0)
    set_run_font(p.add_run(text), size=10.5, bold=bold)

if not any(p.text.strip() == '团队分工' for p in doc.paragraphs):
    chapter1 = next(p for p in doc.paragraphs if p.text.strip().startswith('1 设计的目的'))
    insert_para_before(chapter1).add_run().add_break(WD_BREAK.PAGE)
    insert_para_before(chapter1, '团队分工', WD_ALIGN_PARAGRAPH.CENTER, False, 16, True, '黑体', 12, 12)
    insert_para_before(chapter1, '本课程设计由两名成员协作完成。团队采用“前端检测—后端控制—联合仿真—报告整理”的分工方式，在保持系统总体方案一致的基础上分别完成对应模块，并在后期进行工程合并、仿真验证和结果分析。具体分工如表所示。', WD_ALIGN_PARAGRAPH.JUSTIFY, True, 12, False, '宋体')
    t = doc.add_table(rows=1, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['成员', '主要分工', '具体工作内容']):
        set_cell(t.rows[0].cells[i], h, True)
    rows = [
        ['郑翔元', '整体思路设计、后端电路设计、实验报告撰写', '确定系统总体方案与保护逻辑；完成故障合成、EN 控制、状态等效支路与 Q3 负载断电模块设计；整理仿真结果并完成实验报告撰写。'],
        ['李凯', '前端电路设计、电路仿真', '完成过压、过流、过温检测模块设计；配合完成前后端原理图合并；完成瞬态、温度、参数、最坏条件、Monte Carlo 和 Smoke 等仿真验证。']
    ]
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            set_cell(cells[i], v, False)
    tblPr = t._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    tblPr.append(borders)
    for edge in ['top', 'bottom']:
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '10')
        e.set(qn('w:color'), '000000')
        borders.append(e)
    for edge in ['left', 'right', 'insideH', 'insideV']:
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'nil')
        borders.append(e)
    for c in t.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:color'), '000000')
        tcBorders.append(bottom)
        tcPr.append(tcBorders)
    chapter1._p.addprevious(t._tbl)
    insert_para_before(chapter1).add_run().add_break(WD_BREAK.PAGE)

# Replace the reference list.
ref_heading_index = next(i for i, p in enumerate(doc.paragraphs) if p.text.strip() == '参考文献')
old_ref_paras = [p for p in doc.paragraphs[ref_heading_index + 1:] if p.text.strip().startswith('[')]
if not old_ref_paras:
    raise RuntimeError('No existing references found.')
for i, p in enumerate(old_ref_paras):
    if i < len(NEW_REFS):
        replace_paragraph_text(p, NEW_REFS[i], 10.5)
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.hanging_indent = Pt(21)
    else:
        p._element.getparent().remove(p._element)
anchor = old_ref_paras[min(len(old_ref_paras), len(NEW_REFS)) - 1]._p
for text in NEW_REFS[len(old_ref_paras):]:
    clone = deepcopy(anchor)
    texts = clone.findall('.//' + qn('w:t'))
    if texts:
        texts[0].text = text
        for t in texts[1:]:
            t.text = ''
    else:
        run = OxmlElement('w:r')
        txt = OxmlElement('w:t')
        txt.text = text
        run.append(txt)
        clone.append(run)
    anchor.addnext(clone)
    anchor = clone

for p in doc.paragraphs:
    for r in p.runs:
        r.font.color.rgb = RGBColor(0, 0, 0)

doc.save(OUT)
print(OUT)
