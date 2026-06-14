from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = Path('校园电动车充电安全监测与自动断电设计_课程设计报告_真实截图终稿.docx')
OUT = Path('校园电动车充电安全监测与自动断电设计_课程设计报告_团队分工与参考文献已更新.docx')

doc = Document(SRC)

NEW_REFS = [
    '[1] 欧阳明高，冯旭宁，韩雪冰，等. 锂离子电池全生命周期安全性研究进展[J]. 机械工程学报，2018，54（16）：1-17.',
    '[2] 冯旭宁，欧阳明高，刘晓菊，等. 动力锂离子电池热安全性研究进展[J]. 汽车安全与节能学报，2016，7（4）：350-365.',
    '[3] 王青松，孙金华，褚亚洲，等. 锂离子电池热失控危险性分析[J]. 中国安全科学学报，2011，21（6）：88-92.',
    '[4] 孙逢春，王震坡，朱家琏. 电动汽车动力电池管理系统研究[J]. 汽车工程，2004，26（5）：501-505.',
    '[5] 王震坡，孙逢春. 电动汽车用动力电池管理系统研究[J]. 北京理工大学学报，2004，24（11）：1013-1017.',
    '[6] FENG X, OUYANG M, LIU X, et al. Thermal runaway mechanism of lithium ion battery for electric vehicles: A review[J]. Energy Storage Materials, 2018, 10: 246-267.',
    '[7] WANG Q, PING P, ZHAO X, et al. Thermal runaway caused fire and explosion of lithium ion battery[J]. Journal of Power Sources, 2012, 208: 210-224.',
    '[8] LU L, HAN X, LI J, et al. A review on the key issues for lithium-ion battery management in electric vehicles[J]. Journal of Power Sources, 2013, 226: 272-288.',
    '[9] RAHIMI-EICHI H, OJHA U, BARONTI F, et al. Battery management system: An overview of its application in the smart grid and electric vehicles[J]. IEEE Industrial Electronics Magazine, 2013, 7（2）: 4-16.',
    '[10] CADENCE DESIGN SYSTEMS. PSpice User Guide[EB/OL]. San Jose: Cadence Design Systems, 2024.'
]

P1 = '随着校园电动车数量不断增加，宿舍区、教学区和生活区周边的集中充电需求逐渐增大。电动车充电装置通常需要长时间连续工作，若在充电过程中出现充电电压过高、充电电流过大、电池温度异常升高或散热条件变差等情况，可能导致电池性能下降、器件过应力，甚至引发安全隐患。锂离子电池在过充、高温和滥用条件下可能诱发热失控、燃烧甚至爆炸等安全问题，因此充电过程中的电压、电流和温度监测具有重要意义[1-3,6-7]。现有电动汽车电池管理系统通常通过电压、电流、温度等多参数采集实现状态监测、故障判断和保护控制[4-5,8-9]。因此，充电系统不仅需要完成基本的能量传输功能，还应具备对关键状态量的检测、故障判断和自动切断能力。'
P2 = '电子电路 CAD 技术能够在实际制作电路之前完成原理图绘制、元件参数设置、网络连接检查以及仿真验证。利用 OrCAD Capture 和 PSpice 可以对电路进行瞬态分析、温度分析、参数扫描、容差分析和电应力分析，从而在设计阶段判断电路功能是否正确、参数设置是否合理以及器件是否具有足够安全裕量[10]。本设计以校园电动车充电安全为应用背景，采用比较器、二极管逻辑和晶体管开关等常见模拟电子电路模块，设计一种具有过压、过流、过温检测和自动断电功能的充电安全监测电路。'

def set_run_font(run, size=12, bold=False, cn='宋体'):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def replace_paragraph_text(paragraph, new_text, size=12, bold=False, cn='宋体'):
    for r in paragraph.runs:
        r.text = ''
    if paragraph.runs:
        r = paragraph.runs[0]
    else:
        r = paragraph.add_run()
    r.text = new_text
    set_run_font(r, size=size, bold=bold, cn=cn)

# Restore the cover information so the regenerated document matches the team report.
for table in doc.tables:
    if len(table.rows) >= 2 and len(table.columns) >= 2:
        labels = [row.cells[0].text.strip().replace('：','') for row in table.rows]
        if any('姓名' in x for x in labels) and any('学号' in x for x in labels):
            for row in table.rows:
                label = row.cells[0].text.strip()
                if '姓名' in label:
                    replace_paragraph_text(row.cells[1].paragraphs[0], '郑翔元，李凯', 14)
                elif '学号' in label:
                    replace_paragraph_text(row.cells[1].paragraphs[0], '2410160076\n2410110188', 14)
            break

# Replace only the two paragraphs in 1.1 so all references are cited at their suggested positions.
for p in doc.paragraphs:
    txt = p.text.strip()
    if txt.startswith('随着校园电动车数量'):
        replace_paragraph_text(p, P1, 12)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Pt(24)
        p.paragraph_format.line_spacing = 1.5
    elif txt.startswith('电子电路 CAD 技术能够'):
        replace_paragraph_text(p, P2, 12)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Pt(24)
        p.paragraph_format.line_spacing = 1.5

# Insert one team-division page after the catalog and before Chapter 1.
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

# Avoid duplicate insertion if this script is rerun.
if not any(p.text.strip() == '团队分工' for p in doc.paragraphs):
    chapter1 = next(p for p in doc.paragraphs if p.text.strip().startswith('1 设计的目的'))
    insert_para_before(chapter1).add_run().add_break()  # page break before team page if needed
    title = insert_para_before(chapter1, '团队分工', WD_ALIGN_PARAGRAPH.CENTER, False, 16, True, '黑体', 12, 12)
    insert_para_before(chapter1, '本课程设计由两名成员协作完成。团队采用“前端检测—后端控制—联合仿真—报告整理”的分工方式，在保持系统总体方案一致的基础上分别完成对应模块，并在后期进行工程合并、仿真验证和结果分析。具体分工如表所示。', WD_ALIGN_PARAGRAPH.JUSTIFY, True, 12, False, '宋体')
    # Insert table before Chapter 1.
    t = doc.add_table(rows=1, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ['成员', '主要分工', '具体工作内容']
    for i, h in enumerate(headers):
        set_cell(t.rows[0].cells[i], h, True)
    rows = [
        ['郑翔元', '整体思路设计、后端电路设计、实验报告撰写', '确定系统总体方案与保护逻辑；完成故障合成、EN 控制、状态等效支路与 Q3 负载断电模块设计；整理仿真结果并完成实验报告撰写。'],
        ['李凯', '前端电路设计、电路仿真', '完成过压、过流、过温检测模块设计；配合完成前后端原理图合并；完成瞬态、温度、参数、最坏条件、Monte Carlo 和 Smoke 等仿真验证。']
    ]
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            set_cell(cells[i], v, False)
    # Three-line table style.
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
    insert_para_before(chapter1).add_run().add_break()  # page break after team page

# Replace the reference list, preserving the reference paragraph formatting.
ref_heading_index = next(i for i, p in enumerate(doc.paragraphs) if p.text.strip() == '参考文献')
old_ref_paras = [p for p in doc.paragraphs[ref_heading_index + 1:] if p.text.strip().startswith('[')]
if not old_ref_paras:
    raise RuntimeError('No existing references found.')
for i, p in enumerate(old_ref_paras):
    if i < len(NEW_REFS):
        replace_paragraph_text(p, NEW_REFS[i], 10.5)
        p.paragraph_format.hanging_indent = Pt(21)
        p.paragraph_format.line_spacing = 1.5
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
