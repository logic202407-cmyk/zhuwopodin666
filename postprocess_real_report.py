from pathlib import Path
import importlib
import requests
import fitz
from PIL import Image
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
REAL = ROOT / 'real_figures'
REAL.mkdir(exist_ok=True)

# Generate the full academic report text and initial layout first.
import generate_report  # noqa: F401
SRC = ROOT / '校园电动车充电安全监测与自动断电设计_课程设计报告.docx'
OUT = ROOT / '校园电动车充电安全监测与自动断电设计_课程设计报告_真实截图终稿.docx'

URLS = {
    'front': 'https://api.pdfmaker.to/download/51787e75-f4d3-46a2-b6d9-9396e8974b5f',
    'back': 'https://api.pdfmaker.to/download/43420cab-875c-4c40-9681-3b72f9b50a5c',
    'basic_total': 'https://api.pdfmaker.to/download/553102af-b2b0-4141-a28b-b5c31a4e158d',
    'basic_current': 'https://api.pdfmaker.to/download/c7fb2e1e-3c16-4115-a5be-3c7be690f973',
    'ov_local': 'https://api.pdfmaker.to/download/92d6c5ff-95e4-4da5-8a79-14cdc12b53af',
    'oc_local': 'https://api.pdfmaker.to/download/bd97121b-d037-435d-9bb7-3e24fcb9a2dd',
    'ot_local': 'https://api.pdfmaker.to/download/da200381-1d98-4085-863b-9e9f7f8bec3d',
    'param_normal': 'https://api.pdfmaker.to/download/440cceb7-61b4-4f77-a559-7261e7648ac9',
    'param_critical': 'https://api.pdfmaker.to/download/907ab58e-0b62-4d0e-887c-a7ee8b19b61b',
    'param_current': 'https://api.pdfmaker.to/download/09778e07-d468-4abd-8e86-02e5bc5df107',
    'temp_plot': 'https://api.pdfmaker.to/download/e0ef25e9-46f6-4295-8821-248404bfe0e9',
    'temp_settings': 'https://api.pdfmaker.to/download/81248027-0817-4ad3-967c-41006a2e5e3f',
    'temp_table': 'https://api.pdfmaker.to/download/54dfaef2-3d67-4198-ae26-9e7f6004acd4',
    'worst': 'https://api.pdfmaker.to/download/2b174b88-9864-4706-8631-7f3e41ed6fec',
    'mc': 'https://api.pdfmaker.to/download/ccac1152-b1d8-4220-b984-cf559fdbe95e',
    'smoke_initial': 'https://api.pdfmaker.to/download/7ceab7d0-b036-4cf5-bfc3-6cfe3860a2e9',
    'smoke_final': 'https://api.pdfmaker.to/download/58b27a93-572e-4552-bac9-a6efeed62f92',
}

def render_pdf(name, url):
    pdf_path = REAL / f'{name}.pdf'
    png_path = REAL / f'{name}.png'
    response = requests.get(url, timeout=90)
    response.raise_for_status()
    pdf_path.write_bytes(response.content)
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), alpha=False)
    pix.save(png_path)
    pdf.close()
    return png_path

IMAGES = {name: render_pdf(name, url) for name, url in URLS.items()}

doc = Document(SRC)

# Word font helpers.
def set_font(run, size=10.5, bold=False, cn='宋体'):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def width_for(path, max_cm=15.6, max_h_cm=18.0):
    with Image.open(path) as im:
        w, h = im.size
    cm = max_cm
    if cm * h / w > max_h_cm:
        cm = max_h_cm * w / h
    return cm

def find_para(starts):
    for p in doc.paragraphs:
        if p.text.strip().startswith(starts):
            return p
    raise ValueError(f'Paragraph not found: {starts}')

def add_picture_before(target_para, path, caption):
    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.paragraph_format.keep_with_next = True
    pic.paragraph_format.space_before = Pt(3)
    pic.paragraph_format.space_after = Pt(1)
    pic.add_run().add_picture(str(path), width=Cm(width_for(path)))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.keep_with_next = False
    cap.paragraph_format.space_before = Pt(0)
    cap.paragraph_format.space_after = Pt(6)
    set_font(cap.add_run(caption), 10.5)
    target_para._p.addprevious(pic._p)
    target_para._p.addprevious(cap._p)

def add_body_before(target_para, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run(text), 12)
    target_para._p.addprevious(p._p)

def replace_figure(old_caption_start, path, new_caption):
    cap = find_para(old_caption_start)
    prev = cap._p.getprevious()
    # Old figure is a paragraph in the original report.
    if prev is not None and prev.tag == qn('w:p'):
        prev.getparent().remove(prev)
    cap.text = ''
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(6)
    set_font(cap.add_run(new_caption), 10.5)
    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.paragraph_format.keep_with_next = True
    pic.paragraph_format.space_after = Pt(1)
    pic.add_run().add_picture(str(path), width=Cm(width_for(path)))
    cap._p.addprevious(pic._p)

# Replace every schematic/simulation image already in the report with the user's real screenshot.
replace_figure('图 3.1', IMAGES['front'], '图 3.1 前端过压、过流和过温检测模块原理图')
replace_figure('图 3.2', IMAGES['back'], '图 3.2 故障合成、状态等效支路与自动断电模块原理图')
replace_figure('图 3.3', IMAGES['basic_total'], '图 3.3 基本瞬态仿真总波形')
replace_figure('图 3.4', IMAGES['basic_current'], '图 3.4 等效充电负载电流瞬态波形')
replace_figure('图 3.5', IMAGES['temp_plot'], '图 3.5 −20～85 ℃温度扫描总波形')
replace_figure('图 3.6', IMAGES['param_critical'], '图 3.6 过压参考阈值临界范围参数扫描结果')
replace_figure('图 3.7', IMAGES['worst'], '图 3.7 最坏条件下系统保护波形')
replace_figure('图 3.8', IMAGES['mc'], '图 3.8 最大负载电流的 Monte Carlo 统计分布')

# Add supporting real screenshots to the appropriate sections.
target = find_para('3.8 高温条件分析')
add_body_before(target, '为进一步核对三类故障的动作时序，分别对过压、过流和过温窗口进行局部放大。局部波形表明，2～4 ms、6～8 ms 和 10～12 ms 内，FAULT_RAW 均能升高、EN 均能降低，负载电流同步降至近似零；故障结束后系统自动恢复。')
add_picture_before(target, IMAGES['ov_local'], '图 3.9 过压故障局部保护波形')
add_picture_before(target, IMAGES['oc_local'], '图 3.10 过流故障局部保护波形')
add_picture_before(target, IMAGES['ot_local'], '图 3.11 过温故障局部保护波形')

# Add the temperature setup/data screenshots before parameter sweep.
target = find_para('3.9 参数扫描分析')
add_picture_before(target, IMAGES['temp_settings'], '图 3.12 温度扫描仿真设置')
add_picture_before(target, IMAGES['temp_table'], '图 3.13 温度扫描测量数据')
add_body_before(target, '温度扫描取 −20 ℃、27 ℃、40 ℃、60 ℃和 85 ℃。测量数据表明，正常状态下 I(RLOAD) 约由 22.427 mA 变化至 22.413 mA，变化量仅约 0.014 mA，相对变化约 0.062%；EN 高电平仍保持在 11.24 V 以上，FAULT_RAW 无故障低电平远低于 2.5 V 判定阈值。')

# Add the full parameter-sweep evidence before worst case.
target = find_para('3.10 最坏条件分析')
add_picture_before(target, IMAGES['param_normal'], '图 3.14 过压阈值正常范围参数扫描结果')
add_picture_before(target, IMAGES['param_current'], '图 3.15 临界阈值扫描下的负载电流结果')
add_body_before(target, '在 3.8～4.2 V 的正常阈值范围内，4.5 V 过压输入均能触发保护；在 4.3～4.7 V 临界扫描中，当 VOV 高于 4.5 V 时，2～4 ms 内负载电流不再被切断，说明阈值过高会造成漏保护。')

# Insert the initial and improved Smoke results before section 3.13.
target = find_para('3.13')
add_picture_before(target, IMAGES['smoke_initial'], '图 3.16 初始额定功率设置下的 Smoke 分析结果')
add_body_before(target, '初次 Smoke 分析中，RLOAD、R4 和 R5 的 PDM 功耗项目出现红色过应力提示。RLOAD 峰值功耗约为 110.5824 mW，R4、R5 峰值功耗约为 53.87 mW，但在既定降额规则下安全裕量不足。因此将 RLOAD 的 POWER 属性由 0.25 W 提高至 0.5 W，将 R4、R5 由 0.125 W 提高至 0.25 W。')
add_picture_before(target, IMAGES['smoke_final'], '图 3.17 提高额定功率后的 Smoke 分析结果')
add_body_before(target, '修改后重新运行 Smoke，红色过应力项消失。RLOAD 峰值功耗利用率降至约 83%，R4 和 R5 均降至约 79%，表明提高额定功率后已获得合理功耗裕量。窗口中的部分黄色 100% 项与运算放大器模型额定属性或 RCA、RJC 热参数缺失有关，因此本次结论主要针对已正确配置额定参数的电阻功耗。')

# Correct the abstract and conclusion so they include the completed Smoke re-test.
for p in doc.paragraphs:
    if '初次 Smoke 分析' in p.text and '提出' in p.text:
        old = p.text
        new = old.replace('初次 Smoke 分析发现 RLOAD、R4、R5 的功耗降额裕量不足，据此提出提高额定功率的改进方案。', '初次 Smoke 分析发现 RLOAD、R4、R5 的功耗降额裕量不足；提高额定功率并重新分析后，红色过应力项消失。')
        if new != old:
            p.text = new
            for r in p.runs:
                set_font(r, 12)

# Verify that no generated schematic/simulation caption remains.
for p in doc.paragraphs:
    p.text = p.text.replace('根据最终 OrCAD 工程整理的', '').replace('根据 PSpice 仿真条件整理的', '')
    for r in p.runs:
        r.font.color.rgb = RGBColor(0, 0, 0)

# Metadata and output.
doc.core_properties.title = '校园电动车充电安全监测与自动断电设计'
doc.core_properties.subject = '电子电路 CAD 技术课程设计报告（真实 OrCAD/PSpice 截图版）'
doc.core_properties.author = '郑翔元'
doc.save(OUT)
print(f'Created: {OUT}')
print(f'Real screenshots embedded: {len(IMAGES)}')
