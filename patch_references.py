from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = Path('校园电动车充电安全监测与自动断电设计_课程设计报告_真实截图终稿.docx')
OUT = Path('校园电动车充电安全监测与自动断电设计_课程设计报告_参考文献已更新.docx')

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

def replace_text_in_paragraph(paragraph, old, new):
    full = ''.join(run.text for run in paragraph.runs)
    if old not in full:
        return False
    replaced = full.replace(old, new)
    if paragraph.runs:
        paragraph.runs[0].text = replaced
        for run in paragraph.runs[1:]:
            run.text = ''
    else:
        paragraph.add_run(replaced)
    return True

# Restore the cover information so the regenerated document matches the user's uploaded file.
for table in doc.tables:
    if len(table.rows) >= 2 and len(table.columns) >= 2:
        labels = [row.cells[0].text.strip().replace('：','') for row in table.rows]
        if any('姓名' in x for x in labels) and any('学号' in x for x in labels):
            for row in table.rows:
                label = row.cells[0].text.strip()
                if '姓名' in label:
                    p = row.cells[1].paragraphs[0]
                    if p.runs:
                        p.runs[0].text = '郑翔元，李凯'
                        for r in p.runs[1:]: r.text = ''
                    else:
                        p.add_run('郑翔元，李凯')
                elif '学号' in label:
                    p = row.cells[1].paragraphs[0]
                    if p.runs:
                        p.runs[0].text = '2410160076\n2410110188'
                        for r in p.runs[1:]: r.text = ''
                    else:
                        p.add_run('2410160076\n2410110188')
            break

# Change only the existing citation markers in the body.
for p in doc.paragraphs:
    if '保护功能[1-3]' in p.text:
        replace_text_in_paragraph(p, '[1-3]', '[1-9]')
    if 'Smoke 等分析[4-8]' in p.text:
        replace_text_in_paragraph(p, '[4-8]', '[10]')

# Locate the reference list.
ref_heading_index = next(i for i,p in enumerate(doc.paragraphs) if p.text.strip() == '参考文献')
old_ref_paras = [p for p in doc.paragraphs[ref_heading_index+1:] if p.text.strip().startswith('[')]
if not old_ref_paras:
    raise RuntimeError('No existing references found.')

# Replace the existing reference paragraphs while retaining their formatting.
for i, p in enumerate(old_ref_paras):
    if i < len(NEW_REFS):
        replace_text_in_paragraph(p, p.text, NEW_REFS[i])
    else:
        p._element.getparent().remove(p._element)

# Add any extra references by cloning the last formatted reference paragraph.
anchor = old_ref_paras[min(len(old_ref_paras), len(NEW_REFS)) - 1]._p
for text in NEW_REFS[len(old_ref_paras):]:
    clone = deepcopy(anchor)
    # Replace all text nodes in the cloned paragraph with a single text node.
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

doc.save(OUT)
print(OUT)
