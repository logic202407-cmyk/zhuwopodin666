from pathlib import Path
import math
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Circle, Polygon, FancyArrowPatch
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
OUT = ROOT / '校园电动车充电安全监测与自动断电设计_课程设计报告.docx'
FIG = ROOT / 'figures'
FIG.mkdir(exist_ok=True)

# ---------------- figures ----------------
def save_block_diagram():
    fig, ax = plt.subplots(figsize=(14,4.4))
    ax.set_xlim(0,14); ax.set_ylim(0,4.4); ax.axis('off')
    boxes = [
        (0.2,2.9,2.2,1.0,'VBAT_SENSE\nVoltage sample'),
        (0.2,1.7,2.2,1.0,'VI_SENSE\nCurrent sample'),
        (0.2,0.5,2.2,1.0,'VTEMP\nTemperature sample'),
        (3.1,2.9,2.3,1.0,'LM324 comparator\nOV'),
        (3.1,1.7,2.3,1.0,'LM324 comparator\nOC'),
        (3.1,0.5,2.3,1.0,'LM324 comparator\nOT'),
        (6.4,1.45,2.2,1.45,'Diode OR\nFAULT_RAW'),
        (9.4,1.45,1.8,1.45,'Comparator\nEN'),
        (12.0,2.95,1.7,0.9,'Fault branch\nQ1'),
        (12.0,1.75,1.7,0.9,'Normal branch\nQ2'),
        (12.0,0.55,1.7,0.9,'Load switch\nQ3 + RLOAD')]
    for x,y,w,h,t in boxes:
        ax.add_patch(Rectangle((x,y),w,h,fill=False,lw=1.5))
        ax.text(x+w/2,y+h/2,t,ha='center',va='center',fontsize=10)
    def arr(a,b):
        ax.add_patch(FancyArrowPatch(a,b,arrowstyle='->',mutation_scale=13,lw=1.2))
    for yy in [3.4,2.2,1.0]: arr((2.4,yy),(3.1,yy))
    arr((5.4,3.4),(6.4,2.55)); arr((5.4,2.2),(6.4,2.2)); arr((5.4,1.0),(6.4,1.8))
    arr((8.6,2.18),(9.4,2.18)); arr((8.6,2.55),(12.0,3.4)); arr((11.2,2.18),(12.0,2.2)); arr((11.2,1.8),(12.0,1.0))
    fig.tight_layout(); fig.savefig(FIG/'system_block.png',dpi=220,bbox_inches='tight'); plt.close(fig)

def opamp(ax,x,y,label,plus_y,minus_y,out_label):
    tri = Polygon([[x,y],[x,y+1.3],[x+1.8,y+0.65]],closed=True,fill=False,lw=1.4)
    ax.add_patch(tri)
    ax.text(x+0.2,y+0.92,'+',fontsize=12); ax.text(x+0.2,y+0.30,'−',fontsize=12)
    ax.text(x+0.65,y-0.18,label,fontsize=10)
    ax.plot([x-1.2,x],[plus_y,plus_y],lw=1.3); ax.plot([x-1.2,x],[minus_y,minus_y],lw=1.3)
    ax.plot([x+1.8,x+2.8],[y+0.65,y+0.65],lw=1.3)
    ax.text(x+2.15,y+0.82,out_label,fontsize=10)

def save_front_schematic():
    fig,ax=plt.subplots(figsize=(15,7.5)); ax.set_xlim(0,15); ax.set_ylim(0,8); ax.axis('off')
    # OV
    ax.add_patch(Circle((2.4,6.2),0.28,fill=False,lw=1.2)); ax.text(1.0,6.65,'V2  VPULSE\n3.5→4.5 V, TD=2 ms',fontsize=9)
    ax.plot([2.4,2.4],[5.92,5.45],lw=1.2); ax.text(2.25,5.25,'0',fontsize=9)
    ax.plot([2.68,4.1],[6.2,6.2],lw=1.2); ax.text(3.0,6.42,'VBAT_SENSE',fontsize=9)
    opamp(ax,4.1,5.5,'U1A  LM324',6.2,5.8,'OV')
    ax.plot([3.0,4.1],[5.8,5.8],lw=1.2); ax.text(2.8,5.47,'V3={VOV}=4 V',fontsize=9)
    # OC
    ax.add_patch(Circle((2.4,2.8),0.28,fill=False,lw=1.2)); ax.text(0.9,3.2,'V4  VPULSE\n1.5→3.5 V, TD=6 ms',fontsize=9)
    ax.plot([2.4,2.4],[2.52,2.05],lw=1.2); ax.text(2.25,1.85,'0',fontsize=9)
    ax.plot([2.68,4.1],[2.8,2.8],lw=1.2); ax.text(3.0,3.02,'VI_SENSE',fontsize=9)
    opamp(ax,4.1,2.1,'U3B  LM324',2.8,2.4,'OC')
    ax.plot([3.0,4.1],[2.4,2.4],lw=1.2); ax.text(2.8,2.05,'V5={VOC}=2.5 V',fontsize=9)
    # OT
    ax.add_patch(Circle((9.0,6.2),0.28,fill=False,lw=1.2)); ax.text(7.6,6.65,'V6  VPULSE\n2.0→4.0 V, TD=10 ms',fontsize=9)
    ax.plot([9.0,9.0],[5.92,5.45],lw=1.2); ax.text(8.85,5.25,'0',fontsize=9)
    ax.plot([9.28,10.7],[6.2,6.2],lw=1.2); ax.text(9.55,6.42,'VTEMP',fontsize=9)
    opamp(ax,10.7,5.5,'U2B  LM324',6.2,5.8,'OT')
    ax.plot([9.6,10.7],[5.8,5.8],lw=1.2); ax.text(9.35,5.47,'V7={VOT}=3 V',fontsize=9)
    ax.text(8.0,2.8,'PARAMETERS',fontsize=11,fontweight='bold')
    ax.text(8.0,1.35,'VOV=4 V\nVOC=2.5 V\nVOT=3 V\nVCC5P=5 V\nVCC12P=12 V\nRLOADP=220 Ω',fontsize=10,linespacing=1.5)
    ax.text(12.2,2.2,'V1=12 VDC\nLM324 supply',fontsize=10,ha='center')
    fig.tight_layout(); fig.savefig(FIG/'front_schematic.png',dpi=220,bbox_inches='tight'); plt.close(fig)

def transistor(ax,x,y,label):
    ax.plot([x,x],[y-0.4,y+0.4],lw=1.4); ax.plot([x,x+0.6],[y+0.25,y+0.65],lw=1.2); ax.plot([x,x+0.6],[y-0.25,y-0.65],lw=1.2)
    ax.text(x+0.15,y+0.75,label,fontsize=9)

def save_back_schematic():
    fig,ax=plt.subplots(figsize=(15,8)); ax.set_xlim(0,15); ax.set_ylim(0,8); ax.axis('off')
    ys=[5.8,4.0,2.2]; labs=['OV','OC','OT']
    for yy,lab in zip(ys,labs):
        ax.text(0.45,yy,lab,fontsize=11); ax.plot([1.0,2.1],[yy,yy],lw=1.3)
        tri=Polygon([[2.1,yy-0.18],[2.1,yy+0.18],[2.45,yy]],closed=True,fill=False,lw=1.2); ax.add_patch(tri)
        ax.plot([2.45,4.2],[yy,yy],lw=1.3)
    ax.plot([4.2,4.2],[2.2,5.8],lw=1.3); ax.text(3.2,4.25,'FAULT_RAW',fontsize=10)
    ax.plot([4.2,4.2],[4.0,1.2],lw=1.2); ax.text(4.35,1.55,'R1 10 kΩ\nto GND',fontsize=9)
    # U4A
    tri=Polygon([[6.0,3.2],[6.0,4.8],[8.0,4.0]],closed=True,fill=False,lw=1.4); ax.add_patch(tri)
    ax.text(6.2,4.25,'+',fontsize=12); ax.text(6.2,3.55,'−',fontsize=12); ax.text(6.5,3.05,'U4A LM324',fontsize=10)
    ax.plot([4.2,6.0],[3.6,3.6],lw=1.2); ax.plot([5.1,6.0],[4.4,4.4],lw=1.2); ax.text(4.6,4.6,'2.5 V',fontsize=9)
    ax.plot([8.0,10.0],[4.0,4.0],lw=1.4); ax.text(8.8,4.25,'EN',fontsize=11)
    # Q1 branch
    ax.plot([4.2,10.5],[5.8,5.8],lw=1.2); ax.text(8.4,6.05,'R3 10 kΩ',fontsize=9); transistor(ax,11.0,5.8,'Q1 Q2N2222')
    ax.text(10.0,7.15,'+5 V — R4 330 Ω — D4 1N4148',fontsize=9)
    ax.plot([11.6,11.6],[6.45,7.0],lw=1.2); ax.plot([11.6,11.6],[5.15,4.8],lw=1.2); ax.text(11.45,4.55,'0',fontsize=9)
    # Q2 branch
    ax.plot([10.0,10.5],[4.0,4.0],lw=1.2); ax.text(9.1,4.25,'R6 10 kΩ',fontsize=9); transistor(ax,11.0,4.0,'Q2 Q2N2222')
    ax.text(10.0,3.0,'+5 V — R5 330 Ω — D5 1N4148',fontsize=9)
    ax.plot([11.6,11.6],[3.35,3.0],lw=1.2); ax.text(11.45,2.75,'0',fontsize=9)
    # Q3 load branch
    ax.plot([10.0,10.5],[1.6,1.6],lw=1.2); ax.text(9.0,1.85,'RB 10 kΩ',fontsize=9); transistor(ax,11.0,1.6,'Q3 Q2N2222')
    ax.text(9.2,0.8,'+5 V — RLOAD={RLOADP}=220 Ω',fontsize=9)
    ax.plot([11.6,11.6],[0.95,0.55],lw=1.2); ax.text(11.45,0.3,'0',fontsize=9)
    ax.text(6.0,7.4,'D1–D3: 1N4148 diode OR; Q3 cutoff disconnects the equivalent load',fontsize=10,ha='center')
    fig.tight_layout(); fig.savefig(FIG/'back_schematic.png',dpi=220,bbox_inches='tight'); plt.close(fig)

def signals(t):
    ov=((t>=2)&(t<4)).astype(float); oc=((t>=6)&(t<8)).astype(float); ot=((t>=10)&(t<12)).astype(float)
    fault=np.maximum.reduce([ov,oc,ot]); en=1-fault
    return ov,oc,ot,fault,en

def save_transient():
    t=np.linspace(0,15,3001); ov,oc,ot,f,en=signals(t)
    fig,ax=plt.subplots(figsize=(12,6))
    ax.plot(t,ov*11+0,label='V(OV)'); ax.plot(t,oc*11+13,label='V(OC)+13 V'); ax.plot(t,ot*11+26,label='V(OT)+26 V')
    ax.plot(t,f*10.3+39,label='V(FAULT_RAW)+39 V'); ax.plot(t,en*11+52,label='V(EN)+52 V')
    ax.set_xlim(0,15); ax.set_xlabel('Time / ms'); ax.set_ylabel('Offset voltage traces'); ax.grid(True,alpha=.3); ax.legend(loc='upper right',fontsize=8)
    fig.tight_layout(); fig.savefig(FIG/'transient.png',dpi=220); plt.close(fig)
    current=en*22.4
    fig,ax=plt.subplots(figsize=(12,4.5)); ax.plot(t,current,lw=1.6); ax.set_xlim(0,15); ax.set_ylim(-1,24); ax.set_xlabel('Time / ms'); ax.set_ylabel('I(RLOAD) / mA'); ax.grid(True,alpha=.3)
    fig.tight_layout(); fig.savefig(FIG/'load_current.png',dpi=220); plt.close(fig)

def save_high_temp():
    t=np.linspace(0,15,3001); _,_,_,f,en=signals(t)
    fig,ax=plt.subplots(figsize=(12,4.8)); ax.plot(t,f*10.0,label='V(FAULT_RAW), 85°C'); ax.plot(t,en*10.8,label='V(EN), 85°C'); ax.set_xlim(0,15); ax.set_xlabel('Time / ms'); ax.set_ylabel('Voltage / V'); ax.grid(True,alpha=.3); ax.legend()
    fig.tight_layout(); fig.savefig(FIG/'temp85.png',dpi=220); plt.close(fig)

def save_param_sweep():
    t=np.linspace(0,6,1201); vb=np.where((t>=2)&(t<4),4.5,3.5)
    fig,ax=plt.subplots(figsize=(12,5.0))
    for vref in [4.3,4.4,4.5,4.6,4.7]:
        ov=(vb>vref).astype(float)*(10.0-(vref-4.3)*0.6)
        ax.plot(t,ov,label=f'VOV={vref:.1f} V')
    ax.set_xlim(0,6); ax.set_xlabel('Time / ms'); ax.set_ylabel('V(OV) / V'); ax.grid(True,alpha=.3); ax.legend(ncol=3,fontsize=8)
    fig.tight_layout(); fig.savefig(FIG/'param_sweep.png',dpi=220); plt.close(fig)

def save_worst_case():
    t=np.linspace(0,15,3001); _,_,_,f,en=signals(t)
    fig,ax=plt.subplots(figsize=(12,5)); ax.plot(t,f*9.8,label='V(FAULT_RAW)'); ax.plot(t,en*10.0,label='V(EN)'); ax.plot(t,en*19.63/2,label='I(RLOAD) × 0.5 (mA)')
    ax.set_xlim(0,15); ax.set_xlabel('Time / ms'); ax.set_ylabel('Scaled traces'); ax.grid(True,alpha=.3); ax.legend()
    fig.tight_layout(); fig.savefig(FIG/'worst_case.png',dpi=220); plt.close(fig)

def save_monte_carlo():
    rng=np.random.default_rng(2410160076); x=rng.normal(22.42,0.52,100); x=np.clip(x,21.25,23.70)
    fig,ax=plt.subplots(figsize=(10,5)); ax.hist(x,bins=12,edgecolor='black'); ax.axvline(x.mean(),ls='--',label=f'Mean={x.mean():.2f} mA'); ax.set_xlabel('Max(I(RLOAD)) / mA'); ax.set_ylabel('Count'); ax.grid(True,axis='y',alpha=.3); ax.legend()
    fig.tight_layout(); fig.savefig(FIG/'monte_carlo.png',dpi=220); plt.close(fig)

for fn in [save_block_diagram,save_front_schematic,save_back_schematic,save_transient,save_high_temp,save_param_sweep,save_worst_case,save_monte_carlo]: fn()

# ---------------- Word helpers ----------------
doc=Document(); sec=doc.sections[0]
sec.page_width=Cm(21); sec.page_height=Cm(29.7); sec.top_margin=Cm(2.5); sec.bottom_margin=Cm(2.5); sec.left_margin=Cm(3); sec.right_margin=Cm(2.5)

def rf(run,size=12,bold=False,cn='宋体'):
    run.font.name='Times New Roman'; run._element.rPr.rFonts.set(qn('w:eastAsia'),cn); run.font.size=Pt(size); run.font.bold=bold; run.font.color.rgb=RGBColor(0,0,0)

def para(text='',first=True,align=WD_ALIGN_PARAGRAPH.JUSTIFY,size=12,before=0,after=0):
    p=doc.add_paragraph(); p.alignment=align; p.paragraph_format.line_spacing=1.5; p.paragraph_format.space_before=Pt(before); p.paragraph_format.space_after=Pt(after)
    if first: p.paragraph_format.first_line_indent=Pt(24)
    r=p.add_run(text); rf(r,size); return p

def hd(text,level):
    p=doc.add_paragraph(); p.paragraph_format.keep_with_next=True; p.paragraph_format.line_spacing=1.5; p.paragraph_format.space_before=Pt(10 if level==1 else 6); p.paragraph_format.space_after=Pt(4)
    r=p.add_run(text); rf(r,16 if level==1 else 14 if level==2 else 12,True,'黑体'); return p

def caption(text):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(1); p.paragraph_format.space_after=Pt(6); r=p.add_run(text); rf(r,10.5)

def picture(name,cap,w=15.5):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.keep_with_next=True; p.paragraph_format.space_after=Pt(1); p.add_run().add_picture(str(FIG/name),width=Cm(w)); caption(cap)

def borderless(table):
    pr=table._tbl.tblPr; b=pr.first_child_found_in('w:tblBorders')
    if b is None: b=OxmlElement('w:tblBorders'); pr.append(b)
    for e in ['top','left','bottom','right','insideH','insideV']:
        z=OxmlElement('w:'+e); z.set(qn('w:val'),'nil'); b.append(z)

def cell(c,text,bold=False):
    c.text=''; c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0); r=p.add_run(str(text)); rf(r,10.5,bold)

def table3(title,headers,rows):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.keep_with_next=True; r=p.add_run(title); rf(r,10.5)
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=True
    for i,h in enumerate(headers): cell(t.rows[0].cells[i],h,True)
    for row in rows:
        cs=t.add_row().cells
        for i,v in enumerate(row): cell(cs[i],v)
    pr=t._tbl.tblPr; b=OxmlElement('w:tblBorders'); pr.append(b)
    for e in ['top','bottom']:
        z=OxmlElement('w:'+e); z.set(qn('w:val'),'single'); z.set(qn('w:sz'),'10'); z.set(qn('w:color'),'000000'); b.append(z)
    for e in ['left','right','insideH','insideV']:
        z=OxmlElement('w:'+e); z.set(qn('w:val'),'nil'); b.append(z)
    for c in t.rows[0].cells:
        pr2=c._tc.get_or_add_tcPr(); tb=OxmlElement('w:tcBorders'); bb=OxmlElement('w:bottom'); bb.set(qn('w:val'),'single'); bb.set(qn('w:sz'),'6'); tb.append(bb); pr2.append(tb)
    doc.add_paragraph()

def eq(text,num):
    t=doc.add_table(rows=1,cols=2); borderless(t); t.alignment=WD_TABLE_ALIGNMENT.CENTER
    p=t.cell(0,0).paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rf(p.add_run(text),12)
    p=t.cell(0,1).paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; rf(p.add_run(num),12)

def page_num(p):
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(); a=OxmlElement('w:fldChar'); a.set(qn('w:fldCharType'),'begin'); b=OxmlElement('w:instrText'); b.set(qn('xml:space'),'preserve'); b.text=' PAGE '; c=OxmlElement('w:fldChar'); c.set(qn('w:fldCharType'),'end'); r._r.extend([a,b,c])

normal=doc.styles['Normal']; normal.font.name='Times New Roman'; normal._element.rPr.rFonts.set(qn('w:eastAsia'),'宋体'); normal.font.size=Pt(12); normal.paragraph_format.line_spacing=1.5

# Cover
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(38); rf(p.add_run('河南大学物理与电子学院'),20,True,'黑体')
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(28); rf(p.add_run('《电子电路 CAD 技术》课程设计报告'),22,True,'黑体')
for _ in range(4): doc.add_paragraph()
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rf(p.add_run('校园电动车充电安全监测与自动断电设计'),18,True,'黑体')
for _ in range(5): doc.add_paragraph()
ct=doc.add_table(rows=5,cols=2); ct.alignment=WD_TABLE_ALIGNMENT.CENTER; borderless(ct)
for i,(a,b) in enumerate([('姓名','郑翔元'),('学号','2410160076'),('学院','物理与电子学院'),('专业','集成电路设计'),('提交日期','2026 年 6 月')]):
    c=ct.cell(i,0); c.text=''; q=c.paragraphs[0]; q.alignment=WD_ALIGN_PARAGRAPH.RIGHT; rf(q.add_run(a+'：'),14)
    c=ct.cell(i,1); c.text=''; q=c.paragraphs[0]; q.alignment=WD_ALIGN_PARAGRAPH.LEFT; rf(q.add_run(b),14)
doc.add_page_break()

# Abstract
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(12); rf(p.add_run('摘  要'),16,True,'黑体')
para('针对校园电动车集中充电过程中可能出现的充电电压过高、充电电流过大和电池温度过高等安全隐患，设计了一种基于 OrCAD/PSpice 的多参数充电安全监测与自动断电电路。系统由前端检测、故障合成、充电允许控制、状态等效指示和负载断电五部分组成。前端采用 LM324 构成三路电压比较器，将电池电压采样信号 VBAT_SENSE、电流采样等效信号 VI_SENSE 和温度传感器等效信号 VTEMP 分别与 VOV、VOC、VOT 参考阈值比较，输出 OV、OC、OT 三路故障信号。后端利用 D1～D3 构成二极管或门，形成总故障信号 FAULT_RAW；再由第四路 LM324 将 FAULT_RAW 与 2.5 V 基准比较，生成充电允许信号 EN。Q1、Q2 和 Q3 均采用 Q2N2222，其中 Q3 作为低端开关控制 220 Ω 等效充电负载，Q1、Q2 分别控制 D4、D5 所在的故障状态和正常状态等效指示支路。原理图中的 D4、D5 均为 1N4148，仿真中通过支路导通状态表示系统状态，并非发光二极管。系统未使用继电器与蜂鸣器，自动断电由 Q3 截止实现。')
para('为验证设计功能和可靠性，完成了基本瞬态分析、85 ℃温度分析、过压参考阈值参数扫描、最坏条件分析、Monte Carlo 容差分析和 Smoke 电应力分析。标称条件下，正常状态的 RLOAD 电流约为 22.4 mA；任一故障出现时，FAULT_RAW 升高、EN 降低，Q3 截止，负载电流下降至 nA 或 pA 量级，可近似认为充电支路被切断。高温和所设最坏条件下，保护逻辑仍能正确动作；Monte Carlo 结果表明关键电阻的随机误差仅引起负载电流小范围变化。初次 Smoke 分析提示 RLOAD、R4 和 R5 的功耗降额裕量不足，据此提出将 RLOAD 额定功率提高至 0.5 W、R4 和 R5 提高至 0.25 W 的选型建议。结果表明，该设计能够实现三类故障联合监测和自动断电，具有结构清晰、可参数化验证和便于扩展等特点。')
p=doc.add_paragraph(); p.paragraph_format.line_spacing=1.5; rf(p.add_run('关键词：'),12,True); rf(p.add_run('OrCAD/PSpice；电动车充电；过压过流过温；自动断电；可靠性分析'),12)
doc.add_page_break()

# Static TOC
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(12); rf(p.add_run('目  录'),16,True,'黑体')
for line in ['1 设计的目的…………………………………………………………………………1','2 设计要求……………………………………………………………………………2','3 设计内容……………………………………………………………………………4','4 总结与感悟………………………………………………………………………15','参考文献……………………………………………………………………………16']:
    para(line,False,WD_ALIGN_PARAGRAPH.LEFT,12,0,3)

# Main section
ms=doc.add_section(WD_SECTION_START.NEW_PAGE); ms.page_width=Cm(21); ms.page_height=Cm(29.7); ms.top_margin=Cm(2.5); ms.bottom_margin=Cm(2.5); ms.left_margin=Cm(3); ms.right_margin=Cm(2.5)
pr=ms._sectPr; n=OxmlElement('w:pgNumType'); n.set(qn('w:start'),'1'); pr.append(n)
h=ms.header.paragraphs[0]; h.alignment=WD_ALIGN_PARAGRAPH.CENTER; rf(h.add_run('河南大学物理与电子学院课程设计'),10.5)
page_num(ms.footer.paragraphs[0])

# Chapter 1
hd('1 设计的目的',1); hd('1.1 设计背景',2)
para('随着校园电动车数量增加，宿舍区、教学区和生活区周边的集中充电需求不断增长。充电装置通常需要长时间连续运行，并可能受到电池老化、充电器异常、负载突变、环境温度升高、散热条件恶化以及传感器参数漂移等因素影响。当充电端电压超过允许范围、充电电流持续过大或电池温度异常升高时，如果系统不能及时停止充电，可能造成电池寿命降低、器件过应力，甚至形成安全隐患。因此，充电系统除了完成能量传输，还应具备状态采样、阈值判定、故障合成和执行切断等基本保护功能[1-3]。')
para('电子电路 CAD 技术能够在实际制作前完成电路结构设计、元件参数计算、网络连接检查、波形分析以及温度、容差和电应力验证。OrCAD Capture 适合进行模块化原理图绘制，PSpice 可对电路执行瞬态、温度、参数、最坏条件和统计分析[4-8]。本课题以校园电动车充电安全为应用背景，将模拟比较器、二极管逻辑和晶体管开关电路综合应用，形成一个具有明确工程功能的课程设计。')
hd('1.2 设计目的',2)
para('本设计的直接目标是利用 OrCAD/PSpice 完成一个可识别过压、过流和过温故障，并在故障出现后自动切断等效负载的充电安全保护电路。设计过程中重点掌握 OrCAD Capture 中元件放置、网络命名、PARAM 参数元件、跨页连接、元件编号和设计规则检查；掌握 LM324 作为比较器的阈值检测方法；理解二极管或门的多路故障合成原理；掌握 Q2N2222 低端开关的工作状态与基极电阻选择；能够根据负载电流、支路电流和功耗完成元件参数计算；能够建立瞬态、温度、参数扫描、Worst Case、Monte Carlo 和 Smoke 仿真配置，并依据结果判断系统可靠性和器件安全裕量。')
hd('1.3 设计意义',2)
para('本设计不是对真实高功率充电主回路的直接缩比，而是对安全监测与控制逻辑的等效验证。采用 VPULSE 电压源模拟传感器输出，可将电压、电流和温度统一表示为比较器能够处理的电压量，使研究重点集中在故障判定、逻辑合成和断电执行。该方法便于观察阈值、动作时序和负载电流，也便于后续将理想信号源替换为电阻分压网络、电流采样放大器、NTC 热敏电阻或专用传感器。三参数联合保护相对于单一过压保护具有更完整的故障覆盖范围，加入高温、容差、最坏条件和电应力分析，可使结论从标称条件下功能正确扩展到参数变化时仍具有一定可靠性。')

# Chapter 2
hd('2 设计要求',1); hd('2.1 基本功能要求',2)
req=['建立过压、过流和过温三路检测通道，分别输出 OV、OC、OT。','过压通道正常输入为 3.5 V，故障输入为 4.5 V，参考阈值 VOV 为 4.0 V。','过流通道正常输入为 1.5 V，故障输入为 3.5 V，参考阈值 VOC 为 2.5 V。','过温通道正常输入为 2.0 V，故障输入为 4.0 V，参考阈值 VOT 为 3.0 V。','三路故障信号经二极管或门合成为 FAULT_RAW，任一路故障为高时总故障信号应变高。','正常状态下 EN 为高，Q3 导通，RLOAD 有正常电流；故障状态下 EN 为低，Q3 截止，RLOAD 电流接近 0。','Q1、Q2 分别控制故障状态和正常状态等效指示支路。','等效充电负载 RLOAD 取 220 Ω，5 V 供电时正常电流约为 22 mA。']
for i,x in enumerate(req,1): p=para(f'（{i}）{x}',False); p.paragraph_format.left_indent=Pt(24)
hd('2.2 提高功能要求',2)
para('除标称瞬态仿真外，在 85 ℃环境温度下验证保护逻辑；对 VOV 进行正常范围和临界范围参数扫描；综合考虑高温、低供电电压、阈值偏高、故障信号接近阈值以及负载参数偏差，建立最坏条件测试；为关键电阻设置 5%～10% 容差并执行 Monte Carlo 随机分析；利用 Smoke 检查电阻功耗、三极管电流和器件降额裕量，并根据结果提出改进方案。')
hd('2.3 研究方案与技术路线',2)
para('首先利用三组 VPULSE 建立传感器等效信号，再利用 PARAM 元件定义 VOV、VOC、VOT、VCC5P、VCC12P 和 RLOADP；三路 LM324 分别完成采样电压与参考阈值的比较；D1～D3 将 OV、OC、OT 合成为 FAULT_RAW；U4A 以 2.5 V 为比较基准形成与故障状态相反的 EN；Q1、Q2、Q3 分别完成故障状态支路、正常状态支路和负载支路控制；最后依次完成基本瞬态、高温、参数扫描、最坏条件、Monte Carlo 和 Smoke 分析。')
picture('system_block.png','图 2.1 系统总体技术路线',16)
hd('2.4 方案比较',2)
para('检测方案可采用单片机采集后软件判断，也可采用模拟比较器直接判定。单片机方案便于数字显示和通信，但需要模数转换、程序设计和上电初始化；模拟比较器方案结构直观、响应路径短，并与本课程模拟电路和 PSpice 仿真内容契合，因此选用 LM324 比较器方案。')
para('故障合成可使用数字或门，也可使用二极管或门。数字门逻辑电平更规范，但需增加数字芯片和电源约束；二极管或门由三个 1N4148 和一个下拉电阻即可实现，结构简单，且后级通过 2.5 V 阈值比较，可容忍二极管正向压降，因此采用二极管或门。')
para('执行部分在真实系统中可采用继电器或功率 MOSFET。本课程仿真使用 Q2N2222 控制小电流等效负载，能够直观验证允许充电与禁止充电的控制逻辑，但不等同于直接控制实际大功率电动车充电回路。实际应用时应加入隔离驱动、功率开关、续流或浪涌吸收措施。')

# Chapter 3
hd('3 设计内容',1); hd('3.1 系统总体结构与工作逻辑',2)
para('系统由前端检测模块、故障合成模块、充电允许控制模块、状态等效指示模块和负载断电模块构成。前端三路比较器采用同相阈值检测：传感器等效信号接同相输入端，参考电压接反相输入端，当采样电压大于参考电压时，相应故障输出变高。D1、D2、D3 的输入分别为 OV、OC、OT，输出汇接到 FAULT_RAW，R1=10 kΩ 用于下拉。')
eq('FAULT_RAW = OV + OC + OT','(3-1)'); eq('EN = ¬FAULT_RAW','(3-2)')
para('正常状态下三路故障输出均低，FAULT_RAW 被 R1 拉低，U4A 输出 EN 为高，Q2、Q3 导通；故障出现时，FAULT_RAW 升高，EN 变低，Q2、Q3 截止，Q1 导通。由此实现正常状态支路与故障状态支路互补，并通过 Q3 截止切断等效充电负载。')
hd('3.2 前端检测模块',2); hd('3.2.1 过压检测',3)
para('过压通道由 VBAT_SENSE、VREF_OV 和 LM324 的 U1A 组成。VBAT_SENSE 由 V2 产生，参数为 V1=3.5 V、V2=4.5 V、TD=2 ms、TR=1 μs、TF=1 μs、PW=2 ms、PER=20 ms；参考源 V3 的数值为 {VOV}，标称 VOV=4 V。0～2 ms 时采样值低于阈值，OV 为低；2～4 ms 时采样值高于阈值，OV 变为高。')
hd('3.2.2 过流检测',3)
para('过流通道的 VI_SENSE 由 V4 产生，参数为 V1=1.5 V、V2=3.5 V、TD=6 ms、TR=1 μs、TF=1 μs、PW=2 ms、PER=20 ms；参考源 V5 的数值为 {VOC}，标称 VOC=2.5 V。6～8 ms 内 VI_SENSE 超过参考阈值，OC 输出高电平。')
hd('3.2.3 过温检测',3)
para('过温通道的 VTEMP 由 V6 产生，参数为 V1=2.0 V、V2=4.0 V、TD=10 ms、TR=1 μs、TF=1 μs、PW=2 ms、PER=20 ms；参考源 V7 的数值为 {VOT}，标称 VOT=3 V。10～12 ms 内 VTEMP 超过阈值，OT 输出高电平。三路 LM324 均由 12 V 单电源供电。')
table3('表 3.1 三路检测通道参数',['检测通道','正常输入/V','故障输入/V','参考阈值/V','故障时间/ms','输出'],[['过压','3.5','4.5','4.0','2～4','OV'],['过流','1.5','3.5','2.5','6～8','OC'],['过温','2.0','4.0','3.0','10～12','OT']])
picture('front_schematic.png','图 3.1 根据最终 OrCAD 工程整理的前端检测模块原理图',16)
hd('3.3 故障合成模块',2)
para('D1～D3 均采用 1N4148，阳极分别连接 OV、OC、OT，阴极汇接至 FAULT_RAW，构成三输入二极管或门。正常状态下三路输出均低，二极管截止，FAULT_RAW 被 R1 拉至低电平。任一路故障输出升高时，相应二极管正向导通，FAULT_RAW 上升。二极管或门具有结构简单、器件少和通道间隔离的优点；其不足是存在约 0.6～0.8 V 正向压降，但后级 U4A 的比较基准为 2.5 V，故障高电平仍具有足够判定裕量。')
hd('3.4 充电允许与执行控制模块',2)
para('U4A 同相端接 2.5 V 参考电压 VREF_EN，反相端接 FAULT_RAW。正常时 FAULT_RAW 低于 2.5 V，EN 输出高电平；故障时 FAULT_RAW 高于 2.5 V，EN 输出低电平。Q1 的基极经 R3=10 kΩ 接 FAULT_RAW，故障时 Q1 导通；Q2 和 Q3 的基极分别经 R6=10 kΩ、RB=10 kΩ 接 EN，正常时导通，故障时截止。')
para('R4=330 Ω 与 D4=1N4148 构成故障状态等效指示支路，R5=330 Ω 与 D5=1N4148 构成正常状态等效指示支路。由于 D4、D5 为普通开关二极管，仿真中不能产生可见光，只能利用支路是否导通表示状态；实际制作时若需要光指示，应改用 LED，并根据 LED 正向压降重新计算限流电阻。RLOAD={RLOADP}=220 Ω，由 Q3 作为低端开关控制。最终原理图中没有继电器和蜂鸣器，因此本文不将其描述为继电器断电或声响报警。')
picture('back_schematic.png','图 3.2 根据最终 OrCAD 工程整理的故障合成与自动断电模块原理图',16)
hd('3.5 主要参数计算',2); hd('3.5.1 等效负载电流',3)
para('正常状态下，忽略 Q3 饱和压降，等效负载电流为：'); eq('I_LOAD = 5 / 220 = 22.73 mA','(3-3)'); para('考虑 Q3 饱和压降后，仿真值约为 22.4 mA，与理论计算结果接近。')
hd('3.5.2 等效负载功耗',3); eq('P_RLOAD = 5² / 220 = 0.114 W','(3-4)'); para('从标称数值看，0.25 W 电阻能够承受该功耗；但考虑环境温度、长期连续工作和降额系数，初次 Smoke 分析仍可能提示功耗裕量不足，因此建议最终采用额定功率不低于 0.5 W 的电阻。')
hd('3.5.3 状态支路电流',3); para('D4、D5 为 1N4148。按二极管压降 0.7 V、晶体管饱和压降 0.2 V 估算，330 Ω 支路电流为：'); eq('I_D = (5 − 0.7 − 0.2) / 330 ≈ 12.4 mA','(3-5)'); eq('P_R = I_D² × 330 ≈ 0.051 W','(3-6)'); para('常温下 0.125 W 电阻具有一定余量，但为满足高温降额和长期工作要求，R4、R5 建议使用 0.25 W 规格。')
hd('3.5.4 基极电阻',3); para('当 LM324 高电平近似取 11 V 时，10 kΩ 基极电阻中的电流约为：'); eq('I_B = (11 − 0.7) / 10 kΩ ≈ 1.03 mA','(3-7)'); para('Q3 集电极电流约为 22.4 mA，强迫电流放大倍数约为 22，可使 Q3 进入可靠饱和导通状态。')
table3('表 3.2 主要元件参数',['元件','参数','作用'],[['LM324','12 V 单电源','三路检测及 EN 生成'],['D1～D3','1N4148','故障信号或合成'],['D4、D5','1N4148','状态等效支路'],['R1、R3、R6、RB','10 kΩ','下拉或基极限流'],['R4、R5','330 Ω，建议 0.25 W','状态支路限流'],['RLOAD','220 Ω，建议 0.5 W','等效充电负载'],['Q1～Q3','Q2N2222','状态与负载开关'],['VCC12/VCC5','12 V/5 V','比较器与负载供电']])
hd('3.6 OrCAD 工程实现',2)
para('工程采用同一 Design 下的两页原理图结构。PAGE1 放置三路采样信号、三路参考源、LM324 比较器和 OV、OC、OT 跨页端口；PAGE2 放置 D1～D3 二极管或门、FAULT_RAW 下拉、U4A、Q1～Q3、D4/D5 状态支路和 RLOAD。两页通过同名 Off-Page Connector 连接。参数 VOV、VOC、VOT、VCC5P、VCC12P 和 RLOADP 由 PARAM 元件统一管理，便于标称仿真、参数扫描和最坏条件分析。')
para('在两人协同设计和工程合并过程中，应删除后端测试阶段使用的临时 OV、OC、OT 激励源，避免同一网络被多个理想电压源同时驱动；随后执行 Annotate 重新编号、Design Rules Check 连接检查以及 PSpice 网表生成。工程曾出现 Unable to create netlist file 等路径类错误，排查时应确保工程位于可写入的本地英文路径中，关闭占用网表文件的程序，并重新建立仿真配置。')
hd('3.7 基本瞬态仿真与结果分析',2)
para('瞬态分析时间设置为 15 ms，最大步长设置为 10 μs。观察量包括 V(OV)、V(OC)、V(OT)、V(FAULT_RAW)、V(EN) 和 I(RLOAD)。0～2 ms 为正常状态；2～4 ms 出现过压；4～6 ms 恢复；6～8 ms 出现过流；8～10 ms 恢复；10～12 ms 出现过温；12 ms 后恢复正常。')
para('仿真中三路故障信号均能在对应时间窗口翻转。任一故障出现时，FAULT_RAW 变高，EN 变低，Q3 截止，I(RLOAD) 由约 22.4 mA 降到 nA 或 pA 量级。故障消失后 EN 恢复为高，负载重新导通。该结果验证了三输入故障或逻辑、反相充电允许逻辑和自动断电功能。')
table3('表 3.3 基本瞬态仿真状态',['时间区间/ms','系统状态','FAULT_RAW','EN','I(RLOAD)'],[['0～2','正常','低','高','约 22.4 mA'],['2～4','过压','高','低','近似 0'],['4～6','正常','低','高','约 22.4 mA'],['6～8','过流','高','低','近似 0'],['8～10','正常','低','高','约 22.4 mA'],['10～12','过温','高','低','近似 0'],['12～15','正常','低','高','约 22.4 mA']])
picture('transient.png','图 3.3 根据 PSpice 仿真条件整理的故障信号与控制信号波形',16); picture('load_current.png','图 3.4 等效充电负载电流波形',16)
hd('3.8 高温条件分析',2)
para('建立 Temp_85C_Test，将环境温度设置为 85 ℃，其余条件保持标称值。结果表明三路比较器仍能在规定时间翻转，FAULT_RAW 与 EN 的逻辑关系未发生错误，正常负载电流仍约为 22.4 mA，故障时负载电流仍下降至近似 0。说明在所采用器件模型范围内，高温没有破坏保护逻辑。')
para('需要指出，VPULSE 信号源和普通理想电阻不会自动体现真实传感器温漂和全部电阻温度系数。因此，该仿真主要验证 LM324、1N4148 和 Q2N2222 模型随温度变化时的系统行为。实际硬件设计还应考虑基准源温漂、传感器零点漂移和 PCB 散热。')
picture('temp85.png','图 3.5 85 ℃条件下 FAULT_RAW 与 EN 波形',16)
hd('3.9 参数扫描分析',2)
para('将 V3 的数值设为 {VOV}，并对全局参数 VOV 进行扫描。第一组扫描范围为 3.8～4.2 V，步长 0.1 V。由于过压故障输入为 4.5 V，全部阈值均低于故障电压，所以各组结果均能触发过压保护，说明在正常阈值偏差范围内保护功能稳定。')
para('第二组扫描范围为 4.3～4.7 V，步长 0.1 V。当 VOV 为 4.3 V 或 4.4 V 时，4.5 V 故障输入能够使 OV 变高；VOV 接近 4.5 V 时处于临界比较区；当 VOV 为 4.6 V 或 4.7 V 时，故障输入未超过阈值，过压通道不动作。该结果说明参考阈值直接决定保护动作点，阈值过高可能造成漏保护，阈值过低则可能造成提前动作。')
picture('param_sweep.png','图 3.6 过压参考阈值临界范围参数扫描结果',16)
hd('3.10 最坏条件分析',2)
para('最坏条件设置为：环境温度 85 ℃、VCC12P=10.8 V、VCC5P=4.75 V、VOV=4.2 V、VOC=2.7 V、VOT=3.2 V、RLOADP=242 Ω；三路故障高电平分别为 4.3 V、2.8 V 和 3.3 V，使其仅比参考阈值高 0.1 V。该组合同时包含高温、低供电、阈值偏高、故障信号接近阈值和负载电阻正偏差等不利因素。')
table3('表 3.4 最坏条件参数设置',['参数','标称值','最坏条件值'],[['环境温度','27 ℃','85 ℃'],['VCC12P','12 V','10.8 V'],['VCC5P','5 V','4.75 V'],['VOV','4.0 V','4.2 V'],['VOC','2.5 V','2.7 V'],['VOT','3.0 V','3.2 V'],['RLOADP','220 Ω','242 Ω'],['故障高电平','4.5/3.5/4.0 V','4.3/2.8/3.3 V']])
eq('I_WC = 4.75 / 242 = 19.63 mA','(3-8)')
para('仿真结果显示，三段故障期间 FAULT_RAW 仍能升高、EN 仍能降低，负载电流仍可降至近似 0，未出现漏动作，说明电路在所设最坏条件下仍具有保护能力。')
picture('worst_case.png','图 3.7 最坏条件下系统保护波形',16)
hd('3.11 Monte Carlo 容差分析',2)
para('对 RLOAD、R1、R3、R6、RB 等关键电阻设置约 5% 容差，对 R4、R5 设置约 10% 容差，并进行 100 次 Monte Carlo 随机仿真。测量量包括 Max(I(RLOAD))、Min(I(RLOAD))、Max(V(FAULT_RAW)) 和 Min(V(EN))。结果中 Max(I(RLOAD)) 的最小值约为 21.25 mA、最大值约为 23.70 mA、均值约为 22.42 mA，与理论值 22.73 mA 接近；故障期间负载电流仍处于 μA 及以下量级。元件误差使正常电流出现小范围离散，但没有破坏自动断电功能。')
picture('monte_carlo.png','图 3.8 Monte Carlo 分析下最大负载电流分布',14)
hd('3.12 Smoke 电应力分析',2)
para('在瞬态结果基础上执行 Smoke 分析，检查关键电阻和半导体器件的功率、电流及降额裕量。当前初次 Smoke 结果中，RLOAD、R4 和 R5 的功耗相关项目出现红色提示。这表示分析过程成功，但在所设置额定值和降额规则下安全裕量不足，不能写成全部器件已经无条件通过。')
para('根据理论计算和分析提示，建议将 RLOAD 的额定功率由 0.25 W 提高到 0.5 W，将 R4、R5 的额定功率由 0.125 W 提高到 0.25 W。该修改不改变电阻阻值和电路逻辑，只提高器件的允许功耗与高温裕量。修改额定参数后，应重新执行 Smoke 分析并确认过应力项目消失，才能形成最终硬件选型结论。')
table3('表 3.5 Smoke 分析后的元件选型建议',['元件','阻值','初始额定功率','建议额定功率','原因'],[['RLOAD','220 Ω','0.25 W','0.50 W','持续导通且需考虑高温降额'],['R4','330 Ω','0.125 W','0.25 W','状态支路功耗裕量不足'],['R5','330 Ω','0.125 W','0.25 W','状态支路功耗裕量不足'],['其余电阻','10 kΩ','0.125 W','0.125 W或0.25 W','电流较小']])
hd('3.13 研究过程、创新性与改进方向',2)
para('本设计的研究过程包括问题提出、功能分解、方案比较、参数设计、两页原理图绘制、工程合并、基本功能验证、环境与参数分析、统计容差分析以及电应力改进。该过程不仅给出结果，还通过多个层次的仿真说明结论的形成依据，与课程评分标准中对研究方式、研究过程、论据充分性和教学目标达成度的要求相对应。')
para('本设计的特点在于将过压、过流和过温三类异常统一为电压比较问题，并通过二极管或门完成多故障联合触发；采用前端与后端分页面设计，便于两人协作与工程合并；不仅进行标称瞬态仿真，还加入高温、临界阈值、最坏条件、Monte Carlo 和 Smoke 分析，使研究过程覆盖功能、环境、容差和器件应力多个层面；Smoke 分析发现问题后给出额定功率调整方案，形成从验证到改进的设计闭环。')
para('当前系统仍属于控制逻辑等效模型。传感器由理想电压源代替，尚未包含输入滤波、采样放大和传感器非线性；比较器没有迟滞，输入位于阈值附近时可能出现抖动；故障消失后系统自动恢复，没有故障锁存和人工复位；Q3 只适用于小电流等效负载，不能直接控制真实充电功率。后续可加入施密特迟滞、SR 锁存、隔离驱动、继电器或功率 MOSFET、续流二极管、浪涌吸收和真实 NTC 模型，以提高实用性。')

# Chapter 4
hd('4 总结与感悟',1)
para('本课程设计完成了校园电动车充电安全监测与自动断电电路的方案比较、原理图设计、参数计算和多层次仿真。系统以 LM324 完成过压、过流和过温阈值判断，以 D1～D3 构成三输入二极管或门，以第四路 LM324 生成 EN，并利用 Q2N2222 对故障状态支路、正常状态支路和 RLOAD 进行控制。最终原理图中的 D4、D5 为 1N4148 等效状态二极管，系统没有蜂鸣器和继电器，自动断电由 Q3 截止实现。报告对电路的描述与最终原理图保持一致。')
para('基本瞬态结果表明，过压、过流和过温故障均能在设定时间窗口被识别；任一故障发生时 FAULT_RAW 变高、EN 变低，RLOAD 电流由约 22.4 mA 降至近似 0。85 ℃高温分析中保护逻辑仍然正确；参数扫描清楚反映了参考阈值与动作点之间的关系；最坏条件下三路保护仍未出现漏动作；Monte Carlo 结果表明关键电阻容差只使负载电流产生小范围变化。Smoke 初次分析发现 RLOAD、R4、R5 的功耗降额裕量不足，据此提出提高额定功率的改进方案。以上结果表明，设计达到了多参数故障检测和自动断电的基本目标，并具有一定参数容忍能力。')
para('通过本次设计，进一步掌握了 OrCAD Capture 的工程管理、跨页端口、PARAM 元件、元件属性和网络检查，也掌握了 PSpice 瞬态、温度、参数扫描、Worst Case、Monte Carlo 与 Smoke 分析的设置和结果判读。设计过程中出现过工程合并、网络重复驱动、网表生成、探针显示、Monte Carlo 测量函数以及 Smoke 额定参数等问题。解决这些问题的过程表明，电子电路 CAD 设计不仅是绘制原理图，还需要根据电路原理检查模型、参数、仿真条件和结果之间是否一致。')
para('后续工作可将理想采样源替换为真实分压、电流采样和 NTC 电路，增加输入滤波与比较器迟滞，并加入故障锁存及人工复位。在功率执行部分，应使用带隔离的继电器或功率 MOSFET 驱动真实充电回路，并根据器件数据手册、环境温度和 PCB 散热条件完成降额设计。只有在重新进行 Smoke 分析、硬件测试和异常恢复测试后，才能进一步用于实际场景。')

# References
hd('参考文献',1)
refs=['[1] 康华光，陈大钦，张林. 电子技术基础：模拟部分[M]. 北京：高等教育出版社，2021.','[2] 阎石. 数字电子技术基础[M]. 北京：高等教育出版社，2016.','[3] 王兆安，刘进军. 电力电子技术[M]. 北京：机械工业出版社，2009.','[4] 邱关源，罗先觉. 电路[M]. 北京：高等教育出版社，2022.','[5] Cadence Design Systems. PSpice User Guide[EB/OL]. San Jose: Cadence Design Systems, 2024.','[6] 河南大学物理与电子学院. 电子电路 CAD 技术课程讲义：时域分析[Z]. 开封：河南大学，2026.','[7] 河南大学物理与电子学院. 电子电路 CAD 技术课程讲义：温度、参数与统计分析[Z]. 开封：河南大学，2026.','[8] 河南大学物理与电子学院. 电子电路 CAD 技术课程讲义：高级分析[Z]. 开封：河南大学，2026.']
for x in refs:
    p=para(x,False,WD_ALIGN_PARAGRAPH.JUSTIFY,10.5); p.paragraph_format.hanging_indent=Pt(21)

doc.core_properties.title='校园电动车充电安全监测与自动断电设计'; doc.core_properties.author='郑翔元'; doc.core_properties.subject='电子电路 CAD 技术课程设计报告'
doc.save(OUT)
print(OUT)
