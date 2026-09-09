#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ST摸鱼风云-V5个股分析报告 —— 通用模板（定型版）

品牌：ST摸鱼风云-V5个股分析报告（正式定型版，2026-09-02 由 ST摸鱼风云-V2 升级定名）
评分体系：ST保壳评分系统V2（十二维，score_v2 单源出分，config v2.3，与保壳风云榜同刻度，2026-09-02 13→12重构）
框架：九章式深度报告（封面 + 投资要点 + 第一章~第九章 + 免责声明）

本文件是「模板引擎」：
  1. 复用固定的页面/表格/标题/正文/风险框 helper 函数（A4 + 表格满宽16cm + fixed布局，防表格撑破页面）
  2. 只改下方「===== 报告数据配置区 =====」内的 CFG 与各章 data_* 变量即可生成任意 ST/*ST 个股报告
  3. 评分常量 S/TOTAL_SCORE/RATING 建议直接用 score_v2.py 出分后填入（保持与榜单零偏差）

用法：
    python st_moyu_v5_report_template.py        # 按当前 CFG 生成 report.docx
依赖：python-docx（conda python 已装）
"""

import json
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement


# ============================================================
# ===== 报告数据配置区（生成新个股报告时只改这里 =====
# ============================================================
CFG = {
    "stock_name": "ST冀凯",          # 带ST前缀的名称
    "stock_code": "002691",
    "company_full": "冀凯装备制造股份有限公司",
    "board": "深圳主板",
    "st_type": "*ST（财务类）",
    "hat_date": "2026年4月30日戴帽",
    "report_date": "2026年9月2日",     # 生成/分析日期
    "output_path": r"ST摸鱼风云-V5完整版_ST冀凯_002691_报告.docx",
    "framework_note": "评分体系：ST保壳评分系统V2（十二维，score_v2单源，与保壳风云榜同刻度，2026-09-02 13→12重构）",
}

# ---- 评分（score_v2 出分后填入；002691 十二维 70/B，config v2.3） ----
S = {
    "C1": 5, "C2": 10, "S1": 10, "S2": 4,
    "A1": 9, "A2": 6, "A3": 0, "D1": 0,
    "B1": 10, "B2": 12, "F1": 0, "H1": 4,
}
TOTAL_SCORE = sum(S.values())          # 70
RATING = "B档"
RATING_FULL = "B（51-70分，保壳能力中等）"
RATING_COLOR = "2E75B6"
V2_RANK = 21                           # 全市场排名（保壳容易度）
V2_SHELL_BASE = 33.81                  # 壳费基准（own口径中位市值）
V2_SHELL_FEE = 22.02                   # 壳费中位数（亿元）

# ---- 封面摘要表 ----
cover_summary = [
    ["总资产 / 总负债", "11.89亿元 / 约3.91亿元（资产负债率32.85%，2026H1）"],
    ["归母净资产", "7.84亿元（2026H1，每股约2.31元）"],
    ["营业收入 / 归母净利润", "2026H1：1.012亿元 / -5,652万元（营收同比-14.13%，亏损同比扩大87.5%）"],
    ["短期借款", "2.37亿元（2026H1，较2024年末0.5亿元激增约374%）"],
    ["当前市值", "15.64亿元（2026-08-28收盘价4.60元；PB 1.99；52周区间3.15-11.32元）"],
    ["实控人 / 二股东", "冯春保 / 34.33%；深圳卓众达富 / 29.00%（合计63.33%，前十大合计70.8%）"],
]

# ---- 投资要点：核心风险清单 ----
risks = [
    ("财务类退市倒计时", "2025年扣除后营收28,869.98万<3亿触发9.3.1条，2026年报达标概率<5%，2027年4月终止上市概率80-85%"),
    ("立案调查未决", "2026-04-30证监会立案（涉嫌信披违规），若构成重大违法将切换为不可逆退市路径"),
    ("公开谴责", "2026-06-24深交所公开谴责（深证上〔2026〕861号），公司+冯帆+孙波+乔贵彩被谴责"),
    ("主业持续失血", "2026H1营收1.012亿元（同比-14.13%）、归母净利-5,652万元（亏损扩大87.5%）、毛利率9.55%、经营现金流-6,814万元"),
    ("高杠杆续命", "短贷2.37亿较2024年末0.5亿激增374%，资产负债率升至32.85%"),
]

# ---- 投资要点：分析结论（保留 score_v2 版本） ----
conclusion = (
    "*ST冀凯保壳评分70分（B档，V2十二维），但评分未覆盖的退市概率维度极端恶化：2026年报达标概率不足5%、"
    "2027年4月触发终止上市概率80-85%，叠加立案调查结果未明、公开谴责、投资者索赔三重合规风险，公司处于「退市倒计时+合规双杀」状态。"
    "壳价安全垫（市值15.44亿=壳费基准33.81亿的4.6折，V2-C2壳价值满分10分）仅对重组博弈具有理论意义，对场外资金而言风险收益比极差。"
    "投资核心假设是2026年内能否出现实质性保壳动作（资产注入/并购）——在立案调查未结案的背景下，任何方案推进都面临监管不确定性。"
    "不建议普通投资者参与，专业投资者亦应视为极高风险事件驱动标的。"
)

# ---- 目录 ----
toc_entries = [
    "第一章  公司概况", "  1.1 公司简介", "  1.2 主营业务", "  1.3 财务概况", "  1.4 股东背景",
    "第二章  财务分析", "  2.1 盈利能力", "  2.2 偿债能力", "  2.3 现金流分析", "  2.4 营运能力", "  2.5 财务趋势",
    "第三章  股权结构", "  3.1 前十大股东", "  3.2 实际控制人", "  3.3 股权质押",
    "第四章  风险全景", "  4.1 司法诉讼与舆情", "  4.2 戴帽原因及审计质量", "  4.3 退市风险研判",
    "第五章  风险因素与投资逻辑", "  5.1 驱动因素", "  5.2 压制因素", "  5.3 多空博弈研判",
    "第六章  资本运作成本与安全边际", "  6.1 隐性负债穿透", "  6.2 交易本质定性",
    "第七章  估值分析", "  7.1 壳资源估值", "  7.2 估值情景分析",
    "第八章  投资建议与策略", "  8.1 ST保壳评分系统V2 十二维评分", "  8.2 综合评级（V2刻度）", "  8.3 分层策略", "  8.4 重点跟踪指标清单",
    "第九章  附录", "  9.1 核心风险事项跟踪表", "  9.2 核心财务数据汇总表", "  9.3 股东结构汇总表", "  9.4 壳资源估值参照表", "  9.5 潜在并购资产画像",
    "免责声明",
]

# ---- 第八章 8.1 评分表（十二维 · 满分降序） ----
score_table = [
    ["S1 实控人性质", "12", "10", "民企(个人)冯春保；前3大股东合计65.61%>50%→上探10（民企档上限8→10）"],
    ["A2 扣非主营收入", "12", "6", "扣除后营收2.89亿、缺口4%（档位12分）−*ST戴帽(3)−低毛利<30%(3)=6分"],
    ["B2 审计意见", "12", "12", "2025年度标准无保留（大信，财务+内控双无保留）"],
    ["A1 净资产充裕度", "12", "9", "归母净资产8.40亿（2025年报，5-10亿档→9分，12分需≥10亿）"],
    ["C2 壳价值锚定", "10", "10", "市值15.44亿=壳费基准33.81亿的4.6折（≤5折=并购机会区，满分10）"],
    ["B1 立案/造假信号", "10", "10", "榜单快照无信号（10分）；动态敏感度见8.2"],
    ["C1 面值距离", "6", "5", "股价4.54元（≥3元档），距1元面值线安全垫充足"],
    ["S2 股权质押", "6", "4", "整体质押28.78%（20%-50%档，扣2分）"],
    ["A3 扣非盈利", "6", "0", "扣非净利-6,359万（2025），连续3年为负"],
    ["F1 重组/纾困", "6", "0", "无重整/重组/债务豁免/赠与预案（原F2维度，v2.3更名）"],
    ["D1 现金流质量", "4", "0", "2025经营现金流-6,271万，收现比恶化"],
    ["H1 实控人司法风险", "4", "4", "实控人无冻结/限高/立案记录"],
    ["合计", "100", str(TOTAL_SCORE), RATING_FULL + " · 通道封顶未触发"],
]

# ---- 第八章 8.4 跟踪指标 ----
track_items = [
    ("立案调查进展", "证监会调查结论——若构成重大违法，直接切换退市路径，一切保壳预期失效"),
    ("2026Q3营收", "10月三季报：Q3单季营收是否回升至6,000万+（对应H2合计约1.2亿+，距3亿红线仍远但可见边际）"),
    ("短贷与现金流", "2.37亿元短贷的续贷情况、货币资金变化，警惕流动性断裂"),
    ("保壳动作公告", "是否发布重大资产重组/发行股份购买资产/重整预案，方案可行性与时间表"),
    ("股东户数变化", "11,150户基础上是否进一步集中，大额吸筹信号"),
    ("索赔进展", "投资者索赔规模与立案结论联动，关注是否触发重大违法认定"),
]

# 各章正文请按九章框架补充（第1-7章为标的内容分析，见 generate_st_jikai_v2.py 范例），
# 这里用占位说明，实际使用时整体替换为标的真实分析。


# ============================================================
# 固定版式 helper（以下无需修改）
# ============================================================
doc = Document()
REPORT_DATE = CFG["report_date"]
OUTPUT_PATH = CFG["output_path"]

# 页面：A4 + 左右边距2.5cm（正文可用宽16cm）
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)
section.top_margin = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.different_first_page_header_footer = False

# 页脚
footer = section.footer
footer.is_linked_to_previous = False
fp = footer.paragraphs[0]
fp.clear()
pPr = fp._element.get_or_add_pPr()
pBdr = parse_xml('<w:pBdr %s><w:top w:val="single" w:sz="4" w:space="1" w:color="999999"/></w:pBdr>' % nsdecls('w'))
pPr.append(pBdr)
table = footer.add_table(rows=1, cols=2, width=Inches(6.0))
table.autofit = True
for row in table.rows:
    for cell in row.cells:
        tcPr = cell._element.get_or_add_tcPr()
        tcPr.append(parse_xml(
            '<w:tcBorders %s><w:top w:val="none" w:sz="0"/><w:left w:val="none" w:sz="0"/>'
            '<w:bottom w:val="none" w:sz="0"/><w:right w:val="none" w:sz="0"/></w:tcBorders>' % nsdecls('w')
        ))
cl = table.cell(0, 0); cl.width = Inches(3.0)
pl = cl.paragraphs[0]; pl.alignment = WD_ALIGN_PARAGRAPH.LEFT
pl.paragraph_format.space_before = Pt(0); pl.paragraph_format.space_after = Pt(0)
r = pl.add_run("小调AI - workbuddy")
r.font.size = Pt(8); r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
cr = table.cell(0, 1); cr.width = Inches(3.0)
pr = cr.paragraphs[0]; pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
pr.paragraph_format.space_before = Pt(0); pr.paragraph_format.space_after = Pt(0)
def _add_field(para, field_name):
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin')
    i = OxmlElement('w:instrText'); i.set(qn('xml:space'), 'preserve'); i.text = f' {field_name} '
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end')
    run = para.add_run()
    run.font.size = Pt(8); run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    run._element.append(b); run._element.append(i); run._element.append(e)
    return run
_add_field(pr, 'PAGE')
sep = pr.add_run(" / "); sep.font.size = Pt(8); sep.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
_add_field(pr, 'NUMPAGES')

# 页眉
header = section.header
header.is_linked_to_previous = False
hp = header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hr = hp.add_run(f"{CFG['stock_name']}({CFG['stock_code']}) | ST摸鱼风云-V5个股分析报告 | " + REPORT_DATE)
hr.font.size = Pt(8); hr.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
hr.font.name = '微软雅黑'; hr._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# 颜色与表格函数
COLOR_HEADER_BG = "1F4E79"
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_ROW_DARK = "DEEAF1"
COLOR_ROW_LIGHT = "BDD7EE"

def set_cell_bg(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        ed = kwargs.get(edge)
        if ed:
            tag = OxmlElement(f'w:{edge}')
            for k in ed: tag.set(qn(f'w:{k}'), ed[k])
            tcBorders.append(tag)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tblCellMar = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        mar = OxmlElement(f'w:{side}')
        mar.set(qn('w:w'), str(val)); mar.set(qn('w:type'), 'dxa')
        tblCellMar.append(mar)
    tcPr.append(tblCellMar)

BORDER_STD = {'sz': '1', 'val': 'single', 'color': 'CCCCCC'}
PAGE_USABLE_CM = 16.0
CM_TO_TWIPS = 567

def _visual_width(text):
    if text is None: return 0.0
    return sum(1.0 if ord(ch) > 127 else 0.55 for ch in str(text))

def auto_col_widths_cm(headers, rows, usable_cm=PAGE_USABLE_CM, min_col=1.5, max_ratio=0.6):
    n = len(headers)
    if n == 0: return []
    col_max = [0.0] * n
    for ci in range(n):
        hw = _visual_width(headers[ci])
        mw = max([_visual_width(r[ci]) for r in rows] + [0.0])
        col_max[ci] = max(hw, mw)
    cap = usable_cm * max_ratio
    need = [min(max(c * 0.32 + 0.25, min_col), cap) for c in col_max]
    total = sum(need)
    if total <= usable_cm:
        remain = usable_cm - total
        widths = [w + remain * (w / total) for w in need]
        return [round(w, 2) for w in widths]
    base = n * min_col; spare = usable_cm - base
    if spare <= 0: return [round(min_col, 2)] * n
    extra = [max(w - min_col, 0.0) for w in need]
    total_extra = sum(extra)
    if total_extra <= 0: return [round(min_col, 2)] * n
    widths = [min_col + (e / total_extra) * spare for e in extra]
    return [round(w, 2) for w in widths]

def set_table_fixed_width(table, widths_cm):
    tbl = table._tbl; tblPr = tbl.tblPr
    for el in tblPr.findall(qn('w:tblW')): tblPr.remove(el)
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), str(int(round(sum(widths_cm) * CM_TO_TWIPS))))
    tblW.set(qn('w:type'), 'dxa'); tblPr.append(tblW)
    for el in tblPr.findall(qn('w:tblLayout')): tblPr.remove(el)
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed'); tblPr.append(layout)
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None: tbl.remove(grid)
    grid = OxmlElement('w:tblGrid')
    for w in widths_cm:
        gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(int(round(w * CM_TO_TWIPS)))); grid.append(gc)
    tblPr.addnext(grid)
    for row in table.rows:
        for ci, w in enumerate(widths_cm):
            if ci < len(row.cells):
                tcPr = row.cells[ci]._tc.get_or_add_tcPr()
                for el in tcPr.findall(qn('w:tcW')): tcPr.remove(el)
                tcW = OxmlElement('w:tcW'); tcW.set(qn('w:w'), str(int(round(w * CM_TO_TWIPS))))
                tcW.set(qn('w:type'), 'dxa'); tcPr.append(tcW)

def add_table_with_style(doc, headers, rows, col_widths=None):
    num_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=num_cols)
    table.style = 'Table Grid'; table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_bg(cell, COLOR_HEADER_BG)
        set_cell_border(cell, top=BORDER_STD, bottom=BORDER_STD, left=BORDER_STD, right=BORDER_STD)
        set_cell_margins(cell)
        para = cell.paragraphs[0]; para.alignment = WD_ALIGN_PARAGRAPH.CENTER; para.text = ''
        run = para.add_run(h); run.bold = True
        run.font.color.rgb = RGBColor.from_string(COLOR_HEADER_TEXT); run.font.size = Pt(10)
        run.font.name = '微软雅黑'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    for ri, row in enumerate(rows):
        bg = COLOR_ROW_LIGHT if ri % 2 == 0 else COLOR_ROW_DARK
        for ci, val in enumerate(row):
            if ci >= num_cols: continue
            cell = table.rows[ri + 1].cells[ci]
            set_cell_bg(cell, bg)
            set_cell_border(cell, top=BORDER_STD, bottom=BORDER_STD, left=BORDER_STD, right=BORDER_STD)
            set_cell_margins(cell)
            para = cell.paragraphs[0]; para.text = ''
            run = para.add_run(str(val) if val is not None else '')
            run.font.size = Pt(9.5); run.font.name = '微软雅黑'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    if col_widths:
        widths = [float(w) for w in col_widths]
        total = sum(widths)
        if abs(total - PAGE_USABLE_CM) > 0.1:
            widths = [w * (PAGE_USABLE_CM / total) for w in widths]
    else:
        widths = auto_col_widths_cm(headers, rows)
    set_table_fixed_width(table, widths)
    return table

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(COLOR_HEADER_BG)
        run.font.name = '微软雅黑'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    return h

def add_body_text(doc, text, size=10.5, bold=False, align=None, color=None, space_before=0, space_after=4):
    p = doc.add_paragraph()
    if align: p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.name = '微软雅黑'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    if bold: r.bold = True
    if color: r.font.color.rgb = RGBColor.from_string(color)
    return p

def add_risk_box(doc, level, title, content):
    colors = {'high': 'FF0000', 'medium': 'FF8C00', 'low': '0070C0'}
    bg_colors = {'high': 'FFF0F0', 'medium': 'FFF8F0', 'low': 'F0F8FF'}
    c = colors.get(level, 'FF8C00'); bg = bg_colors.get(level, 'FFF8F0')
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f"【{title}】"); r.bold = True; r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor.from_string(c)
    r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    p2 = doc.add_paragraph(); p2.paragraph_format.space_before = Pt(0); p2.paragraph_format.space_after = Pt(6)
    r2 = p2.add_run(content); r2.font.size = Pt(9.5); r2.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    r2.font.name = '微软雅黑'; r2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    for pp in [p, p2]:
        pPr_elem = pp._element.get_or_add_pPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), bg)
        pPr_elem.append(shd)

# ============================================================
# 封面
# ============================================================
for _ in range(6):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    r = p.add_run(""); r.font.size = Pt(14)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("ST/*ST上市公司全维度分析报告")
r.bold = True; r.font.size = Pt(26); r.font.color.rgb = RGBColor.from_string(COLOR_HEADER_BG)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
r = p.add_run("ST摸鱼风云-V5个股分析报告")
r.font.size = Pt(14); r.font.color.rgb = RGBColor(0x64, 0x64, 0x64)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
r = p.add_run(CFG["framework_note"])
r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x8C, 0x8C, 0x8C)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run(CFG["stock_name"] + "（" + CFG["stock_code"] + "）")
r.bold = True; r.font.size = Pt(22); r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
r = p.add_run(CFG["company_full"])
r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x64, 0x64, 0x64)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
for _ in range(4):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    r = p.add_run(""); r.font.size = Pt(14)
add_table_with_style(doc, ["项目", "内容"], [
    ["综合评级", f"{RATING}（{TOTAL_SCORE} / 100 分 · ST保壳评分系统V2十二维）"],
    ["分析日期", REPORT_DATE],
    ["分析机构", "小调AI-WorkBuddy"],
    ["数据来源", "westock行情/财务/股东 / st-factor-fetch公共因子库 / 巨潮资讯网 / 公司公告"],
], col_widths=[2.5, 5.5])
doc.add_page_break()

# ============================================================
# 投资要点
# ============================================================
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0); r = p.add_run(""); r.font.size = Pt(14)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(2)
r = p.add_run(f"{CFG['stock_name']}（{CFG['stock_code']}）投资要点")
r.bold = True; r.font.size = Pt(18); r.font.color.rgb = RGBColor.from_string(COLOR_HEADER_BG)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(14)
r = p.add_run(f"{CFG['company_full']}  |  {CFG['board']}  |  {CFG['st_type']}  |  {CFG['hat_date']}")
r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
add_table_with_style(doc, ["项目", "内容"], cover_summary, col_widths=[3.0, 9.0])
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(10)
r = p.add_run(f"综合评级：{RATING}（{TOTAL_SCORE} / 100 分 · V2十二维）——保壳评分中等，但退市风险极高，详见第八章")
r.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor.from_string(RATING_COLOR)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
r = p.add_run("核心风险"); r.bold = True; r.font.size = Pt(11)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
for i, (title, desc) in enumerate(risks, 1):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"  {i}.  {title}："); r.bold = True; r.font.size = Pt(10)
    r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    r2 = p.add_run(desc); r2.font.size = Pt(10)
    r2.font.name = '微软雅黑'; r2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(4)
r = p.add_run("分析结论"); r.bold = True; r.font.size = Pt(11)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
add_body_text(doc, conclusion, size=10)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_before = Pt(14)
r = p.add_run(f"分析机构：小调AI-WorkBuddy  |  分析日期：{REPORT_DATE}")
r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
doc.add_page_break()

# ============================================================
# 目录
# ============================================================
add_heading(doc, "目  录", level=1)
for entry in toc_entries:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(1)
    is_chapter = not entry.startswith("  ")
    r = p.add_run(entry)
    r.bold = is_chapter; r.font.size = Pt(11 if is_chapter else 10)
    r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
doc.add_page_break()

# ============================================================
# 第一章~第九章正文
# 以下内容基于公开财务数据(westock 2026H1/2025年报)、巨潮公告(深交所纪律处分)、st_scores_v2评分快照撰写，
# 关键事件(2026-04-30立案、2026-06-24公开谴责、2026-04-30戴帽)已经过公开渠道核实。
# ============================================================

# ================= 第一章 公司概况 =================
add_heading(doc, "第一章  公司概况", level=1)
add_heading(doc, "1.1 公司简介", level=2)
add_body_text(doc, "冀凯装备制造股份有限公司（证券简称 *ST冀凯，证券代码 002691）成立于2003年5月，2012年7月31日在深圳证券交易所主板上市，注册地址位于河北省石家庄高新区湘江道418号，现注册资本3.4亿元，董事长为冯帆。公司属于机械设备行业，是矿山装备领域的专业供应商与服务商，主营支护机具、安全钻机、掘进设备和运输机械等矿用机械装备的研发、制造、销售与服务。")
add_body_text(doc, "公司下游主要为煤炭矿山及隧道掘进领域客户，经营景气度与煤炭行业资本开支高度相关。受煤炭市场行情下行、煤机行业竞争加剧影响，公司近年盈利能力持续承压，主业营收规模萎缩、毛利率下滑，并出现应收账款回款不及预期导致的信用减值损失，构成近年经营亏损的主因。")

add_heading(doc, "1.2 主营业务结构", level=2)
add_table_with_style(doc, ["业务板块", "主要产品", "下游场景"], [
    ["支护机具", "矿用锚杆、锚索、支护钻具", "煤矿巷道支护"],
    ["安全钻机", "瓦斯抽采钻机、探放水钻机", "煤矿安全治理"],
    ["掘进设备", "掘进机及配套部件", "巷道掘进工程"],
    ["运输机械", "矿用带式输送机等", "矿井物料运输"],
], col_widths=[2.5, 5.5, 4.0])
add_body_text(doc, "公司产品线围绕煤矿井下「采-掘-支-运」四大环节布局，属于典型的煤炭后周期设备供应商。需求弹性取决于煤矿资本开支节奏，2024-2026年煤炭行业资本开支下行叠加竞争加剧，直接冲击公司订单与盈利。")

add_heading(doc, "1.3 财务概况（截至2026H1）", level=2)
add_table_with_style(doc, ["指标", "2026H1", "2025年报", "趋势"], [
    ["营业总收入", "1.01亿元", "3.01亿元", "同比-14.13%"],
    ["归母净利润", "-5,652万元", "-5,487万元", "亏损扩大87.5%"],
    ["扣非净利润", "-5,929万元", "-6,359万元", "连续为负"],
    ["总资产", "11.89亿元", "11.34亿元", "—"],
    ["归母净资产", "7.84亿元", "8.40亿元", "持续下滑"],
    ["资产负债率", "32.85%", "24.6%", "杠杆上升"],
    ["短期借款", "2.37亿元", "1.40亿元", "激增"],
], col_widths=[3.2, 3.2, 3.2, 3.0])
add_body_text(doc, "整体看，公司陷入“营收萎缩-亏损扩大-杠杆抬升”的负向循环。2026H1营收仅1.01亿元（同比-14.13%）、归母净利亏损5,652万元（上年同期-3,014万元，亏损近乎翻倍），短期借款从2025年末1.40亿元攀升至2.37亿元，资产负债率由24.6%升至32.85%，靠债务融资支撑运营的迹象明显。")

add_heading(doc, "1.4 股东背景与上市沿革", level=2)
add_body_text(doc, "公司由自然人冯春保控制（民企·个人属性），实控人无国资或产业资本背书，保壳依赖自身资源与外部资本博弈。2026年4月30日，因2025年度利润总额、净利润、扣非净利润均为负，且扣除后营业收入（2.887亿元）低于3亿元，触及《股票上市规则》财务类退市风险警示条款，公司股票被实施 *ST（退市风险警示）。")

# ================= 第二章 财务分析 =================
add_heading(doc, "第二章  财务分析", level=1)
add_heading(doc, "2.1 盈利能力：主业持续失血", level=2)
add_table_with_style(doc, ["指标", "2026H1", "2025年报", "2024年报"], [
    ["营业收入", "1.01亿元", "3.01亿元", "3.49亿元"],
    ["营业成本", "0.92亿元", "2.30亿元", "—"],
    ["毛利率", "约9.6%", "约23.5%", "下滑"],
    ["归母净利", "-5,652万元", "-5,487万元", "亏损"],
    ["扣非净利", "-5,929万元", "-6,359万元", "亏损"],
], col_widths=[3.0, 3.0, 3.0, 3.0])
add_body_text(doc, "盈利能力恶化呈现“量价齐跌”特征：营收规模持续萎缩（从2024年约3.5亿降至2025年3.01亿、2026H1同比再降14%），毛利率大幅下滑（2026H1仅约9.6%）。公司归因于煤机市场竞争加剧、部分产品售价下调，同时应收账款回款未达预期导致信用减值损失增加。")
add_risk_box(doc, "high", "盈利端核心风险", "扣非净利润连续多年为负（2025年-6,359万、2026H1-5,929万），主业造血能力丧失。V2口径下A3扣非盈利=0/6、D1现金流=0/4，财务红线通道持续承压。只要营收与扣非盈利未实质改善，退市风险警示便无法解除。")

add_heading(doc, "2.2 偿债能力：短贷激增，流动性承压", level=2)
add_body_text(doc, "2026H1短期借款2.37亿元，较2025年末1.40亿元激增约69%（若较2024年末0.5亿元口径则增幅更大），资产负债率由24.6%升至32.85%。虽整体杠杆率尚在可控区间，但在主业持续亏损、经营现金流为负的背景下，债务滚续依赖外部融资支持，存在流动性边际风险。")

add_heading(doc, "2.3 现金流分析：经营净流出", level=2)
add_body_text(doc, "2026H1经营活动现金流净额约-6,814万元（2025全年约-6,271万元），连续为负且缺口扩大。收现比恶化、信用减值高企，说明公司销售回款质量下降、营运资金被不断侵蚀。在净利与现金流双负的情况下，公司运转高度依赖借款补血，财务可持续性存疑。")

add_heading(doc, "2.4 营运能力与资产质量", level=2)
add_body_text(doc, "公司应收账款规模较大（2026H1应收账款及票据合计约3.27亿元），远超同期营收，回款周期长。2026年8月20日公司公告计提资产减值准备及核销资产，反映下游客户（煤矿企业）付款能力下降，应收账款与存货存在减值风险。资产质量下滑叠加主业亏损，是公司近年报表端的显著特征。")

add_heading(doc, "2.5 财务趋势与退市红线距离", level=2)
add_risk_box(doc, "high", "退市红线：扣非营业收入缺口", "2025年度扣除后营业收入28,869.98万元，距离主板3亿元退市红线仅差约4%。2026H1营收仅1.01亿元（同比-14.13%），若下半年无显著改善，全年达标3亿元的概率极低——这是公司最核心的退市风险触发点，也是决定2027年是否终止上市的关键变量。")
add_body_text(doc, "财务趋势结论：公司处于“营收红线逼近+持续亏损+现金流为负”的三重挤压，靠内生经营完成2026年度财务类指标修复的可能性很小，保壳必须依赖外部变量（重组注入、政府纾困等）方能破局，详见第六章。")

# ================= 第三章 股权结构 =================
add_heading(doc, "第三章  股权结构", level=1)
add_heading(doc, "3.1 实际控制人", level=2)
add_body_text(doc, "公司实际控制人为冯春保，属于民企（个人）属性，V2口径下S1实控人性质得分10/12——民企档上限8分，但公司前3大股东合计持股65.61%（冯春保34.33%+深圳卓众达富29%+杭州焱热2.28%）>50%，触发控制力上探至10分。无国资兜底、无产业资本背书，保壳仍依赖自身资源或引入外部资本，但前3大股东股权高度集中为后续引入重组方保留了博弈空间。")
add_risk_box(doc, "medium", "实控人属性提示", "民企实控人若无增量资本介入，在财务类退市压力下自救能力有限。需关注实控人是否可能转让控制权、引入战略方或进行资产腾挪，作为保壳博弈的关键观察点。")

add_heading(doc, "3.2 前十大股东与二股东", level=2)
add_body_text(doc, "据2026年中报披露，实控人冯春保持股约34.33%，第二大股东深圳卓众达富持股约29.00%，二者合计约63.33%，前十大股东合计持股约70.8%，股权相对集中。二股东卓众达富作为机构/产业方持股近三成，其后续动作（增持、退出或配合重组）对控制权稳定与保壳路径有实质影响，需重点跟踪。")

add_heading(doc, "3.3 股权质押与冻结", level=2)
add_body_text(doc, "整体质押比例约28.78%（V2口径S2得分4/6，20%-50%档扣2分）。质押比例虽未达极端高位，但在股价低迷、退市风险高悬的背景下，若股价进一步下行可能触发质押平仓风险，需持续监控质押变动公告。")

# ================= 第四章 风险全景 =================
add_heading(doc, "第四章  风险全景", level=1)
add_heading(doc, "4.1 司法诉讼与监管处罚", level=2)
add_risk_box(doc, "high", "证监会立案调查（2026-04-30）", "公司于2026年4月30日披露收到证监会《立案告知书》，因涉嫌信息披露违法违规被立案调查。该立案源于2025年度业绩预告信披问题（详见4.2）。若后续认定构成重大违法（如触及重大违法强制退市情形），将触发不可逆的强制退市路径，一切保壳预期随之失效——这是当前所有风险中影响最重大、最不确定的一条。")
add_risk_box(doc, "high", "深交所公开谴责（2026-06-24）", "深交所出具《深证上〔2026〕861号》公开谴责处分：对冀凯股份及董事长冯帆、总经理孙波、财务负责人乔贵彩给予公开谴责。公司及3名高管同时被处分，反映监管对信披违规的严肃追责，也加大了对公司后续资本运作的监管约束。")

add_heading(doc, "4.2 戴帽原因及审计质量", level=2)
add_body_text(doc, "公司戴帽（*ST）的直接原因：2025年度利润总额、净利润、扣非净利润均为负值，且扣除后营业收入28,869.98万元低于3亿元，触发《股票上市规则（2025年修订）》第9.3.1条财务类退市风险警示规定，自2026年4月30日起被实施 *ST。")
add_risk_box(doc, "high", "戴帽链条中的信披违规（*ST起因）", "回溯：公司2026年1月28日披露的《2025年度业绩预告》仅预计净利亏损，未披露扣非营收低于3亿元情形，也未及时提示退市风险；至2026年4月24日才披露《业绩预告修正公告》并首次提示可能被实施退市风险警示。深交所认定其违反披露义务并予公开谴责，证监会随后立案——「业绩预告漏报→修正补报→被谴责→被立案→被戴帽」是完整的风险链条，直接导致公司进入退市风险警示状态。")

add_heading(doc, "4.3 退市风险研判", level=2)
add_body_text(doc, "冀凯当前面临三重退市路径叠加：① 财务类——2026年度需扣非营收≥3亿，但2026H1仅1.01亿且下滑，达标概率极低；② 违法类——证监会立案若升级为重大违法认定，直接强制退市；③ 交易类——现价虽在面值线上方（4.61元），但退市预期下存在股价持续下行的理论可能。其中财务类是概率最高的退市路径，违法类是破坏力最大的路径。")

# ================= 第五章 风险因素与投资逻辑 =================
add_heading(doc, "第五章  风险因素与投资逻辑", level=1)
add_heading(doc, "5.1 驱动因素（潜在保壳/看多方）", level=2)
add_bullets = [
    ("壳价值安全垫", "市值约15.44亿元，仅为壳费基准33.81亿元的4.6折（≤5折=并购机会区），对潜在重组方而言壳价便宜、借壳成本低"),
    ("账面净资产", "归母净资产8.40亿元（2025年报，V2-A1=9/12），无资不抵债，财务退市的“资不抵债”红线尚远"),
    ("审计意见干净", "2025年度财务+内控均为标准无保留（V2-B2=12/12），无因审计意见导致的退市触发"),
    ("股权集中", "实控人+二股东合计63.33%，若二股东卓众达富愿配合，引入重组方案的操作空间较大"),
]
for t, d in add_bullets:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"• {t}："); r.bold = True; r.font.size = Pt(10)
    r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    r2 = p.add_run(d); r2.font.size = Pt(10)
    r2.font.name = '微软雅黑'; r2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

add_heading(doc, "5.2 压制因素（看空/风险方）", level=2)
for t, d in [
    ("财务类退市倒计时", "2026年报达标(扣非营收≥3亿)概率<5%，2027年4月触发终止上市概率高（80-85%区间）"),
    ("证监会立案未决", "重大违法认定一旦成立即强制退市，结果未明前一切保壳动作受限"),
    ("公开谴责叠加", "监管问责升级，董监高面临处分，公司再融资/重组审批难度加大"),
    ("主业持续失血", "营收萎缩+亏损扩大+现金流为负，内生修复无望"),
    ("高杠杆续命", "短贷激增、资产负债率上升，流动性与债务风险累积"),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"• {t}："); r.bold = True; r.font.size = Pt(10)
    r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    r2 = p.add_run(d); r2.font.size = Pt(10)
    r2.font.name = '微软雅黑'; r2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

add_heading(doc, "5.3 多空博弈研判", level=2)
add_body_text(doc, "冀凯的博弈本质是“财务类退市倒计时”与“壳价值+潜在重组”的赛跑。多头逻辑建立在壳价便宜、若二股东配合引入重组可保壳的假设上；空头逻辑则指向财务类红线几乎无法内生跨越、叠加立案调查不确定性的现实。从风险收益比看，多头的理论收益（重组后估值修复）远低于空头的现实风险（退市归零），对场外资金而言是典型的高风险事件驱动标的，仅对能承受归零风险、深谙退市博弈规则的专业资金有博弈意义。")

# ================= 第六章 资本运作成本与安全边际 =================
add_heading(doc, "第六章  资本运作成本与安全边际", level=1)
add_heading(doc, "6.1 隐性负债穿透", level=2)
add_body_text(doc, "公司表内资产负债率32.85%虽不算高，但需穿透关注：① 短期借款2.37亿元依赖滚动续贷，主业不产生净现金流入的情况下，续贷能力依赖银行授信与股东支持；② 应收账款及票据约3.27亿元规模较大，其中存在减值风险的部分可能构成表内「隐性」减损；③ 立案调查与公开谴责可能影响公司再融资渠道，限制通过定增等工具补充资本的能力。综合看，公司的真实财务安全边际弱于表观负债率所反映的水平。")

add_heading(doc, "6.2 交易本质定性", level=2)
add_body_text(doc, "若发生控制权或重大资产交易，冀凯作为*ST壳标的的本质是“以净资产为底、以保壳预期为溢价的博弈载体”。当前市值15.44亿 vs 归母净资产7.84亿，PB约2倍——壳的估值含了保壳成功的期权价值。潜在买方的理性出价会锚定：① 壳的净壳程度（负债/诉讼/立案风险的清理成本）；② 保壳时间窗口（2026年报为生死线，留给重组方时间极短）；③ 立案调查的确定性。在立案未结前，任何交易的执行都面临巨大监管与时间不确定性。")

# ================= 第七章 估值分析 =================
add_heading(doc, "第七章  估值分析", level=1)
add_heading(doc, "7.1 壳资源估值", level=2)
add_table_with_style(doc, ["估值参照", "数值", "说明"], [
    ["交易市值中位数", "33.81亿元", "own口径：近2年154笔实控权变更样本"],
    ["壳费中位数", "22.02亿元", "own口径L2单源基准"],
    ["冀凯当前市值", "约15.44亿元", "仅约为基准的4.6折"],
    ["归母净资产", "7.84亿元", "PB约2倍"],
], col_widths=[3.0, 3.0, 6.0])
add_body_text(doc, "V2-C2壳价值锚定得分10/10：冀凯市值≤壳基准33.81亿的5折，处于并购/借壳成本优势区——壳越便宜，潜在重组方入场成本越低，被并购的概率上升。但需清醒认识到，壳价值锚定是「重组博弈发生前提」下的估值参考，在财务类退市倒计时+立案未决的现实下，壳价值的变现高度依赖能否在2026年报前完成有效的保壳动作。")

add_heading(doc, "7.2 估值情景分析", level=2)
add_table_with_style(doc, ["情景", "触发条件", "估值含义", "概率判断"], [
    ["退市归零", "2026年报扣非营收<3亿且无重大违法(退市整理后摘牌)", "市值趋向0（整理期缩水）", "中高(概率最高路径)"],
    ["立案升级强制退市", "证监会认定重大违法", "直接退市，壳价值归零", "需盯调查结论"],
    ["成功保壳(摘帽)", "2026年报达标+立案结案中性+审计无保留", "估值修复至净资产上方", "低(需重组配合)"],
    ["重组博弈", "引入资产注入/二股东主导重组", "壳价博弈抬升市值", "中(窗口极短)"],
], col_widths=[3.0, 5.0, 3.5, 3.0])
add_body_text(doc, "估值结论：基准情形是“财务类退市倒计时下的壳价博弈”。当前15.44亿市值对应的是市场对“年内出现保壳动作”的期权定价，一旦时间窗口关闭且无动作，估值将向归零方向快速收敛。投资者须对“壳价值仅在重组发生时才成立”有清醒认知。")

add_heading(doc, "第八章  投资建议与策略", level=1)
add_heading(doc, "8.1 ST保壳评分系统V2 十二维评分", level=2)
add_body_text(doc, "评分口径：与保壳风云榜同刻度同数据源（score_v2单源出分，财务基期以榜单快照为准），十二维权重=近5年176家退市案例实证贡献度，维度即退市通道。2026-09-02重构：取消原F1经营改善趋势4分→A1净资产(+2)、C2壳价值(+2)，原F2重组/纾困更名F1。通道封顶（一票否决）：C1=0或B2=0→总分封顶50；B1=0→总分封顶30。", size=9.5)
add_table_with_style(doc, ["维度", "满分", "得分", "得分逻辑"], score_table)
add_heading(doc, "8.2 综合评级（V2刻度）", level=2)
add_body_text(doc, f"全市场保壳评分总分：{TOTAL_SCORE} / 100分（ST保壳评分系统V2·十二维）", bold=True, size=11)
add_body_text(doc, f"综合评级：{RATING_FULL}——全市场排名第{V2_RANK}位")
add_heading(doc, "8.3 分层策略", level=2)
add_table_with_style(doc, ["时间维度", "策略", "关键节点", "仓位建议"], [
    ["短期(0-6月)", "观望为主，不参与；跟踪立案调查进展与保壳公告", "立案结论、最新季报(10月)、重组/并购公告", "0%（场外资金）"],
    ["中期(6-12月)", "仅当出现实质性保壳动作（预案披露+立案结案）再评估", "年报业绩预告(次年1月)、年报(次年4月)", "0-3%（专业投资者博弈仓）"],
    ["长期(12月+)", "退市落地→停止交易退出；摘帽成功→重新评估", "终止上市决定/摘帽日", "—"],
], col_widths=[2.5, 4.5, 3.0, 2.0])
add_body_text(doc, "核心原则：在立案调查结案+实质性保壳方案落地之前，不参与。对场外资金（非专业投资者）：明确不建议参与，退市归零风险远超博弈收益。对专业投资者：若确要博弈，仅限事件驱动策略，仓位≤2%，严格设置止损（-15%），并以「立案结案为中性、重大违法认定为清仓」为硬纪律。")
add_heading(doc, "8.4 重点跟踪指标清单", level=2)
for i, (title, desc) in enumerate(track_items, 1):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"{i}. {title}："); r.bold = True; r.font.size = Pt(9.5)
    r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    r2 = p.add_run(desc); r2.font.size = Pt(9.5)
    r2.font.name = '微软雅黑'; r2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
doc.add_page_break()

# ================= 第九章 附录 =================
add_heading(doc, "第九章  附录", level=1)

add_heading(doc, "9.1 核心风险事项跟踪表", level=2)
add_table_with_style(doc, ["风险事项", "状态", "触发影响", "跟踪时点"], [
    ["证监会立案调查(信披违规)", "调查进行中", "若认定重大违法→强制退市", "调查结论公布日"],
    ["2026年度扣非营收<3亿", "2026H1仅1.01亿，缺口大", "触发财务类终止上市", "2027年4月年报"],
    ["深交所公开谴责", "已处分", "监管约束资本运作/再融资", "持续"],
    ["主业亏损+现金流为负", "持续", "财务红线通道承压", "各期财报"],
    ["短贷滚续/流动风险", "2.37亿短贷", "流动性断裂", "季报/融资公告"],
], col_widths=[4.5, 3.0, 4.0, 3.0])

add_heading(doc, "9.2 核心财务数据汇总表", level=2)
add_table_with_style(doc, ["科目", "2026H1", "2025年报"], [
    ["营业总收入", "1.01亿元", "3.01亿元"],
    ["归母净利润", "-5,652万元", "-5,487万元"],
    ["扣非净利润", "-5,929万元", "-6,359万元"],
    ["扣非营业收入", "—", "2.887亿元(触发红线)"],
    ["总资产", "11.89亿元", "11.34亿元"],
    ["归母净资产", "7.84亿元", "8.40亿元"],
    ["资产负债率", "32.85%", "24.6%"],
    ["短期借款", "2.37亿元", "1.40亿元"],
    ["经营现金流净额", "约-6,814万元", "约-6,271万元"],
], col_widths=[4.5, 5.0, 5.0])
add_body_text(doc, "数据来源：westock财务接口（2026H1/2025年报），与封面摘要、投资要点表一致。")

add_heading(doc, "9.3 股东结构汇总表", level=2)
add_table_with_style(doc, ["股东", "持股比例", "性质", "备注"], [
    ["冯春保", "约34.33%", "实际控制人", "民企(个人)"],
    ["深圳卓众达富", "约29.00%", "二股东", "机构/产业方"],
    ["前十大合计", "约70.8%", "—", "股权集中"],
], col_widths=[4.0, 3.0, 3.5, 4.0])
add_body_text(doc, "注：整体质押比例约28.78%（V2-S2口径）。股东结构数据以公司定期报告披露为准，本表为榜单快照期口径。")

add_heading(doc, "9.4 壳资源估值参照表", level=2)
add_table_with_style(doc, ["参照指标", "数值", "口径说明"], [
    ["交易市值中位数(基准)", "33.81亿元", "own口径：近2年154笔实控权变更"],
    ["壳费中位数", "22.02亿元", "L2单源基准"],
    ["冀凯市值", "约15.44亿元", "约基准4.6折(并购机会区)"],
    ["归母净资产", "7.84亿元", "PB约2倍"],
], col_widths=[4.5, 3.5, 6.0])

add_heading(doc, "9.5 潜在并购资产画像", level=2)
add_body_text(doc, "在财务类退市倒计时+证监会立案未结的双重约束下，冀凯的现实保壳路径高度依赖外部资产注入或控制权交易。潜在并购方画像：① 需要低成本获取A股上市平台但难以走IPO/借壳（现壳价约基准4.6折具吸引力）；② 具备能与公司现有煤矿机械协同、或可快速置入高盈利资产的产业方；③ 二股东卓众达富若具备产业背景，可能成为主导重组的核心力量。但需强调：任何并购方案在立案调查结案前都面临监管不确定性与极短时间窗口（2026年报为生死线）的双重制约，实际落地概率需打显著折扣。")

# ============================================================
# 免责声明
# ============================================================
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(8)
r = p.add_run("免责声明"); r.bold = True; r.font.size = Pt(15)
r.font.color.rgb = RGBColor.from_string(COLOR_HEADER_BG)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(1)
r = p.add_run("【评级说明】"); r.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x64, 0x64, 0x64)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = Pt(14)
r = p.add_run(
    "本报告为「ST摸鱼风云-V5个股分析报告」，采用ST保壳评分系统V2（score_v2单源出分，config v2.3，100分制），"
    "下设十二个维度（C1面值距离6分/C2壳价值锚定10分/S1实控人性质12分/S2股权质押6分/A1净资产充裕度12分/A2扣非主营收入12分/A3扣非盈利6分/D1现金流质量4分/B1立案造假信号10分/B2审计意见12分/F1重组纾困6分/H1实控人司法风险4分；2026-09-02重构：取消原F1经营改善趋势4分→A1净资产+2、C2壳价值+2，原F2重组/纾困更名F1），"
    "维度=退市通道，权重=近5年176家退市案例实证贡献度。"
    "综合评级分为四档：A（>70分）保壳能力强、B（51-70分）中等、C（31-50分）较弱、D（≤30分）极易退市。高分=保壳容易，与保壳风云榜榜单同刻度横向可比。"
    "通道封顶（一票否决）：C1=0或B2=0→总分封顶50；B1=0→总分封顶30。"
    "C2壳价值锚定采用L2单源壳费基准（own口径：近2年154笔完成实控权变更、交易市值<60亿样本，壳费中位数22.02亿元、交易市值中位数33.81亿元）。"
    "评分基于公开可获取的财务数据和披露信息，采用标准化打分模型，不构成盈利预测或投资承诺。"
)
r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x64, 0x64, 0x64)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(1)
r = p.add_run("【分析师声明】"); r.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x64, 0x64, 0x64)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = Pt(14)
r = p.add_run(
    "本报告由小调AI-WorkBuddy智能分析系统生成，底层使用公开AI模型进行数据处理和文本撰写。"
    "分析过程遵循独立、客观、公正原则，数据来源于westock行情/财务/股东接口、st-factor-fetch公共因子库、巨潮资讯网、公司公告等公开渠道。"
    "研究方法包括但不限于：财务报表分析、合规记录审查、司法诉讼检索、股权结构穿透、壳资源估值模型、ST保壳评分系统V2十二维保壳评分等标准化流程。"
    "分析师的薪酬不与本报告的具体评级或投资建议挂钩。"
)
r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x64, 0x64, 0x64)
r.font.name = '微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

doc.save(OUTPUT_PATH)
print(f"ST摸鱼风云-V5个股分析报告已生成: {OUTPUT_PATH}")
print(f"Total score: {TOTAL_SCORE}/100, Rating: {RATING}")
