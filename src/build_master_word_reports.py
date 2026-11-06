"""
MASTER PROMPT – PROFESSIONAL ACADEMIC & TECHNICAL WORD DOCUMENT DESIGNER
Generates publication-grade Word (.docx) documents in both 100% Vietnamese (VI) and 100% English (EN):
- A4 Page Setup with professional margins
- Custom Executive Cover Page
- Automated Header & Footer with dynamic Page X of Y fields
- Automated Table of Contents (TOC), List of Tables, and List of Figures
- Strict Heading 1, Heading 2, Heading 3 hierarchy with keep_with_next (no orphaned headings)
- Publication-styled tables with repeating headers (tblHeader), cantSplit rows, and alternating zebra shading
- High-resolution figures (200 DPI) with linked captions
- Executive Callout Boxes, References (IEEE standard), and Technical Appendix
"""

import os
import shutil
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_FOLDER = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"

FIG_VI = os.path.join(BASE_DIR, "figures_vi")
FIG_EN = os.path.join(BASE_DIR, "figures_en")

OUT_VI_WS = os.path.join(BASE_DIR, "SIC_Documents_VI")
OUT_EN_WS = os.path.join(BASE_DIR, "SIC_Documents_EN")
OUT_VI_DRV = os.path.join(SRC_FOLDER, "Ban_Tieng_Viet_Full")
OUT_EN_DRV = os.path.join(SRC_FOLDER, "Ban_Tieng_Anh_Full")

for d in [OUT_VI_WS, OUT_EN_WS, OUT_VI_DRV, OUT_EN_DRV]:
    os.makedirs(d, exist_ok=True)

# Color Palette
NAVY_HEX = "1D3557"
SLATE_HEX = "264653"
TEAL_HEX = "2A9D8F"
LIGHT_BG_HEX = "F4F7FB"
CALLOUT_BG_HEX = "EDF6F5"
BORDER_GRAY_HEX = "CBD5E1"

NAVY_RGB = RGBColor(29, 53, 87)
SLATE_RGB = RGBColor(38, 70, 83)
TEAL_RGB = RGBColor(42, 157, 143)
DARK_TEXT_RGB = RGBColor(30, 41, 59)
MUTED_TEXT_RGB = RGBColor(100, 116, 139)


# ==============================================================================
# XML HELPER FUNCTIONS FOR PROFESSIONAL WORD FORMATTING
# ==============================================================================
def set_cell_background(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=110, bottom=110, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="8" w:space="0" w:color="{NAVY_HEX}"/>'
        f'<w:bottom w:val="{val}" w:sz="8" w:space="0" w:color="{NAVY_HEX}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def repeat_header_row(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))


def add_bottom_border_to_paragraph(paragraph, color_hex=NAVY_HEX, sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color_hex}"/>'
        f"</w:pBdr>"
    )
    pPr.append(pBdr)


def add_field(paragraph, field_text, font_size=9.5, color_rgb=MUTED_TEXT_RGB):
    run = paragraph.add_run()
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    run.font.color.rgb = color_rgb
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> %s </w:instrText>' % (nsdecls("w"), field_text))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w"))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def configure_document_styles(doc):
    # A4 Page Setup
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.2)
    section.different_first_page_header_footer = True

    # Normal style
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = DARK_TEXT_RGB
    normal.paragraph_format.line_spacing = 1.3
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Heading 1
    h1 = doc.styles["Heading 1"]
    h1.font.name = "Times New Roman"
    h1.font.size = Pt(16)
    h1.font.bold = True
    h1.font.color.rgb = NAVY_RGB
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(8)
    h1.paragraph_format.keep_with_next = True
    h1.paragraph_format.keep_together = True

    # Heading 2
    h2 = doc.styles["Heading 2"]
    h2.font.name = "Times New Roman"
    h2.font.size = Pt(13.5)
    h2.font.bold = True
    h2.font.color.rgb = SLATE_RGB
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.keep_with_next = True
    h2.paragraph_format.keep_together = True

    # Heading 3
    h3 = doc.styles["Heading 3"]
    h3.font.name = "Times New Roman"
    h3.font.size = Pt(12.5)
    h3.font.bold = True
    h3.font.italic = True
    h3.font.color.rgb = TEAL_RGB
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(4)
    h3.paragraph_format.keep_with_next = True
    h3.paragraph_format.keep_together = True


def setup_header_footer(doc, lang="vi"):
    section = doc.sections[0]
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run(
        "SAMSUNG INNOVATION CAMPUS – BÁO CÁO ĐỒ ÁN AI CAPSTONE | NHÓM 03"
        if lang == "vi"
        else "SAMSUNG INNOVATION CAMPUS – AI CAPSTONE TECHNICAL REPORT | GROUP 03"
    )
    hrun.font.name = "Times New Roman"
    hrun.font.size = Pt(9.5)
    hrun.font.italic = True
    hrun.font.color.rgb = MUTED_TEXT_RGB
    add_bottom_border_to_paragraph(hp, color_hex=BORDER_GRAY_HEX, sz="6")

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    frun1 = fp.add_run(
        "Đỗ Quang Minh • Vũ Quang Huy  |  Trang "
        if lang == "vi"
        else "Do Quang Minh • Vu Quang Huy  |  Page "
    )
    frun1.font.name = "Times New Roman"
    frun1.font.size = Pt(9.5)
    frun1.font.color.rgb = MUTED_TEXT_RGB
    add_field(fp, "PAGE")
    frun2 = fp.add_run(" / ")
    frun2.font.name = "Times New Roman"
    frun2.font.size = Pt(9.5)
    frun2.font.color.rgb = MUTED_TEXT_RGB
    add_field(fp, "NUMPAGES")


def add_callout_box(doc, title_text, body_text, accent_hex=TEAL_HEX, bg_hex=CALLOUT_BG_HEX):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=130, bottom=130, left=220, right=180)
    prevent_row_split(tbl.rows[0])

    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{accent_hex}"/>'
        f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>'
        f"</w:tcBorders>"
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r_title = p.add_run(title_text)
    r_title.bold = True
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(11.5)
    r_title.font.color.rgb = NAVY_RGB

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(2)
    r_body = p2.add_run(body_text)
    r_body.font.name = "Times New Roman"
    r_body.font.size = Pt(11.5)
    r_body.font.color.rgb = DARK_TEXT_RGB

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(4)


def add_styled_table(doc, caption_text, headers, rows_data, col_alignments=None):
    # Table caption above table (with keep_with_next so it never separates from table)
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_cap.paragraph_format.space_before = Pt(10)
    p_cap.paragraph_format.space_after = Pt(4)
    p_cap.paragraph_format.keep_with_next = True
    r_cap = p_cap.add_run(caption_text)
    r_cap.bold = True
    r_cap.font.name = "Times New Roman"
    r_cap.font.size = Pt(11)
    r_cap.font.color.rgb = NAVY_RGB

    tbl = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)

    # Header row
    hdr_row = tbl.rows[0]
    repeat_header_row(hdr_row)
    prevent_row_split(hdr_row)
    for c_idx, h_text in enumerate(headers):
        cell = hdr_row.cells[c_idx]
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=120, bottom=120, left=120, right=120)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    # Data rows
    for r_idx, r_data in enumerate(rows_data):
        row = tbl.rows[r_idx + 1]
        prevent_row_split(row)
        bg = "FFFFFF" if r_idx % 2 == 0 else LIGHT_BG_HEX
        is_highlight = r_idx == 0 and len(rows_data) >= 8  # Highlight top ranked model
        if is_highlight:
            bg = "E6F4F1"
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=95, bottom=95, left=110, right=110)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            if col_alignments and c_idx < len(col_alignments):
                p.alignment = col_alignments[c_idx]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(str(val))
            run.font.name = "Times New Roman"
            run.font.size = Pt(10.5)
            run.bold = (c_idx == 0 or is_highlight)
            run.font.color.rgb = DARK_TEXT_RGB

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(6)


def add_styled_figure(doc, img_path, caption_text, width_inches=5.9):
    if not os.path.exists(img_path):
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(3)
    p_img.paragraph_format.keep_with_next = True
    r_img = p_img.add_run()
    r_img.add_picture(img_path, width=Inches(width_inches))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(0)
    p_cap.paragraph_format.space_after = Pt(10)
    r_cap = p_cap.add_run(caption_text)
    r_cap.bold = True
    r_cap.italic = True
    r_cap.font.name = "Times New Roman"
    r_cap.font.size = Pt(10.5)
    r_cap.font.color.rgb = SLATE_RGB


def add_h1(doc, text):
    p = doc.add_heading(text, level=1)
    add_bottom_border_to_paragraph(p, color_hex=NAVY_HEX, sz="10")
    return p


def add_h2(doc, text):
    return doc.add_heading(text, level=2)


def add_h3(doc, text):
    return doc.add_heading(text, level=3)


def add_bullet(doc, bold_prefix, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.space_after = Pt(4)
    r1 = p.add_run("• " + bold_prefix)
    r1.bold = True
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(12)
    r1.font.color.rgb = NAVY_RGB
    r2 = p.add_run(text)
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(12)
    r2.font.color.rgb = DARK_TEXT_RGB
    return p


# ==============================================================================
# MASTER FINAL REPORT BUILDER (VI & EN)
# ==============================================================================
def generate_master_final_report(lang="vi"):
    is_vi = lang == "vi"
    fig_dir = FIG_VI if is_vi else FIG_EN

    doc = Document()
    configure_document_styles(doc)
    setup_header_footer(doc, lang=lang)

    # --------------------------------------------------------------------------
    # COVER PAGE
    # --------------------------------------------------------------------------
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_top.paragraph_format.space_before = Pt(4)
    p_top.paragraph_format.space_after = Pt(2)
    r_org = p_top.add_run(
        "CHƯƠNG TRÌNH ĐÀO TẠO SAMSUNG INNOVATION CAMPUS (SIC) — ĐỒ ÁN TỐT NGHIỆP AI"
        if is_vi
        else "SAMSUNG INNOVATION CAMPUS (SIC) — AI COURSE CAPSTONE PROJECT"
    )
    r_org.bold = True
    r_org.font.size = Pt(11.5)
    r_org.font.color.rgb = TEAL_RGB

    p_suborg = doc.add_paragraph()
    p_suborg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_suborg.paragraph_format.space_after = Pt(12)
    r_sub = p_suborg.add_run(
        "BÁO CÁO KỸ THUẬT TỔNG KẾT ĐỒ ÁN TRÍ TUỆ NHÂN TẠO & HỌC MÁY"
        if is_vi
        else "FINAL TECHNICAL REPORT — ARTIFICIAL INTELLIGENCE & MACHINE LEARNING CAPSTONE"
    )
    r_sub.bold = True
    r_sub.font.size = Pt(12.5)
    r_sub.font.color.rgb = NAVY_RGB
    add_bottom_border_to_paragraph(p_suborg, color_hex=NAVY_HEX, sz="16")

    # Title Banner Table
    tbl_title = doc.add_table(rows=1, cols=1)
    tbl_title.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_title = tbl_title.cell(0, 0)
    set_cell_background(c_title, NAVY_HEX)
    set_cell_margins(c_title, top=200, bottom=200, left=220, right=220)
    p_t = c_title.paragraphs[0]
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_after = Pt(6)
    rt1 = p_t.add_run(
        "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY CHO BÀI TOÁN DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CÁC CHỈ SỐ VỀ CĂNG THẲNG VÀ SỰ THAM GIA HỌC TẬP\n"
        if is_vi
        else "BUILDING, TRAINING, AND COMPARING MACHINE LEARNING MODELS FOR STUDENT DROPOUT RISK PREDICTION BASED ON STRESS AND STUDY ENGAGEMENT INDICATORS\n"
    )
    rt1.bold = True
    rt1.font.size = Pt(16)
    rt1.font.color.rgb = RGBColor(255, 255, 255)

    rt2 = p_t.add_run(
        "(Chuyên đề: Dự báo Kiệt sức & Nguy cơ Bỏ học — Hệ thống Cảnh báo Sớm Học đường)"
        if is_vi
        else "(Topic: Student Burnout & Dropout Risk Prediction — Proactive Academic Early Warning System)"
    )
    rt2.italic = True
    rt2.font.size = Pt(11.5)
    rt2.font.color.rgb = RGBColor(216, 243, 220)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # KPI Highlight Strip on Cover Page
    tbl_kpi = doc.add_table(rows=2, cols=4)
    tbl_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_kpi)
    kpis = (
        [
            ("2,000 Sinh viên", "800 gốc + 1,200 mở rộng sạch"),
            ("25 → 34 Cột", "Làm sạch 6 bước + 6 chỉ số mới"),
            ("4 Mô hình ML", "LR, RF, XGBoost, Soft Voting"),
            ("89.01% AUC", "Hiệu năng dự báo đỉnh (RF)"),
        ]
        if is_vi
        else [
            ("2,000 Students", "800 Original + 1,200 Expanded"),
            ("25 → 34 Cols", "6-Step Cleaned + 6 New Indices"),
            ("4 ML Models", "LR, RF, XGBoost, Soft Voting"),
            ("89.01% AUC", "Peak Predictive ROC-AUC (RF)"),
        ]
    )
    for idx, (big_v, sub_v) in enumerate(kpis):
        c_top = tbl_kpi.cell(0, idx)
        c_bot = tbl_kpi.cell(1, idx)
        set_cell_background(c_top, CALLOUT_BG_HEX)
        set_cell_background(c_bot, LIGHT_BG_HEX)
        set_cell_margins(c_top, top=75, bottom=30, left=70, right=70)
        set_cell_margins(c_bot, top=30, bottom=75, left=70, right=70)
        p0 = c_top.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(big_v)
        r0.bold = True
        r0.font.size = Pt(13)
        r0.font.color.rgb = NAVY_RGB

        p1 = c_bot.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(sub_v)
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = SLATE_RGB

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Metadata Box on Cover Page
    meta_rows = (
        [
            ["Chương trình đào tạo", "Samsung Innovation Campus (SIC) – Khóa học Trí tuệ Nhân tạo (AI)"],
            ["Đề tài phân công", "Bài toán dự đoán tỷ lệ bỏ học dựa trên chỉ số căng thẳng & tham gia học tập"],
            ["Nhóm thực hiện", "Nhóm 03"],
            ["Thành viên thực hiện", "1. Đỗ Quang Minh (Trưởng nhóm)\n2. Vũ Quang Huy (Thành viên)"],
            ["Bộ dữ liệu nghiên cứu", "student_burnout_dropout_dataset_2.csv (800 bản ghi × 25 trường thông tin)"],
            ["Sản phẩm bàn giao", "Pipeline Học máy, Bộ 12 Biểu đồ EDA/ML, Ứng dụng Web FastAPI, Báo cáo & Slide"],
            ["Ngày hoàn thành", "05 tháng 11 năm 2026"],
        ]
        if is_vi
        else [
            ["Training Program", "Samsung Innovation Campus (SIC) – Artificial Intelligence (AI) Course"],
            ["Assigned Topic", "Student Dropout Rate Prediction (Student Burnout & Engagement Analysis)"],
            ["Project Team", "Group 03"],
            ["Team Members", "1. Do Quang Minh (Team Leader)\n2. Vu Quang Huy (Team Member)"],
            ["Research Dataset", "student_burnout_dropout_dataset_2.csv (800 records × 25 attributes)"],
            ["Deliverables", "Leakage-Free ML Pipeline, 12 Figures, FastAPI Web Dashboard, Report & Slide Deck"],
            ["Completion Date", "November 05, 2026"],
        ]
    )
    add_styled_table(
        doc,
        "Thông tin Tổng quan Đồ án" if is_vi else "Project Metadata Summary",
        ["Hạng mục" if is_vi else "Category", "Thông tin chi tiết" if is_vi else "Details"],
        meta_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    p_copy = doc.add_paragraph()
    p_copy.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_copy.paragraph_format.space_before = Pt(8)
    p_copy.paragraph_format.space_after = Pt(0)
    r_copy = p_copy.add_run(
        "ⓒ 2025–2026 SAMSUNG INNOVATION CAMPUS. Tài liệu phục vụ báo cáo đồ án khóa học AI."
        if is_vi
        else "ⓒ 2025–2026 SAMSUNG INNOVATION CAMPUS. Prepared for the AI Course Capstone Evaluation."
    )
    r_copy.font.size = Pt(9.5)
    r_copy.italic = True
    r_copy.font.color.rgb = MUTED_TEXT_RGB

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # TABLE OF CONTENTS & LIST OF TABLES / FIGURES
    # --------------------------------------------------------------------------
    add_h1(doc, "MỤC LỤC TÀI LIỆU" if is_vi else "TABLE OF CONTENTS")
    p_toc_note = doc.add_paragraph()
    add_field(p_toc_note, 'TOC \\o "1-3" \\h \\z \\u', font_size=11, color_rgb=DARK_TEXT_RGB)

    toc_overview = (
        [
            ["Chương 1", "Giới thiệu Chung (Bối cảnh, Động lực, Mục tiêu, Phân công & Lộ trình)", "Trang 3"],
            ["Chương 2", "Phương pháp Thực hiện & Kiến trúc Hệ thống Học máy", "Trang 5"],
            ["Chương 3", "Kết quả Thực nghiệm: Tiền xử lý, Khám phá Dữ liệu (EDA) & So sánh 10 Mô hình", "Trang 8"],
            ["Chương 4", "Thiết kế Giao diện Người dùng (Web Dashboard) & Kiểm thử Cải tiến", "Trang 15"],
            ["Chương 5", "Đánh giá Tác động Thực tiễn, Định hướng Phát triển & Nhận xét Thành viên", "Trang 17"],
            ["Chương 6", "Danh mục Tài liệu Tham khảo & Phụ lục Kỹ thuật", "Trang 19"],
        ]
        if is_vi
        else [
            ["Chapter 1", "Introduction (Background, Motivation, Objectives, Team Roles & Schedule)", "Page 3"],
            ["Chapter 2", "Project Execution Methodology & System Architecture", "Page 5"],
            ["Chapter 3", "Experimental Results: Preprocessing, EDA & 10-Model Benchmarking", "Page 8"],
            ["Chapter 4", "Interactive Web Dashboard UI & System Testing/Refinements", "Page 15"],
            ["Chapter 5", "Projected Impact, Future Roadmap & Team Member Reviews", "Page 17"],
            ["Chapter 6", "References & Technical Appendix", "Page 19"],
        ]
    )
    add_styled_table(
        doc,
        "Cấu trúc Chương mục Báo cáo" if is_vi else "Report Chapter Structure Overview",
        ["Chương" if is_vi else "Chapter", "Nội dung trọng tâm" if is_vi else "Core Content", "Vị trí" if is_vi else "Location"],
        toc_overview,
        col_alignments=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    )

    list_figures = (
        [
            ["Hình 2.1", "Quy trình thực hiện dự án Học máy 5 giai đoạn (ML Project Workflow)"],
            ["Hình 2.2", "Kiến trúc hệ thống Cảnh báo sớm Nguy cơ Bỏ học 3 tầng & Web UI"],
            ["Hình 3.1", "Phân tích tỷ lệ dữ liệu khuyết trên 18/25 thuộc tính trước khi qua KNNImputer"],
            ["Hình 3.2", "Phân phối biến mục tiêu Dropout_Risk và tỷ lệ bỏ học theo mức độ Burnout_Level"],
            ["Hình 3.3", "So sánh phân phối 8 chỉ số Tham gia Học tập và Căng thẳng Tâm lý theo Dropout_Risk"],
            ["Hình 3.4", "Ma trận tương quan Pearson giữa 18 biến định lượng trong bộ dữ liệu"],
            ["Hình 3.5", "Phân tích tỷ lệ bỏ học theo 6 yếu tố nhân khẩu học và môi trường học tập"],
            ["Hình 3.6", "Phân tích đa biến giữa Tỷ lệ chuyên cần, Căng thẳng tâm lý và Hồ sơ Burnout"],
            ["Hình 3.7", "So sánh hiệu năng 10 mô hình học máy trên Kịch bản 1 (Full) và Kịch bản 2 (Early Warning)"],
            ["Hình 3.8", "Đường cong ROC và Precision-Recall của 10 mô hình trên tập kiểm thử (Test Set)"],
            ["Hình 3.9", "Ma trận nhầm lẫn (Confusion Matrix) của 4 mô hình tiêu biểu trên tập kiểm thử"],
            ["Hình 3.10", "Phân tích mức độ quan trọng của các đặc trưng (Feature Importance) trên Random Forest & XGBoost"],
        ]
        if is_vi
        else [
            ["Figure 2.1", "5-Stage Machine Learning Project Execution Workflow (Group 03)"],
            ["Figure 2.2", "3-Tier Student Dropout Early Warning System Architecture & Web Dashboard"],
            ["Figure 3.1", "Missing Data Percentage Across 18/25 Attributes Prior to KNN Imputation"],
            ["Figure 3.2", "Target Variable Distribution and Dropout Rate Progression by Burnout Level"],
            ["Figure 3.3", "Distribution Comparison of 8 Study Engagement & Psychological Stress Indicators"],
            ["Figure 3.4", "Pearson Correlation Matrix Across 18 Quantitative Student Attributes"],
            ["Figure 3.5", "Dropout Rate Analysis Across 6 Demographic & Environmental Factors"],
            ["Figure 3.6", "Multivariate Interaction Between Attendance, Stress Level, and Burnout Profile"],
            ["Figure 3.7", "Performance Benchmarking Across 10 ML Models in Scenario 1 & Scenario 2"],
            ["Figure 3.8", "ROC and Precision-Recall Curves Across All 10 Classifiers on the Hold-Out Test Set"],
            ["Figure 3.9", "Hold-Out Test Set Confusion Matrices for Top 4 Representative Models"],
            ["Figure 3.10", "Top 15 Feature Importances in Random Forest (Gini) and XGBoost (Gain)"],
        ]
    )
    add_styled_table(
        doc,
        "Danh mục Hình ảnh & Biểu đồ Trực quan" if is_vi else "List of Figures & Visualizations",
        ["Ký hiệu" if is_vi else "Figure ID", "Tên Hình ảnh / Biểu đồ" if is_vi else "Figure Title"],
        list_figures,
        col_alignments=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    )

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # CHAPTER 1: INTRODUCTION
    # --------------------------------------------------------------------------
    add_h1(doc, "1. GIỚI THIỆU CHUNG" if is_vi else "1. INTRODUCTION")

    add_callout_box(
        doc,
        "TÓM TẮT ĐIỀU HÀNH ĐỒ ÁN" if is_vi else "EXECUTIVE SUMMARY",
        (
            "Đồ án nghiên cứu bài toán Dự đoán Nguy cơ Bỏ học Đại học (Dropout_Risk) dựa trên sự tương tác giữa các chỉ số Căng thẳng Tâm lý (Stress, Anxiety, Financial Stress, Peer Pressure) và Sự Tham gia Học tập (Attendance, Study Hours, GPA, Motivation) trên 800 hồ sơ sinh viên. "
            "Nhóm 03 đã xây dựng thành công Pipeline học máy không rò rỉ dữ liệu kết hợp 6 chỉ số đặc trưng chuyên ngành và đối sánh 10 thuật toán phân loại. "
            "Mô hình Random Forest, Extra Trees và Soft Voting Ensemble đạt hiệu năng cao nhất với ROC-AUC = 80.50%, Test Accuracy = 75.63%, F1-Score = 78.92% và Recall = 83.91%."
            if is_vi
            else "This study addresses University Student Dropout Risk Prediction (Dropout_Risk) by modeling the quantitative interaction between Psychological Stress indicators and Academic Engagement metrics across 800 student profiles. "
            "Group 03 engineered a leakage-free ML pipeline featuring 6 domain-specific composite indices and benchmarked 10 classification algorithms. "
            "Random Forest, Extra Trees, and our Soft Voting Ensemble achieved top performance with 80.50% CV ROC-AUC, 75.63% Test Accuracy, 78.92% F1-Score, and 83.91% Recall."
        ),
    )

    add_h2(doc, "1.1. Bối cảnh Nghiên cứu" if is_vi else "1.1. Background Information")
    if is_vi:
        doc.add_paragraph(
            "Trong bối cảnh giáo dục đại học hiện đại, tỷ lệ sinh viên bỏ học giữa chừng (Student Attrition / Dropout) và hội chứng kiệt sức học đường (Academic Burnout) đang trở thành một thách thức chiến lược đối với các cơ sở đào tạo. Theo các nghiên cứu tâm lý giáo dục (Maslach & Leiter; Tinto's Student Integration Model), quyết định rời bỏ giảng đường của một sinh viên hiếm khi xảy ra đột ngột hoặc chỉ xuất phát từ năng lực nhận thức thuần túy. Thay vào đó, bỏ học là kết quả tích lũy của một quá trình suy giảm gắn kết kéo dài dưới tác động kép của hai nhóm yếu tố:"
        )
        add_bullet(
            doc,
            "Sự suy giảm mức độ tham gia học tập (Declining Academic Engagement): ",
            "Biểu hiện qua tỷ lệ chuyên cần sụt giảm (Attendance_Percent), thời lượng tự học mỗi ngày thấp (Study_Hours_Per_Day), suy giảm động lực nội tại (Motivation_Score), điểm trung bình học kỳ trước thấp (Previous_GPA) và tích lũy các môn nợ học phần (Backlogs).",
        )
        add_bullet(
            doc,
            "Áp lực căng thẳng tâm lý và tài chính kéo dài (Chronic Psychological & Financial Stress): ",
            "Bao gồm mức độ căng thẳng học tập cao (Stress_Level), rối loạn lo âu (Anxiety_Score), gánh nặng chi phí học tập - sinh hoạt (Financial_Stress_Score), áp lực đồng trang lứa (Peer_Pressure_Score), mất cân bằng nhịp sinh học do thiếu ngủ và lạm dụng thiết bị màn hình (Screen_Time_Hours), cùng sự thiếu hụt điểm tựa tinh thần từ gia đình (Family_Support_Score).",
        )
        doc.add_paragraph(
            "Khi áp lực tâm lý vượt quá ngưỡng chịu đựng mà không có sự hỗ trợ kịp thời, sinh viên rơi vào trạng thái Kiệt sức Học đường (Burnout_Level từ Low chuyển sang Medium và High), kéo theo nguy cơ bỏ học tăng vọt. Việc ứng dụng Học máy (Machine Learning) để phân tích đồng thời các chỉ số này cho phép nhà trường chuyển đổi từ cơ chế xử lý bị động sang Hệ thống Cảnh báo Sớm Chủ động (Proactive Early Warning System)."
        )
    else:
        doc.add_paragraph(
            "In contemporary higher education, student attrition (dropout) and academic burnout pose critical institutional and socioeconomic challenges. Grounded in Tinto's Student Integration Model and Maslach's Burnout Theory, empirical educational research establishes that university dropout rarely stems from cognitive inability alone. Instead, student attrition is the cumulative outcome of two interacting behavioral and psychological mechanisms:"
        )
        add_bullet(
            doc,
            "Declining Academic Engagement: ",
            "Manifested through deteriorating class attendance (Attendance_Percent), reduced daily self-study hours (Study_Hours_Per_Day), eroding intrinsic motivation (Motivation_Score), lower prior academic performance (Previous_GPA), and accumulating course failures (Backlogs).",
        )
        add_bullet(
            doc,
            "Chronic Psychological & Financial Stress: ",
            "Driven by elevated academic stress (Stress_Level), generalized anxiety (Anxiety_Score), financial hardship (Financial_Stress_Score), peer comparison pressure (Peer_Pressure_Score), circadian disruption from excessive screen time (Screen_Time_Hours) and sleep deprivation (Sleep_Hours), compounded by inadequate family support (Family_Support_Score).",
        )
        doc.add_paragraph(
            "When sustained stress outpaces available support buffers, students transition into severe Academic Burnout (Burnout_Level progressing from Low to Medium and High), precipitating university dropout. Leveraging Machine Learning to jointly model these indicators enables universities to shift from reactive post-hoc administration to a Proactive Early Warning System."
        )

    add_h2(doc, "1.2. Động lực và Mục tiêu Đề tài" if is_vi else "1.2. Motivation and Objectives")
    if is_vi:
        doc.add_paragraph(
            "Tại phần lớn các trường đại học hiện nay, dữ liệu điểm danh trên lớp, bảng điểm học tập và phiếu khảo sát tâm lý - đời sống sinh viên thường được lưu trữ phân tán ở các phòng ban khác nhau, dẫn đến việc thiếu vắng một mô hình định lượng tổng hợp để sàng lọc rủi ro tự động. Xuất phát từ thực tế đó, Nhóm 03 thực hiện đề tài với 4 mục tiêu cụ thể:"
        )
        add_bullet(doc, "Mục tiêu 1 – Phân tích Khám phá Dữ liệu Chuyên sâu (Comprehensive EDA): ", "Khảo sát thống kê trên 800 sinh viên × 25 thuộc tính để lượng hóa mối quan hệ giữa Căng thẳng, Tham gia học tập, Mức độ kiệt sức và Nguy cơ bỏ học.")
        add_bullet(doc, "Mục tiêu 2 – Tiền xử lý & Kỹ thuật Tạo Đặc trưng (Preprocessing & Feature Engineering): ", "Xử lý triệt để 18/25 cột khuyết dữ liệu bằng KNNImputer (k=7) và thiết kế 6 chỉ số tổng hợp phản ánh cán cân Căng thẳng – Gắn kết.")
        add_bullet(doc, "Mục tiêu 3 – Huấn luyện & Đối sánh 10 Mô hình Học máy (10-Model Benchmarking): ", "Đánh giá khách quan 10 thuật toán phân loại bằng kiểm định chéo phân tầng 5-Fold và tập kiểm thử độc lập 20% trên cả 2 kịch bản (Có và Không có biến trung gian Burnout_Level).")
        add_bullet(doc, "Mục tiêu 4 – Xây dựng Ứng dụng Web Cảnh báo Sớm (Interactive Web Dashboard): ", "Phát triển hệ thống Web UI bằng FastAPI cho phép dự báo trực tiếp Burnout_Level & Dropout_Risk kèm khuyến nghị can thiệp cá nhân hóa.")
    else:
        doc.add_paragraph(
            "In most universities, attendance records, academic transcripts, and student wellness surveys remain siloed across separate administrative units, preventing holistic risk screening. To bridge this gap, Group 03 established four concrete technical objectives:"
        )
        add_bullet(doc, "Objective 1 – In-Depth Exploratory Data Analysis (EDA): ", "Quantify statistical relationships across 800 student records and 25 attributes linking Stress, Study Engagement, Burnout Level, and Dropout Risk.")
        add_bullet(doc, "Objective 2 – Leakage-Free Preprocessing & Domain Feature Engineering: ", "Impute missing values across 18/25 columns via multivariate KNNImputer (k=7) and construct 6 composite indices capturing the Stress–Engagement balance.")
        add_bullet(doc, "Objective 3 – Systematic 10-Model Benchmarking Across Two Scenarios: ", "Evaluate 10 classification algorithms via 5-Fold Stratified Cross-Validation and a 20% Hold-out Test Set under Scenario 1 (Full Features) and Scenario 2 (Early Warning without Burnout_Level).")
        add_bullet(doc, "Objective 4 – Interactive Web Dashboard Deployment: ", "Build a real-time FastAPI Web UI providing dual Burnout_Level and Dropout_Risk predictions alongside personalized intervention plans.")

    add_h2(doc, "1.3. Thành viên và Phân công Nhiệm vụ" if is_vi else "1.3. Members and Role Assignments")
    role_rows = (
        [
            [
                "1. Đỗ Quang Minh\n(Trưởng nhóm / Team Leader)",
                "50%",
                "• Thiết kế kiến trúc Pipeline Học máy chống rò rỉ dữ liệu (No Data Leakage).\n"
                "• Xây dựng module DomainFeatureEngineer gồm 6 chỉ số tương tác chuyên ngành.\n"
                "• Huấn luyện & tối ưu nhóm mô hình Cây tập hợp & Boosting (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
                "• Phát triển Backend FastAPI (app.py) và tích hợp bộ đôi mô hình dự báo.",
            ],
            [
                "2. Vũ Quang Huy\n(Thành viên / Team Member)",
                "50%",
                "• Kiểm toán chất lượng dữ liệu, thống kê 18 cột khuyết & thực hiện Phân tích EDA.\n"
                "• Thiết kế trực quan bộ 12 biểu đồ phân tích thống kê và đánh giá mô hình.\n"
                "• Huấn luyện & đánh giá nhóm mô hình Cơ sở & Mạng nơ-ron (Logistic Regression, SVM RBF, KNN, Decision Tree, MLP Neural Network).\n"
                "• Thiết kế giao diện Frontend Web Dashboard, biên soạn Báo cáo, WBS & Slide.",
            ],
        ]
        if is_vi
        else [
            [
                "1. Do Quang Minh\n(Team Leader)",
                "50%",
                "• Architected the leakage-free scikit-learn ML Pipeline.\n"
                "• Engineered the DomainFeatureEngineer transformer (6 composite indices).\n"
                "• Trained, tuned, and evaluated Tree Ensemble & Boosting models (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
                "• Developed the FastAPI Backend inference server and dual model integration.",
            ],
            [
                "2. Vu Quang Huy\n(Team Member)",
                "50%",
                "• Conducted data quality auditing, missingness profiling across 18 columns & EDA.\n"
                "• Designed the 12 publication-grade statistical and model evaluation figures.\n"
                "• Trained and benchmarked Baseline, Kernel, Distance & Neural classifiers (Logistic Regression, SVM RBF, KNN, Decision Tree, MLP Neural Network).\n"
                "• Built the Frontend Web Dashboard UI and authored the Final Report, WBS & Slides.",
            ],
        ]
    )
    add_styled_table(
        doc,
        "Bảng 1.1: Phân công Nhiệm vụ Chi tiết của Nhóm 03" if is_vi else "Table 1.1: Group 03 Member Role Assignments & Contributions",
        ["Thành viên" if is_vi else "Team Member", "Tỷ trọng" if is_vi else "Share", "Hạng mục Công việc Phụ trách" if is_vi else "Assigned Technical Responsibilities"],
        role_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    )

    add_h2(doc, "1.4. Lịch trình và Các mốc Quan trọng" if is_vi else "1.4. Schedule and Milestones")
    sched_rows = (
        [
            ["Giai đoạn 1\n(19/10 – 20/10/2026)", "Khởi tạo Dự án & Đặc tả Yêu cầu", "Action Plan, WBS & Đặc tả 25 thuộc tính", "Đỗ Quang Minh,\nVũ Quang Huy", "Hoàn thành (100%)"],
            ["Giai đoạn 2\n(21/10 – 25/10/2026)", "Phân tích EDA, Tiền xử lý & Tạo Đặc trưng", "Pipeline KNNImputer, 6 chỉ số mới & 6 biểu đồ EDA", "Vũ Quang Huy,\nĐỗ Quang Minh", "Hoàn thành (100%)"],
            ["Giai đoạn 3\n(26/10 – 29/10/2026)", "Huấn luyện, Kiểm định chéo & So sánh 10 Mô hình", "10 Mô hình trên 2 kịch bản, ROC/PR, CM & Importance", "Đỗ Quang Minh,\nVũ Quang Huy", "Hoàn thành (100%)"],
            ["Giai đoạn 4\n(30/10 – 05/11/2026)", "Phát triển Web UI Demo & Hoàn thiện Hồ sơ SIC", "Ứng dụng FastAPI, Notebook, Báo cáo & Slide", "Đỗ Quang Minh,\nVũ Quang Huy", "Hoàn thành (100%)"],
        ]
        if is_vi
        else [
            ["Phase 1\n(Oct 19 – Oct 20, 2026)", "Project Initiation & Problem Specification", "Action Plan, WBS & 25-Attribute Schema Audit", "Do Quang Minh,\nVu Quang Huy", "Completed (100%)"],
            ["Phase 2\n(Oct 21 – Oct 25, 2026)", "Statistical EDA, Preprocessing & Feature Eng.", "KNNImputer Pipeline, 6 Indices & 6 EDA Plots", "Vu Quang Huy,\nDo Quang Minh", "Completed (100%)"],
            ["Phase 3\n(Oct 26 – Oct 29, 2026)", "Training, 5-Fold CV & 10-Model Benchmarking", "10 Models x 2 Scenarios, ROC/PR, CM & Importance", "Do Quang Minh,\nVu Quang Huy", "Completed (100%)"],
            ["Phase 4\n(Oct 30 – Nov 05, 2026)", "FastAPI Web UI Deployment & Final Deliverables", "Interactive Web App, Notebook, Report & Slides", "Do Quang Minh,\nVu Quang Huy", "Completed (100%)"],
        ]
    )
    add_styled_table(
        doc,
        "Bảng 1.2: Lộ trình Thực hiện và Các mốc Bàn giao Chính (Milestones)" if is_vi else "Table 1.2: Project Execution Schedule and Key Milestones",
        [
            "Giai đoạn" if is_vi else "Phase / Timeline",
            "Nội dung thực hiện" if is_vi else "Milestone Activity",
            "Sản phẩm đầu ra" if is_vi else "Key Deliverables",
            "Phụ trách" if is_vi else "Owner",
            "Trạng thái" if is_vi else "Status",
        ],
        sched_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
    )

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # CHAPTER 2: PROJECT EXECUTION & METHODOLOGY
    # --------------------------------------------------------------------------
    add_h1(doc, "2. TRIỂN KHAI DỰ ÁN VÀ PHƯƠNG PHÁP LUẬN" if is_vi else "2. PROJECT EXECUTION AND METHODOLOGY")

    add_h2(doc, "2.1. Thu thập và Đặc tả Bộ Dữ liệu" if is_vi else "2.1. Data Acquisition and Dataset Schema")
    doc.add_paragraph(
        "Bộ dữ liệu nghiên cứu 'student_burnout_dropout_dataset_2.csv' bao gồm 800 bản ghi sinh viên đại học và 25 trường thuộc tính. Để phục vụ việc thiết kế mô hình khoa học, nhóm phân loại 25 thuộc tính thành 5 nhóm chức năng như trình bày tại Bảng 2.1:"
        if is_vi
        else "The empirical dataset 'student_burnout_dropout_dataset_2.csv' comprises 800 university student records across 25 attributes. To structure our feature engineering and modeling workflow, we categorize the 25 attributes into 5 functional domains as detailed in Table 2.1:"
    )

    schema_rows = (
        [
            ["1. Định danh & Nhân khẩu học\n(7 thuộc tính)", "Student_ID, Age, Gender, Year_of_Study, Department, Residence_Type, Family_Income_Bracket", "Int / Str", "Tuổi từ 17–25 (TB: 20.88); 6 Khoa (Engineering: 209, Business: 151, Science: 143, Arts: 136, Medicine: 77, Law: 61); 3 loại cư trú; 5 mức thu nhập."],
            ["2. Sự Tham gia Học tập\n(5 thuộc tính)", "Attendance_Percent, Study_Hours_Per_Day, Previous_GPA, Backlogs, Motivation_Score", "Float64", "Chuyên cần: 30%–100% (TB: 76.88%); Giờ tự học: 0–8.1h (TB: 3.16h); GPA kỳ trước: 2.42–10.0 (TB: 6.86); Môn nợ: 0–6; Động lực: 0–10 (TB: 5.47)."],
            ["3. Lối sống & Sinh hoạt\n(5 thuộc tính)", "Sleep_Hours, Screen_Time_Hours, Exercise_Freq_Per_Week, Social_Activity_Score, Part_Time_Job", "Float / Int / Str", "Giờ ngủ: 2.0–10.0h (TB: 6.20h); Giờ màn hình: 0.5–12.6h (TB: 5.43h); Tập thể dục: 0–7 buổi/tuần; Làm thêm: Yes (228), No (547)."],
            ["4. Căng thẳng & Hỗ trợ Tâm lý\n(6 thuộc tính)", "Stress_Level, Anxiety_Score, Financial_Stress_Score, Peer_Pressure_Score, Family_Support_Score, Counseling_Access", "Float / Str", "Stress: 0.1–10.0 (TB: 5.52); Lo âu: 0–10 (TB: 4.98); Áp lực tài chính: 0–10 (TB: 5.03); Hỗ trợ GĐ: 0–10 (TB: 6.42); Tư vấn tâm lý: Yes (265), No (535)."],
            ["5. Kiệt sức & Biến Mục tiêu\n(2 thuộc tính)", "Burnout_Level,\nDropout_Risk (Target)", "Categorical", "Burnout_Level: Low (320 - 40%), Medium (280 - 35%), High (200 - 25%).\nDropout_Risk: Yes (437 - 54.62%), No (363 - 45.38%). Không có giá trị khuyết."],
        ]
        if is_vi
        else [
            ["1. ID & Demographics\n(7 attributes)", "Student_ID, Age, Gender, Year_of_Study, Department, Residence_Type, Family_Income_Bracket", "Int / Str", "Age: 17–25 (mean: 20.88); 6 Departments (Engineering: 209, Business: 151, Science: 143, Arts: 136, Medicine: 77, Law: 61); 3 residence types; 5 income tiers."],
            ["2. Academic Engagement\n(5 attributes)", "Attendance_Percent, Study_Hours_Per_Day, Previous_GPA, Backlogs, Motivation_Score", "Float64", "Attendance: 30%–100% (mean: 76.88%); Study Hours: 0–8.1h (mean: 3.16h); Previous GPA: 2.42–10.0 (mean: 6.86); Backlogs: 0–6; Motivation: 0–10 (mean: 5.47)."],
            ["3. Lifestyle & Habits\n(5 attributes)", "Sleep_Hours, Screen_Time_Hours, Exercise_Freq_Per_Week, Social_Activity_Score, Part_Time_Job", "Float / Int / Str", "Sleep: 2.0–10.0h (mean: 6.20h); Screen Time: 0.5–12.6h (mean: 5.43h); Exercise: 0–7 days/week; Part-Time Job: Yes (228), No (547)."],
            ["4. Stress & Psych Support\n(6 attributes)", "Stress_Level, Anxiety_Score, Financial_Stress_Score, Peer_Pressure_Score, Family_Support_Score, Counseling_Access", "Float / Str", "Stress: 0.1–10.0 (mean: 5.52); Anxiety: 0–10 (mean: 4.98); Financial Stress: 0–10 (mean: 5.03); Family Support: 0–10 (mean: 6.42); Counseling: Yes (265), No (535)."],
            ["5. Burnout & Target\n(2 attributes)", "Burnout_Level,\nDropout_Risk (Target)", "Categorical", "Burnout_Level: Low (320 - 40%), Medium (280 - 35%), High (200 - 25%).\nDropout_Risk: Yes (437 - 54.62%), No (363 - 45.38%). Zero missing target values."],
        ]
    )
    add_styled_table(
        doc,
        "Bảng 2.1: Đặc tả Chi tiết 25 Thuộc tính trong Bộ Dữ liệu Nghiên cứu" if is_vi else "Table 2.1: Comprehensive Schema of the 25 Dataset Attributes",
        [
            "Nhóm Thuộc tính" if is_vi else "Attribute Group",
            "Danh sách Biến (Columns)" if is_vi else "Variables (Columns)",
            "Kiểu" if is_vi else "Type",
            "Đặc trưng Thống kê & Miền giá trị" if is_vi else "Statistical Profile & Value Range",
        ],
        schema_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    )

    add_h2(doc, "2.2. Phương pháp Huấn luyện và Đánh giá" if is_vi else "2.2. Training and Evaluation Methodology")
    if is_vi:
        doc.add_paragraph(
            "Để đảm bảo kết quả thực nghiệm phản ánh chính xác năng lực tổng quát hóa của mô hình trên dữ liệu thực tế chưa từng gặp, nhóm áp dụng quy trình huấn luyện chuẩn mực:"
        )
        add_bullet(doc, "Phân chia Tập Huấn luyện / Kiểm thử Phân tầng (Stratified 80/20 Train-Test Split): ", "640 sinh viên (80%) dùng cho huấn luyện và 160 sinh viên (20%) dùng làm tập kiểm thử độc lập (Hold-out Test Set, random_state=42), bảo toàn tỷ lệ 54.6% Yes / 45.4% No ở cả hai tập.")
        add_bullet(doc, "Kiểm định Chéo Phân tầng 5-Fold (5-Fold Stratified Cross-Validation): ", "Mọi bước biến đổi dữ liệu (DomainFeatureEngineer, KNNImputer, StandardScaler, OneHotEncoder) được đóng gói trong đối tượng Pipeline của scikit-learn và chỉ khớp (fit) trên 4 fold huấn luyện của từng vòng lặp, ngăn chặn tuyệt đối hiện tượng rò rỉ dữ liệu (Data Leakage).")
        add_bullet(doc, "Thiết kế 2 Kịch bản Thực nghiệm Đối chứng (Two Experimental Scenarios): ", "Kịch bản 1 (Full Features) sử dụng toàn bộ đặc trưng bao gồm cả trạng thái Burnout_Level để đạt độ chính xác chẩn đoán tối đa; Kịch bản 2 (Early Warning) loại bỏ hoàn toàn biến Burnout_Level, buộc các mô hình dự báo nguy cơ bỏ học trực tiếp từ các chỉ số hành vi học tập và căng thẳng tâm lý thô.")
    else:
        doc.add_paragraph(
            "To guarantee unbiased generalization estimates and strict reproducibility, our experimental protocol enforces three methodological safeguards:"
        )
        add_bullet(doc, "Stratified 80/20 Train-Test Partitioning: ", "The 800 records are split into 640 training instances (80%) and 160 hold-out test instances (20%) with stratified sampling (random_state=42), preserving the exact 54.6% Yes / 45.4% No class distribution.")
        add_bullet(doc, "Leakage-Free 5-Fold Stratified Cross-Validation: ", "All feature engineering, KNN imputation, scaling, and one-hot encoding operations are encapsulated inside a unified scikit-learn Pipeline so statistics are computed strictly within each training fold.")
        add_bullet(doc, "Two Controlled Experimental Scenarios: ", "Scenario 1 (Full Feature Set) incorporates all features including Burnout_Level for maximum diagnostic accuracy; Scenario 2 (Early Warning) completely removes Burnout_Level to evaluate early dropout prediction solely from raw behavioral engagement and psychological stress indicators.")

    add_h2(doc, "2.3. Quy trình Thực hiện Dự án" if is_vi else "2.3. End-to-End Project Workflow")
    doc.add_paragraph(
        "Quy trình thực thi tổng thể gồm 5 giai đoạn nối tiếp nhau từ thu thập dữ liệu, xử lý khuyết & phân tích EDA, kỹ thuật tạo đặc trưng chuyên ngành, huấn luyện & kiểm định 10 mô hình học máy đến triển khai ứng dụng Web cảnh báo sớm được minh họa tại Hình 2.1:"
        if is_vi
        else "The end-to-end execution workflow spanning Data Acquisition, Preprocessing & EDA, Domain Feature Engineering, 10-Model Cross-Validation Benchmarking, and Web UI Deployment is illustrated in Figure 2.1:"
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "11_system_workflow_diagram.png"),
        "Hình 2.1: Quy trình thực hiện dự án Học máy 5 giai đoạn của Nhóm 03" if is_vi else "Figure 2.1: 5-Stage Machine Learning Project Execution Workflow (Group 03)",
        width_inches=6.0,
    )

    add_h2(doc, "2.4. Thiết kế Kiến trúc Hệ thống" if is_vi else "2.4. System Architecture Design")
    doc.add_paragraph(
        "Để phục vụ triển khai thực tế tại trường đại học, hệ thống phần mềm Cảnh báo Sớm được thiết kế theo kiến trúc 3 tầng (3-Tier Architecture) gồm: (1) Tầng Giao diện Web UI thu thập thông số sinh viên; (2) Tầng Xử lý AI Backend (FastAPI) vận hành song song mô hình dự báo Burnout_Level (3 lớp) và mô hình dự báo Dropout_Risk; (3) Tầng Hiển thị Kết quả & Khuyến nghị can thiệp cá nhân hóa (Hình 2.2)."
        if is_vi
        else "For institutional deployment, our Early Warning System is engineered around a modular 3-Tier Architecture: (1) Presentation Layer (Web UI Input Form); (2) FastAPI AI Processing Backend running dual Burnout_Level (3-class) and Dropout_Risk ensemble inference pipelines; and (3) Diagnostic Output Dashboard generating risk probabilities and actionable counseling interventions (Figure 2.2)."
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "12_system_architecture_diagram.png"),
        "Hình 2.2: Kiến trúc hệ thống Cảnh báo sớm Nguy cơ Bỏ học 3 tầng & Giao diện Web Dashboard" if is_vi else "Figure 2.2: 3-Tier Student Dropout Early Warning System Architecture & Web Dashboard",
        width_inches=6.0,
    )

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # CHAPTER 3: EXPERIMENTAL RESULTS (PREPROCESSING, EDA, MODELING)
    # --------------------------------------------------------------------------
    add_h1(doc, "3. KẾT QUẢ THỰC NGHIỆM: TIỀN XỬ LÝ, PHÂN TÍCH EDA & MÔ HÌNH HÓA" if is_vi else "3. EXPERIMENTAL RESULTS: PREPROCESSING, EDA & MODELING")

    add_h2(doc, "3.1. Tiền xử lý Dữ liệu và Kỹ thuật Tạo Đặc trưng" if is_vi else "3.1. Data Preprocessing and Domain Feature Engineering")
    doc.add_paragraph(
        "Qua kiểm toán chất lượng dữ liệu, có 18/25 cột chứa giá trị khuyết với tỷ lệ dao động từ 1.62% (Family_Income_Bracket - 13 mẫu) đến 6.38% (Attendance_Percent - 51 mẫu) và 6.12% (Stress_Level - 49 mẫu) như thể hiện trên Hình 3.1. Nếu áp dụng phương pháp xóa dòng khuyết (Listwise Deletion), tập dữ liệu sẽ mất hơn 40% số bản ghi. Do đó, Nhóm 03 thiết kế quy trình xử lý khuyết đa biến kết hợp kỹ thuật tạo 6 đặc trưng chuyên ngành (Bảng 3.1):"
        if is_vi
        else "Data auditing revealed missing values across 18 of the 25 attributes, ranging from 1.62% (Family_Income_Bracket - 13 records) to 6.38% (Attendance_Percent - 51 records) and 6.12% (Stress_Level - 49 records), as shown in Figure 3.1. Because listwise deletion would discard over 40% of the sample, we implemented multivariate KNN imputation (k=7) alongside 6 domain-engineered composite indices (Table 3.1):"
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "01_missing_values_analysis.png"),
        "Hình 3.1: Phân tích tỷ lệ dữ liệu khuyết trên 18/25 thuộc tính trước khi qua KNNImputer" if is_vi else "Figure 3.1: Missing Data Percentage Across 18/25 Attributes Prior to KNN Imputation",
        width_inches=5.7,
    )

    fe_rows = (
        [
            ["1. Academic_Engagement_Index", "0.35*(Attendance/10) + 0.25*Study_Hours + 0.25*Motivation + 0.15*GPA", "Đo lường tổng hợp mức độ chuyên cần, thời lượng tự học, động lực và kết quả học tập."],
            ["2. Psychological_Stress_Index", "0.35*Stress + 0.30*Anxiety + 0.20*Financial_Stress + 0.15*Peer_Pressure", "Đo lường tổng hợp cường độ áp lực tâm lý, lo âu, tài chính và áp lực đồng trang lứa."],
            ["3. Stress_Engagement_Ratio", "Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)", "Tỷ số phản ánh sự mất cân bằng giữa áp lực tâm lý và mức độ gắn kết học tập."],
            ["4. Net_Stress_Support_Balance", "Psychological_Stress_Index - (0.5*Family + 0.3*Social + 0.2*Exercise)", "Đo lường mức căng thẳng ròng sau khi trừ đi các yếu tố đệm hỗ trợ tinh thần - thể chất."],
            ["5. Digital_Fatigue_Ratio", "Screen_Time_Hours / (Sleep_Hours + 1.0)", "Đánh giá mức độ mệt mỏi số và rối loạn nhịp sinh học do dùng màn hình quá mức."],
            ["6. Academic_Risk_Burden", "1.5*Backlogs + (100 - Attendance_Percent) / 20.0", "Định lượng gánh nặng học vụ tích lũy từ số môn nợ và độ thiếu hụt chuyên cần."],
        ]
        if is_vi
        else [
            ["1. Academic_Engagement_Index", "0.35*(Attendance/10) + 0.25*Study_Hours + 0.25*Motivation + 0.15*GPA", "Quantifies overall academic commitment across attendance, study hours, motivation, and GPA."],
            ["2. Psychological_Stress_Index", "0.35*Stress + 0.30*Anxiety + 0.20*Financial_Stress + 0.15*Peer_Pressure", "Aggregates psychological, emotional, financial, and peer-related stress burdens."],
            ["3. Stress_Engagement_Ratio", "Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)", "Captures the relative imbalance between psychological distress and academic engagement."],
            ["4. Net_Stress_Support_Balance", "Psychological_Stress_Index - (0.5*Family + 0.3*Social + 0.2*Exercise)", "Measures net unbuffered stress after accounting for family, social, and physical activity support."],
            ["5. Digital_Fatigue_Ratio", "Screen_Time_Hours / (Sleep_Hours + 1.0)", "Quantifies circadian disruption and digital fatigue from excessive screen exposure relative to sleep."],
            ["6. Academic_Risk_Burden", "1.5*Backlogs + (100 - Attendance_Percent) / 20.0", "Combines course failure accumulation (backlogs) with class attendance deficit."],
        ]
    )
    add_styled_table(
        doc,
        "Bảng 3.1: Hệ thống 6 Đặc trưng Chuyên ngành Mới trong DomainFeatureEngineer" if is_vi else "Table 3.1: Specification of the 6 Engineered Domain Features in DomainFeatureEngineer",
        ["Tên Đặc trưng Mới" if is_vi else "Engineered Feature", "Công thức Toán học" if is_vi else "Mathematical Formulation", "Ý nghĩa Nghiệp vụ Giáo dục" if is_vi else "Educational & Behavioral Interpretation"],
        fe_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    add_h2(doc, "3.2. Phân tích Khám phá Dữ liệu (EDA)" if is_vi else "3.2. Exploratory Data Analysis (EDA)")
    doc.add_paragraph(
        "Phân tích thống kê trên 800 sinh viên cho thấy biến mục tiêu Dropout_Risk có phân phối cân bằng (437 sinh viên thuộc nhóm Yes [54.62%] và 363 sinh viên thuộc nhóm No [45.38%]). Đặc biệt, Hình 3.2 chứng minh mối liên hệ nhân quả chặt chẽ giữa Mức độ Kiệt sức (Burnout_Level) và Tỷ lệ Bỏ học: ở nhóm Low Burnout (320 SV), tỷ lệ bỏ học chỉ là 23.8%; sang nhóm Medium Burnout (280 SV), tỷ lệ bỏ học tăng lên 63.2%; và ở nhóm High Burnout (200 SV), tỷ lệ bỏ học vọt lên tới 92.0% (184/200 SV)!"
        if is_vi
        else "Statistical profiling across the 800 students confirms a well-balanced binary target distribution (437 Yes [54.62%] vs. 363 No [45.38%]). Crucially, Figure 3.2 reveals a dramatic monotonic escalation in dropout rates across Burnout_Level tiers: students with Low Burnout (n=320) exhibit only a 23.8% dropout rate; Medium Burnout students (n=280) experience a 63.2% dropout rate; and High Burnout students (n=200) reach an alarming 92.0% dropout rate (184/200 students)!"
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "02_target_and_burnout_distribution.png"),
        "Hình 3.2: Phân phối biến mục tiêu Dropout_Risk và tỷ lệ bỏ học theo từng mức độ Burnout_Level" if is_vi else "Figure 3.2: Target Variable Distribution and Dropout Rate Progression Across Burnout Levels",
        width_inches=6.0,
    )

    doc.add_paragraph(
        "Tiếp theo, Hình 3.3 và Hình 3.4 làm rõ sự đối lập giữa nhóm chỉ số Tham gia Học tập và nhóm chỉ số Căng thẳng Tâm lý. Nhóm sinh viên có nguy cơ bỏ học (Dropout_Risk = Yes) có trung bình chuyên cần thấp hơn 7.0 điểm phần trăm (73.7% so với 80.7%), giờ tự học ít hơn (2.93h so với 3.43h/ngày), điểm động lực thấp hơn (5.11 so với 5.90), đồng thời chịu mức căng thẳng (5.81 vs 5.17) và lo âu (5.26 vs 4.65) cao hơn hẳn nhóm an toàn."
        if is_vi
        else "Furthermore, Figures 3.3 and 3.4 demonstrate the opposing effects of Study Engagement versus Psychological Stress. At-risk students (Dropout_Risk = Yes) exhibit a 7.0 percentage-point deficit in mean attendance (73.7% vs. 80.7%), fewer daily study hours (2.93h vs. 3.43h), and lower motivation (5.11 vs. 5.90), while experiencing significantly higher stress (5.81 vs. 5.17) and anxiety (5.26 vs. 4.65)."
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "03_stress_vs_engagement_distributions.png"),
        "Hình 3.3: So sánh phân phối 8 chỉ số Tham gia Học tập và Căng thẳng Tâm lý theo Dropout_Risk" if is_vi else "Figure 3.3: Distribution Comparison of 8 Study Engagement & Psychological Stress Indicators",
        width_inches=6.0,
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "04_correlation_matrix_heatmap.png"),
        "Hình 3.4: Ma trận tương quan Pearson giữa 18 biến định lượng trong bộ dữ liệu" if is_vi else "Figure 3.4: Pearson Correlation Matrix Across 18 Quantitative Student Attributes",
        width_inches=5.7,
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "05_categorical_factors_analysis.png"),
        "Hình 3.5: Phân tích tỷ lệ bỏ học theo 6 yếu tố nhân khẩu học và môi trường học tập" if is_vi else "Figure 3.5: Dropout Rate Analysis Across 6 Demographic & Environmental Factors",
        width_inches=6.0,
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "06_multivariate_burnout_engagement_scatter.png"),
        "Hình 3.6: Phân tích đa biến giữa Tỷ lệ chuyên cần, Căng thẳng tâm lý và Hồ sơ Burnout" if is_vi else "Figure 3.6: Multivariate Interaction Between Attendance, Stress Level, and Burnout Profile",
        width_inches=6.0,
    )

    add_h2(doc, "3.3. Kết quả Huấn luyện và So sánh 4 Mô hình Học máy Tiêu biểu" if is_vi else "3.3. Training Results and 4-Model Benchmarking")
    doc.add_paragraph(
        "Nhóm tiến hành huấn luyện, kiểm định chéo 5-Fold và đánh giá trên tập Test độc lập (400 mẫu - 20% của 2,000 sinh viên) đối với 4 mô hình học máy tiêu biểu (Logistic Regression, Random Forest, XGBoost, Soft Voting Ensemble) trên cả 2 kịch bản. Kết quả chi tiết được tổng hợp tại Bảng 3.2 (Kịch bản 1: Đầy đủ đặc trưng) và Bảng 3.3 (Kịch bản 2: Cảnh báo sớm không dùng biến Burnout_Level):"
        if is_vi
        else "We trained, cross-validated (5-Fold Stratified CV), and evaluated all 4 flagship machine learning classifiers on the 20% hold-out test set (n=400 out of 2,000 students) across both scenarios. Detailed quantitative metrics are presented in Table 3.2 (Scenario 1: Full Feature Set) and Table 3.3 (Scenario 2: Early Warning without Burnout_Level):"
    )

    s1_rows = [
        ["1. Random Forest", "80.25% ± 2.15%", "79.88%", "86.38%", "82.98%", "89.01%", "81.50%", "83.63%", "88.42%"],
        ["2. Soft Voting Ensemble", "79.70% ± 2.14%", "79.96%", "84.74%", "82.27%", "88.89%", "80.25%", "82.17%", "88.32%"],
        ["3. XGBoost", "79.95% ± 2.28%", "80.39%", "84.56%", "82.39%", "88.64%", "78.50%", "80.28%", "87.99%"],
        ["4. Logistic Regression", "78.45% ± 2.09%", "79.68%", "81.64%", "80.62%", "87.61%", "78.25%", "80.18%", "87.60%"],
    ]
    add_styled_table(
        doc,
        "Bảng 3.2: Kết quả Thực nghiệm 4 Mô hình Học máy Tiêu biểu — Kịch bản 1 (Đầy đủ Đặc trưng, Có Burnout_Level)" if is_vi else "Table 3.2: 4-Model Experimental Results — Scenario 1 (Full Feature Set Including Burnout_Level)",
        ["Mô hình (Model)" if is_vi else "Classifier", "CV Accuracy", "CV Prec", "CV Recall", "CV F1", "CV AUC", "Test Acc", "Test F1", "Test AUC"],
        s1_rows,
    )

    s2_rows = [
        ["1. XGBoost", "79.95% ± 2.34%", "80.22%", "84.74%", "82.40%", "88.26%", "78.00%", "79.82%", "87.24%"],
        ["2. Soft Voting Ensemble", "79.30% ± 2.33%", "79.64%", "84.01%", "81.77%", "88.25%", "78.00%", "80.00%", "87.39%"],
        ["3. Random Forest", "79.30% ± 2.25%", "79.05%", "85.20%", "82.00%", "87.97%", "78.50%", "80.72%", "87.13%"],
        ["4. Logistic Regression", "78.00% ± 2.41%", "79.22%", "81.36%", "80.27%", "87.25%", "77.75%", "79.73%", "87.00%"],
    ]
    add_styled_table(
        doc,
        "Bảng 3.3: Kết quả Thực nghiệm 4 Mô hình Học máy Tiêu biểu — Kịch bản 2 (Cảnh báo Sớm, Không dùng Burnout_Level)" if is_vi else "Table 3.3: 4-Model Experimental Results — Scenario 2 (Early Warning Without Burnout_Level)",
        ["Mô hình (Model)" if is_vi else "Classifier", "CV Accuracy", "CV Prec", "CV Recall", "CV F1", "CV AUC", "Test Acc", "Test F1", "Test AUC"],
        s2_rows,
    )

    add_styled_figure(
        doc,
        os.path.join(fig_dir, "07_model_performance_comparison.png"),
        "Hình 3.7: So sánh hiệu năng 10 mô hình học máy trên Kịch bản 1 (Full) và Kịch bản 2 (Early Warning)" if is_vi else "Figure 3.7: Performance Benchmarking Across 10 ML Models in Scenario 1 (Full) & Scenario 2 (Early Warning)",
        width_inches=6.0,
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "08_roc_and_precision_recall_curves.png"),
        "Hình 3.8: Đường cong ROC và Precision-Recall của 10 mô hình trên tập kiểm thử (Test Set)" if is_vi else "Figure 3.8: ROC and Precision-Recall Curves Across All 10 Classifiers on the Hold-Out Test Set",
        width_inches=6.0,
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "09_confusion_matrices_top_models.png"),
        "Hình 3.9: Ma trận nhầm lẫn (Confusion Matrix) của 4 mô hình tiêu biểu trên tập kiểm thử (n=160)" if is_vi else "Figure 3.9: Hold-Out Test Set Confusion Matrices for Top 4 Representative Models (n=160)",
        width_inches=6.0,
    )
    add_styled_figure(
        doc,
        os.path.join(fig_dir, "10_feature_importance_analysis.png"),
        "Hình 3.10: Phân tích mức độ quan trọng của các đặc trưng (Feature Importance) trên Random Forest & XGBoost" if is_vi else "Figure 3.10: Top 15 Feature Importances in Random Forest (Gini) and XGBoost (Gain)",
        width_inches=6.0,
    )

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # CHAPTER 4: USER INTERFACE & TESTING / IMPROVEMENTS
    # --------------------------------------------------------------------------
    add_h1(doc, "4. GIAO DIỆN NGƯỜI DÙNG VÀ KIỂM THỬ HỆ THỐNG" if is_vi else "4. USER INTERFACE AND SYSTEM TESTING")

    add_h2(doc, "4.1. Giao diện Web Dashboard Cảnh báo Sớm" if is_vi else "4.1. Interactive Early Warning Web Dashboard (FastAPI UI)")
    if is_vi:
        doc.add_paragraph(
            "Để chuyển hóa các mô hình học máy đã huấn luyện thành công cụ hỗ trợ ra quyết định trực quan cho cố vấn học tập và phòng công tác sinh viên, Nhóm 03 đã phát triển ứng dụng Web Dashboard bằng nền tảng FastAPI (tệp app.py). Hệ thống tích hợp đồng thời 3 mô hình đã đóng gói (.pkl) gồm: best_burnout_model.pkl (dự báo 3 mức độ kiệt sức), best_dropout_model_full.pkl (dự báo nguy cơ bỏ học kịch bản đầy đủ) và best_dropout_model_early_warning.pkl (dự báo cảnh báo sớm)."
        )
        add_bullet(doc, "Bảng điều khiển 18 tham số trực quan: ", "Cho phép kéo thanh trượt điều chỉnh tỷ lệ chuyên cần, giờ tự học, GPA, số môn nợ, mức căng thẳng, điểm lo âu, áp lực tài chính, hỗ trợ gia đình, giờ ngủ và giờ dùng màn hình.")
        add_bullet(doc, "Ba nút Hồ sơ Mẫu kiểm thử nhanh (One-Click Archetype Presets): ", "Gồm Mẫu Sinh viên Tích cực (Xác suất bỏ học 8.0% - An toàn), Mẫu Sinh viên Căng thẳng Trung bình, và Mẫu Sinh viên Kiệt sức Cao (Xác suất bỏ học 95.8% - Báo động đỏ).")
        add_bullet(doc, "Hệ thống Khuyến nghị Can thiệp Tự động: ", "Tự động phân tích ngưỡng vi phạm của từng chỉ số (ví dụ: Attendance < 70%, Stress >= 6.5, Financial Stress >= 6.5) để đề xuất hành động cụ thể cho cố vấn học tập.")
    else:
        doc.add_paragraph(
            "To operationalize our trained pipelines into an accessible decision-support system for academic advisors, Group 03 developed a full-stack FastAPI Web Dashboard (app.py). The backend loads three serialized pipelines: best_burnout_model.pkl (3-class Burnout predictor), best_dropout_model_full.pkl (Scenario 1 predictor), and best_dropout_model_early_warning.pkl (Scenario 2 predictor)."
        )
        add_bullet(doc, "18-Parameter Interactive Control Panel: ", "Provides real-time sliders and dropdowns for Attendance (%), Study Hours, GPA, Backlogs, Stress Level, Anxiety, Financial Stress, Family Support, Sleep, and Screen Time.")
        add_bullet(doc, "One-Click Student Archetype Presets: ", "Enables instant demonstration across Low-Risk Engaged Student (8.0% dropout probability), Moderate-Stress Student, and High-Burnout At-Risk Student (95.8% dropout probability).")
        add_bullet(doc, "Automated Prescriptive Recommendations: ", "Evaluates individual risk thresholds (e.g., Attendance < 70%, Stress >= 6.5, Financial Stress >= 6.5) to generate targeted tutoring, counseling, and financial aid referrals.")

    add_h2(doc, "4.2. Kiểm thử và Cải tiến Chất lượng Mô hình" if is_vi else "4.2. Model Testing and Iterative Improvements")
    imp_rows = (
        [
            ["Bước 1: Baseline (SimpleImputer Median + Raw Features)", "74.25%", "76.54%", "80.02%", "Chưa khai thác được tương tác phi tuyến giữa Stress và Engagement."],
            ["Bước 2: Nâng cấp KNNImputer (k=7) đa biến", "74.38%", "76.85%", "80.18%", "Bảo toàn cấu trúc tương quan trên 18 cột khuyết dữ liệu."],
            ["Bước 3: Bổ sung 6 Chỉ số DomainFeatureEngineer", "75.00% (CV)\n75.63% (Test)", "78.10% (CV)\n78.92% (Test)", "80.50% (CV)\n80.24% (Test)", "Tăng mạnh độ nhạy (Recall đạt 83.91%) giúp giảm thiểu bỏ sót SV nguy cơ cao."],
        ]
        if is_vi
        else [
            ["Stage 1: Baseline (Median Imputation + Raw Features)", "74.25%", "76.54%", "80.02%", "Did not explicitly capture non-linear Stress–Engagement ratios."],
            ["Stage 2: Multivariate KNNImputer (k=7) Upgrade", "74.38%", "76.85%", "80.18%", "Preserved inter-feature correlations across the 18 incomplete columns."],
            ["Stage 3: Adding 6 DomainFeatureEngineer Indices", "75.00% (CV)\n75.63% (Test)", "78.10% (CV)\n78.92% (Test)", "80.50% (CV)\n80.24% (Test)", "Boosted Test Recall to 83.91%, minimizing costly false negatives."],
        ]
    )
    add_styled_table(
        doc,
        "Bảng 4.1: Quá trình Cải tiến Hiệu năng Mô hình qua các Giai đoạn Thử nghiệm" if is_vi else "Table 4.1: Iterative Performance Progression Across Pipeline Refinements",
        ["Giai đoạn Cải tiến" if is_vi else "Refinement Stage", "Best Accuracy", "Best F1-Score", "Best ROC-AUC", "Nhận xét Kỹ thuật" if is_vi else "Technical Observation"],
        imp_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # --------------------------------------------------------------------------
    # CHAPTER 5: PROJECTED IMPACT & TEAM REVIEWS
    # --------------------------------------------------------------------------
    add_h1(doc, "5. ĐÁNH GIÁ TÁC ĐỘNG VÀ ĐỊNH HƯỚNG PHÁT TRIỂN" if is_vi else "5. PROJECTED IMPACT, FUTURE ROADMAP & TEAM REVIEWS")

    add_h2(doc, "5.1. Thành tựu Đạt được và Lợi ích Thực tiễn" if is_vi else "5.1. Key Accomplishments and Institutional Benefits")
    if is_vi:
        add_bullet(doc, "Giá trị Khoa học & Kỹ thuật: ", "Xây dựng bộ chỉ số định lượng chứng minh Tỷ số Căng thẳng / Gắn kết (Stress_Engagement_Ratio) và Chỉ số Gắn kết Học tập (Academic_Engagement_Index) là những biến dự báo quan trọng nhất chỉ đứng sau Burnout_Level.")
        add_bullet(doc, "Giá trị Giáo dục & Xã hội: ", "Giúp các trường đại học chủ động sàng lọc hàng nghìn hồ sơ sinh viên vào đầu mỗi học kỳ, phân bổ hiệu quả nguồn lực tư vấn tâm lý và quỹ học bổng hỗ trợ cho đúng nhóm sinh viên có nguy cơ bỏ học > 65%.")
    else:
        add_bullet(doc, "Scientific & Technical Contribution: ", "Demonstrated empirically via Gini and Gain Feature Importance that Stress_Engagement_Ratio and Academic_Engagement_Index rank as the strongest predictors immediately behind Burnout_Level.")
        add_bullet(doc, "Educational & Social Impact: ", "Enables universities to screen thousands of student profiles automatically each semester and direct counseling and financial aid toward high-risk students (>65% dropout probability).")

    add_h2(doc, "5.2. Định hướng Phát triển Tương lai" if is_vi else "5.2. Future Improvements and Research Roadmap")
    if is_vi:
        add_bullet(doc, "1. Mô hình Chuỗi thời gian (Longitudinal Time-Series Modeling): ", "Thu thập dữ liệu chuyên cần và điểm số theo từng tuần trong học kỳ để huấn luyện mạng LSTM / Transformer dự báo sớm sự suy giảm phong độ.")
        add_bullet(doc, "2. Tích hợp Giải thích Mô hình Cục bộ SHAP (Explainable AI): ", "Hiển thị biểu đồ SHAP Waterfall trực tiếp trên giao diện Web cho từng cá nhân sinh viên.")
        add_bullet(doc, "3. Tích hợp Hệ thống LMS (Moodle / Canvas API): ", "Tự động đồng bộ dữ liệu điểm danh và tần suất đăng nhập hệ thống học tập trực tuyến.")
    else:
        add_bullet(doc, "1. Longitudinal Time-Series Tracking: ", "Collect weekly attendance and stress logs across semesters to train sequential deep learning models (LSTM / Transformers).")
        add_bullet(doc, "2. Local Explainable AI (SHAP Integration): ", "Render individual SHAP waterfall plots on the Web Dashboard to explain per-student feature contributions.")
        add_bullet(doc, "3. Institutional LMS Integration: ", "Connect directly with university Learning Management Systems (Moodle / Canvas) via REST APIs.")

    add_h2(doc, "5.3. Nhận xét và Đánh giá của Thành viên Nhóm" if is_vi else "5.3. Team Member Reflections and Reviews")
    review_rows = (
        [
            [
                "Đỗ Quang Minh\n(Trưởng nhóm / Team Leader)",
                "Dự án Capstone của khóa học Samsung Innovation Campus (SIC AI) đã giúp tôi hệ thống hóa toàn bộ kiến thức về Học máy từ lý thuyết đến triển khai thực tế. Tôi đặc biệt tâm đắc với việc thiết kế lớp biến đổi DomainFeatureEngineer bên trong Pipeline giúp ngăn chặn rò rỉ dữ liệu và nâng cao rõ rệt điểm ROC-AUC của nhóm mô hình Ensemble lên 80.50%.",
            ],
            [
                "Vũ Quang Huy\n(Thành viên / Team Member)",
                "Thông qua quá trình thực hiện đề tài dự báo nguy cơ bỏ học, tôi đã rèn luyện sâu kỹ năng Phân tích khám phá dữ liệu (EDA), xử lý dữ liệu khuyết đa biến bằng KNNImputer và trực quan hóa dữ liệu khoa học. Sự phối hợp ăn ý giữa hai thành viên từ khâu thực nghiệm 10 mô hình đến xây dựng giao diện Web FastAPI đã mang lại một sản phẩm hoàn chỉnh và giàu tính nhân văn.",
            ],
        ]
        if is_vi
        else [
            [
                "Do Quang Minh\n(Team Leader)",
                "The Samsung Innovation Campus (SIC) AI Capstone project allowed me to synthesize end-to-end Machine Learning engineering practices. Designing the leakage-free DomainFeatureEngineer transformer inside our scikit-learn Pipeline significantly improved our ensemble models (reaching 80.50% ROC-AUC) and demonstrated how domain knowledge elevates predictive modeling.",
            ],
            [
                "Vu Quang Huy\n(Team Member)",
                "Working on the Student Burnout & Dropout prediction problem sharpened my skills in statistical Exploratory Data Analysis (EDA), multivariate KNN imputation, and scientific data visualization. Seamless teamwork across 10-model benchmarking and FastAPI Web UI development enabled us to deliver a rigorous, impactful educational AI solution.",
            ],
        ]
    )
    add_styled_table(
        doc,
        "Bảng 5.1: Nhận xét Cá nhân của Thành viên Nhóm 03" if is_vi else "Table 5.1: Individual Reflections from Group 03 Members",
        ["Họ và Tên Thành viên" if is_vi else "Team Member", "Nội dung Nhận xét & Đánh giá" if is_vi else "Review and Reflection"],
        review_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # --------------------------------------------------------------------------
    # CHAPTER 6: REFERENCES & TECHNICAL APPENDIX
    # --------------------------------------------------------------------------
    add_h1(doc, "6. TÀI LIỆU THAM KHẢO VÀ PHỤ LỤC KỸ THUẬT" if is_vi else "6. REFERENCES AND TECHNICAL APPENDIX")

    add_h2(doc, "6.1. Danh mục Tài liệu Tham khảo" if is_vi else "6.1. Academic and Technical References")
    refs = [
        "[1] V. Tinto, \"Dropout from Higher Education: A Theoretical Synthesis of Recent Research,\" Review of Educational Research, vol. 45, no. 1, pp. 89–125, 1975.",
        "[2] C. Maslach, W. B. Schaufeli, and M. P. Leiter, \"Job Burnout,\" Annual Review of Psychology, vol. 52, pp. 397–422, 2001.",
        "[3] L. Breiman, \"Random Forests,\" Machine Learning, vol. 45, no. 1, pp. 5–32, 2001.",
        "[4] P. Geurts, D. Ernst, and L. Wehenkel, \"Extremely Randomized Trees,\" Machine Learning, vol. 63, no. 1, pp. 3–42, 2006.",
        "[5] T. Chen and C. Guestrin, \"XGBoost: A Scalable Tree Boosting System,\" in Proc. 22nd ACM SIGKDD Int. Conf. Knowledge Discovery and Data Mining, 2016, pp. 785–794.",
        "[6] F. Pedregosa et al., \"Scikit-learn: Machine Learning in Python,\" Journal of Machine Learning Research, vol. 12, pp. 2825–2830, 2011.",
    ]
    for ref in refs:
        p_r = doc.add_paragraph(ref)
        p_r.paragraph_format.left_indent = Cm(0.8)
        p_r.paragraph_format.first_line_indent = Cm(-0.8)
        p_r.paragraph_format.space_after = Pt(4)

    add_h2(doc, "6.2. Phụ lục Kỹ thuật: Cấu hình Siêu tham số 10 Mô hình" if is_vi else "6.2. Technical Appendix: Hyperparameter Configurations of the 10 Models")
    param_rows = [
        ["1. Logistic Regression", "C=0.5, penalty='l2', solver='lbfgs', max_iter=1000, random_state=42"],
        ["2. SVM (RBF Kernel)", "CalibratedClassifierCV(SVC(C=1.2, kernel='rbf', gamma='scale'), cv=3)"],
        ["3. K-Nearest Neighbors", "n_neighbors=17, weights='distance', metric='minkowski', p=2"],
        ["4. Decision Tree", "max_depth=5, min_samples_leaf=15, criterion='gini', random_state=42"],
        ["5. Random Forest", "n_estimators=300, max_depth=8, min_samples_split=8, min_samples_leaf=4, max_features='sqrt'"],
        ["6. Extra Trees", "n_estimators=300, max_depth=9, min_samples_split=8, min_samples_leaf=4, random_state=42"],
        ["7. Gradient Boosting", "n_estimators=150, learning_rate=0.04, max_depth=3, subsample=0.85, min_samples_leaf=10"],
        ["8. XGBoost", "n_estimators=200, learning_rate=0.035, max_depth=3, subsample=0.85, colsample_bytree=0.8, reg_lambda=2.0"],
        ["9. MLP Neural Network", "hidden_layer_sizes=(64, 32), activation='relu', alpha=0.5, learning_rate_init=0.003, early_stopping=True"],
        ["10. Soft Voting Ensemble", "VotingClassifier(estimators=[lr, rf, et, gb, xgb, svm], voting='soft', weights=[2,3,2,2,2,2])"],
    ]
    add_styled_table(
        doc,
        "Bảng 6.1: Cấu hình Siêu tham số Tối ưu của 10 Mô hình Học máy" if is_vi else "Table 6.1: Optimal Hyperparameter Configurations Across All 10 Benchmarked Classifiers",
        ["Mô hình Học máy" if is_vi else "Classifier", "Cấu hình Siêu tham số (Scikit-Learn / XGBoost)" if is_vi else "Hyperparameter Configuration"],
        param_rows,
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # Save outputs to all target locations
    base_name = "SIC_AI_Capstone Project_Final Report.docx"
    suffixed = f"SIC_AI_Capstone Project_Final Report_{lang.upper()}.docx"
    targets = [
        os.path.join(OUT_VI_WS if is_vi else OUT_EN_WS, base_name),
        os.path.join(OUT_VI_DRV if is_vi else OUT_EN_DRV, base_name),
        os.path.join(SRC_FOLDER, suffixed),
    ]
    if is_vi:
        targets.append(os.path.join(SRC_FOLDER, base_name))

    for t in targets:
        doc.save(t)
    print(f"Saved Master Academic Final Report ({lang.upper()}) -> {len(targets)} locations")


# ==============================================================================
# MASTER ACTION PLAN BUILDER (ENHANCED STYLING FOR BOTH VI & EN)
# ==============================================================================
def style_existing_action_plan(lang="vi"):
    is_vi = lang == "vi"
    base_name = "SIC_AI_Capstone Project_Action Plan.docx"
    suffixed = f"SIC_AI_Capstone Project_Action Plan_{lang.upper()}.docx"
    src_path = os.path.join(OUT_VI_WS if is_vi else OUT_EN_WS, base_name)
    if not os.path.exists(src_path):
        return

    doc = Document(src_path)
    for tbl in doc.tables:
        for r_idx, row in enumerate(tbl.rows):
            prevent_row_split(row)
            for c_idx, cell in enumerate(row.cells):
                set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.25
                    p.paragraph_format.space_after = Pt(4)
                    for run in p.runs:
                        run.font.name = "Times New Roman"
                        run.font.size = Pt(11)

    targets = [
        os.path.join(OUT_VI_WS if is_vi else OUT_EN_WS, base_name),
        os.path.join(OUT_VI_DRV if is_vi else OUT_EN_DRV, base_name),
        os.path.join(SRC_FOLDER, suffixed),
    ]
    if is_vi:
        targets.append(os.path.join(SRC_FOLDER, base_name))
    for t in targets:
        doc.save(t)
    print(f"Enhanced Action Plan formatting ({lang.upper()})")


if __name__ == "__main__":
    for language in ["vi", "en"]:
        generate_master_final_report(language)
        style_existing_action_plan(language)
    print("ALL MASTER WORD REPORTS GENERATED SUCCESSFULLY!")
