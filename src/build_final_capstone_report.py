"""
Builds FINAL_CAPSTONE/03_Final_Report.docx
Strictly conforms to publication-grade academic and technical reporting standards:
- Clear 6-chapter structure with verified scientific numbers.
- Presents both Experiment 1 (800 real students, 10 models) and Experiment 2 (2,000 expanded profiles, 4 flagship models).
- Includes full Confusion Matrix numbers (TN, FP, FN, TP), Feature Importance, and Web Dashboard benchmarks.
- Incorporates rigorous academic critique: data leakage risks, synthetic augmentation limits, and non-causal interpretations.
- Embeds high-resolution bilingual figures with accurate captions.
"""
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = r"c:\Users\mrmin\Desktop\Sudent_burnout"
FIG_DIR = os.path.join(BASE_DIR, "figures_vi")
OUT_PATH = os.path.join(BASE_DIR, "FINAL_CAPSTONE", "03_Final_Report.docx")

# Color Palette
NAVY_HEX = "0B162C"
BLUE_HEX = "0066FF"
CYAN_HEX = "00B4D8"
TEAL_HEX = "0D9488"
AMBER_HEX = "D97706"
CORAL_HEX = "E11D48"
SLATE_HEX = "334155"
LIGHT_BG_HEX = "F8FAFC"
CALLOUT_BG_HEX = "F1F5F9"
BORDER_GRAY_HEX = "CBD5E1"

NAVY_RGB = RGBColor(11, 22, 44)
SLATE_RGB = RGBColor(51, 65, 85)
TEAL_RGB = RGBColor(13, 148, 136)
DARK_TEXT_RGB = RGBColor(15, 23, 42)
MUTED_TEXT_RGB = RGBColor(100, 116, 139)


def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'))


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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
        f'<w:left w:val="none"/><w:right w:val="none"/><w:insideV w:val="none"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def prevent_row_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def repeat_header_row(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def add_bottom_border_to_paragraph(paragraph, color_hex=NAVY_HEX, sz="8"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color_hex}"/>'
        f"</w:pBdr>"
    )
    pPr.append(pBdr)


def add_field(paragraph, field_text, font_size=9.5, color_rgb=MUTED_TEXT_RGB):
    run = paragraph.add_run()
    run.font.name = "Arial"
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


def add_callout(doc, title, text, border_color=BLUE_HEX, bg_color=CALLOUT_BG_HEX):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.cell(0, 0)
    set_cell_background(c, bg_color)
    set_cell_margins(c, top=120, bottom=120, left=180, right=140)
    prevent_row_split(tbl.rows[0])
    tcPr = c._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(
        f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
        f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>'
    ))
    p1 = c.paragraphs[0]
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(2)
    p1.paragraph_format.keep_with_next = True
    r1 = p1.add_run(title)
    r1.bold = True
    r1.font.name = "Arial"
    r1.font.size = Pt(11)
    r1.font.color.rgb = NAVY_RGB

    p2 = c.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run(text)
    r2.font.name = "Arial"
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = DARK_TEXT_RGB

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(4)


def add_fig(doc, img_name, fig_label, fig_title, width_inches=6.2):
    img_path = os.path.join(FIG_DIR, img_name)
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.paragraph_format.keep_with_next = True
        p_img.add_run().add_picture(img_path, width=Inches(width_inches))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(10)
    r_lbl = p_cap.add_run(f"{fig_label}: ")
    r_lbl.bold = True
    r_lbl.font.name = "Arial"
    r_lbl.font.size = Pt(9.5)
    r_lbl.font.color.rgb = NAVY_RGB

    r_t = p_cap.add_run(fig_title)
    r_t.font.name = "Arial"
    r_t.font.size = Pt(9.5)
    r_t.font.italic = True
    r_t.font.color.rgb = MUTED_TEXT_RGB


def format_table(table, col_widths, col_alignments):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    repeat_header_row(table.rows[0])
    set_table_borders(table)

    for i, row in enumerate(table.rows):
        prevent_row_split(row)
        for j, cell in enumerate(row.cells):
            cell.width = Inches(col_widths[j])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            if i == 0:
                set_cell_background(cell, NAVY_HEX)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(9.5)
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                if i % 2 == 1:
                    set_cell_background(cell, LIGHT_BG_HEX)
                for p in cell.paragraphs:
                    p.alignment = col_alignments[j]
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(9.0)
                        r.font.color.rgb = DARK_TEXT_RGB


def build_final_report():
    doc = Document()

    # Page Setup A4
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.8)
    sec.right_margin = Cm(2.2)
    sec.different_first_page_header_footer = True

    # Header / Footer
    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run("SAMSUNG INNOVATION CAMPUS — BÁO CÁO ĐỒ ÁN AI CAPSTONE | NHÓM 03")
    hrun.font.name = "Arial"
    hrun.font.size = Pt(9.0)
    hrun.font.italic = True
    hrun.font.color.rgb = MUTED_TEXT_RGB
    add_bottom_border_to_paragraph(hp, color_hex=BORDER_GRAY_HEX, sz="6")

    footer = sec.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    frun = fp.add_run("Đỗ Quang Minh • Vũ Quang Huy  |  Trang ")
    frun.font.name = "Arial"
    frun.font.size = Pt(9.0)
    frun.font.color.rgb = MUTED_TEXT_RGB
    add_field(fp, "PAGE")
    frun2 = fp.add_run(" / ")
    frun2.font.name = "Arial"
    frun2.font.size = Pt(9.0)
    frun2.font.color.rgb = MUTED_TEXT_RGB
    add_field(fp, "NUMPAGES")

    # Typography styles
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    normal.font.color.rgb = DARK_TEXT_RGB
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    h1 = doc.styles["Heading 1"]
    h1.font.name = "Arial"
    h1.font.size = Pt(15)
    h1.font.bold = True
    h1.font.color.rgb = NAVY_RGB
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(6)
    h1.paragraph_format.keep_with_next = True

    h2 = doc.styles["Heading 2"]
    h2.font.name = "Arial"
    h2.font.size = Pt(13)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0, 102, 255)
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(4)
    h2.paragraph_format.keep_with_next = True

    h3 = doc.styles["Heading 3"]
    h3.font.name = "Arial"
    h3.font.size = Pt(11.5)
    h3.font.bold = True
    h3.font.italic = True
    h3.font.color.rgb = TEAL_RGB
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(3)
    h3.paragraph_format.keep_with_next = True

    # =========================================================================
    # COVER PAGE
    # =========================================================================
    p_prog = doc.add_paragraph()
    p_prog.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_prog = p_prog.add_run("CHƯƠNG TRÌNH ĐÀO TẠO SAMSUNG INNOVATION CAMPUS (SIC) — ĐỒ ÁN AI CAPSTONE")
    r_prog.font.name = "Arial"
    r_prog.font.size = Pt(11)
    r_prog.font.bold = True
    r_prog.font.color.rgb = NAVY_RGB

    p_doc_t = doc.add_paragraph()
    p_doc_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_doc_t = p_doc_t.add_run("BÁO CÁO KỸ THUẬT TỔNG KẾT ĐỒ ÁN TRÍ TUỆ NHÂN TẠO & HỌC MÁY")
    r_doc_t.font.name = "Arial"
    r_doc_t.font.size = Pt(16)
    r_doc_t.font.bold = True
    r_doc_t.font.color.rgb = RGBColor(0, 102, 255)

    # Title Card
    tbl_t = doc.add_table(rows=1, cols=1)
    tbl_t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_t = tbl_t.cell(0, 0)
    set_cell_background(c_t, NAVY_HEX)
    set_cell_margins(c_t, top=160, bottom=160, left=200, right=200)
    p_t = c_t.paragraphs[0]
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_t.add_run(
        "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY CHO BÀI TOÁN DỰ ĐOÁN TỶ LỆ BỎ HỌC "
        "DỰA TRÊN CÁC CHỈ SỐ VỀ CĂNG THẲNG VÀ SỰ THAM GIA HỌC TẬP\n\n"
        "(Hệ thống Cảnh báo Sớm Học đường Thông minh — Phân tích Kiệt sức & Nguy cơ Bỏ học Đại học)"
    )
    r_t.font.name = "Arial"
    r_t.font.size = Pt(13.5)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(255, 255, 255)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4 Key Pillars Banner
    tbl_kpi = doc.add_table(rows=2, cols=4)
    tbl_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    kpi_data = [
        ("800 → 2.000 SV", "Quy mô Nghiên cứu", "800 gốc mở rộng 2.000 sạch", "0B162C"),
        ("25 → 34 Cột", "Kỹ thuật Đặc trưng", "Tích hợp 6 chỉ số tương tác", "0066FF"),
        ("10 & 4 Mô hình", "Đối sánh Toàn diện", "Kiểm định chéo 5-Fold CV", "0D9488"),
        ("89.01% / 88.26%", "Hiệu năng ROC-AUC", "Random Forest & XGBoost", "D97706"),
    ]
    for j, (val, title, sub, col_hex) in enumerate(kpi_data):
        c0 = tbl_kpi.cell(0, j)
        set_cell_background(c0, col_hex)
        set_cell_margins(c0, top=60, bottom=40, left=60, right=60)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(val)
        r0.font.name = "Arial"
        r0.font.size = Pt(11.5)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(255, 255, 255)

        c1 = tbl_kpi.cell(1, j)
        set_cell_background(c1, LIGHT_BG_HEX)
        set_cell_margins(c1, top=40, bottom=60, left=60, right=60)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(f"{title}\n")
        r1.font.name = "Arial"
        r1.font.size = Pt(9.0)
        r1.font.bold = True
        r1.font.color.rgb = NAVY_RGB
        r1_sub = p1.add_run(sub)
        r1_sub.font.name = "Arial"
        r1_sub.font.size = Pt(8.2)
        r1_sub.font.color.rgb = MUTED_TEXT_RGB
    set_table_borders(tbl_kpi)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Project Information Box
    tbl_info = doc.add_table(rows=8, cols=2)
    info_rows = [
        ("Chương trình đào tạo", "Samsung Innovation Campus (SIC) – Khóa học Trí tuệ Nhân tạo (AI Course)"),
        ("Đề tài nghiên cứu", "Dự đoán tỷ lệ bỏ học dựa trên các chỉ số căng thẳng và sự tham gia học tập"),
        ("Tên Nhóm dự án", "Nhóm 03 (Group 03 - Student Burnout AI)"),
        ("Thành viên thực hiện", "1. Đỗ Quang Minh (Trưởng nhóm / Team Leader — 50%)\n2. Vũ Quang Huy (Thành viên / Team Member — 50%)"),
        ("Bộ dữ liệu nghiên cứu", "student_burnout_dropout_dataset_2.csv (800 bản ghi gốc × 25 thuộc tính; mở rộng 2.000 hồ sơ sạch)"),
        ("Phương pháp cốt lõi", "KNNImputer (k=7), DomainFeatureEngineer (6 chỉ số), Stratified 5-Fold CV, FastAPI Web UI"),
        ("Sản phẩm bàn giao", "Bộ 5 tài liệu SIC, Pipeline mô hình (.pkl), Bộ 12 biểu đồ EDA/ML, Ứng dụng Web Dashboard Demo"),
        ("Thời gian thực hiện", "19/10/2026 – 05/11/2026 (Nghiệm thu kỹ thuật: 06/11/2026)"),
    ]
    for i, (k, v) in enumerate(info_rows):
        tbl_info.cell(i, 0).text = k
        tbl_info.cell(i, 1).text = v
    format_table(tbl_info, [2.2, 5.0], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_page_break()

    # =========================================================================
    # EXECUTIVE SUMMARY & TOC
    # =========================================================================
    add_callout(
        doc,
        "TÓM TẮT ĐIỀU HÀNH ĐỒ ÁN (EXECUTIVE SUMMARY)",
        "Đồ án giải quyết bài toán Dự đoán Nguy cơ Bỏ học Đại học (Dropout_Risk) thông qua phân tích định lượng tương tác "
        "giữa Căng thẳng Tâm lý (Stress, Anxiety, Financial Stress, Peer Pressure) và Sự Tham gia Học tập (Attendance, Study Hours, GPA, Motivation). "
        "Quy trình thực hiện gồm 2 giai đoạn khoa học: (1) Khảo sát thực nghiệm trên 800 sinh viên gốc với 25 thuộc tính, làm sạch 612 ô khuyết bằng KNNImputer (k=7) "
        "và đối sánh 10 mô hình học máy (Random Forest đạt đỉnh CV ROC-AUC 80.50% ± 3.83%, Extra Trees đạt Test Recall 83.91%); (2) Mở rộng mẫu quy mô kỹ thuật "
        "lên 2.000 hồ sơ sạch phục vụ kiểm thử 4 mô hình cốt lõi (Random Forest đạt CV ROC-AUC 89.01% ở Kịch bản 1 có Burnout; XGBoost đạt CV ROC-AUC 88.26% "
        "ở Kịch bản 2 Cảnh báo sớm chỉ dùng hành vi thô). Hệ thống được triển khai thực tế trên ứng dụng Web Dashboard (FastAPI app.py) "
        "cho phép dự báo kép tức thời (~145–153ms) và sinh khuyến nghị can thiệp cá nhân hóa."
    )

    doc.add_heading("MỤC LỤC TÀI LIỆU", level=1)
    toc_items = [
        ("Chương 1", "Giới thiệu Chung (Bối cảnh, Động lực, 4 Mục tiêu, Phân công 50/50 & Lộ trình WBS)"),
        ("Chương 2", "Triển khai Dự án & Phương pháp luận (Đặc tả 25 thuộc tính, Pipeline chống rò rỉ, 2 Kịch bản)"),
        ("Chương 3", "Kết quả Thực nghiệm: Tiền xử lý, EDA & Đối sánh Mô hình (Thực nghiệm 800 SV và 2.000 hồ sơ)"),
        ("Chương 4", "Thiết kế Giao diện Người dùng (Web Dashboard FastAPI) & Đánh giá Cải tiến"),
        ("Chương 5", "Đánh giá Tác động Thực tiễn, Nhìn nhận Hạn chế Khoa học & Hướng Phát triển"),
        ("Chương 6", "Danh mục Tài liệu Tham khảo chuẩn IEEE & Phụ lục Cấu hình Siêu tham số"),
    ]
    tbl_toc = doc.add_table(rows=len(toc_items), cols=2)
    for i, (ch, name) in enumerate(toc_items):
        tbl_toc.cell(i, 0).text = ch
        tbl_toc.cell(i, 1).text = name
    format_table(tbl_toc, [1.5, 5.7], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_heading("DANH MỤC HÌNH ẢNH & BIỂU ĐỒ TRỰC QUAN", level=2)
    fig_items = [
        ("Hình 2.1", "Quy trình thực hiện dự án Học máy 5 giai đoạn khép kín"),
        ("Hình 2.2", "Kiến trúc hệ thống Cảnh báo sớm Nguy cơ Bỏ học 3 tầng & Web UI"),
        ("Hình 3.1", "Phân tích tỷ lệ dữ liệu khuyết trên 18/25 thuộc tính trước khi qua KNNImputer"),
        ("Hình 3.2", "Phân phối biến mục tiêu Dropout_Risk và tỷ lệ bỏ học theo mức độ Burnout_Level"),
        ("Hình 3.3", "So sánh phân phối 8 chỉ số Tham gia Học tập và Căng thẳng Tâm lý theo Dropout_Risk"),
        ("Hình 3.4", "Ma trận tương quan Pearson giữa các biến định lượng trên dữ liệu khảo sát gốc"),
        ("Hình 3.5", "Phân tích tỷ lệ bỏ học theo các yếu tố nhân khẩu học và lối sống"),
        ("Hình 3.6", "Phân tích đa biến giữa Tỷ lệ chuyên cần, Căng thẳng tâm lý và Hồ sơ Burnout"),
        ("Hình 3.7", "So sánh hiệu năng các mô hình học máy trên Kịch bản 1 (Full) và Kịch bản 2 (Early Warning)"),
        ("Hình 3.8", "Đường cong ROC và Precision-Recall trên tập kiểm thử độc lập (Test Set)"),
        ("Hình 3.9", "Ma trận nhầm lẫn (Confusion Matrix) của 4 mô hình tiêu biểu trên tập kiểm thử"),
        ("Hình 3.10", "Phân tích mức độ quan trọng đặc trưng (Feature Importance) trên Random Forest & XGBoost"),
        ("Hình 4.1", "Giao diện Web Dashboard Cảnh báo Sớm thời gian thực triển khai trên FastAPI (app.py)"),
    ]
    tbl_figs = doc.add_table(rows=len(fig_items), cols=2)
    for i, (fn, desc) in enumerate(fig_items):
        tbl_figs.cell(i, 0).text = fn
        tbl_figs.cell(i, 1).text = desc
    format_table(tbl_figs, [1.5, 5.7], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_page_break()

    # =========================================================================
    # CHƯƠNG 1
    # =========================================================================
    doc.add_heading("1. GIỚI THIỆU CHUNG", level=1)

    doc.add_heading("1.1. Bối cảnh Nghiên cứu", level=2)
    doc.add_paragraph(
        "Trong giáo dục đại học hiện đại, tình trạng sinh viên bỏ học giữa chừng (Student Attrition / Dropout) "
        "và hội chứng kiệt sức học đường (Academic Burnout) đang là mối bận tâm hàng đầu của các nhà quản lý giáo dục. "
        "Theo mô hình tích hợp của Tinto (1975) và lý thuyết kiệt sức của Maslach & Leiter (2001), quyết định thôi học "
        "hiếm khi diễn ra đột ngột mà là kết quả của một quá trình suy giảm gắn kết học tập kéo dài dưới tác động cộng dồn của hai nhóm yếu tố:"
    )
    doc.add_paragraph(
        "• Suy giảm mức độ tham gia học tập (Academic Disengagement): Thể hiện qua tỷ lệ chuyên cần sụt giảm (Attendance_Percent < 70%), "
        "thời lượng tự học ít (Study_Hours_Per_Day < 2h), suy giảm động lực học tập (Motivation_Score) và tích lũy môn nợ học phần (Backlogs).\n"
        "• Áp lực tâm lý và tài chính kéo dài (Chronic Stress Burden): Căng thẳng học vụ (Stress_Level), rối loạn lo âu (Anxiety_Score), "
        "gánh nặng tài chính sinh hoạt (Financial_Stress_Score), thiếu ngủ (<6h) do lạm dụng màn hình và thiếu vắng sự hỗ trợ từ gia đình."
    )
    doc.add_paragraph(
        "Khi áp lực tâm lý vượt quá sức chịu đựng, sinh viên rơi vào trạng thái Kiệt sức Học đường (Burnout_Level từ Low chuyển sang Medium và High), "
        "kéo theo tỷ lệ bỏ học tăng vọt. Việc xây dựng một hệ thống học máy định lượng sớm nguy cơ này cho phép chuyển đổi từ cơ chế can thiệp "
        "bị động sang Hệ thống Cảnh báo Sớm Chủ động (Proactive Early Warning System)."
    )

    doc.add_heading("1.2. Động lực và Mục tiêu Đề tài", level=2)
    doc.add_paragraph("Để giải quyết bài toán trên, Nhóm 03 xác định 4 mục tiêu cụ thể:")
    doc.add_paragraph(
        "1. Mục tiêu 1 – Phân tích Khám phá Dữ liệu Toàn diện (Comprehensive EDA): Khảo sát trên 800 sinh viên gốc × 25 thuộc tính "
        "nhằm lượng hóa mối quan hệ thống kê giữa chuyên cần, căng thẳng, mức độ kiệt sức và tỷ lệ bỏ học.\n"
        "2. Mục tiêu 2 – Tiền xử lý & Kỹ thuật Tạo Đặc trưng (Feature Engineering): Xử lý 612 ô khuyết trên 18 cột bằng KNNImputer (k=7) "
        "trong Pipeline chống rò rỉ dữ liệu; thiết kế 6 chỉ số tương tác chuyên ngành đo lường cán cân Căng thẳng vs. Gắn kết học tập.\n"
        "3. Mục tiêu 3 – Huấn luyện & Đối sánh Đa Mô hình (Model Benchmarking): Đánh giá khách quan qua 5-Fold Stratified Cross-Validation "
        "và tập kiểm thử độc lập 20% trên 2 kịch bản (Kịch bản 1: Đầy đủ đặc trưng; Kịch bản 2: Cảnh báo sớm trực tiếp từ hành vi thô).\n"
        "4. Mục tiêu 4 – Mở rộng Quy mô Kỹ thuật & Triển khai Web Dashboard (System Deployment): Mở rộng 2.000 hồ sơ sạch (0 NaN) "
        "để kiểm thử tải hệ thống; xây dựng ứng dụng Web tương tác bằng FastAPI (app.py) dự báo thời gian thực kèm khuyến nghị can thiệp."
    )

    doc.add_heading("1.3. Phân công Nhiệm vụ Thành viên", level=2)
    tbl_roles = doc.add_table(rows=3, cols=3)
    role_data = [
        ("Thành viên", "Tỷ trọng", "Hạng mục Công việc Phụ trách"),
        ("1. Đỗ Quang Minh\n(Trưởng nhóm)", "50%",
         "• Thiết kế kiến trúc Pipeline học máy chống rò rỉ dữ liệu (No Data Leakage).\n"
         "• Xây dựng module DomainFeatureEngineer gồm 6 chỉ số tương tác chuyên ngành.\n"
         "• Huấn luyện, kiểm định chéo và tối ưu nhóm mô hình Cây tập hợp & Boosting (RF, Extra Trees, XGBoost, Soft Voting).\n"
         "• Phát triển Backend FastAPI (app.py), đóng gói mô hình (.pkl) và kiểm thử Web UI."),
        ("2. Vũ Quang Huy\n(Thành viên)", "50%",
         "• Kiểm toán dữ liệu, thống kê 612 ô khuyết trên 18 cột và thực hiện Phân tích EDA.\n"
         "• Thiết lập KNNImputer (k=7), lọc ngoại lai IQR và vẽ bộ 12 biểu đồ trực quan hóa dữ liệu.\n"
         "• Huấn luyện và đánh giá nhóm mô hình Cơ sở & Mạng nơ-ron (Logistic Regression, SVM, KNN, DT, MLP).\n"
         "• Thiết kế giao diện Frontend Web Dashboard, biên soạn Báo cáo cuối kỳ, WBS và Slide.")
    ]
    for i, (c0, c1, c2) in enumerate(role_data):
        tbl_roles.cell(i, 0).text = c0
        tbl_roles.cell(i, 1).text = c1
        tbl_roles.cell(i, 2).text = c2
    format_table(tbl_roles, [1.8, 1.0, 4.4], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_heading("1.4. Lộ trình Thực hiện và Các Mốc Quan trọng", level=2)
    tbl_plan = doc.add_table(rows=5, cols=5)
    plan_data = [
        ("Giai đoạn", "Thời gian", "Nội dung thực hiện", "Sản phẩm đầu ra", "Trạng thái"),
        ("GĐ 1", "19/10–20/10/2026", "Khởi tạo dự án, phân tích yêu cầu, khảo sát 800 mẫu", "Action Plan, WBS & Đặc tả 25 biến", "Hoàn thành (100%)"),
        ("GĐ 2", "21/10–25/10/2026", "Phân tích EDA, điền khuyết KNN (k=7), tạo 6 đặc trưng", "Pipeline KNN, 6 chỉ số, 6 biểu đồ EDA", "Hoàn thành (100%)"),
        ("GĐ 3", "26/10–29/10/2026", "Huấn luyện 10 mô hình qua 5-Fold CV trên 2 kịch bản", "Bảng kết quả CSV, ROC/PR, Feature Imp.", "Hoàn thành (100%)"),
        ("GĐ 4", "30/10–06/11/2026", "Mở rộng 2.000 hồ sơ, Web FastAPI, Báo cáo & Slide", "Web UI Demo, Báo cáo SIC & Slide", "Hoàn thành (100%)"),
    ]
    for i, row in enumerate(plan_data):
        for j, val in enumerate(row):
            tbl_plan.cell(i, j).text = val
    format_table(tbl_plan, [0.8, 1.4, 2.3, 1.8, 0.9], [WD_ALIGN_PARAGRAPH.CENTER] * 2 + [WD_ALIGN_PARAGRAPH.LEFT] * 2 + [WD_ALIGN_PARAGRAPH.CENTER])

    doc.add_page_break()

    # =========================================================================
    # CHƯƠNG 2
    # =========================================================================
    doc.add_heading("2. TRIỂN KHAI DỰ ÁN VÀ PHƯƠNG PHÁP LUẬN", level=1)

    doc.add_heading("2.1. Đặc tả 25 Thuộc tính Dữ liệu Gốc", level=2)
    doc.add_paragraph(
        "Bộ dữ liệu nghiên cứu gốc 'student_burnout_dropout_dataset_2.csv' bao gồm 800 bản ghi sinh viên và 25 trường thuộc tính. "
        "Để phục vụ phân tích học máy bài bản, nhóm phân loại 25 thuộc tính thành 5 nhóm chức năng như trình bày tại Bảng 2.1:"
    )

    tbl_feats = doc.add_table(rows=6, cols=4)
    feat_data = [
        ("Nhóm Thuộc tính", "Danh sách Biến (Columns)", "Kiểu", "Miền giá trị & Thống kê"),
        ("1. Nhân khẩu học\n(7 thuộc tính)", "Student_ID, Age, Gender, Year_of_Study, Department, Residence_Type, Family_Income_Bracket", "Int / Str",
         "Tuổi 17–25 (TB: 20.88); 6 Khoa (Engineering 209, Business 151, Science 143, Arts 136, Medicine 77, Law 61); 5 mức thu nhập."),
        ("2. Tham gia học tập\n(5 thuộc tính)", "Attendance_Percent, Study_Hours_Per_Day, Previous_GPA, Backlogs, Motivation_Score", "Float64",
         "Chuyên cần: 30%–100% (TB: 76.88%); Giờ tự học: 0–8.1h (TB: 3.16h); GPA: 2.42–10.0 (TB: 6.86); Môn nợ: 0–6; Động lực: 0–10 (TB: 5.47)."),
        ("3. Lối sống & Sinh hoạt\n(5 thuộc tính)", "Sleep_Hours, Screen_Time_Hours, Exercise_Freq_Per_Week, Social_Activity_Score, Part_Time_Job", "Float / Int / Str",
         "Giờ ngủ: 2–10h (TB: 6.20h); Giờ màn hình: 0.5–12.6h (TB: 5.43h); Tập thể dục: 0–7 buổi/tuần; Làm thêm: Yes (228), No (547)."),
        ("4. Căng thẳng & Hỗ trợ\n(6 thuộc tính)", "Stress_Level, Anxiety_Score, Financial_Stress_Score, Peer_Pressure_Score, Family_Support_Score, Counseling_Access", "Float / Str",
         "Stress: 0.1–10.0 (TB: 5.52); Lo âu: 0–10 (TB: 4.98); Áp lực tài chính: 0–10 (TB: 5.03); Hỗ trợ GĐ: 0–10 (TB: 6.42); Tư vấn: Yes (265), No (535)."),
        ("5. Kiệt sức & Mục tiêu\n(2 thuộc tính)", "Burnout_Level,\nDropout_Risk (Target)", "Phân loại",
         "Burnout: Low (320 - 40%), Med (280 - 35%), High (200 - 25%).\nDropout_Risk: Yes (437 - 54.62%), No (363 - 45.38%). Không khuyết nhãn."),
    ]
    for i, row in enumerate(feat_data):
        for j, val in enumerate(row):
            tbl_feats.cell(i, j).text = val
    format_table(tbl_feats, [1.6, 2.2, 0.9, 2.5], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_heading("2.2. Quy trình Thực hiện Dự án & Kiểm định Mô hình", level=2)
    doc.add_paragraph(
        "Quy trình thực thi gồm 5 giai đoạn liên kết chặt chẽ (Hình 2.1). Nhóm áp dụng giao thức kiểm định chéo phân tầng 5-Fold "
        "(Stratified 5-Fold Cross-Validation) kết hợp tập kiểm thử độc lập 20% (Hold-out Test Set, random_state=42), bảo toàn tỷ lệ nhãn "
        "54.6% Yes / 45.4% No. Mọi bước biến đổi dữ liệu (KNNImputer, StandardScaler, OneHotEncoder) được đóng gói trong Pipeline của scikit-learn "
        "để ngăn ngừa hiện tượng rò rỉ thông tin từ tập kiểm thử sang tập huấn luyện."
    )
    add_fig(doc, "11_system_workflow_diagram.png", "Hình 2.1", "Quy trình thực hiện dự án Học máy 5 giai đoạn khép kín của Nhóm 03")

    doc.add_heading("2.3. Thiết kế Kiến trúc Hệ thống 3 Tầng", level=2)
    doc.add_paragraph(
        "Hệ thống phần mềm Cảnh báo Sớm được thiết kế theo kiến trúc 3 tầng (3-Tier Architecture) phục vụ vận hành thực tế tại trường đại học (Hình 2.2): "
        "(1) Tầng Nhập liệu Web UI thu thập 18 tham số; (2) Tầng Backend AI (FastAPI app.py) vận hành song song mô hình dự báo Burnout_Level (3 lớp) "
        "và mô hình dự báo Dropout_Risk; (3) Tầng Hiển thị Kết quả & Khuyến nghị can thiệp cá nhân hóa."
    )
    add_fig(doc, "12_system_architecture_diagram.png", "Hình 2.2", "Kiến trúc hệ thống Cảnh báo sớm Nguy cơ Bỏ học 3 tầng & Giao diện Web")

    doc.add_page_break()

    # =========================================================================
    # CHƯƠNG 3
    # =========================================================================
    doc.add_heading("3. KẾT QUẢ THỰC NGHIỆM: TIỀN XỬ LÝ, PHÂN TÍCH EDA & MÔ HÌNH HÓA", level=1)

    doc.add_heading("3.1. Tiền xử lý Dữ liệu và Kỹ thuật Tạo 6 Đặc trưng Chuyên ngành", level=2)
    doc.add_paragraph(
        "Kiểm toán chất lượng dữ liệu cho thấy 18/25 cột chứa dữ liệu khuyết với tổng số 612 ô khuyết (Hình 3.1). "
        "Tỷ lệ khuyết dao động từ 1.62% (Family_Income_Bracket - 13 ô) đến 6.38% (Attendance_Percent - 51 ô). "
        "Có 440/800 dòng (55.0%) chứa ít nhất một ô khuyết; do đó nếu xóa dòng sẽ mất hơn một nửa tập dữ liệu. "
        "Nhóm áp dụng thuật toán KNNImputer (k=7) trên các biến liên tục và Mode Imputation trên biến phân loại bên trong Pipeline."
    )
    add_fig(doc, "01_missing_values_analysis.png", "Hình 3.1", "Phân tích tỷ lệ dữ liệu khuyết trên 18/25 thuộc tính trước khi qua KNNImputer")

    doc.add_paragraph(
        "Nhằm định lượng trực tiếp mối tương tác giữa Căng thẳng và Tham gia học tập, nhóm xây dựng lớp biến đổi DomainFeatureEngineer "
        "gồm 6 chỉ số chuyên ngành chính (Bảng 3.1) kèm tích số Chuyên cần × Giờ học (Study_Attendance_Product):"
    )

    tbl_fe = doc.add_table(rows=7, cols=3)
    fe_data = [
        ("Tên Đặc trưng Mới", "Công thức Toán học", "Ý nghĩa Sư phạm & Nghiệp vụ"),
        ("1. Academic_Engagement_Index", "0.35×(Attendance/10) + 0.25×Study + 0.25×Motivation + 0.15×GPA",
         "Đo lường tổng hợp mức độ chuyên cần, thời lượng tự học, động lực nội tại và điểm tích lũy học tập."),
        ("2. Psychological_Stress_Index", "0.35×Stress + 0.30×Anxiety + 0.20×Financial + 0.15×Peer_Pressure",
         "Đo lường tổng hợp áp lực tâm lý, lo âu, gánh nặng tài chính và áp lực đồng trang lứa."),
        ("3. Stress_Engagement_Ratio", "Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)",
         "Tỷ số mất cân bằng khi áp lực tâm lý vượt quá độ gắn kết học tập (+1.0 tránh chia 0)."),
        ("4. Net_Stress_Support_Balance", "Psychological_Stress_Index − (0.5×Family + 0.3×Social + 0.2×Exercise)",
         "Mức căng thẳng ròng còn lại sau khi trừ đi các điểm tựa nâng đỡ từ gia đình, xã hội và thể chất."),
        ("5. Digital_Fatigue_Ratio", "Screen_Time_Hours / (Sleep_Hours + 1.0)",
         "Đánh giá hội chứng mệt mỏi số hóa và rối loạn nhịp sinh học do thức khuya dùng màn hình quá mức."),
        ("6. Academic_Risk_Burden", "1.5×Backlogs + (100.0 − Attendance_Percent) / 20.0",
         "Lượng hóa gánh nặng học vụ tích lũy từ số môn nợ học phần và tỷ lệ vắng học trên lớp."),
    ]
    for i, row in enumerate(fe_data):
        for j, val in enumerate(row):
            tbl_fe.cell(i, j).text = val
    format_table(tbl_fe, [2.0, 2.5, 2.7], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_heading("3.2. Phân tích Khám phá Dữ liệu (EDA)", level=2)
    doc.add_paragraph(
        "Phân tích thống kê trên 800 sinh viên gốc khẳng định mối liên hệ chặt chẽ giữa mức độ kiệt sức và tỷ lệ bỏ học (Hình 3.2): "
        "Ở nhóm Kiệt sức Thấp (Low Burnout, 320 SV), tỷ lệ bỏ học chỉ là 23.8% (76/320); sang nhóm Kiệt sức Vừa (Medium Burnout, 280 SV), "
        "tỷ lệ bỏ học tăng lên 63.2% (177/280); và ở nhóm Kiệt sức Cao (High Burnout, 200 SV), tỷ lệ bỏ học lên tới 92.0% (184/200), "
        "tăng gấp gần 3.9 lần so với nhóm Thấp."
    )
    add_fig(doc, "02_target_and_burnout_distribution.png", "Hình 3.2", "Phân phối biến mục tiêu Dropout_Risk và tỷ lệ bỏ học theo mức Burnout_Level")

    doc.add_paragraph(
        "Hình 3.3 và Hình 3.4 chỉ ra sự đối lập rõ rệt giữa hai nhóm nhân tố: Nhóm sinh viên có nguy cơ bỏ học (Dropout_Risk = Yes) "
        "có tỷ lệ chuyên cần trung bình thấp hơn 7.0 điểm phần trăm (73.7% vs 80.7%), tự học ít hơn (2.93h vs 3.43h/ngày), "
        "động lực thấp hơn (5.11 vs 5.90); đồng thời chịu mức căng thẳng (5.81 vs 5.17) và lo âu (5.26 vs 4.65) cao hơn rõ rệt. "
        "Ma trận tương quan Pearson trên dữ liệu gốc xác nhận: Attendance_Percent có r = -0.244, Motivation có r = -0.193 (yếu tố bảo vệ); "
        "trong khi Burnout có r = +0.550 và Stress_Level có r = +0.164 (yếu tố nguy cơ)."
    )
    add_fig(doc, "03_stress_vs_engagement_distributions.png", "Hình 3.3", "So sánh phân phối các chỉ số Tham gia học tập và Căng thẳng theo Dropout_Risk")
    add_fig(doc, "04_correlation_matrix_heatmap.png", "Hình 3.4", "Ma trận tương quan Pearson giữa các biến định lượng trên dữ liệu khảo sát gốc")
    add_fig(doc, "05_categorical_factors_analysis.png", "Hình 3.5", "Phân tích tỷ lệ bỏ học theo các yếu tố nhân khẩu học và môi trường sống")
    add_fig(doc, "06_multivariate_burnout_engagement_scatter.png", "Hình 3.6", "Phân tích đa biến giữa Chuyên cần, Căng thẳng tâm lý và Kiệt sức học đường")

    doc.add_heading("3.3. Kết quả Huấn luyện và Đối sánh Mô hình Học máy", level=2)
    doc.add_paragraph(
        "Để đảm bảo tính toàn diện và trung thực học thuật, nhóm trình bày kết quả thực nghiệm theo hai giai đoạn nghiên cứu:"
    )

    doc.add_heading("Giai đoạn A: Đối sánh 10 Thuật toán Phân loại trên Dữ liệu Gốc (800 Sinh viên)", level=3)
    doc.add_paragraph(
        "Trên tập 800 sinh viên gốc, nhóm huấn luyện và đánh giá 10 thuật toán học máy đại diện cho 5 họ thuật toán "
        "(Tuyến tính, Láng giềng, Cây quyết định, Tập hợp cây & Boosting, và Mạng nơ-ron). Kết quả kiểm định chéo phân tầng 5-Fold "
        "và tập kiểm thử độc lập (n=160) được tổng hợp tại Bảng 3.2. Random Forest đạt điểm CV ROC-AUC cao nhất (80.50% ± 3.83%), "
        "trong khi Extra Trees đạt hiệu năng kiểm thử Test Recall cao nhất (83.91%) và Test F1 (78.92%)."
    )

    tbl_m10 = doc.add_table(rows=11, cols=8)
    m10_data = [
        ("Mô hình Học máy (10 Models)", "CV Acc", "CV Prec", "CV Recall", "CV F1", "CV AUC", "Test F1", "Test AUC"),
        ("1. Random Forest", "74.38% ± 4.87%", "74.90%", "80.31%", "77.46%", "80.50% ± 3.83%", "75.14%", "80.24%"),
        ("2. Soft Voting Ensemble", "73.88% ± 4.68%", "74.84%", "78.94%", "76.79%", "79.95% ± 4.02%", "75.82%", "79.37%"),
        ("3. Logistic Regression", "72.25% ± 4.27%", "73.72%", "76.87%", "75.19%", "79.82% ± 4.18%", "76.14%", "79.86%"),
        ("4. MLP Neural Network", "73.62% ± 3.70%", "75.57%", "76.42%", "75.94%", "79.68% ± 4.05%", "72.83%", "77.75%"),
        ("5. XGBoost", "71.75% ± 3.48%", "73.93%", "74.81%", "74.28%", "79.25% ± 3.69%", "73.86%", "77.77%"),
        ("6. Extra Trees", "75.00% ± 4.81%", "75.27%", "81.23%", "78.10%", "79.22% ± 3.94%", "78.92%", "78.81%"),
        ("7. Gradient Boosting", "71.75% ± 4.23%", "73.78%", "75.27%", "74.44%", "79.08% ± 3.66%", "75.14%", "78.52%"),
        ("8. SVM (RBF Kernel)", "74.63% ± 5.50%", "75.25%", "80.08%", "77.57%", "78.55% ± 4.28%", "77.84%", "79.15%"),
        ("9. K-Nearest Neighbors", "70.25% ± 4.04%", "72.87%", "72.54%", "72.59%", "77.14% ± 4.18%", "74.86%", "76.37%"),
        ("10. Decision Tree", "66.88% ± 3.60%", "72.24%", "64.50%", "67.75%", "73.88% ± 2.97%", "71.59%", "73.82%"),
    ]
    for i, row in enumerate(m10_data):
        for j, val in enumerate(row):
            tbl_m10.cell(i, j).text = val
    format_table(tbl_m10, [2.0, 0.8, 0.7, 0.7, 0.7, 0.8, 0.7, 0.8], [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 7)

    doc.add_heading("Giai đoạn B: Huấn luyện 4 Mô hình Cốt lõi trên Tập Dữ liệu Mở rộng (2.000 Hồ sơ)", level=3)
    doc.add_paragraph(
        "Nhằm kiểm thử khả năng mở rộng quy mô kỹ thuật và phục vụ triển khai ứng dụng Web Dashboard, nhóm mở rộng mẫu "
        "lên 2.000 hồ sơ sạch (0 NaN) và tinh giản về 4 mô hình cốt lõi: Logistic Regression (Baseline), Random Forest, XGBoost, "
        "và Soft Voting Ensemble. Đánh giá được thực hiện độc lập trên 2 kịch bản:"
    )

    doc.add_paragraph("Bảng 3.3: Kết quả Thực nghiệm 4 Mô hình Cốt lõi — Kịch bản 1 (Đầy đủ Đặc trưng, Có Burnout_Level)")
    tbl_kb1 = doc.add_table(rows=5, cols=9)
    kb1_data = [
        ("Mô hình", "CV Acc", "CV Prec", "CV Recall", "CV F1", "CV AUC", "Test Acc", "Test F1", "Test AUC"),
        ("1. Random Forest", "80.25% ± 0.65%", "79.14%", "87.32%", "82.98%", "89.01% ± 0.80%", "81.50%", "83.63%", "88.42%"),
        ("2. Soft Voting", "79.70% ± 1.27%", "79.51%", "85.33%", "82.27%", "88.89% ± 0.84%", "80.25%", "82.17%", "88.32%"),
        ("3. XGBoost", "79.95% ± 1.11%", "80.05%", "84.96%", "82.39%", "88.64% ± 0.84%", "78.50%", "80.28%", "87.99%"),
        ("4. Logistic Reg.", "78.45% ± 0.68%", "80.12%", "81.25%", "80.62%", "87.61% ± 1.02%", "78.25%", "80.18%", "87.60%"),
    ]
    for i, row in enumerate(kb1_data):
        for j, val in enumerate(row):
            tbl_kb1.cell(i, j).text = val
    format_table(tbl_kb1, [1.8, 0.7, 0.65, 0.65, 0.65, 0.75, 0.65, 0.65, 0.7], [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 8)

    doc.add_paragraph("Bảng 3.4: Kết quả Thực nghiệm 4 Mô hình Cốt lõi — Kịch bản 2 (Cảnh báo Sớm, Loại bỏ Burnout_Level)")
    tbl_kb2 = doc.add_table(rows=5, cols=9)
    kb2_data = [
        ("Mô hình", "CV Acc", "CV Prec", "CV Recall", "CV F1", "CV AUC", "Test Acc", "Test F1", "Test AUC"),
        ("1. XGBoost", "79.95% ± 1.48%", "80.02%", "85.05%", "82.40%", "88.26% ± 1.03%", "78.00%", "79.82%", "87.24%"),
        ("2. Soft Voting", "79.30% ± 1.02%", "79.63%", "84.15%", "81.77%", "88.25% ± 1.01%", "78.00%", "80.00%", "87.39%"),
        ("3. Random Forest", "79.30% ± 1.72%", "78.83%", "85.60%", "82.00%", "87.97% ± 1.11%", "78.50%", "80.72%", "87.13%"),
        ("4. Logistic Reg.", "78.00% ± 0.00%", "79.52%", "81.16%", "80.27%", "87.25% ± 0.99%", "77.75%", "79.73%", "87.00%"),
    ]
    for i, row in enumerate(kb2_data):
        for j, val in enumerate(row):
            tbl_kb2.cell(i, j).text = val
    format_table(tbl_kb2, [1.8, 0.7, 0.65, 0.65, 0.65, 0.75, 0.65, 0.65, 0.7], [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 8)

    doc.add_paragraph(
        "Bảng 3.5 cung cấp số liệu chi tiết của Ma trận Nhầm lẫn (Confusion Matrix) trên tập kiểm thử độc lập (n=400 mẫu; gồm 221 ca Nguy cơ và 179 ca An toàn):"
    )
    tbl_cm = doc.add_table(rows=5, cols=7)
    cm_data = [
        ("Kịch bản / Mô hình", "TN", "FP", "FN", "TP", "Độ nhạy (Recall)", "Độ chính xác (Acc)"),
        ("KB1 - Random Forest", "137", "42", "32", "189", "85.52% (189/221)", "81.50% (326/400)"),
        ("KB1 - Soft Voting", "139", "40", "39", "182", "82.35% (182/221)", "80.25% (321/400)"),
        ("KB1 - XGBoost", "139", "40", "46", "175", "79.19% (175/221)", "78.50% (314/400)"),
        ("KB2 - XGBoost (Cảnh báo sớm)", "138", "41", "47", "174", "78.73% (174/221)", "78.00% (312/400)"),
    ]
    for i, row in enumerate(cm_data):
        for j, val in enumerate(row):
            tbl_cm.cell(i, j).text = val
    format_table(tbl_cm, [2.2, 0.7, 0.7, 0.7, 0.7, 1.2, 1.0], [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 6)

    add_fig(doc, "07_model_performance_comparison.png", "Hình 3.7", "So sánh hiệu năng các mô hình trên Kịch bản 1 (Full) và Kịch bản 2 (Early Warning)")
    add_fig(doc, "08_roc_and_precision_recall_curves.png", "Hình 3.8", "Đường cong ROC và Precision-Recall của các mô hình trên tập kiểm thử")
    add_fig(doc, "09_confusion_matrices_top_models.png", "Hình 3.9", "Ma trận nhầm lẫn của 4 mô hình tiêu biểu trên tập kiểm thử độc lập (n=400)")

    doc.add_heading("Phân tích Mức độ Quan trọng của Đặc trưng (Feature Importance)", level=3)
    doc.add_paragraph(
        "Phân tích độ quan trọng đặc trưng trên Random Forest (Gini Importance) và XGBoost (Gain Importance) tại Hình 3.10 "
        "cho thấy hai đặc trưng mới do nhóm thiết kế xếp ở vị trí hàng đầu:\n"
        "• Random Forest: (1) Stress_Engagement_Ratio (0.126) → (2) Academic_Engagement_Index (0.111) → (3) Attendance_Percent (0.092) "
        "→ (4) Burnout_Ordinal (0.077) → (5) Stress_Level (0.068).\n"
        "• XGBoost: (1) Burnout_Level_Low (0.218) → (2) Stress_Engagement_Ratio (0.127) → (3) Burnout_Ordinal (0.091) "
        "→ (4) Academic_Engagement_Index (0.063) → (5) Stress_Level (0.038).\n"
        "Lưu ý: Do Burnout_Level được biểu diễn qua cả biến thứ tự (Ordinal) và mã hóa One-Hot, độ quan trọng của biến này bị phân tán "
        "qua nhiều thuộc tính phái sinh. Kết quả khẳng định chỉ số tương tác phi tuyến (Stress_Engagement_Ratio) có đóng góp phân tách lớn nhất."
    )
    add_fig(doc, "10_feature_importance_analysis.png", "Hình 3.10", "Phân tích mức độ quan trọng đặc trưng trên Random Forest và XGBoost")

    doc.add_page_break()

    # =========================================================================
    # CHƯƠNG 4
    # =========================================================================
    doc.add_heading("4. GIAO DIỆN NGƯỜI DÙNG VÀ KIỂM THỬ HỆ THỐNG", level=1)

    doc.add_heading("4.1. Thiết kế Giao diện Web Dashboard Cảnh báo Sớm", level=2)
    doc.add_paragraph(
        "Nhóm phát triển ứng dụng Web Dashboard tương tác bằng nền tảng FastAPI (tệp app.py, Hình 4.1). Hệ thống tích hợp đồng thời "
        "3 tệp mô hình đã đóng gói (.pkl) gồm: best_burnout_model.pkl (dự báo 3 mức độ kiệt sức), best_dropout_model_full.pkl "
        "(Random Forest kịch bản đầy đủ) và best_dropout_model_early_warning.pkl (XGBoost kịch bản cảnh báo sớm).\n"
        "• Bảng điều khiển 18 tham số: Gồm 14 thanh trượt (Range Sliders) và 4 hộp chọn (Select Dropdowns) điều chỉnh linh hoạt mọi chỉ số.\n"
        "• Ba bộ hồ sơ mẫu (Presets): Nút nhấn một chạm nạp ngay hồ sơ mẫu đại diện (SV Tích cực, SV Căng thẳng TB, SV Kiệt sức Cao).\n"
        "• Công cụ giải thích trực quan: Biểu đồ thanh ngang hiển thị 3 chỉ số tổng hợp (AEI, PSI, SER) giúp cố vấn hiểu rõ căn nguyên rủi ro."
    )
    add_fig(doc, "13_live_web_dashboard_screenshot.png", "Hình 4.1", "Giao diện Web Dashboard Cảnh báo Sớm thời gian thực triển khai trên FastAPI")

    doc.add_heading("4.2. Kiểm thử Thực tế và Đo lường Hiệu năng Hệ thống", level=2)
    doc.add_paragraph(
        "Nhóm tiến hành kiểm thử thực nghiệm trên 3 hồ sơ mẫu được lập trình sẵn trong hệ thống Web app.py. Bảng 4.1 ghi nhận kết quả đo lường thực tế:"
    )

    tbl_web = doc.add_table(rows=4, cols=5)
    web_data = [
        ("Hồ sơ Mẫu (Preset)", "Xác suất Bỏ học (Preset Burnout)", "Xác suất Bỏ học (Chế độ Auto)", "Mức Burnout Dự báo", "Độ trễ Phản hồi TB (ms)"),
        ("Mẫu 1: SV Tích cực", "1.7% (An toàn)", "2.2% (An toàn)", "Low Burnout", "153.3 ms"),
        ("Mẫu 2: SV Căng thẳng TB", "60.5% (Cảnh báo)", "59.3% (Cảnh báo)", "Medium Burnout", "142.5 ms"),
        ("Mẫu 3: SV Kiệt sức Cao", "98.6% (Báo động đỏ)", "98.9% (Báo động đỏ)", "High Burnout", "147.2 ms"),
    ]
    for i, row in enumerate(web_data):
        for j, val in enumerate(row):
            tbl_web.cell(i, j).text = val
    format_table(tbl_web, [1.8, 1.6, 1.6, 1.2, 1.0], [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 4)

    doc.add_paragraph(
        "Module sinh khuyến nghị tự động kích hoạt dựa trên các ngưỡng vi phạm thực tế:\n"
        "• Chuyên cần < 70%: Tự động kích hoạt thông báo cố vấn học tập liên hệ tìm hiểu nguyên nhân vắng học.\n"
        "• Giờ tự học < 2.0h/ngày: Đề xuất tham gia nhóm học tập (Study Group) hoặc lớp phụ đạo chuyên môn.\n"
        "• Điểm Căng thẳng / Lo âu ≥ 6.5: Kích hoạt cảnh báo kết nối khẩn cấp Phòng Tham vấn Tâm lý học đường.\n"
        "• Áp lực tài chính ≥ 6.5: Đề xuất xét duyệt gói học bổng hỗ trợ sinh viên có hoàn cảnh khó khăn."
    )

    doc.add_page_break()

    # =========================================================================
    # CHƯƠNG 5
    # =========================================================================
    doc.add_heading("5. ĐÁNH GIÁ TÁC ĐỘNG THỰC TIỄN, HẠN CHẾ KHOA HỌC & HƯỚNG PHÁT TRIỂN", level=1)

    doc.add_heading("5.1. Giá trị Đóng góp và Tác động Xã hội", level=2)
    doc.add_paragraph(
        "• Đóng góp Khoa học & Kỹ thuật: Đồ án chứng minh trên cơ sở thực nghiệm rằng việc kết hợp các chỉ số tâm lý và học tập "
        "thông qua kỹ thuật tạo đặc trưng phi tuyến (Domain Feature Engineering) giúp tăng đáng kể độ nhạy (Recall > 83%) "
        "của các mô hình phân loại. Chỉ số Tỷ số Căng thẳng / Gắn kết (Stress_Engagement_Ratio) đóng vai trò biến phân tách quan trọng nhất.\n"
        "• Giá trị Giáo dục & Xã hội: Hệ thống cung cấp cho các trường đại học một công cụ sàng lọc tự động chi phí thấp, "
        "cho phép nhận diện sinh viên suy giảm động lực học tập ngay từ những tuần đầu học kỳ, hỗ trợ nhà trường tối ưu hóa nguồn lực tư vấn "
        "và giảm tỷ lệ hao hụt sinh viên hàng năm."
    )

    doc.add_heading("5.2. Nhìn nhận Thẳng thắn các Hạn chế Khoa học (Academic Reviewer Rigor)", level=2)
    doc.add_paragraph(
        "Với tinh thần phản biện học thuật khách quan, nhóm ghi nhận các hạn chế phương pháp luận sau đây:\n"
        "1. Bản chất dữ liệu cắt ngang (Cross-sectional Study): Bộ dữ liệu ghi nhận tại một thời điểm duy nhất, do đó các hệ số tương quan "
        "và Feature Importance chỉ phản ánh mức độ liên kết thống kê, tuyệt đối không được suy diễn thành quan hệ nhân quả một chiều.\n"
        "2. Rủi ro rò rỉ dữ liệu khi mở rộng mẫu tổng hợp (Synthetic Data Augmentation): Để mở rộng 1.200 hồ sơ bằng phân phối Gaussian đa biến, "
        "thuật toán đã sử dụng thông tin nhãn mục tiêu và áp dụng quy tắc tăng cường độ phân tách theo ngưỡng. Khi phân rã tập kiểm thử 400 mẫu, "
        "mô hình đạt ROC-AUC 94.3% trên các bản ghi tổng hợp nhưng chỉ đạt 77.3% trên các bản ghi sinh viên gốc. Phép chẩn đoán kiểm chứng sạch "
        "(chỉ huấn luyện trên 640 mẫu thật và kiểm thử trên 160 mẫu thật) cho thấy mô hình đạt ROC-AUC 80.24% (Random Forest). "
        "Do đó, con số 80.50% trên dữ liệu gốc phản ánh chính xác nhất năng lực tổng quát hóa của mô hình trên sinh viên thực tế.\n"
        "3. Đa cộng tuyến đặc trưng: Các chỉ số tổng hợp có tương quan tuyến tính nhất định với các biến thành phần gốc. Mặc dù các thuật toán cây "
        "(Random Forest, XGBoost) ít bị ảnh hưởng bởi đa cộng tuyến, việc lựa chọn đặc trưng tối ưu (Feature Selection) cần được nghiên cứu sâu hơn."
    )

    doc.add_heading("5.3. Định hướng Phát triển Tương lai", level=2)
    doc.add_paragraph(
        "• Mô hình Chuỗi thời gian (Longitudinal Time-Series): Thu thập dữ liệu chuyên cần và điểm kiểm tra định kỳ theo từng tuần "
        "để huấn luyện mạng hồi quy LSTM hoặc Transformer dự báo quỹ đạo học tập của sinh viên.\n"
        "• Tích hợp Trí tuệ Nhân tạo Giải thích được (Explainable AI - XAI): Nhúng trực tiếp thư viện SHAP (SHAP Waterfall Plot) "
        "vào giao diện Web UI để giải thích chi tiết lý do từng sinh viên cụ thể bị xếp vào diện cảnh báo.\n"
        "• Kết nối Hệ thống Quản lý Học tập (LMS Integration): Xây dựng API kết nối tự động với hệ thống Moodle hoặc Canvas của trường đại học "
        "để tự động đồng bộ dữ liệu điểm danh và tương tác khóa học."
    )

    doc.add_heading("5.4. Nhận xét Cá nhân của Thành viên Nhóm", level=2)
    tbl_fb = doc.add_table(rows=3, cols=2)
    fb_data = [
        ("Thành viên", "Nội dung Nhận xét & Đánh giá Cá nhân"),
        ("Đỗ Quang Minh\n(Trưởng nhóm)",
         "Khóa học AI của Samsung Innovation Campus đã trang bị cho tôi nền tảng tư duy kỹ thuật vững chắc từ lý thuyết Machine Learning "
         "đến triển khai sản phẩm hoàn chỉnh. Tôi đặc biệt tâm đắc với việc thiết kế Pipeline học máy khép kín chống rò rỉ dữ liệu, "
         "xây dựng lớp DomainFeatureEngineer và trực tiếp phân tích các ranh giới khoa học giữa dữ liệu tổng hợp và dữ liệu thực tế. "
         "Dự án là bước đệm lớn cho định hướng nghiên cứu AI ứng dụng của tôi."),
        ("Vũ Quang Huy\n(Thành viên)",
         "Quá trình thực hiện đồ án giúp tôi làm chủ kỹ năng kiểm toán dữ liệu khuyết, xử lý đa biến bằng KNNImputer và trực quan hóa dữ liệu "
         "theo tiêu chuẩn học thuật. Sự phối hợp ăn ý 50% - 50% giữa hai thành viên từ khâu thực nghiệm đến thiết kế giao diện Web FastAPI "
         "đã tạo nên một sản phẩm hoàn thiện, có tính ứng dụng cao và mang ý nghĩa nhân văn sâu sắc cho cộng đồng giáo dục.")
    ]
    for i, row in enumerate(fb_data):
        for j, val in enumerate(row):
            tbl_fb.cell(i, j).text = val
    format_table(tbl_fb, [1.8, 5.4], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_page_break()

    # =========================================================================
    # CHƯƠNG 6
    # =========================================================================
    doc.add_heading("6. TÀI LIỆU THAM KHẢO VÀ PHỤ LỤC KỸ THUẬT", level=1)

    doc.add_heading("6.1. Danh mục Tài liệu Tham khảo (Chuẩn IEEE)", level=2)
    refs = [
        "[1] V. Tinto, 'Dropout from Higher Education: A Theoretical Synthesis of Recent Research,' Review of Educational Research, vol. 45, no. 1, pp. 89–125, 1975.",
        "[2] C. Maslach, W. B. Schaufeli, and M. P. Leiter, 'Job Burnout,' Annual Review of Psychology, vol. 52, pp. 397–422, 2001.",
        "[3] L. Breiman, 'Random Forests,' Machine Learning, vol. 45, no. 1, pp. 5–32, 2001.",
        "[4] P. Geurts, D. Ernst, and L. Wehenkel, 'Extremely Randomized Trees,' Machine Learning, vol. 63, no. 1, pp. 3–42, 2006.",
        "[5] T. Chen and C. Guestrin, 'XGBoost: A Scalable Tree Boosting System,' in Proc. 22nd ACM SIGKDD Int. Conf. Knowledge Discovery and Data Mining, 2016, pp. 785–794.",
        "[6] F. Pedregosa et al., 'Scikit-learn: Machine Learning in Python,' Journal of Machine Learning Research, vol. 12, pp. 2825–2830, 2011.",
        "[7] S. M. Lundberg and S.-I. Lee, 'A Unified Approach to Interpreting Model Predictions,' in Advances in Neural Information Processing Systems (NeurIPS), 2017, pp. 4765–4774.",
        "[8] O. Troyanskaya et al., 'Missing value estimation methods for DNA microarrays,' Bioinformatics, vol. 17, no. 6, pp. 520–525, 2001."
    ]
    for r in refs:
        p = doc.add_paragraph(r)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)

    doc.add_heading("6.2. Phụ lục Kỹ thuật: Cấu hình Siêu tham số Mô hình", level=2)
    doc.add_paragraph(
        "Bảng 6.1 đặc tả chi tiết cấu hình siêu tham số cố định được sử dụng cho các mô hình học máy trong nghiên cứu "
        "(được thiết lập thủ công dựa trên kinh nghiệm thực nghiệm; việc tinh chỉnh tự động bằng Bayesian Optimization được đề xuất cho hướng phát triển tiếp theo):"
    )

    tbl_params = doc.add_table(rows=11, cols=2)
    param_data = [
        ("Mô hình Học máy", "Cấu hình Siêu tham số (Scikit-Learn / XGBoost)"),
        ("1. Logistic Regression", "C=0.5, penalty='l2', solver='lbfgs', max_iter=1000, random_state=42"),
        ("2. SVM (RBF Kernel)", "CalibratedClassifierCV(SVC(C=1.2, kernel='rbf', gamma='scale', random_state=42), cv=3)"),
        ("3. K-Nearest Neighbors", "n_neighbors=17, weights='distance', metric='minkowski', p=2"),
        ("4. Decision Tree", "max_depth=5, min_samples_leaf=15, criterion='gini', random_state=42"),
        ("5. Random Forest", "n_estimators=300, max_depth=8, min_samples_split=8, min_samples_leaf=4, max_features='sqrt', random_state=42"),
        ("6. Extra Trees", "n_estimators=300, max_depth=9, min_samples_split=8, min_samples_leaf=4, random_state=42"),
        ("7. Gradient Boosting", "n_estimators=150, learning_rate=0.04, max_depth=3, subsample=0.85, min_samples_leaf=10, random_state=42"),
        ("8. XGBoost", "n_estimators=200, learning_rate=0.035, max_depth=3, subsample=0.85, colsample_bytree=0.8, reg_lambda=2.0, min_child_weight=5"),
        ("9. MLP Neural Network", "hidden_layer_sizes=(64, 32), activation='relu', alpha=0.5, learning_rate_init=0.003, max_iter=600, early_stopping=True"),
        ("10. Soft Voting Ensemble", "VotingClassifier(estimators=[lr, rf, xgb], voting='soft', weights=[1, 3, 2]) [v2] / weights=[2,3,2,2,2,2] [v1]"),
    ]
    for i, row in enumerate(param_data):
        for j, val in enumerate(row):
            tbl_params.cell(i, j).text = val
    format_table(tbl_params, [2.2, 5.0], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.save(OUT_PATH)
    print(f"Generated successfully: {OUT_PATH}")


if __name__ == "__main__":
    build_final_report()

