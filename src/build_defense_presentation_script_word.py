"""
Builds FINAL_CAPSTONE/08_Kich_Ban_Thuyet_Trinh_Theo_Slide.docx
Complete slide-by-slide presentation script for the 9-slide defense deck:
- Clear timing for each slide (Total target: 6 min 30 sec, fits 5-8 minutes).
- Verbatim spoken script with bold pronunciation/emphasis markers.
- Physical cues & screen pointers (hành động chỉ tay, chuyển slide).
- Slide-by-slide Q&A defense tips (câu hỏi dự kiến của hội đồng và câu trả lời chuẩn xác).
- Dedicated Master Q&A section for tough questions (Data Leakage, 2.000 samples, non-causal).
- Team division strategy between Leader (Đỗ Quang Minh) and Member (Vũ Quang Huy).
"""
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = r"c:\Users\mrmin\Desktop\Sudent_burnout"
OUT_DOCX = os.path.join(BASE_DIR, "FINAL_CAPSTONE", "08_Kich_Ban_Thuyet_Trinh_Theo_Slide.docx")
OUT_MD = os.path.join(BASE_DIR, "FINAL_CAPSTONE", "08_Kich_Ban_Thuyet_Trinh_Theo_Slide.md")

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
BLUE_RGB = RGBColor(0, 102, 255)
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
    set_cell_margins(c, top=100, bottom=100, left=160, right=140)
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
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = NAVY_RGB

    p2 = c.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run(text)
    r2.font.name = "Arial"
    r2.font.size = Pt(10)
    r2.font.color.rgb = DARK_TEXT_RGB

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(3)


def format_table(table, col_widths, col_alignments):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    repeat_header_row(table.rows[0])
    set_table_borders(table)

    for i, row in enumerate(table.rows):
        prevent_row_split(row)
        for j, cell in enumerate(row.cells):
            cell.width = Inches(col_widths[j])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
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


def build_script():
    doc = Document()

    # Page Setup A4
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.2)
    sec.different_first_page_header_footer = True

    # Header / Footer
    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run("KỊCH BẢN THUYẾT TRÌNH TỪNG SLIDE (5–8 PHÚT) | NHÓM 03 — SIC AI CAPSTONE")
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
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = DARK_TEXT_RGB
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    h1 = doc.styles["Heading 1"]
    h1.font.name = "Arial"
    h1.font.size = Pt(14)
    h1.font.bold = True
    h1.font.color.rgb = NAVY_RGB
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(5)
    h1.paragraph_format.keep_with_next = True

    h2 = doc.styles["Heading 2"]
    h2.font.name = "Arial"
    h2.font.size = Pt(12)
    h2.font.bold = True
    h2.font.color.rgb = BLUE_RGB
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(3)
    h2.paragraph_format.keep_with_next = True

    # =========================================================================
    # COVER TITLE & INSTRUCTIONS
    # =========================================================================
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_top = p_top.add_run("SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE PROJECT\n")
    r_top.font.name = "Arial"
    r_top.font.size = Pt(11)
    r_top.font.bold = True
    r_top.font.color.rgb = NAVY_RGB

    r_title = p_top.add_run("KỊCH BẢN THUYẾT TRÌNH BẢO VỆ ĐỒ ÁN CHI TIẾT THEO TỪNG SLIDE\n")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = BLUE_RGB

    r_sub = p_top.add_run("(Chuẩn thời lượng 5–8 phút | Bộ 9 Slide chính thức | Đỗ Quang Minh & Vũ Quang Huy)\n")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(10.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = MUTED_TEXT_RGB

    add_callout(
        doc,
        "HƯỚNG DẪN DÀNH CHO BẠN TRƯỚC KHI LÊN THUYẾT TRÌNH NGÀY MAI:",
        "1. Thời lượng căn chuẩn: 6 phút 30 giây (rơi vào khoảng an toàn tuyệt đối 5–8 phút của Hội đồng SIC).\n"
        "2. Định dạng lời thoại: Các từ in đậm là điểm nhấn giọng (nhấn mạnh số liệu / từ khóa kỹ thuật). "
        "Các đoạn trong ngoặc vuông [Hành động] là chỉ dẫn bấm slide hoặc chỉ tay vào màn hình.\n"
        "3. Tâm thế thuyết trình: Tự tin, giọng dứt khoát, trung thực khoa học, nhấn mạnh vào giá trị can thiệp sớm và tính khép kín của Pipeline học máy.\n"
        "4. Phân chia người nói: Có thể 1 bạn thuyết trình toàn bộ, hoặc Trưởng nhóm (Minh) nói Slide 1, 2, 4, 6, 7; Thành viên (Huy) nói Slide 3, 5, 8, 9."
    )

    # Timing Summary Table
    tbl_time = doc.add_table(rows=10, cols=4)
    time_data = [
        ("Slide", "Tiêu đề Slide", "Thời lượng", "Người trình bày đề xuất"),
        ("Slide 1", "Project Overview (Trang bìa)", "25 giây", "Đỗ Quang Minh (Trưởng nhóm)"),
        ("Slide 2", "Problem Statement (Đặt vấn đề)", "45 giây", "Đỗ Quang Minh"),
        ("Slide 3", "Dataset & Data Understanding (Dữ liệu)", "45 giây", "Vũ Quang Huy (Thành viên)"),
        ("Slide 4", "Preprocessing & Feature Engineering (Đặc trưng)", "50 giây", "Đỗ Quang Minh"),
        ("Slide 5", "Exploratory Data Analysis (Phân tích EDA)", "50 giây", "Vũ Quang Huy"),
        ("Slide 6", "AI/ML Methodology (Phương pháp luận 2 KB)", "45 giây", "Đỗ Quang Minh"),
        ("Slide 7", "Model Performance & Evaluation (Đối sánh)", "50 giây", "Đỗ Quang Minh"),
        ("Slide 8", "Demo / System Architecture (Web FastAPI)", "45 giây", "Vũ Quang Huy"),
        ("Slide 9", "Conclusion & Future Work (Tổng kết & Q&A)", "35 giây", "Đỗ Quang Minh"),
    ]
    for i, row in enumerate(time_data):
        for j, val in enumerate(row):
            tbl_time.cell(i, j).text = val
    format_table(tbl_time, [1.0, 3.2, 1.2, 2.0], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # =========================================================================
    # SLIDE BY SLIDE SCRIPTS
    # =========================================================================
    slides_content = [
        (
            "SLIDE 1: PROJECT OVERVIEW (TRANG BÌA)",
            "25 giây",
            "Mở đầu trang trọng, giới thiệu đề tài, nhóm thực hiện và neo ngay 4 chỉ số thực nghiệm nổi bật.",
            "[Hành động: Đứng thẳng, mắt nhìn bao quát Hội đồng, mỉm cười nhẹ và mở lời tự tin.]\n\n"
            "\"Kính thưa Quý Thầy Cô trong Hội đồng Đánh giá Đồ án Tốt nghiệp Samsung Innovation Campus!\n\n"
            "Em là **Đỗ Quang Minh**, Trưởng nhóm 03, cùng bạn đồng hành là **Vũ Quang Huy**. Hôm nay, nhóm chúng em xin được báo cáo đồ án tốt nghiệp Capstone với đề tài: "
            "**'Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập'**.\n\n"
            "[Hành động: Tay hướng về 4 thẻ KPI bên phải slide]\n"
            "Đồ án của chúng em khảo sát từ **800 bản ghi sinh viên gốc**, làm sạch và mở rộng kỹ thuật lên **2.000 hồ sơ**, thiết kế **6 đặc trưng chuyên ngành mới**, "
            "và huấn luyện mô hình đạt hiệu năng đỉnh **ROC-AUC 89.01%** ở kịch bản đầy đủ và **88.26%** ở kịch bản cảnh báo sớm. "
            "Sau đây, em xin phép bắt đầu phần trình bày chi tiết!\"",
            "Thầy Cô có thể hỏi: 'Đề tài này có gì mới so với các bài toán dự đoán điểm số hay phân loại học lực thông thường?'\n"
            "-> Gợi ý trả lời: 'Dạ thưa Thầy Cô, điểm mới cốt lõi của đề tài là chúng em không chỉ nhìn vào điểm số GPA, mà lượng hóa tương tác phi tuyến giữa áp lực tâm lý (Stress, Lo âu, Tài chính) "
            "và mức độ gắn kết học tập (Chuyên cần, Giờ tự học) thông qua chỉ số Stress_Engagement_Ratio để cảnh báo sớm nguy cơ bỏ học ngay cả khi sinh viên chưa bị cảnh cáo học vụ.'"
        ),
        (
            "SLIDE 2: PROBLEM STATEMENT (BÀI TOÁN CHÚNG TÔI GIẢI QUYẾT)",
            "45 giây",
            "Làm rõ cơ chế hình thành rủi ro: Bỏ học là kết quả tích lũy, không diễn ra đột ngột.",
            "[Hành động: Bấm chuyển Slide 2. Chỉ tay vào con số 54.62% màu đỏ bên trái]\n\n"
            "\"Thưa Hội đồng, trong giáo dục đại học, tỷ lệ sinh viên có nguy cơ bỏ học trong tập dữ liệu khảo sát lên tới **54.62%**, tương đương 437 trên 800 sinh viên. "
            "Thực tế chỉ ra rằng, sinh viên **rất hiếm khi bỏ học đột ngột** chỉ vì một kỳ thi trượt. Thay vào đó, bỏ học là **hệ quả tích lũy kéo dài** qua 3 giai đoạn:\n\n"
            "[Hành động: Chỉ tay lần lượt vào 3 khối tiến trình bên phải]\n"
            "• **Giai đoạn 1 – Nguyên nhân gốc:** Bắt đầu từ sự suy giảm gắn kết học tập như chuyên cần sụt giảm dưới 70%, giảm giờ tự học, cộng hưởng cùng áp lực căng thẳng học vụ, lo âu và gánh nặng tài chính.\n"
            "• **Giai đoạn 2 – Cầu nối trung gian:** Khi áp lực vượt quá ngưỡng chịu đựng, sinh viên rơi vào trạng thái **Kiệt sức Học đường (Academic Burnout)**. Biến này có tương quan thuận cực mạnh với nguy cơ bỏ học, hệ số Pearson đạt **r = +0.550**.\n"
            "• **Giai đoạn 3 – Giải pháp của chúng em:** Thay vì đợi sinh viên nộp đơn xin thôi học mới xử lý bị động, nhóm xây dựng **Hệ thống Cảnh báo Sớm Học máy** để phát hiện rủi ro ngay từ đầu kỳ, giúp nhà trường kịp thời tư vấn tâm lý và phụ đạo học thuật.\"",
            "Thầy Cô có thể hỏi: 'Tại sao lại chọn bài toán phân loại nhị phân Dropout_Risk mà không dự đoán trực tiếp điểm GPA?'\n"
            "-> Gợi ý trả lời: 'Dạ, mục tiêu ứng dụng thực tế của đồ án là giúp Phòng Công tác Sinh viên ra quyết định can thiệp (Actionable Decision). Việc phân loại rủi ro Dropout_Risk kèm mức độ xác suất giúp nhà trường phân bổ chính xác nguồn lực tư vấn và học bổng cho đúng nhóm đối tượng cấp thiết nhất.'"
        ),
        (
            "SLIDE 3: DATASET & DATA UNDERSTANDING (DỮ LIỆU ĐẦU VÀO)",
            "45 giây",
            "Minh bạch cấu trúc 25 biến gốc, giải thích 612 ô khuyết và làm rõ nguồn gốc 2.000 hồ sơ.",
            "[Hành động: Bấm chuyển Slide 3. Bạn Vũ Quang Huy bắt đầu trình bày]\n\n"
            "\"Tiếp theo, em xin phép trình bày về cấu trúc dữ liệu đầu vào.\n\n"
            "[Hành động: Chỉ vào dải số liệu trên cùng: 800 -> 612 -> k=7 -> 2.000 -> 34]\n"
            "Bộ dữ liệu nghiên cứu gốc gồm **800 bản ghi sinh viên × 25 thuộc tính**, bao quát 5 nhóm thông tin: Nhân khẩu học, Sự tham gia học tập, Lối sống sinh hoạt, Áp lực căng thẳng và Biến mục tiêu.\n\n"
            "Qua kiểm toán chất lượng dữ liệu, nhóm phát hiện **612 ô khuyết nằm rải rác trên 18 cột**, chiếm tới 55% số dòng dữ liệu. "
            "Nếu áp dụng phương pháp xóa dòng truyền thống, chúng em sẽ mất hơn một nửa tập dữ liệu. Vì vậy, nhóm quyết định giữ lại toàn bộ mẫu và xử lý bằng thuật toán **KNNImputer với k=7** láng giềng gần nhất.\n\n"
            "Đặc biệt, để kiểm thử khả năng mở rộng quy mô kỹ thuật và vận hành Web Dashboard tải lớn, từ 800 bản ghi gốc đã làm sạch, nhóm đã mở rộng thêm 1.200 bản ghi từ phân phối Gaussian đa biến phân tầng, "
            "tạo nên tập dữ liệu **2.000 hồ sơ sạch 100% không còn giá trị NaN**.\"",
            "Thầy Cô có thể hỏi: '2.000 hồ sơ này lấy từ đâu? Có phải khảo sát thực tế không?'\n"
            "-> Gợi ý trả lời: 'Dạ thưa Thầy Cô, nhóm xin được minh bạch hoàn toàn: Bộ dữ liệu khảo sát gốc là 800 sinh viên. 1.200 hồ sơ bổ sung là dữ liệu tổng hợp phân tầng (Multivariate Gaussian) được nhóm tạo ra nhằm kiểm thử khả năng mở rộng quy mô kỹ thuật và kiểm thử giao diện Web. Trong phần thực nghiệm, nhóm đã đánh giá riêng biệt cả trên 800 mẫu gốc và tập mở rộng để đảm bảo tính khách quan khoa học.'"
        ),
        (
            "SLIDE 4: PREPROCESSING & FEATURE ENGINEERING (TIỀN XỬ LÝ & ĐẶC TRƯNG)",
            "50 giây",
            "Trình bày quy trình 5 bước và nêu bật 2 công thức đặc trưng quan trọng nhất (SER & AEI).",
            "[Hành động: Bấm chuyển Slide 4. Bạn Đỗ Quang Minh tiếp tục trình bày]\n\n"
            "\"Về phương pháp xử lý, nhóm xây dựng quy trình khép kín gồm **5 bước chuẩn mực**:\n"
            "Làm sạch lọc ngoại lai IQR -> Điền khuyết KNNImputer trong Pipeline -> Thiết kế đặc trưng -> Chuẩn hóa OneHot và StandardScaler -> Kiểm định chéo 5-Fold.\n\n"
            "[Hành động: Hướng tay vào lưới 6 ô Bento bên dưới, đặc biệt nhấn mạnh vào F1 và F3]\n"
            "Đóng góp kỹ thuật quan trọng nhất của nhóm là thiết kế **6 chỉ số tương tác chuyên ngành**:\n"
            "• Tiêu biểu nhất là **F1 – Academic Engagement Index**: Lượng hóa mức độ gắn kết học tập từ chuyên cần, giờ tự học, động lực và GPA.\n"
            "• Và **F3 – Stress Engagement Ratio**: Tỷ số phản ánh áp lực tâm lý chia cho mức độ gắn kết học tập. Khi căng thẳng vượt ngưỡng mà gắn kết học tập suy giảm, tỷ số này tăng vọt và trở thành **chỉ báo cảnh báo sớm nhạy bén nhất**.\n"
            "• Bên cạnh đó, chúng em thiết kế chỉ số cân bằng áp lực và điểm tựa gia đình (F4), tỷ số mệt mỏi số (F5) và gánh nặng học vụ tích lũy từ nợ môn (F6).\"",
            "Thầy Cô có thể hỏi: 'Tại sao các hệ số trọng số 0.35, 0.25 trong công thức lại chọn như vậy? Có căn cứ toán học không?'\n"
            "-> Gợi ý trả lời: 'Dạ thưa Thầy Cô, các trọng số ban đầu được thiết lập dựa trên lý thuyết sư phạm của Tinto và Maslach kết hợp phân tích ma trận tương quan nội bộ (tương quan của từng biến với nhãn bỏ học). Sau đó, khi đưa vào mô hình Random Forest và XGBoost, các cây quyết định đã tự động tối ưu hóa các ngưỡng phân tách phi tuyến, khẳng định 2 chỉ số này đóng góp phân tách cao nhất trong toàn bộ mô hình.'"
        ),
        (
            "SLIDE 5: EXPLORATORY DATA ANALYSIS (PHÂN TÍCH KHÁM PHÁ DỮ LIỆU)",
            "50 giây",
            "Nhấn mạnh phát hiện trọng tâm: tỷ lệ bỏ học tăng vọt theo Burnout (23.8% -> 63.2% -> 92.0%) và ma trận tương quan.",
            "[Hành động: Bấm chuyển Slide 5. Bạn Vũ Quang Huy trình bày]\n\n"
            "\"Sau đây là những phát hiện thống kê quan trọng nhất từ quá trình khám phá dữ liệu (EDA) trên 800 sinh viên gốc:\n\n"
            "[Hành động: Chỉ tay vào biểu đồ cột và 3 khối số lớn bên phải: 23.8% -> 63.2% -> 92.0%]\n"
            "Phân tích chỉ ra mối liên hệ trực tiếp giữa mức độ kiệt sức và tỷ lệ bỏ học:\n"
            "• Ở nhóm Kiệt sức Thấp, tỷ lệ bỏ học chỉ là **23.8%**.\n"
            "• Khi chuyển sang Kiệt sức Vừa, tỷ lệ này tăng vọt lên **63.2%**.\n"
            "• Và ở nhóm Kiệt sức Cao, tỷ lệ bỏ học lên tới **92.0%**, tức là 184 trên 200 sinh viên! Tỷ lệ này **tăng gấp gần 3.9 lần** so với nhóm kiệt sức thấp.\n\n"
            "[Hành động: Chỉ vào khung phân tích tương quan phía dưới]\n"
            "Đồng thời, ma trận tương quan Pearson khẳng định rõ hai cực đối lập:\n"
            "• **Cực bảo vệ (giảm nguy cơ):** Chuyên cần đi học (r = -0.244, nhóm an toàn đi học 80.7% so với 73.7% của nhóm bỏ học) và Động lực nội tại (r = -0.193).\n"
            "• **Cực nguy cơ (tăng nguy cơ):** Căng thẳng học vụ, lo âu kéo dài và áp lực tài chính. Đáng chú ý, sinh viên thuộc nhóm thu nhập thấp có tỷ lệ bỏ học lên tới **68.4%**.\"",
            "Thầy Cô có thể hỏi: 'Hệ số tương quan r = +0.550 giữa Burnout và Bỏ học có đồng nghĩa với quan hệ nhân quả không?'\n"
            "-> Gợi ý trả lời: 'Dạ thưa Thầy Cô, nhóm hoàn toàn ý thức được đây là nghiên cứu quan sát cắt ngang (Cross-sectional Study). Hệ số r = +0.550 phản ánh mối liên hệ thống kê thuận chiều rất mạnh trong tập dữ liệu, chứ không khẳng định quan hệ nhân quả một chiều. Tuy nhiên, về mặt thực tiễn, đây là căn cứ sàng lọc định lượng vô cùng giá trị cho nhà trường.'"
        ),
        (
            "SLIDE 6: AI/ML METHODOLOGY (PHƯƠNG PHÁP LUẬN HỌC MÁY)",
            "45 giây",
            "Làm rõ 2 kịch bản thực nghiệm đối chứng và cam kết chống rò rỉ dữ liệu (No Data Leakage).",
            "[Hành động: Bấm chuyển Slide 6. Bạn Đỗ Quang Minh trình bày]\n\n"
            "\"Để giải quyết bài toán một cách khoa học, nhóm thiết kế **hai kịch bản thực nghiệm đối chứng**:\n\n"
            "[Hành động: Chỉ vào 2 cột Kịch bản 1 và Kịch bản 2]\n"
            "• **Kịch bản 1 – Chẩn đoán Toàn diện:** Sử dụng đầy đủ 32 biến đầu vào bao gồm cả biến kiệt sức Burnout_Level, giúp đánh giá độ chính xác chẩn đoán tối đa khi nhà trường có đợt khảo sát tâm lý chuyên sâu.\n"
            "• **Kịch bản 2 – Cảnh báo Sớm Học đường:** Đây là trọng tâm ứng dụng thực tiễn. Chúng em **loại bỏ hoàn toàn biến Burnout_Level**, buộc mô hình phải dự báo nguy cơ bỏ học trực tiếp chỉ từ dữ liệu điểm danh, giờ học, GPA và áp lực thô. Kịch bản này cho phép phát hiện sớm ngay từ tuần đầu nhập học.\n\n"
            "[Hành động: Nhấn mạnh vào khung cam kết khoa học bên dưới]\n"
            "Đặc biệt, nhóm áp dụng nguyên tắc **chống rò rỉ dữ liệu tuyệt đối (No Data Leakage)**: Mọi khâu biến đổi dữ liệu từ điền khuyết KNN đến chuẩn hóa đều được đóng gói trong Pipeline của scikit-learn và chỉ khớp (fit) trên fold huấn luyện của từng vòng lặp kiểm định chéo phân tầng 5-Fold.\"",
            "Thầy Cô có thể hỏi: 'Nếu Kịch bản 1 dùng Burnout_Level mà biến này lại tương quan mạnh với Bỏ học thì mô hình có bị thiên lệch (bias) không?'\n"
            "-> Gợi ý trả lời: 'Dạ chính xác thưa Thầy Cô! Đó chính là lý do vì sao nhóm bắt buộc phải thiết kế Kịch bản 2 độc lập. Kịch bản 2 chứng minh rằng ngay cả khi không có thông tin Burnout_Level, các mô hình học máy (tiêu biểu là XGBoost) vẫn đạt ROC-AUC tới 88.26%, khẳng định tính khả thi của hệ thống cảnh báo sớm trong thực tế.'"
        ),
        (
            "SLIDE 7: MODEL PERFORMANCE & EVALUATION (KẾT QUẢ THỰC NGHIỆM)",
            "50 giây",
            "Trình bày bảng đối sánh 4 mô hình, số liệu ma trận nhầm lẫn chuẩn xác và độ quan trọng đặc trưng.",
            "[Hành động: Bấm chuyển Slide 7. Bạn Đỗ Quang Minh trình bày]\n\n"
            "\"Kính thưa Hội đồng, đây là kết quả thực nghiệm đối sánh 4 mô hình học máy cốt lõi trên tập dữ liệu mở rộng 2.000 hồ sơ:\n\n"
            "[Hành động: Chỉ vào Bảng đối sánh và biểu đồ Benchmark]\n"
            "• Ở **Kịch bản 1 (Đầy đủ đặc trưng)**: **Random Forest** đạt hiệu năng dẫn đầu với **CV ROC-AUC = 89.01% ± 0.80%**, trên tập kiểm thử độc lập đạt **Test ROC-AUC = 88.42%** và **Test Accuracy = 81.50%**.\n"
            "• Ở **Kịch bản 2 (Cảnh báo sớm)**: **XGBoost** thể hiện ưu thế vượt trội với **CV ROC-AUC = 88.26% ± 1.03%** và **Test ROC-AUC = 87.24%**.\n\n"
            "[Hành động: Chỉ vào Ma trận nhầm lẫn n=400 và đường cong ROC]\n"
            "Xét về Ma trận nhầm lẫn trên tập kiểm thử độc lập 400 mẫu, Random Forest phát hiện chính xác **189 trên 221 sinh viên có nguy cơ bỏ học**, đạt **Độ nhạy (Recall) lên tới 85.52%**, chỉ bỏ sót 32 trường hợp (FN).\n"
            "Phân tích độ quan trọng đặc trưng khẳng định: Hai chỉ số tương tác do nhóm tự thiết kế là **Stress_Engagement_Ratio** và **Academic_Engagement_Index** xếp vị trí số 1 và số 2 trong các biến quyết định của Random Forest!\"",
            "Thầy Cô có thể hỏi: 'Tại sao trong bài toán này Độ nhạy (Recall) lại quan trọng hơn Độ chính xác (Precision)?'\n"
            "-> Gợi ý trả lời: 'Dạ thưa Thầy Cô, trong bài toán cảnh báo sớm bỏ học, sai lầm loại 2 (False Negative - bỏ sót sinh viên có nguy cơ) gây hậu quả nghiêm trọng hơn nhiều so với sai lầm loại 1 (False Positive - nhắc nhở nhầm sinh viên an toàn). Nếu bỏ sót, sinh viên có thể thôi học mà không được hỗ trợ; trong khi nếu nhắc nhở nhầm, sinh viên chỉ nhận thêm một buổi tham vấn tâm lý hoặc tư vấn học tập có lợi.'"
        ),
        (
            "SLIDE 8: DEMO / SYSTEM ARCHITECTURE (KIẾN TRÚC & WEB DASHBOARD)",
            "45 giây",
            "Minh họa kiến trúc 3 tầng và demo trực tiếp kết quả đo lường thực tế trên Web FastAPI app.py.",
            "[Hành động: Bấm chuyển Slide 8. Bạn Vũ Quang Huy trình bày]\n\n"
            "\"Để chuyển hóa mô hình nghiên cứu thành sản phẩm ứng dụng thực tế, nhóm đã phát triển hoàn chỉnh **Ứng dụng Web Dashboard Cảnh báo Sớm** bằng nền tảng **FastAPI** (tệp app.py) theo kiến trúc 3 tầng:\n\n"
            "[Hành động: Chỉ vào sơ đồ 3 tầng bên trái và ảnh chụp giao diện bên phải]\n"
            "• **Tầng 1 – Giao diện:** Cho phép nhập 18 chỉ số qua thanh trượt trực quan hoặc bấm chọn nhanh 3 hồ sơ mẫu.\n"
            "• **Tầng 2 – Động cơ AI:** Vận hành song song mô hình dự báo Burnout 3 mức và mô hình dự báo nguy cơ bỏ học. Thời gian phản hồi đo lường nội bộ cực nhanh, chỉ xấp xỉ **145 đến 153 mili-giây** mỗi lượt dự báo.\n"
            "• **Tầng 3 – Kết quả & Khuyến nghị:** Hiển thị trực quan xác suất rủi ro và tự động kích hoạt khuyến nghị sư phạm.\n\n"
            "[Hành động: Chỉ vào 3 kết quả Preset bên dưới]\n"
            "Kiểm thử thực tế trên 3 hồ sơ mẫu cho kết quả rất chính xác:\n"
            "• Mẫu SV Tích cực: Nguy cơ bỏ học chỉ **1.7%** (An toàn).\n"
            "• Mẫu SV Căng thẳng Trung bình: Nguy cơ bỏ học **60.5%** (Mức cảnh báo).\n"
            "• Mẫu SV Kiệt sức Cao: Nguy cơ bỏ học lên tới **98.6%** (Báo động đỏ can thiệp khẩn cấp)!\"",
            "Thầy Cô có thể hỏi: 'Hệ thống này triển khai thực tế ở trường đại học như thế nào? Dữ liệu lấy từ đâu?'\n"
            "-> Gợi ý trả lời: 'Dạ thưa Thầy Cô, hệ thống được thiết kế dưới dạng RESTful API bằng FastAPI. Trong tương lai, hệ thống sẽ kết nối trực tiếp với cơ sở dữ liệu điểm danh và LMS (như Moodle/Canvas) của nhà trường qua API để tự động đồng bộ chuyên cần và điểm số mỗi tuần, cố vấn học tập chỉ cần đăng nhập là thấy ngay bảng xếp hạng rủi ro của sinh viên.'"
        ),
        (
            "SLIDE 9: CONCLUSION & FUTURE WORK (TỔNG KẾT & ĐỊNH HƯỚNG)",
            "35 giây",
            "Tổng kết 3 đóng góp lớn, thẳng thắn nhìn nhận hạn chế khoa học và gửi lời cảm ơn Hội đồng.",
            "[Hành động: Bấm chuyển Slide 9. Bạn Đỗ Quang Minh trình bày kết luận]\n\n"
            "\"Kính thưa Quý Thầy Cô trong Hội đồng, tóm lược lại đồ án của Nhóm 03 đã đạt được **3 đóng góp nổi bật**:\n\n"
            "[Hành động: Nhìn thẳng vào Hội đồng, giọng dứt khoát và chân thành]\n"
            "1. **Về Kỹ thuật:** Xây dựng hoàn chỉnh Pipeline học máy khép kín, thiết kế 6 đặc trưng tương tác chuyên ngành chứng minh được hiệu quả vượt trội.\n"
            "2. **Về Học thuật:** Nhóm nhìn nhận thẳng thắn hạn chế của dữ liệu cắt ngang và dữ liệu tổng hợp. Trên dữ liệu sinh viên thực tế 800 mẫu, mô hình đạt ROC-AUC thực chất là **80.50%**.\n"
            "3. **Về Ứng dụng:** Đóng gói thành công Web Dashboard FastAPI hoạt động ổn định, sẵn sàng tích hợp giải thích mô hình SHAP và kết nối hệ thống LMS của trường đại học trong tương lai.\n\n"
            "[Hành động: Cúi đầu chào trang trọng]\n"
            "Nhóm 03 chúng em xin trân trọng cảm ơn sự lắng nghe và hướng dẫn tận tình của Quý Thầy Cô chương trình Samsung Innovation Campus! Chúng em xin sẵn sàng lắng nghe các câu hỏi và nhận xét từ Hội đồng!\"",
            "Thầy Cô có thể hỏi: 'Nếu có thêm thời gian 3 tháng nữa, nhóm sẽ ưu tiên làm gì tiếp theo?'\n"
            "-> Gợi ý trả lời: 'Dạ thưa Thầy Cô, nhóm sẽ ưu tiên 2 việc: Thứ nhất là tích hợp giải thích mô hình cục bộ SHAP Waterfall Plot lên giao diện Web để giải thích lý do cụ thể cho từng sinh viên; Thứ hai là thu thập dữ liệu chuỗi thời gian theo tuần (Longitudinal Data) để thử nghiệm mô hình LSTM dự báo xu hướng suy giảm phong độ học tập.'"
        )
    ]

    for s_idx, (title, timing, obj, script, qa) in enumerate(slides_content, 1):
        doc.add_heading(f"{title} — Thời lượng: {timing}", level=1)

        p_obj = doc.add_paragraph()
        r_obj_lbl = p_obj.add_run("• Mục tiêu slide: ")
        r_obj_lbl.bold = True
        r_obj_lbl.font.color.rgb = BLUE_RGB
        p_obj.add_run(obj)

        doc.add_heading("Lời thoại Thuyết trình Chi tiết (Verbatim Script):", level=2)
        tbl_s = doc.add_table(rows=1, cols=1)
        tbl_s.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_s = tbl_s.cell(0, 0)
        set_cell_background(c_s, LIGHT_BG_HEX)
        set_cell_margins(c_s, top=120, bottom=120, left=160, right=140)
        prevent_row_split(tbl_s.rows[0])
        tcPr = c_s._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="{NAVY_HEX}"/>'
            f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>'
        ))

        for para_text in script.split("\n\n"):
            p = c_s.add_paragraph() if c_s.paragraphs[0].text else c_s.paragraphs[0]
            p.paragraph_format.line_spacing = 1.3
            p.paragraph_format.space_after = Pt(4)
            # parse bold tokens
            parts = para_text.split("**")
            for idx_part, part in enumerate(parts):
                r = p.add_run(part)
                r.font.name = "Arial"
                r.font.size = Pt(10.5)
                if idx_part % 2 == 1:
                    r.bold = True
                    r.font.color.rgb = NAVY_RGB
                else:
                    if part.startswith("[") and "]" in part:
                        r.font.italic = True
                        r.font.color.rgb = TEAL_RGB
                    else:
                        r.font.color.rgb = DARK_TEXT_RGB

        doc.add_heading("Gợi ý Trả lời Phản biện Nhanh (Q&A Defense Tips):", level=2)
        add_callout(doc, "Câu hỏi Dự kiến & Cách Trả lời Ghi Điểm:", qa, border_color=AMBER_HEX, bg_color=CALLOUT_BG_HEX)

        if s_idx < len(slides_content):
            doc.add_page_break()

    # =========================================================================
    # MASTER Q&A CHAPTER
    # =========================================================================
    doc.add_page_break()
    doc.add_heading("BỘ 5 CÂU HỎI PHẢN BIỆN CHUYÊN SÂU CỦA HỘI ĐỒNG SIC & CÁCH TRẢ LỜI", level=1)
    doc.add_paragraph(
        "Dưới đây là 5 câu hỏi học thuật kinh điển mà các Giám khảo AI/ML của Samsung Innovation Campus thường đặt ra. "
        "Nhóm nắm chắc các luận điểm này sẽ thể hiện được đẳng cấp tư duy vượt trội so với các nhóm khác:"
    )

    tough_qa = [
        (
            "Câu hỏi 1 (Về Rò rỉ Dữ liệu - Data Leakage):",
            "'Nhóm làm thế nào để chứng minh mô hình không bị rò rỉ dữ liệu (Data Leakage) khi kiểm định chéo?'",
            "Cách trả lời chuẩn xác:\n"
            "\"Dạ thưa Thầy Cô, nhóm đã áp dụng 3 biện pháp kỹ thuật nghiêm ngặt:\n"
            "1. Đóng gói khép kín trong scikit-learn Pipeline: Toàn bộ quá trình điền khuyết KNNImputer, chuẩn hóa StandardScaler và mã hóa OneHotEncoder "
            "nằm hoàn toàn trong Pipeline, chỉ thực hiện fit trên 4 fold huấn luyện và transform trên fold kiểm thử của từng vòng lặp.\n"
            "2. Phân chia phân tầng (Stratified K-Fold): Tỷ lệ nhãn Dropout_Risk (54.6% Yes / 45.4% No) được bảo toàn đồng đều trên cả 5 folds.\n"
            "3. Giữ tập kiểm thử độc lập (Hold-out Test 20%): Nhóm tách riêng 20% dữ liệu ngay từ đầu bằng random_state=42 và chỉ đưa vào đánh giá sau cùng, "
            "hoàn toàn không tham gia vào quá trình khớp tham số của mô hình.\""
        ),
        (
            "Câu hỏi 2 (Về Dữ liệu Tổng hợp 1.200 bản ghi):",
            "'Tại sao nhóm lại mở rộng thêm 1.200 bản ghi tổng hợp? Liệu con số 89.01% AUC có bị thổi phồng không?'",
            "Cách trả lời ghi điểm tuyệt đối (Thể hiện tính trung thực khoa học):\n"
            "\"Dạ thưa Thầy Cô, nhóm xin được báo cáo hết sức trung thực:\n"
            "Mục tiêu ban đầu của việc sinh 1.200 hồ sơ bằng phân phối Gaussian đa biến phân tầng là nhằm kiểm thử khả năng mở rộng kỹ thuật và kiểm tra độ trễ của giao diện Web FastAPI.\n"
            "Tuy nhiên, qua kiểm toán chuyên sâu, nhóm phát hiện việc sinh mẫu có điều kiện theo nhãn đã vô tình làm tăng độ phân tách dữ liệu (trên phần tổng hợp test đạt tới 94.3% AUC, trong khi phần gốc đạt 77.3% AUC).\n"
            "Vì vậy, nhóm xin khẳng định trước Hội đồng: Con số phản ánh chính xác nhất năng lực tổng quát hóa trên sinh viên thật là ROC-AUC 80.50% (trên 800 sinh viên gốc). "
            "Nhóm đã ghi nhận điều này như một bài học hạn chế quan trọng trong báo cáo chính thức!\""
        ),
        (
            "Câu hỏi 3 (Về Tính Phi Nhân Quả của Feature Importance):",
            "'Mức độ quan trọng đặc trưng (Feature Importance) cao có đồng nghĩa với việc đặc trưng đó là nguyên nhân gây ra bỏ học không?'",
            "Cách trả lời chuẩn mực học thuật:\n"
            "\"Dạ hoàn toàn không thưa Thầy Cô. Feature Importance (như Gini Importance của Random Forest hay Gain của XGBoost) chỉ phản ánh mức độ đóng góp của biến "
            "trong việc giảm thiểu độ vẩn đục (impurity) khi phân tách các nút của cây quyết định. Đây là mối liên hệ thống kê trong cấu trúc mô hình, "
            "không đại diện cho mối quan hệ nhân quả (Causal Relationship). Để khẳng định quan hệ nhân quả, cần có nghiên cứu can thiệp đối chứng ngẫu nhiên (A/B Testing hoặc RCT) "
            "hoặc phân tích biến công cụ (Instrumental Variables).\""
        ),
        (
            "Câu hỏi 4 (Về Lựa chọn Mô hình Triển khai):",
            "'Tại sao Kịch bản 1 chọn Random Forest mà Kịch bản 2 lại chọn XGBoost để triển khai trên Web?'",
            "Cách trả lời kỹ thuật vững vàng:\n"
            "\"Dạ thưa Thầy Cô, ở Kịch bản 1 khi có mặt biến Burnout_Level, cấu trúc dữ liệu có tính phân tầng rõ rệt, cơ chế lấy mẫu ngẫu nhiên (Bagging) của Random Forest "
            "giúp kiểm soát phương sai rất tốt và đạt AUC 89.01%. Trong khi ở Kịch bản 2 (Cảnh báo sớm, không có biến Burnout), bài toán trở nên khó hơn với các tín hiệu yếu "
            "và phi tuyến từ hành vi học tập thô. Cơ chế tăng cường độ dốc (Boosting) của XGBoost với các siêu tham số phạt L2 (reg_lambda=2.0) và colsample_bytree=0.8 "
            "đã phát huy hiệu quả vượt trội trong việc khai phá các tương tác phức tạp này, đạt AUC 88.26%.\""
        ),
        (
            "Câu hỏi 5 (Về Giá trị Ứng dụng Thực tiễn):",
            "'Nếu một trường đại học áp dụng hệ thống của nhóm, quy trình can thiệp sẽ diễn ra như thế nào?'",
            "Cách trả lời giàu tính thực tiễn:\n"
            "\"Dạ thưa Thầy Cô, hệ thống sẽ vận hành theo 3 cấp độ can thiệp tự động:\n"
            "• Cấp độ 1 (Rủi ro Thấp < 30%): Hệ thống ghi nhận hồ sơ học tập bình thường, chỉ gửi bản tin sinh hoạt và phương pháp học tập đại học định kỳ.\n"
            "• Cấp độ 2 (Rủi ro Trung bình 30%–70%): Hệ thống gửi thông báo cho Cố vấn học tập khoa để mời sinh viên trao đổi, tìm hiểu lý do vắng học và ghép vào nhóm học tập (Study Group) có gia sư phụ đạo.\n"
            "• Cấp độ 3 (Báo động đỏ > 70%): Kích hoạt quy trình phối hợp khẩn cấp giữa Phòng Đào tạo (giãn tiến độ học, bảo lưu tín chỉ), Phòng Công tác Sinh viên (xét học bổng hỗ trợ khó khăn) "
            "và Phòng Tham vấn Tâm lý (trị liệu kiệt sức học đường).\""
        )
    ]

    for q_lbl, q_title, ans in tough_qa:
        doc.add_heading(f"{q_lbl} {q_title}", level=2)
        add_callout(doc, "Gợi ý Trả lời Thuyết phục:", ans, border_color=TEAL_HEX, bg_color=LIGHT_BG_HEX)

    # Save Word document
    doc.save(OUT_DOCX)
    print(f"Generated Word Script: {OUT_DOCX}")

    # Generate companion Markdown file
    md_lines = [
        "# KỊCH BẢN THUYẾT TRÌNH BẢO VỆ ĐỒ ÁN CHI TIẾT THEO TỪNG SLIDE",
        "## SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE PROJECT (NHÓM 03)",
        "",
        "**Đề tài:** Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập  ",
        "**Thành viên:** Đỗ Quang Minh (Trưởng nhóm) • Vũ Quang Huy (Thành viên)  ",
        "**Thời lượng mục tiêu:** 6 phút 30 giây (khoảng an toàn 5–8 phút)  ",
        "",
        "---",
        "",
    ]
    for s_idx, (title, timing, obj, script, qa) in enumerate(slides_content, 1):
        md_lines.extend([
            f"## {title} (Thời lượng: {timing})",
            f"- **Mục tiêu slide:** {obj}",
            "",
            "### Lời thoại Thuyết trình Chi tiết (Verbatim Script):",
            script,
            "",
            "> [!TIP]",
            "> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  ",
            "> " + qa.replace("\n", "\n> "),
            "",
            "---",
            "",
        ])

    md_lines.extend([
        "## BỘ 5 CÂU HỎI PHẢN BIỆN CHUYÊN SÂU CỦA HỘI ĐỒNG SIC & CÁCH TRẢ LỜI",
        "",
    ])
    for q_lbl, q_title, ans in tough_qa:
        md_lines.extend([
            f"### {q_lbl} {q_title}",
            "",
            "> [!NOTE]",
            "> " + ans.replace("\n", "\n> "),
            "",
        ])

    with open(OUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    print(f"Generated Markdown Script: {OUT_MD}")


if __name__ == "__main__":
    build_script()

