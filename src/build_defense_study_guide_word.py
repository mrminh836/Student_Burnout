"""
Generates a comprehensive, easy-to-study Word document (.docx) specifically designed
for Group 03 (Đỗ Quang Minh, Vũ Quang Huy) to study and prepare for their oral defense
and live demo in front of the faculty evaluation committee.

Output:
- Tai_Lieu_On_Tap_Va_Bao_Ve_Do_An_Nhom03.docx (in project root and SIC_Documents_VI/)
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
FIG_DIR = os.path.join(BASE_DIR, "figures")
OUT_ROOT = os.path.join(BASE_DIR, "Tai_Lieu_On_Tap_Va_Bao_Ve_Do_An_Nhom03.docx")
OUT_SIC_VI = os.path.join(BASE_DIR, "SIC_Documents_VI", "Tai_Lieu_On_Tap_Va_Bao_Ve_Do_An_Nhom03.docx")

NAVY_HEX = "1D3557"
SLATE_HEX = "264653"
TEAL_HEX = "2A9D8F"
AMBER_HEX = "E76F51"
LIGHT_BG_HEX = "F4F7FB"
CALLOUT_BG_HEX = "EDF6F5"
WARN_BG_HEX = "FFF8E7"
BORDER_GRAY_HEX = "CBD5E1"

NAVY_RGB = RGBColor(29, 53, 87)
SLATE_RGB = RGBColor(38, 70, 83)
TEAL_RGB = RGBColor(42, 157, 143)
RED_RGB = RGBColor(214, 40, 40)
DARK_TEXT_RGB = RGBColor(30, 41, 59)
MUTED_TEXT_RGB = RGBColor(100, 116, 139)


def set_cell_background(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


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


def configure_styles(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = DARK_TEXT_RGB
    normal.paragraph_format.line_spacing = 1.28
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for lvl, sz, col in [(1, 15.5, NAVY_RGB), (2, 13.5, SLATE_RGB), (3, 12.5, TEAL_RGB)]:
        h = doc.styles[f"Heading {lvl}"]
        h.font.name = "Times New Roman"
        h.font.size = Pt(sz)
        h.font.bold = True
        h.font.color.rgb = col
        h.paragraph_format.space_before = Pt(14 if lvl == 1 else 10)
        h.paragraph_format.space_after = Pt(5)
        h.paragraph_format.keep_with_next = True


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
    p.paragraph_format.left_indent = Cm(0.5)
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


def add_callout(doc, title_text, body_text, accent_hex=TEAL_HEX, bg_hex=CALLOUT_BG_HEX):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=120, bottom=120, left=200, right=160)
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
    p.paragraph_format.space_after = Pt(3)
    rt = p.add_run(title_text)
    rt.bold = True
    rt.font.name = "Times New Roman"
    rt.font.size = Pt(11.5)
    rt.font.color.rgb = NAVY_RGB

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p2.paragraph_format.space_after = Pt(2)
    rb = p2.add_run(body_text)
    rb.font.name = "Times New Roman"
    rb.font.size = Pt(11.5)
    rb.font.color.rgb = DARK_TEXT_RGB

    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)


def add_table(doc, caption, headers, rows, highlight_first=False):
    p_cap = doc.add_paragraph()
    p_cap.paragraph_format.space_before = Pt(8)
    p_cap.paragraph_format.space_after = Pt(3)
    p_cap.paragraph_format.keep_with_next = True
    rc = p_cap.add_run(caption)
    rc.bold = True
    rc.font.size = Pt(11)
    rc.font.color.rgb = NAVY_RGB

    tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)

    hdr = tbl.rows[0]
    repeat_header_row(hdr)
    prevent_row_split(hdr)
    for c_idx, h_text in enumerate(headers):
        cell = hdr.cells[c_idx]
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h_text)
        r.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, r_data in enumerate(rows):
        row = tbl.rows[r_idx + 1]
        prevent_row_split(row)
        bg = "E6F4F1" if (highlight_first and r_idx == 0) else ("FFFFFF" if r_idx % 2 == 0 else LIGHT_BG_HEX)
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=85, bottom=85, left=100, right=100)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx in [0, len(r_data) - 1] and len(str(val)) > 25 else WD_ALIGN_PARAGRAPH.CENTER
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(val))
            r.font.size = Pt(10.5)
            r.bold = (c_idx == 0 or (highlight_first and r_idx == 0))
            r.font.color.rgb = DARK_TEXT_RGB

    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)


def add_figure_with_speaking_guide(doc, fig_filename, caption, speaking_tip, width_in=5.8):
    fig_path = os.path.join(FIG_DIR, fig_filename)
    if not os.path.exists(fig_path):
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    p_img.add_run().add_picture(fig_path, width=Inches(width_in))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_after = Pt(4)
    rc = p_cap.add_run(caption)
    rc.bold = True
    rc.italic = True
    rc.font.size = Pt(10.5)
    rc.font.color.rgb = SLATE_RGB

    add_callout(doc, "🎤 CÁCH THUYẾT TRÌNH HÌNH NÀY TRƯỚC THẦY CÔ:", speaking_tip, accent_hex=AMBER_HEX, bg_hex=WARN_BG_HEX)


def build_study_guide():
    doc = Document()
    configure_styles(doc)

    # BANNER TITLE
    tbl_top = doc.add_table(rows=1, cols=1)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0 = tbl_top.cell(0, 0)
    set_cell_background(c0, NAVY_HEX)
    set_cell_margins(c0, top=180, bottom=180, left=200, right=200)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = p0.add_run(
        "CẨM NANG ÔN TẬP & BẢO VỆ ĐỒ ÁN TRƯỚC HỘI ĐỒNG (FULL SYSTEM GUIDE)\n"
    )
    r_main.bold = True
    r_main.font.size = Pt(15.5)
    r_main.font.color.rgb = RGBColor(255, 255, 255)

    r_sub = p0.add_run(
        "Đề tài: Dự đoán Tỷ lệ Bỏ học dựa trên các Chỉ số Căng thẳng và Sự Tham gia Học tập (Student Burnout)\n"
        "Thực hiện: Nhóm 03 (Đỗ Quang Minh – Vũ Quang Huy) | Samsung Innovation Campus (SIC AI)"
    )
    r_sub.font.size = Pt(11.5)
    r_sub.font.color.rgb = RGBColor(216, 243, 220)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ==========================================================================
    # PHẦN 1: TỔNG QUAN & CÁC CON SỐ VÀNG CẦN NHỚ THUỘC LÒNG
    # ==========================================================================
    add_h1(doc, "PHẦN 1: TỔNG QUAN ĐỀ TÀI & BẢNG CON SỐ 'VÀNG' CẦN NHỚ")

    add_callout(
        doc,
        "🎤 LỜI MỞ ĐẦU MẪU KHI LÊN BÁO CÁO (ĐỌC TRONG 45 GIÂY ĐẦU):",
        "“Em xin kính chào thầy/cô trong hội đồng chấm đồ án ạ. Chúng em là Nhóm 03 gồm Đỗ Quang Minh và Vũ Quang Huy. "
        "Hôm nay nhóm em xin trình bày đề tài: Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán Dự đoán Tỷ lệ Bỏ học "
        "dựa trên các chỉ số về Căng thẳng tâm lý và Sự tham gia học tập (Student Burnout). "
        "Mục tiêu của hệ thống là giúp nhà trường phát hiện sớm những sinh viên có nguy cơ kiệt sức và bỏ học ngay từ trong học kỳ "
        "thông qua cả quá trình phân tích dữ liệu, huấn luyện 4 mô hình học máy tiêu biểu và triển khai thành ứng dụng Web cảnh báo sớm thời gian thực.”",
    )

    golden_rows = [
        ["1. Bộ dữ liệu gốc (Raw Data)", "data/student_burnout_dropout_dataset_2.csv", "800 sinh viên × 25 cột (có 18/25 cột bị khuyết tổng cộng 612 ô NaN). Đã giữ nguyên để đối chứng."],
        ["2. Bộ dữ liệu mở rộng & làm sạch", "data/student_burnout_dropout_expanded.csv", "2,000 sinh viên × 25 cột (gồm 800 SV gốc đã làm sạch + 1,200 SV mở rộng). Sạch 100% (0 giá trị NaN)."],
        ["3. Bộ dữ liệu đầy đủ 34 cột", "data/student_burnout_dropout_expanded_cleaned.csv", "2,000 sinh viên × 34 cột (25 cột chuẩn + 3 cột mã hóa số + 6 chỉ số đặc trưng mới)."],
        ["4. Phân chia Train / Test", "80% Train – 20% Test + 5-Fold CV", "1,600 sinh viên tập Huấn luyện (Train) và 400 sinh viên tập Kiểm thử (Test), kiểm định chéo 5-Fold."],
        ["5. 4 Mô hình so sánh chính", "src/train_and_compare_models.py", "1. Logistic Regression\n2. Random Forest (300 cây)\n3. XGBoost (200 cây)\n4. Soft Voting Ensemble (kết hợp LR + RF + XGBoost)"],
        ["6. Mô hình chiến thắng (Top 1)", "Random Forest (Kịch bản 1)\nXGBoost / Soft Voting (Kịch bản 2)", "• Kịch bản 1 (Đầy đủ): Random Forest đạt CV ROC-AUC = 89.01%, Test Accuracy = 81.50%, Test F1 = 83.63%.\n• Kịch bản 2 (Cảnh báo sớm): XGBoost & Soft Voting đạt CV ROC-AUC = 88.26%."],
        ["7. Ứng dụng Web Demo", "app.py (FastAPI + Dashboard UI)", "Chạy bằng lệnh: python app.py -> Mở trình duyệt tại địa chỉ: http://localhost:8000"],
    ]
    add_table(
        doc,
        "Bảng 1: Tóm tắt 7 thông số cốt lõi của hệ thống (Chỉ cần nhớ bảng này là nắm trọn toàn bộ đồ án)",
        ["Hạng mục", "Tên File / Cấu hình", "Con số & Ý nghĩa cốt lõi cần nhớ"],
        golden_rows,
    )

    # ==========================================================================
    # PHẦN 2: GIẢI THÍCH DỮ LIỆU & 6 BƯỚC LÀM SẠCH DỮ LIỆU
    # ==========================================================================
    add_h1(doc, "PHẦN 2: GIẢI THÍCH BỘ DỮ LIỆU & QUY TRÌNH LÀM SẠCH 6 BƯỚC")

    add_h2(doc, "2.1. Cấu trúc 25 thuộc tính đầu vào (Chia làm 5 nhóm dễ nhớ)")
    doc.add_paragraph(
        "Nếu thầy cô hỏi: 'Dữ liệu đầu vào của em gồm những thông tin gì?', bạn không cần đọc thuộc lòng từng tên cột tiếng Anh mà chỉ cần chia thành 5 nhóm logic sau:"
    )
    add_bullet(doc, "Nhóm 1 – Nhân khẩu học & Bối cảnh (7 cột): ", "Mã SV (Student_ID), Tuổi (Age: 17–25), Giới tính (Gender), Năm học (Year_of_Study: năm 1 đến năm 4), Khoa/Ngành học (Department: 6 ngành), Loại cư trú (Residence_Type: Ký túc xá, Ở trọ, Ở cùng gia đình), Thu nhập gia đình (Family_Income_Bracket: 5 mức từ Low đến High).")
    add_bullet(doc, "Nhóm 2 – Sự tham gia học tập / Academic Engagement (5 cột): ", "Tỷ lệ chuyên cần đi học (Attendance_Percent: 35%–100%), Giờ tự học mỗi ngày (Study_Hours_Per_Day), Điểm trung bình kỳ trước (Previous_GPA: thang 10), Số môn nợ học phần (Backlogs: 0–6 môn), Điểm động lực học tập (Motivation_Score: 0–10).")
    add_bullet(doc, "Nhóm 3 – Lối sống & Sinh hoạt (5 cột): ", "Số giờ ngủ/ngày (Sleep_Hours), Thời gian dùng màn hình điện thoại/máy tính (Screen_Time_Hours), Số buổi tập thể dục/tuần (Exercise_Freq_Per_Week), Điểm hoạt động xã hội (Social_Activity_Score), Đi làm thêm (Part_Time_Job: Yes/No).")
    add_bullet(doc, "Nhóm 4 – Chỉ số Căng thẳng & Hỗ trợ tâm lý (6 cột): ", "Mức độ căng thẳng (Stress_Level: 0–10), Điểm lo âu (Anxiety_Score: 0–10), Áp lực tài chính (Financial_Stress_Score: 0–10), Áp lực đồng trang lứa (Peer_Pressure_Score: 0–10), Sự hỗ trợ từ gia đình (Family_Support_Score: 0–10), Tiếp cận phòng tư vấn tâm lý (Counseling_Access: Yes/No).")
    add_bullet(doc, "Nhóm 5 – Trạng thái Kiệt sức & Nhãn mục tiêu (2 cột): ", "Mức độ kiệt sức học đường (Burnout_Level: Low, Medium, High) và Nguy cơ bỏ học (Dropout_Risk: Yes = Có nguy cơ bỏ học [54.6%], No = An toàn [45.4%]).")

    add_h2(doc, "2.2. Quy trình Làm sạch & Tiền xử lý dữ liệu đủ 6 Bước")
    doc.add_paragraph(
        "Trong file dữ liệu gốc 800 dòng (student_burnout_dropout_dataset_2.csv), có 18/25 cột bị khuyết dữ liệu (NaN) với tỷ lệ từ 1.62% đến 6.38% (tổng cộng 612 ô bị trống). Nhóm đã thực hiện làm sạch qua 6 bước chuẩn mực trong file src/generate_expanded_dataset.py và Pipeline trong src/train_and_compare_models.py:"
    )

    clean_steps = [
        [
            "Bước 1:\nKiểm tra trùng lặp & Loại bỏ cột thừa",
            "drop_duplicates()\ndrop('Student_ID')",
            "• Kiểm tra tính duy nhất của mã sinh viên (Student_ID).\n• Loại bỏ cột Student_ID trước khi đưa vào mô hình học máy vì mã số định danh không mang ý nghĩa dự báo (tránh mô hình học vẹt theo số thứ tự).",
        ],
        [
            "Bước 2:\nXử lý dữ liệu khuyết (Missing Values Imputation)",
            "KNNImputer(n_neighbors=7)\n&\nSimpleImputer('most_frequent')",
            "• Với 13 cột số bị khuyết: Dùng thuật toán KNNImputer (k=7) tìm 7 sinh viên có hồ sơ giống nhất để nội suy giá trị bị thiếu (chính xác hơn nhiều so với điền trung bình cộng).\n• Với 5 cột phân loại (chữ) bị khuyết: Điền bằng nhãn phổ biến nhất (Mode).\n• Lý do không xóa dòng (dropna): Nếu xóa mọi dòng có ít nhất 1 ô trống thì sẽ mất hơn 40% dữ liệu!",
        ],
        [
            "Bước 3:\nXử lý ngoại lai (Outliers) & Chuẩn hóa kiểu",
            "Phương pháp IQR\n(Q1 - 1.5*IQR,\nQ3 + 1.5*IQR)",
            "• Dùng khoảng tứ phân vị IQR kết hợp miền giá trị thực tế để cắt (clip) các giá trị dị biệt (ví dụ GPA phải trong [2.5, 10.0], Chuyên cần trong [35%, 100%]).\n• Chuẩn hóa kiểu số nguyên (int) cho Age, Year_of_Study, Exercise_Freq và làm tròn chuẩn xác.",
        ],
        [
            "Bước 4:\nMã hóa biến phân loại (Categorical Encoding)",
            "Ordinal Mapping\n&\nOneHotEncoder",
            "• Mã hóa thứ tự (Ordinal) cho biến có cấp bậc: Family_Income_Bracket (Low=0 -> High=4) và Burnout_Level (Low=0, Medium=1, High=2).\n• Mã hóa nhị phân One-Hot Encoding cho các biến danh định không thứ tự: Gender, Department, Residence_Type, Part_Time_Job, Counseling_Access.",
        ],
        [
            "Bước 5:\nTạo 6 đặc trưng chuyên ngành mới (Feature Eng.)",
            "Class DomainFeatureEngineer\n(6 chỉ số tổng hợp)",
            "Tạo 6 cột chỉ số mới phản ánh sự tương tác giữa Căng thẳng và Học tập:\n1. Academic_Engagement_Index (Chỉ số gắn kết học tập)\n2. Psychological_Stress_Index (Chỉ số căng thẳng tâm lý)\n3. Stress_Engagement_Ratio (Tỷ số Căng thẳng / Gắn kết)\n4. Net_Stress_Support_Balance (Căng thẳng ròng sau khi trừ hỗ trợ GĐ)\n5. Digital_Fatigue_Ratio (Mệt mỏi số = Giờ màn hình / Giờ ngủ)\n6. Academic_Risk_Burden (Gánh nặng học vụ từ môn nợ & vắng học)",
        ],
        [
            "Bước 6:\nChuẩn hóa thang đo & Chia tập dữ liệu",
            "StandardScaler()\n&\nStratified 80/20 Split",
            "• Đưa mọi biến số về phân phối chuẩn Z-score (trung bình = 0, độ lệch chuẩn = 1) bằng StandardScaler.\n• Chia dữ liệu: 80% Train (1,600 SV) và 20% Test (400 SV) kết hợp kiểm định chéo 5-Fold Stratified CV.\n• Toàn bộ nằm trong scikit-learn Pipeline để chống rò rỉ dữ liệu (No Data Leakage).",
        ],
    ]
    add_table(
        doc,
        "Bảng 2: Chi tiết 6 Bước Làm sạch & Tiền xử lý dữ liệu (Câu trả lời trọng tâm khi thầy cô hỏi về Data)",
        ["Bước thực hiện", "Công cụ / Hàm sử dụng", "Giải thích chi tiết cách làm & Lý do chọn"],
        clean_steps,
    )

    add_h2(doc, "2.3. Cách mở rộng dữ liệu từ 800 lên 2,000 sinh viên")
    doc.add_paragraph(
        "Trong file src/generate_expanded_dataset.py, sau khi làm sạch 800 bản ghi gốc, nhóm giữ nguyên 800 bản ghi này và sinh thêm 1,200 bản ghi mới (STU00801 đến STU02000) bằng cách ước lượng phân phối chuẩn đa biến (Multivariate Gaussian Distribution: vector kỳ vọng và ma trận hiệp phương sai) trên từng nhóm phân tầng (Burnout_Level × Dropout_Risk) của dữ liệu gốc. Nhờ đó, bộ dữ liệu mở rộng 2,000 sinh viên sạch 100% (0 NaN), giữ nguyên tỷ lệ nhãn (54.6% Yes / 45.4% No) và giúp các mô hình học máy hội tụ ổn định, tăng độ chính xác ROC-AUC từ ~80.5% lên 89.01%."
    )

    doc.add_page_break()

    # ==========================================================================
    # PHẦN 3: GIẢI THÍCH 4 MÔ HÌNH THUẬT TOÁN TRONG ĐỒ ÁN
    # ==========================================================================
    add_h1(doc, "PHẦN 3: GIẢI THÍCH 4 MÔ HÌNH THUẬT TOÁN (DỄ HIỂU – DỄ TRẢ LỜI)")

    doc.add_paragraph(
        "Trong file src/train_and_compare_models.py (hàm get_models), nhóm lựa chọn 4 mô hình đại diện cho 4 trường phái thuật toán kinh điển và mạnh nhất đối với dữ liệu dạng bảng (Tabular Data):"
    )

    model_explanations = [
        [
            "1. Logistic Regression\n(Hồi quy Logistic)",
            "Mô hình Tuyến tính Cơ sở (Baseline)",
            "• Nguyên lý: Tính tổng có trọng số các đặc trưng đầu vào rồi đưa qua hàm Sigmoid để nén giá trị về khoảng xác suất (0, 1).\n• Vai trò trong đồ án: Làm mốc chuẩn (Baseline) để chứng minh các mô hình phi tuyến phức tạp hơn (như Random Forest, XGBoost) thực sự mang lại hiệu quả vượt trội.",
        ],
        [
            "2. Random Forest\n(Rừng Ngẫu nhiên – 300 cây)",
            "Cây Tập hợp Song song (Bagging Ensemble)",
            "• Nguyên lý: Xây dựng 300 cây quyết định (Decision Trees) độc lập nhau. Mỗi cây chỉ được học trên một mẫu con ngẫu nhiên của dữ liệu và một tập con ngẫu nhiên các cột đặc trưng (max_depth=8). Khi dự đoán, cả 300 cây cùng bỏ phiếu để lấy xác suất trung bình.\n• Ưu điểm: Chống học vẹt (Overfitting) cực tốt, bắt được các ngưỡng tương tác phi tuyến (ví dụ: Chuyên cần < 70% kết hợp Stress > 6.5).\n• Kết quả: Đạt Top 1 ở Kịch bản 1 (CV ROC-AUC = 89.01%, Test Accuracy = 81.50%, Test F1 = 83.63%).",
        ],
        [
            "3. XGBoost\n(Extreme Gradient Boosting – 200 cây)",
            "Cây Tăng cường Nối tiếp (Boosting Ensemble)",
            "• Nguyên lý: Khác với Random Forest xây các cây độc lập song song, XGBoost xây 200 cây quyết định nối tiếp nhau. Cây thứ hai được huấn luyện để sửa sai số dư (residual errors) của cây thứ nhất, cây thứ ba sửa sai cho cây thứ hai... kết hợp điều chuẩn L2 (reg_lambda=2.0) để tránh quá khớp.\n• Kết quả: Đạt Top 1 ở Kịch bản 2 Cảnh báo sớm (CV ROC-AUC = 88.26%, CV F1 = 82.40%).",
        ],
        [
            "4. Soft Voting Ensemble\n(Tổ hợp Bỏ phiếu Mềm)",
            "Mô hình Hội đồng Kết hợp Đa thuật toán",
            "• Nguyên lý: Kết hợp sức mạnh của cả 3 thuật toán trên (Logistic Regression + Random Forest + XGBoost) theo tỷ lệ trọng số [1, 3, 2]. Thay vì đếm phiếu cứng (Yes/No), Soft Voting lấy trung bình có trọng số các con số phần trăm xác suất của cả 3 mô hình.\n• Ưu điểm: Bù trừ điểm yếu cho nhau (kết hợp cả góc nhìn tuyến tính, Bagging và Boosting), cho xác suất dự báo rất mượt và ổn định ở cả 2 kịch bản (ROC-AUC ~88.89%).",
        ],
    ]
    add_table(
        doc,
        "Bảng 3: Nguyên lý hoạt động & Vai trò của 4 Mô hình Học máy trong Đồ án",
        ["Tên Mô hình", "Trường phái Thuật toán", "Cách giải thích ngắn gọn trước Hội đồng"],
        model_explanations,
    )

    # ==========================================================================
    # PHẦN 4: 2 KỊCH BẢN THỰC NGHIỆM & CÁCH ĐỌC CÁC BIỂU ĐỒ KẾT QUẢ
    # ==========================================================================
    add_h1(doc, "PHẦN 4: KẾT QUẢ THỰC NGHIỆM 2 KỊCH BẢN & HƯỚNG DẪN ĐỌC BIỂU ĐỒ")

    add_h2(doc, "4.1. Tại sao lại chia làm 2 Kịch bản thực nghiệm?")
    add_bullet(
        doc,
        "Kịch bản 1 – Đầy đủ đặc trưng (Full Features - Có cột Burnout_Level): ",
        "Giả định nhà trường đã tổ chức khảo sát tâm lý định kỳ và biết sinh viên đang ở mức Burnout nào (Low, Medium, High). Mục tiêu là đạt độ chính xác chẩn đoán cao nhất. Mô hình đứng đầu là Random Forest (CV ROC-AUC = 89.01%, Test Accuracy = 81.50%, Test F1 = 83.63%).",
    )
    add_bullet(
        doc,
        "Kịch bản 2 – Cảnh báo sớm (Early Warning - Bỏ cột Burnout_Level): ",
        "Giả định ngay đầu học kỳ, nhà trường chưa kịp khảo sát Burnout mà chỉ có dữ liệu học tập và lối sống thô (Điểm danh, Giờ học, GPA, Stress...). Khi bỏ cột Burnout_Level đi, mô hình XGBoost và Soft Voting Ensemble vẫn dự báo chính xác nguy cơ bỏ học với CV ROC-AUC = 88.26%.",
    )

    s1_table = [
        ["1. Random Forest (Lưu .pkl Kịch bản 1)", "80.25%", "82.98%", "89.01%", "81.50%", "83.63%", "88.42%"],
        ["2. Soft Voting Ensemble (LR+RF+XGB)", "79.70%", "82.27%", "88.89%", "80.25%", "82.17%", "88.32%"],
        ["3. XGBoost", "79.95%", "82.39%", "88.64%", "78.50%", "80.28%", "87.99%"],
        ["4. Logistic Regression (Baseline)", "78.45%", "80.62%", "87.61%", "78.25%", "80.18%", "87.60%"],
    ]
    add_table(
        doc,
        "Bảng 4: Kết quả thực nghiệm Kịch bản 1 (Đầy đủ đặc trưng – Tập dữ liệu 2,000 sinh viên)",
        ["Mô hình Học máy", "CV Acc", "CV F1", "CV ROC-AUC", "Test Acc", "Test F1", "Test ROC-AUC"],
        s1_table,
        highlight_first=True,
    )

    s2_table = [
        ["1. XGBoost (Lưu .pkl Kịch bản 2)", "79.95%", "82.40%", "88.26%", "78.00%", "79.82%", "87.24%"],
        ["2. Soft Voting Ensemble (LR+RF+XGB)", "79.30%", "81.77%", "88.25%", "78.00%", "80.00%", "87.39%"],
        ["3. Random Forest", "79.30%", "82.00%", "87.97%", "78.50%", "80.72%", "87.13%"],
        ["4. Logistic Regression (Baseline)", "78.00%", "80.27%", "87.25%", "77.75%", "79.73%", "87.00%"],
    ]
    add_table(
        doc,
        "Bảng 5: Kết quả thực nghiệm Kịch bản 2 (Cảnh báo sớm không dùng biến Burnout_Level)",
        ["Mô hình Học máy", "CV Acc", "CV F1", "CV ROC-AUC", "Test Acc", "Test F1", "Test ROC-AUC"],
        s2_table,
        highlight_first=True,
    )

    add_h2(doc, "4.2. Các biểu đồ quan trọng nhất & Cách đọc trước Hội đồng")

    add_figure_with_speaking_guide(
        doc,
        "02_target_and_burnout_distribution.png",
        "Hình 1: Phân phối biến mục tiêu Dropout_Risk và Tỷ lệ bỏ học theo mức độ Burnout_Level",
        "“Thưa thầy cô, nhìn vào biểu đồ bên phải, chúng em phát hiện ra quy luật cực kỳ rõ nét: ở nhóm sinh viên Kiệt sức Thấp (Low Burnout), tỷ lệ bỏ học chỉ là 23.8%; sang nhóm Trung bình (Medium), tỷ lệ bỏ học tăng lên 63.2%; và đặc biệt ở nhóm Kiệt sức Cao (High Burnout), tỷ lệ bỏ học vọt lên tới 92.0%! Điều này chứng minh Burnout chính là nguyên nhân trực tiếp dẫn đến bỏ học.”",
    )

    add_figure_with_speaking_guide(
        doc,
        "07_model_performance_comparison.png",
        "Hình 2: So sánh hiệu năng 4 Mô hình Học máy Tiêu biểu trên 2 Kịch bản",
        "“Trên biểu đồ so sánh 4 mô hình bằng kiểm định chéo 5-Fold, cả 3 mô hình tập hợp cây (Random Forest, Soft Voting Ensemble, XGBoost) đều vượt trội hơn hẳn mô hình tuyến tính cơ sở Logistic Regression. Ở Kịch bản 1, Random Forest dẫn đầu với ROC-AUC đạt 89.0% và F1-Score đạt 83.0%. Ở Kịch bản 2 khi chưa có thông tin Burnout, XGBoost và Soft Voting vẫn duy trì ROC-AUC đạt 88.3%.”",
    )

    add_figure_with_speaking_guide(
        doc,
        "09_confusion_matrices_top_models.png",
        "Hình 3: Ma trận Nhầm lẫn (Confusion Matrix) của 4 Mô hình trên Tập Kiểm thử độc lập (n=400)",
        "“Trên tập kiểm thử độc lập gồm 400 sinh viên (20% dữ liệu không dùng để train), mô hình Random Forest đạt độ chính xác Test Accuracy 81.5% và F1-Score 83.6%. Đặc biệt, ô dưới cùng bên phải (True Positive) cho thấy mô hình phát hiện đúng 86.2% số sinh viên thực sự có nguy cơ bỏ học, giúp nhà trường không bỏ sót các em cần hỗ trợ.”",
    )

    add_figure_with_speaking_guide(
        doc,
        "10_feature_importance_analysis.png",
        "Hình 4: Top 15 Đặc trưng Quan trọng nhất (Feature Importance) trên Random Forest và XGBoost",
        "“Biểu đồ Feature Importance chứng minh hiệu quả của bước Tạo đặc trưng (Feature Engineering): bên cạnh biến Burnout_Ordinal, các chỉ số tổng hợp do nhóm tự thiết kế như Stress_Engagement_Ratio (Tỷ số Căng thẳng / Gắn kết), Psychological_Stress_Index, Academic_Engagement_Index và Attendance_Percent đều nằm trong Top 5 yếu tố quyết định lớn nhất đến nguy cơ bỏ học.”",
    )

    doc.add_page_break()

    # ==========================================================================
    # PHẦN 5: CƠ CHẾ HOẠT ĐỘNG WEB APP (app.py) & KỊCH BẢN DEMO TRỰC TIẾP
    # ==========================================================================
    add_h1(doc, "PHẦN 5: CƠ CHẾ HOẠT ĐỘNG CỦA WEB APP & KỊCH BẢN DEMO")

    add_h2(doc, "5.1. Tại sao trong thư mục models/ có 3 file .pkl và chúng phối hợp thế nào?")
    doc.add_paragraph(
        "Khi bạn chạy lệnh python app.py, ứng dụng Web FastAPI sẽ nạp cả 3 file mô hình trong thư mục models/ để phục vụ 2 nhiệm vụ trên giao diện Dashboard:"
    )
    add_bullet(
        doc,
        "File 1 – best_burnout_model.pkl (Thuật toán Soft Voting Ensemble 3 lớp): ",
        "Nhận các chỉ số học tập và căng thẳng của sinh viên để dự đoán xác suất sinh viên rơi vào 3 mức độ kiệt sức: Low (Thấp), Medium (Trung bình), High (Cao) và vẽ 3 thanh phần trăm màu Xanh / Cam / Đỏ trên giao diện Web.",
    )
    add_bullet(
        doc,
        "File 2 – best_dropout_model_full.pkl (Thuật toán Random Forest – Top 1 Kịch bản 1): ",
        "Nhận toàn bộ thông tin sinh viên CỘNG THÊM trạng thái Burnout_Level để tính xác suất bỏ học chính xác nhất (prob_full).",
    )
    add_bullet(
        doc,
        "File 3 – best_dropout_model_early_warning.pkl (Thuật toán XGBoost – Top 1 Kịch bản 2): ",
        "Tính xác suất bỏ học trực tiếp từ các chỉ số Stress & Học tập thô mà không cần biến Burnout_Level (prob_early).",
    )
    add_bullet(
        doc,
        "Công thức kết hợp thông minh trên Web (Dòng 92 file app.py): ",
        "Nếu người dùng chọn sẵn mức Burnout (Low/Medium/High), Web lấy 100% kết quả từ File 2 (Random Forest). Nếu người dùng để chế độ 'Auto' (AI tự đoán Burnout), Web kết hợp cả hai theo công thức: Xác suất bỏ học = 0.65 * prob_full + 0.35 * prob_early để kết quả ổn định nhất.",
    )

    add_figure_with_speaking_guide(
        doc,
        "12_system_architecture_diagram.png",
        "Hình 5: Kiến trúc 3 tầng của Hệ thống Cảnh báo Sớm Nguy cơ Bỏ học (Web UI + FastAPI AI Backend)",
        "“Hệ thống Web của nhóm được thiết kế theo kiến trúc 3 tầng: Tầng Giao diện thu thập các chỉ số học tập và tâm lý; Tầng Backend FastAPI chạy Pipeline tiền xử lý, tính toán các chỉ số tổng hợp và gọi các mô hình AI; Tầng Kết quả trả về xác suất bỏ học, phân bố mức độ Burnout và danh sách khuyến nghị can thiệp cá nhân hóa cho cố vấn học tập.”",
    )

    add_h2(doc, "5.2. Các bước thao tác Demo Web trực tiếp trước Thầy Cô (Mất 1 phút)")
    demo_steps = [
        ["Bước 1: Mở Terminal chạy Web", "python app.py", "Terminal hiện thông báo: Uvicorn running on http://127.0.0.1:8000"],
        ["Bước 2: Mở trình duyệt", "http://localhost:8000", "Giao diện Dashboard hiện ra với 18 thanh trượt/menu bên trái và bảng kết quả AI bên phải."],
        ["Bước 3: Bấm nút Xanh lá (Mẫu SV Tích cực)", "Click 'Mẫu SV Tích cực (Nguy cơ Thấp)'", "Chuyên cần 93%, Giờ học 4.8h, Stress 3.1 -> AI báo màu Xanh: Nguy cơ bỏ học rất thấp (~13%), trạng thái An toàn."],
        ["Bước 4: Bấm nút Đỏ (Mẫu SV Kiệt sức)", "Click 'Mẫu SV Kiệt sức (Nguy cơ Cao)'", "Chuyên cần tụt xuống 52%, Nợ 3 môn, Stress 8.6, Ngủ 4.5h -> AI đổi sang cảnh báo Đỏ: Nguy cơ bỏ học >90%, kèm 5 khuyến nghị can thiệp khẩn cấp."],
        ["Bước 5: Đổi sang chế độ Auto & Kéo thử thanh trượt", "Chọn Burnout_Level = 'AI Tự động Dự báo (Auto)' rồi kéo thanh Chuyên cần / Stress", "Bấm nút 'Phân tích & Dự đoán' để chứng minh cho thầy cô thấy mô hình chạy tính toán thật 100% theo thời gian thực."],
    ]
    add_table(
        doc,
        "Bảng 6: Kịch bản 5 bước Demo Web trực tiếp trên máy tính",
        ["Thao tác", "Lệnh / Nút bấm", "Kết quả hiển thị để chỉ cho Thầy Cô xem"],
        demo_steps,
    )

    # ==========================================================================
    # PHẦN 6: BẢN ĐỒ MÃ NGUỒN (THẦY CÔ HỎI ĐOẠN NÀO – MỞ ĐÚNG FILE ĐÓ)
    # ==========================================================================
    add_h1(doc, "PHẦN 6: BẢN ĐỒ CODE (NẾU THẦY CÔ YÊU CẦU MỞ CODE GIẢI THÍCH)")

    code_map = [
        [
            "src/generate_expanded_dataset.py",
            "Làm sạch 6 bước & Mở rộng dữ liệu lên 2,000 SV",
            "• Hàm clean_base_25_columns() (dòng 38): Điền khuyết KNNImputer(k=7), SimpleImputer và cắt ngoại lai bằng IQR.\n• Hàm generate_clean_augmented_records() (dòng 77): Sinh thêm 1,200 sinh viên sạch từ phân phối chuẩn đa biến.\n• Hàm add_engineered_features_and_encodings() (dòng 174): Tạo 6 chỉ số mới.",
        ],
        [
            "src/eda_analysis.py",
            "Phân tích Khám phá Dữ liệu (EDA) & Vẽ 6 biểu đồ đầu",
            "• Chứa 6 hàm vẽ biểu đồ từ Hình 01 đến Hình 06: Phân tích dữ liệu khuyết, Phân phối Burnout vs Dropout, Violin/Boxplot, Ma trận tương quan Pearson, Biến phân loại và Scatter đa biến.",
        ],
        [
            "src/train_and_compare_models.py",
            "Tạo đặc trưng, Pipeline & Huấn luyện 4 mô hình chính",
            "• Class DomainFeatureEngineer (dòng 65–150): Tính 6 chỉ số tổng hợp bên trong Pipeline.\n• Hàm build_preprocessor() (dòng 153–177): Đóng gói KNNImputer + StandardScaler + OneHotEncoder.\n• Hàm get_models() (dòng 180–220): Khai báo 4 mô hình (Logistic Regression, Random Forest, XGBoost, Soft Voting Ensemble).\n• Hàm evaluate_models_scenario() (dòng 223–300): Chia Train/Test 80/20, chạy 5-Fold CV và tính Accuracy, Precision, Recall, F1, ROC-AUC.",
        ],
        [
            "src/train_burnout_and_diagrams.py",
            "Huấn luyện mô hình Burnout 3 lớp & Vẽ sơ đồ hệ thống",
            "• Hàm train_burnout_model() (dòng 34): Huấn luyện mô hình phân loại 3 mức độ Burnout (Low/Medium/High) lưu vào models/best_burnout_model.pkl.\n• Hàm draw_workflow_and_system_diagrams() (dòng 68): Vẽ sơ đồ quy trình Hình 11 và kiến trúc Hình 12.",
        ],
        [
            "app.py",
            "Ứng dụng Web FastAPI & Giao diện Dashboard",
            "• Dòng 31–33: Load 3 file mô hình .pkl.\n• Hàm predict_student() (dòng 62–133): API POST /api/predict nhận dữ liệu JSON từ giao diện, chạy mô hình dự đoán xác suất bỏ học, mức độ Burnout và sinh danh sách khuyến nghị.\n• Biến HTML_PAGE (dòng 136–482): Giao diện HTML/CSS/JS.",
        ],
    ]
    add_table(
        doc,
        "Bảng 7: Bản đồ vị trí các hàm quan trọng trong Code (Dùng khi vấn đáp trực tiếp trên IDE)",
        ["Tên File Code", "Chức năng chính", "Các hàm & Dòng code quan trọng cần chỉ cho Thầy Cô"],
        code_map,
    )

    doc.add_page_break()

    # ==========================================================================
    # PHẦN 7: 12 CÂU HỎI PHẢN BIỆN HỘI ĐỒNG THƯỜNG HỎI NHẤT & ĐÁP ÁN MẪU
    # ==========================================================================
    add_h1(doc, "PHẦN 7: BỘ 12 CÂU HỎI PHẢN BIỆN HỘI ĐỒNG & ĐÁP ÁN MẪU ĐIỂM A+")

    qas = [
        (
            "Câu 1: Tại sao khi xử lý dữ liệu bị khuyết (NaN), nhóm không dùng lệnh dropna() xóa dòng hoặc điền bằng trung bình cộng (Mean), mà lại dùng KNNImputer(n_neighbors=7)?",
            "Trả lời: Thưa thầy/cô, trong dữ liệu gốc có tới 18/25 cột chứa giá trị khuyết rải rác ở các dòng khác nhau. Nếu dùng dropna() xóa dòng thì sẽ mất hơn 40% số lượng sinh viên. Còn nếu chỉ điền bằng trung bình cộng (Mean) của toàn trường thì sẽ làm phẳng biến thiên dữ liệu. Vì vậy nhóm dùng KNNImputer(k=7): khi một sinh viên bị thiếu điểm Stress hoặc Chuyên cần, thuật toán sẽ tìm 7 sinh viên có hồ sơ tương đồng nhất để nội suy giá trị, giúp bảo toàn mối tương quan tự nhiên giữa các biến.",
        ),
        (
            "Câu 2: Làm thế nào nhóm đảm bảo không bị Rò rỉ dữ liệu (Data Leakage) trong quá trình tiền xử lý và kiểm định chéo 5-Fold?",
            "Trả lời: Nhóm đóng gói toàn bộ quá trình tạo đặc trưng (DomainFeatureEngineer), điền khuyết (KNNImputer), chuẩn hóa (StandardScaler) và mã hóa (OneHotEncoder) vào bên trong đối tượng Pipeline của thư viện scikit-learn. Khi chạy 5-Fold Cross-Validation hoặc chia Train/Test, các tham số (như trung vị, trung bình, độ lệch chuẩn) chỉ được học (fit) trên tập Train và chỉ áp dụng biến đổi (transform) lên tập Test, tuyệt đối không để thông tin từ tập Test lọt vào quá trình huấn luyện.",
        ),
        (
            "Câu 3: Tại sao nhóm lại chọn 4 mô hình là Logistic Regression, Random Forest, XGBoost và Soft Voting Ensemble?",
            "Trả lời: Nhóm chọn bộ 4 mô hình này vì chúng đại diện cho đầy đủ các cấp độ thuật toán từ cơ bản đến nâng cao trên dữ liệu bảng: (1) Logistic Regression làm mô hình tuyến tính cơ sở (Baseline); (2) Random Forest đại diện cho nhánh Bagging (xây 300 cây song song giúp giảm phương sai, chống học vẹt); (3) XGBoost đại diện cho nhánh Gradient Boosting (xây 200 cây nối tiếp giúp giảm độ chệch, tối ưu sai số); và (4) Soft Voting Ensemble kết hợp cả 3 mô hình trên để tận dụng ưu điểm của từng thuật toán.",
        ),
        (
            "Câu 4: Sự khác nhau cốt lõi giữa Random Forest và XGBoost là gì?",
            "Trả lời: Cả hai đều dựa trên Cây quyết định nhưng khác nhau ở cách kết hợp các cây: Random Forest sử dụng kỹ thuật Bagging — nghĩa là xây dựng 300 cây độc lập song song trên các mẫu con ngẫu nhiên rồi lấy trung bình kết quả để chống quá khớp (Overfitting). Ngược lại, XGBoost sử dụng kỹ thuật Boosting — xây dựng các cây tuần tự nối tiếp nhau, trong đó cây sau học cách sửa sai số (residual error) của các cây đứng trước.",
        ),
        (
            "Câu 5: Soft Voting Ensemble trong bài hoạt động như thế nào? Tại sao nhóm đặt trọng số weights = [1, 3, 2]?",
            "Trả lời: Soft Voting Ensemble lấy xác suất dự đoán (%) của 3 mô hình con là Logistic Regression (trọng số 1), Random Forest (trọng số 3) và XGBoost (trọng số 2), sau đó tính trung bình cộng có trọng số để ra xác suất cuối cùng. Nhóm đặt trọng số [1, 3, 2] vì thực nghiệm cho thấy Random Forest và XGBoost là hai mô hình phi tuyến có độ chính xác độc lập cao hơn hẳn Logistic Regression, nên được ưu tiên tiếng nói lớn hơn trong hội đồng.",
        ),
        (
            "Câu 6: Trong bài toán dự đoán sinh viên bỏ học, giữa Precision và Recall thì chỉ số nào quan trọng hơn? Tại sao?",
            "Trả lời: Trong bài toán cảnh báo sớm học đường, chỉ số Recall (Độ nhạy) và F1-Score quan trọng hơn. Bởi vì nếu mô hình dự báo nhầm một sinh viên an toàn thành có nguy cơ (False Positive), nhà trường chỉ tốn thêm một buổi hỏi thăm của cố vấn học tập. Nhưng nếu bỏ sót một sinh viên thực sự sắp bỏ học (False Negative — tức Recall thấp), sinh viên đó sẽ nghỉ học thật và không được cứu kịp thời. Mô hình Random Forest của nhóm đạt Test F1 = 83.63% và Recall rất cao (~86%), đáp ứng tốt yêu cầu này.",
        ),
        (
            "Câu 7: Tại sao nhóm lại chia thực nghiệm thành 2 kịch bản (Kịch bản 1 có Burnout_Level và Kịch bản 2 không có Burnout_Level)?",
            "Trả lời: Vì trong thực tế quản lý đại học có 2 thời điểm khác nhau: (1) Khi nhà trường đã tổ chức khảo sát tâm lý và biết mức độ Burnout_Level của sinh viên, ta dùng Kịch bản 1 (Random Forest) để đạt độ chính xác cao nhất (ROC-AUC = 89.01%); (2) Khi mới đầu học kỳ chưa kịp khảo sát tâm lý Burnout, nhà trường chỉ có dữ liệu điểm danh, GPA, giờ học... lúc này Kịch bản 2 (XGBoost / Soft Voting) chứng minh hệ thống vẫn dự báo sớm được nguy cơ bỏ học với ROC-AUC đạt 88.26%.",
        ),
        (
            "Câu 8: Theo kết quả của mô hình, những yếu tố nào ảnh hưởng mạnh nhất đến việc một sinh viên bỏ học?",
            "Trả lời: Dựa trên phân tích tương quan và biểu đồ Feature Importance (Hình 10), 5 yếu tố quyết định lớn nhất là: (1) Mức độ kiệt sức (Burnout_Level — nhóm High bỏ học tới 92%); (2) Tỷ số mất cân bằng giữa Căng thẳng và Gắn kết học tập (Stress_Engagement_Ratio); (3) Tỷ lệ chuyên cần đi học (Attendance_Percent); (4) Chỉ số gắn kết học tập tổng hợp (Academic_Engagement_Index: kết hợp chuyên cần, giờ tự học, động lực và GPA); và (5) Chỉ số căng thẳng tâm lý (Psychological_Stress_Index: Stress, Lo âu, Áp lực tài chính).",
        ),
        (
            "Câu 9: Nhóm đã tạo ra 6 đặc trưng mới (Feature Engineering) như thế nào và chúng có tác dụng gì?",
            "Trả lời: Thay vì chỉ đưa các cột rời rạc vào máy học, nhóm dựa trên lý thuyết tâm lý giáo dục để tạo 6 chỉ số tổng hợp trong class DomainFeatureEngineer: ví dụ Academic_Engagement_Index tổng hợp sự chăm chỉ học tập; Psychological_Stress_Index tổng hợp áp lực tâm lý; và đặc biệt là Stress_Engagement_Ratio lấy Căng thẳng chia cho Gắn kết học tập. Các biến mới này giúp mô hình nhận diện ngay những sinh viên vừa chịu áp lực cao vừa suy giảm việc học, giúp tăng khoảng 3–4% điểm ROC-AUC.",
        ),
        (
            "Câu 10: Nhóm mở rộng dữ liệu từ 800 lên 2,000 bản ghi bằng phương pháp gì để đảm bảo tính khoa học?",
            "Trả lời: Nhóm giữ nguyên 800 bản ghi gốc làm nền tảng, sau đó ước lượng phân phối chuẩn đa biến (Multivariate Gaussian Distribution gồm vector trung bình và ma trận hiệp phương sai của 16 biến định lượng) kết hợp phân phối xác suất thực nghiệm của các biến phân loại trên từng nhóm phân tầng (Burnout_Level × Dropout_Risk). Nhờ vậy, 1,200 bản ghi mở rộng tuân thủ chặt chẽ cấu trúc tương quan và miền giá trị của dữ liệu gốc.",
        ),
    ]

    for q_text, a_text in qas:
        add_callout(doc, "❓ " + q_text, a_text, accent_hex=NAVY_HEX, bg_hex=CALLOUT_BG_HEX)

    doc.save(OUT_ROOT)
    os.makedirs(os.path.dirname(OUT_SIC_VI), exist_ok=True)
    shutil.copy2(OUT_ROOT, OUT_SIC_VI)
    print(f"Successfully created Defense Study Guide Word document at:\n  1. {OUT_ROOT}\n  2. {OUT_SIC_VI}")


if __name__ == "__main__":
    build_study_guide()

