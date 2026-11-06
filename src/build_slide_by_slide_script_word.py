"""
Generates a complete Slide-by-Slide Presentation Script Word Document (.docx)
matching all 14 slides of `SIC_AI_Capstone Project_Presentation Slide Template.pptx`.

Outputs:
1. Kich_Ban_Thuyet_Trinh_Tung_Slide_Nhom03.docx (in project root)
2. SIC_Documents_VI/Kich_Ban_Thuyet_Trinh_Tung_Slide_Nhom03.docx
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
FIG_VI = os.path.join(BASE_DIR, "figures_vi")
OUT_ROOT = os.path.join(BASE_DIR, "Kich_Ban_Thuyet_Trinh_Tung_Slide_Nhom03.docx")
OUT_SIC_VI = os.path.join(BASE_DIR, "SIC_Documents_VI", "Kich_Ban_Thuyet_Trinh_Tung_Slide_Nhom03.docx")

NAVY_HEX = "1D3557"
SLATE_HEX = "264653"
TEAL_HEX = "2A9D8F"
AMBER_HEX = "E76F51"
LIGHT_BG_HEX = "F4F7FB"
SCRIPT_BG_HEX = "EDF6F5"
TIP_BG_HEX = "FFF8E7"

NAVY_RGB = RGBColor(29, 53, 87)
SLATE_RGB = RGBColor(38, 70, 83)
TEAL_RGB = RGBColor(42, 157, 143)
DARK_TEXT_RGB = RGBColor(30, 41, 59)


def set_cell_background(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=110, bottom=110, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def configure_doc(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.3)
    section.right_margin = Cm(2.0)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = DARK_TEXT_RGB
    normal.paragraph_format.line_spacing = 1.3
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def add_slide_section(
    doc,
    slide_no,
    slide_title,
    speaker,
    duration,
    visual_layout_points,
    spoken_paragraphs,
    pointer_tip,
    fig_name=None,
):
    # Slide Header Banner
    tbl_hdr = doc.add_table(rows=1, cols=1)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl_hdr.cell(0, 0)
    set_cell_background(c, NAVY_HEX)
    set_cell_margins(c, top=110, bottom=110, left=160, right=160)
    prevent_row_split(tbl_hdr.rows[0])

    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f"SLIDE {slide_no:02d} / 14: {slide_title}\n")
    r1.bold = True
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(13.5)
    r1.font.color.rgb = RGBColor(255, 255, 255)

    r2 = p.add_run(f"👤 Người trình bày gợi ý: {speaker}   |   ⏱ Thời lượng: {duration}")
    r2.italic = True
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = RGBColor(216, 243, 220)

    # On-screen visual layout summary
    p_vis = doc.add_paragraph()
    p_vis.paragraph_format.space_before = Pt(6)
    p_vis.paragraph_format.space_after = Pt(3)
    rv = p_vis.add_run("🖥 Nội dung hiển thị trên màn hình Slide:")
    rv.bold = True
    rv.font.color.rgb = SLATE_RGB

    for pt in visual_layout_points:
        pb = doc.add_paragraph()
        pb.paragraph_format.left_indent = Cm(0.5)
        pb.paragraph_format.space_after = Pt(2)
        rb = pb.add_run("• " + pt)
        rb.font.size = Pt(11.5)

    # Optional embedded figure from the slide
    if fig_name:
        fig_path = os.path.join(FIG_VI, fig_name)
        if os.path.exists(fig_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(4)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.add_run().add_picture(fig_path, width=Inches(5.3))

    # Verbatim Script Box
    tbl_script = doc.add_table(rows=1, cols=1)
    tbl_script.alignment = WD_TABLE_ALIGNMENT.CENTER
    cs = tbl_script.cell(0, 0)
    set_cell_background(cs, SCRIPT_BG_HEX)
    set_cell_margins(cs, top=130, bottom=130, left=200, right=170)
    prevent_row_split(tbl_script.rows[0])

    tcPr = cs._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{TEAL_HEX}"/>'
        f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>'
        f"</w:tcBorders>"
    )
    tcPr.append(borders)

    ps_hdr = cs.paragraphs[0]
    ps_hdr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ps_hdr.paragraph_format.space_after = Pt(4)
    rs_h = ps_hdr.add_run("🎤 LỜI THOẠI THUYẾT TRÌNH (ĐỌC TRỰC TIẾP):")
    rs_h.bold = True
    rs_h.font.name = "Times New Roman"
    rs_h.font.size = Pt(12)
    rs_h.font.color.rgb = NAVY_RGB

    for sp_text in spoken_paragraphs:
        ps_body = cs.add_paragraph()
        ps_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        ps_body.paragraph_format.space_after = Pt(5)
        rs_b = ps_body.add_run(sp_text)
        rs_b.font.name = "Times New Roman"
        rs_b.font.size = Pt(12)
        rs_b.font.color.rgb = DARK_TEXT_RGB

    # Pointer Tip Box
    tbl_tip = doc.add_table(rows=1, cols=1)
    tbl_tip.alignment = WD_TABLE_ALIGNMENT.CENTER
    ct = tbl_tip.cell(0, 0)
    set_cell_background(ct, TIP_BG_HEX)
    set_cell_margins(ct, top=90, bottom=90, left=200, right=160)
    prevent_row_split(tbl_tip.rows[0])

    tcPr2 = ct._tc.get_or_add_tcPr()
    borders2 = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:left w:val="single" w:sz="20" w:space="0" w:color="{AMBER_HEX}"/>'
        f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>'
        f"</w:tcBorders>"
    )
    tcPr2.append(borders2)

    pt_p = ct.paragraphs[0]
    pt_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pt_p.paragraph_format.space_after = Pt(0)
    rt1 = pt_p.add_run("👉 Thao tác chỉ tay & Lưu ý trên Slide: ")
    rt1.bold = True
    rt1.font.name = "Times New Roman"
    rt1.font.size = Pt(11.5)
    rt1.font.color.rgb = NAVY_RGB
    rt2 = pt_p.add_run(pointer_tip)
    rt2.font.name = "Times New Roman"
    rt2.font.size = Pt(11.5)
    rt2.font.color.rgb = DARK_TEXT_RGB

    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(10)


def build_slide_script_doc():
    doc = Document()
    configure_doc(doc)

    # COVER BANNER
    tbl_title = doc.add_table(rows=1, cols=1)
    tbl_title.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0 = tbl_title.cell(0, 0)
    set_cell_background(c0, SLATE_HEX)
    set_cell_margins(c0, top=180, bottom=180, left=200, right=200)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_m = p0.add_run("KỊCH BẢN THUYẾT TRÌNH CHI TIẾT THEO TỪNG SLIDE POWERPOINT (SLIDE 01 → SLIDE 14)\n")
    r_m.bold = True
    r_m.font.name = "Times New Roman"
    r_m.font.size = Pt(15)
    r_m.font.color.rgb = RGBColor(255, 255, 255)

    r_s = p0.add_run(
        "Khớp 100% với file trình chiếu: SIC_Documents_VI/SIC_AI_Capstone Project_Presentation Slide Template.pptx\n"
        "Nhóm 03: Đỗ Quang Minh • Vũ Quang Huy  |  Thời lượng tổng: ~10–12 phút thuyết trình + 2 phút Demo Web"
    )
    r_s.font.name = "Times New Roman"
    r_s.font.size = Pt(11.5)
    r_s.font.color.rgb = RGBColor(216, 243, 220)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # SLIDE 01
    add_slide_section(
        doc,
        slide_no=1,
        slide_title="TRANG BÌA (COVER SLIDE) & 4 CHỈ SỐ TỔNG QUAN",
        speaker="Đỗ Quang Minh (hoặc Vũ Quang Huy)",
        duration="45 giây",
        visual_layout_points=[
            "Tên đề tài lớn ở trung tâm + khung giới thiệu 2 thành viên Nhóm 03 (Đỗ Quang Minh, Vũ Quang Huy).",
            "4 hộp KPI nổi bật ở hàng dưới cùng: [2,000 Hồ sơ Sinh viên (Sạch 100%)] | [25 → 34 Thuộc tính & 6 Chỉ số Mới] | [4 Mô hình: LR, RF, XGBoost, Soft Voting] | [89.01% Đỉnh ROC-AUC].",
        ],
        spoken_paragraphs=[
            "“Em xin kính chào các thầy cô trong Hội đồng chấm đồ án tốt nghiệp khóa học Trí tuệ Nhân tạo – Samsung Innovation Campus ạ! Chúng em là Nhóm 03, gồm hai thành viên là Đỗ Quang Minh và Vũ Quang Huy.”",
            "“Hôm nay, nhóm em xin phép được trình bày kết quả đồ án với đề tài: Xây dựng, huấn luyện và so sánh các mô hình học máy dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập (Student Burnout).”",
            "“Nhìn vào 4 ô chỉ số ở cuối trang bìa, thầy cô có thể thấy bức tranh tổng quan của hệ thống: Từ 800 hồ sơ sinh viên gốc, nhóm đã làm sạch qua 6 bước và mở rộng lên 2,000 hồ sơ sạch 100%; phát triển từ 25 thuộc tính ban đầu lên 34 trường thông tin nhờ 6 chỉ số chuyên ngành mới; đối sánh 4 mô hình học máy tiêu biểu và đạt độ chính xác ROC-AUC đỉnh là 89.01%.”",
        ],
        pointer_tip="Chỉ tay dọc theo 4 hộp KPI ở dưới đáy Slide 01 từ trái qua phải (2,000 SV -> 25->34 cột -> 4 Mô hình -> 89.01%) rồi bấm sang Slide 02.",
    )

    # SLIDE 02
    add_slide_section(
        doc,
        slide_no=2,
        slide_title="CẤU TRÚC BÀI THUYẾT TRÌNH (AGENDA 6 PHẦN)",
        speaker="Đỗ Quang Minh",
        duration="35 giây",
        visual_layout_points=[
            "6 thẻ nội dung chia thành 2 hàng × 3 cột từ Mục 01 đến Mục 06.",
        ],
        spoken_paragraphs=[
            "“Bài thuyết trình của nhóm em hôm nay gồm 6 phần nối tiếp nhau: Phần 1 là Giới thiệu bối cảnh và cơ chế dẫn đến việc bỏ học đại học; Phần 2 trình bày Bộ dữ liệu và quy trình làm sạch 6 bước; Phần 3 giới thiệu Kỹ thuật tạo 6 đặc trưng chuyên ngành; Phần 4 là các phát hiện quan trọng từ Phân tích khám phá dữ liệu (EDA); Phần 5 trình bày kết quả Huấn luyện và đối sánh 4 mô hình học máy trên 2 kịch bản; và cuối cùng Phần 6 là Kiến trúc ứng dụng Web cảnh báo sớm kèm phần Demo trực tiếp.”",
        ],
        pointer_tip="Lướt nhanh qua 6 ô màu trên Slide 02, không cần đọc kỹ từng dòng nhỏ trong ô vì các slide sau sẽ đi vào chi tiết.",
    )

    # SLIDE 03
    add_slide_section(
        doc,
        slide_no=3,
        slide_title="ĐẶT VẤN ĐỀ & CƠ CHẾ HÌNH THÀNH NGUY CƠ BỎ HỌC",
        speaker="Đỗ Quang Minh",
        duration="60 giây",
        visual_layout_points=[
            "Cột trái: 2 hộp nguyên nhân gốc (1. Sự tham gia học tập suy giảm & 2. Áp lực căng thẳng tâm lý - xã hội).",
            "Cột giữa (màu cam): Trạng thái trung gian Kiệt sức học đường (Burnout_Level: Low 40%, Medium 35%, High 25%).",
            "Cột phải (màu xanh nhạt): Biến mục tiêu Dropout_Risk (Yes 54.6% / No 45.4%) & Giải pháp AI 2 kịch bản.",
        ],
        spoken_paragraphs=[
            "“Trước hết, về mặt lý thuyết giáo dục, một sinh viên quyết định bỏ học đại học hiếm khi xảy ra đột ngột, mà là kết quả tích lũy của hai nhóm nguyên nhân ở cột bên trái màn hình: Thứ nhất là sự suy giảm tham gia học tập — thể hiện qua vắng học nhiều, giờ tự học thấp, mất động lực và nợ môn. Thứ hai là áp lực căng thẳng kéo dài — gồm điểm Stress, lo âu, gánh nặng tài chính, áp lực bạn bè và thiếu ngủ.”",
            "“Khi hai nhóm áp lực này vượt quá sức chịu đựng mà thiếu sự hỗ trợ từ gia đình hay nhà trường, sinh viên sẽ rơi vào trạng thái trung gian ở cột giữa là Hội chứng Kiệt sức Học đường (Burnout_Level chia làm 3 mức: Low, Medium và High), từ đó dẫn thẳng tới nguy cơ bỏ học (Dropout_Risk) ở cột bên phải.”",
            "“Chính vì cơ chế này, nhóm em thiết kế giải pháp AI hoạt động trên cả 2 kịch bản: vừa chẩn đoán chính xác khi đã biết mức độ Burnout (đạt ROC-AUC 89.01%), vừa cảnh báo sớm ngay từ các chỉ số học tập và căng thẳng ban đầu khi chưa kịp khảo sát Burnout (đạt ROC-AUC 88.26%).”",
        ],
        pointer_tip="Chỉ từ Cột trái (Nguyên nhân) -> Cột giữa (Trạng thái Kiệt sức Burnout) -> Cột phải (Kết quả Bỏ học Dropout).",
    )

    # SLIDE 04
    add_slide_section(
        doc,
        slide_no=4,
        slide_title="MỤC TIÊU ĐỀ TÀI, PHÂN CÔNG NHÓM 03 & LỘ TRÌNH 4 GIAI ĐOẠN",
        speaker="Đỗ Quang Minh",
        duration="45 giây",
        visual_layout_points=[
            "Khung lớn bên trái: 4 Mục tiêu cốt lõi của đồ án.",
            "2 khung bên phải: Phân công công việc 50% – 50% giữa Đỗ Quang Minh và Vũ Quang Huy.",
            "4 hộp ngang ở dưới đáy: Lộ trình 4 giai đoạn từ 19/10 đến 05/11 (Hoàn thành 100%).",
        ],
        spoken_paragraphs=[
            "“Từ bài toán trên, nhóm đặt ra 4 mục tiêu cụ thể: Một là làm sạch dữ liệu qua 6 bước và mở rộng từ 800 lên 2,000 hồ sơ sinh viên; Hai là thiết kế 6 chỉ số đặc trưng mới; Ba là huấn luyện và kiểm định chéo 5-Fold trên 4 mô hình tiêu biểu; và Bốn là xây dựng ứng dụng Web FastAPI hoàn chỉnh.”",
            "“Về phân công nhiệm vụ, hai thành viên chia sẻ khối lượng công việc đồng đều 50–50: Bạn Đỗ Quang Minh phụ trách thiết kế Pipeline học máy, tạo đặc trưng, huấn luyện nhóm mô hình cây và phát triển Backend FastAPI; bạn Vũ Quang Huy phụ trách kiểm toán dữ liệu, phân tích thống kê EDA, xuất 12 biểu đồ trực quan và thiết kế giao diện Frontend cùng tài liệu báo cáo. Cả 4 giai đoạn của dự án đều đã hoàn thành đúng tiến độ 100%.”",
        ],
        pointer_tip="Chỉ vào khung 4 mục tiêu bên trái rồi chỉ sang 2 hộp phân công thành viên bên phải.",
    )

    # SLIDE 05
    add_slide_section(
        doc,
        slide_no=5,
        slide_title="KIẾN TRÚC BỘ DỮ LIỆU & XỬ LÝ DỮ LIỆU KHUYẾT ĐA BIẾN",
        speaker="Vũ Quang Huy (hoặc Đỗ Quang Minh)",
        duration="60 giây",
        visual_layout_points=[
            "Khung bên trái: Tóm tắt 5 nhóm thuộc tính (25 biến) và lý do chọn KNNImputer(k=7).",
            "Biểu đồ bên phải (Hình 01): Tỷ lệ dữ liệu khuyết trên 18/25 cột của tập dữ liệu gốc.",
        ],
        spoken_paragraphs=[
            "“Bước sang phần Dữ liệu và Tiền xử lý: Bộ dữ liệu gốc gồm 25 thuộc tính chia thành 5 nhóm chức năng: Nhân khẩu học (7 biến), Sự tham gia học tập (5 biến), Lối sống sinh hoạt (5 biến), Căng thẳng & Hỗ trợ tâm lý (6 biến) và 2 biến mục tiêu là Burnout_Level cùng Dropout_Risk.”",
            "“Như thầy cô thấy trên biểu đồ bên phải, khi kiểm toán tập dữ liệu gốc 800 sinh viên, có tới 18 trên 25 cột bị khuyết dữ liệu (NaN) với tỷ lệ từ 1.6% đến 6.4%, cao nhất là Chuyên cần (51 mẫu khuyết) và Mức căng thẳng (49 mẫu khuyết). Nếu chúng ta dùng lệnh xóa dòng (dropna), tập dữ liệu sẽ mất hơn 40% số sinh viên.”",
            "“Vì vậy, nhóm áp dụng quy trình làm sạch 6 bước: dùng thuật toán KNNImputer với k=7 láng giềng gần nhất để nội suy các biến số từ 7 sinh viên có hồ sơ tương đồng nhất, kết hợp lọc ngoại lai bằng phương pháp khoảng tứ phân vị IQR và mở rộng theo phân phối chuẩn đa biến lên 2,000 bản ghi sạch 100% (không còn bất kỳ giá trị NaN nào) để phục vụ huấn luyện.”",
        ],
        pointer_tip="Chỉ vào 2 thanh dài nhất trên biểu đồ bên phải (Attendance_Percent 6.4% và Stress_Level 6.1%) để giải thích tại sao phải dùng KNNImputer.",
        fig_name="01_missing_values_analysis.png",
    )

    # SLIDE 06
    add_slide_section(
        doc,
        slide_no=6,
        slide_title="QUY TRÌNH HỌC MÁY 5 BƯỚC & KỸ THUẬT TẠO 6 ĐẶC TRƯNG MỚI",
        speaker="Đỗ Quang Minh",
        duration="65 giây",
        visual_layout_points=[
            "Nửa trên: Sơ đồ quy trình 5 bước (Hình 11: Thu thập & Mở rộng -> Làm sạch 6 bước -> Tạo đặc trưng -> Huấn luyện 4 mô hình -> Web UI).",
            "Nửa dưới: 6 hộp công thức của 6 chỉ số mới trong lớp DomainFeatureEngineer.",
        ],
        spoken_paragraphs=[
            "“Trên Slide 06 là quy trình thực hiện dự án gồm 5 bước khép kín. Điểm sáng tạo kỹ thuật lớn nhất của nhóm nằm ở Bước 3 — Kỹ thuật tạo đặc trưng chuyên ngành (DomainFeatureEngineer) được hiển thị ở 6 ô phía dưới.”",
            "“Thay vì chỉ đưa các cột rời rạc vào máy học, nhóm đã tổng hợp 6 chỉ số mới mang ý nghĩa giáo dục sâu sắc: Thứ nhất là Academic_Engagement_Index tổng hợp mức độ chăm chỉ từ chuyên cần, giờ tự học, động lực và GPA; Thứ hai là Psychological_Stress_Index tổng hợp áp lực từ Stress, lo âu, tài chính và bạn bè; Thứ ba — và cũng là biến quan trọng nhất — là Stress_Engagement_Ratio lấy Căng thẳng chia cho Gắn kết học tập; cùng 3 chỉ số về Cân bằng hỗ trợ gia đình, Tỷ lệ mỏi mệt số (Giờ màn hình trên Giờ ngủ) và Gánh nặng học vụ.”",
            "“Toàn bộ lớp tạo đặc trưng này được nhúng trực tiếp vào bên trong Pipeline của scikit-learn để đảm bảo không xảy ra hiện tượng rò rỉ dữ liệu (No Data Leakage) giữa tập Train và tập Test.”",
        ],
        pointer_tip="Chỉ vào hộp số 1 (Academic_Engagement_Index), hộp số 2 (Psychological_Stress_Index) và hộp số 3 (Stress_Engagement_Ratio).",
        fig_name="11_system_workflow_diagram.png",
    )

    doc.add_page_break()

    # SLIDE 07
    add_slide_section(
        doc,
        slide_no=7,
        slide_title="PHÁT HIỆN EDA 1: TÁC ĐỘNG QUYẾT ĐỊNH CỦA BURNOUT LÊN BỎ HỌC",
        speaker="Vũ Quang Huy",
        duration="55 giây",
        visual_layout_points=[
            "3 hộp KPI ở hàng trên: Kiệt sức Thấp (23.8% bỏ học) -> Kiệt sức Vừa (63.2% bỏ học) -> Kiệt sức Cao (92.0% bỏ học).",
            "Biểu đồ 3 cột ở dưới (Hình 02): Phân phối Dropout_Risk, phân phối Burnout_Level và biểu đồ chồng tỷ lệ bỏ học theo Burnout.",
        ],
        spoken_paragraphs=[
            "“Tiếp theo, em xin trình bày 3 phát hiện quan trọng nhất từ quá trình Phân tích khám phá dữ liệu (EDA). Phát hiện đầu tiên trên Slide 07 trả lời câu hỏi: Kiệt sức học đường ảnh hưởng mạnh tới mức nào đến quyết định bỏ học?”",
            "“Thầy cô có thể nhìn vào biểu đồ ngoài cùng bên phải và 3 hộp chỉ số phía trên: Ở nhóm sinh viên có mức kiệt sức Thấp (Low Burnout), tỷ lệ bỏ học chỉ là 23.8%. Khi chuyển sang mức Trung bình (Medium Burnout), tỷ lệ bỏ học tăng vọt lên 63.2%. Và đặc biệt ở nhóm Kiệt sức Cao (High Burnout), tỷ lệ bỏ học lên tới 92.0% — tức là cứ 10 sinh viên bị kiệt sức nặng thì có hơn 9 em có nguy cơ bỏ học! Điều này khẳng định việc phát hiện sớm Burnout là chìa khóa để ngăn chặn bỏ học.”",
        ],
        pointer_tip="Chỉ vào cột màu cam đậm (92.0%) ở biểu đồ thứ 3 bên phải màn hình.",
        fig_name="02_target_and_burnout_distribution.png",
    )

    # SLIDE 08
    add_slide_section(
        doc,
        slide_no=8,
        slide_title="PHÁT HIỆN EDA 2: ĐỐI SÁNH THAM GIA HỌC TẬP & CĂNG THẲNG TÂM LÝ",
        speaker="Vũ Quang Huy",
        duration="50 giây",
        visual_layout_points=[
            "Bên trái (Hình 03): 8 biểu đồ Violin kết hợp Boxplot (4 biến học tập ở hàng trên, 4 biến tâm lý ở hàng dưới).",
            "Khung bên phải: Bảng đối chiếu giá trị trung bình (μ) giữa nhóm Không bỏ học (No) và Có bỏ học (Yes).",
        ],
        spoken_paragraphs=[
            "“Sang Slide 08, nhóm sử dụng biểu đồ kết hợp Violin và Boxplot để so sánh trực tiếp 8 chỉ số giữa nhóm sinh viên An toàn (màu xanh) và nhóm Nguy cơ bỏ học (màu cam).”",
            "“Nhìn vào bảng tổng hợp bên phải, chúng ta thấy sự đối lập rất rõ ràng: Nhóm có nguy cơ bỏ học có tỷ lệ chuyên cần trung bình thấp hơn tới 7 điểm phần trăm (chỉ 73.7% so với 80.7% của nhóm an toàn), thời gian tự học ít hơn (2.93 giờ so với 3.43 giờ/ngày), điểm động lực thấp hơn (5.11 so với 5.90); trong khi điểm Căng thẳng (5.81), Lo âu (5.26) và Áp lực tài chính đều cao hơn hẳn nhóm an toàn.”",
        ],
        pointer_tip="Chỉ vào hàng biểu đồ phía trên (màu xanh cao hơn màu cam - Học tập) và hàng dưới (màu cam cao hơn màu xanh - Căng thẳng).",
        fig_name="03_stress_vs_engagement_distributions.png",
    )

    # SLIDE 09
    add_slide_section(
        doc,
        slide_no=9,
        slide_title="PHÁT HIỆN EDA 3: MA TRẬN TƯƠNG QUAN PEARSON & NHÂN KHẨU HỌC",
        speaker="Vũ Quang Huy",
        duration="50 giây",
        visual_layout_points=[
            "Bên trái (Hình 04): Ma trận nhiệt tương quan Pearson giữa 18 biến định lượng.",
            "Bên phải: Hộp trên là 5 Yếu tố bảo vệ (tương quan âm, màu xanh); Hộp dưới là 5 Yếu tố nguy cơ (tương quan dương, màu đỏ).",
        ],
        spoken_paragraphs=[
            "“Trên Slide 09 là Ma trận tương quan Pearson giúp lượng hóa mức độ tác động của từng biến. Nhóm đã đúc kết thành 2 nhóm yếu tố ở bên phải màn hình:”",
            "“Nhóm thứ nhất là 5 Yếu tố bảo vệ (mang dấu âm — tức là càng cao thì càng ít bỏ học), đứng đầu là Tỷ lệ chuyên cần (r = -0.244), Điểm động lực (r = -0.193), Giờ tự học (r = -0.160) và Sự hỗ trợ từ gia đình (r = -0.137). Nhóm thứ hai là 5 Yếu tố nguy cơ (mang dấu dương — càng cao càng dễ bỏ học), đứng đầu là Mức độ Burnout (r = +0.550), Điểm Stress, Điểm Lo âu, Thời gian dùng màn hình điện thoại/máy tính và Áp lực tài chính ở nhóm gia đình thu nhập thấp.”",
        ],
        pointer_tip="Chỉ vào cột đầu tiên của ma trận nhiệt bên trái rồi chỉ sang 2 hộp tổng kết Xanh / Đỏ bên phải.",
        fig_name="04_correlation_matrix_heatmap.png",
    )

    # SLIDE 10
    add_slide_section(
        doc,
        slide_no=10,
        slide_title="SO SÁNH HIỆU NĂNG 4 MÔ HÌNH HỌC MÁY TIÊU BIỂU TRÊN 2 KỊCH BẢN",
        speaker="Đỗ Quang Minh",
        duration="65 giây",
        visual_layout_points=[
            "Biểu đồ ngang lớn ở giữa (Hình 07): So sánh 4 mô hình (Random Forest, Soft Voting Ensemble, XGBoost, Logistic Regression) trên Kịch bản 1 và Kịch bản 2.",
            "2 hộp tóm tắt ở dưới đáy: Kết quả Kịch bản 1 (bên trái) và Kịch bản 2 (bên phải).",
        ],
        spoken_paragraphs=[
            "“Bước sang Phần 4 — Huấn luyện và Đối sánh Mô hình: Nhóm lựa chọn 4 mô hình đại diện cho các trường phái mạnh nhất trên dữ liệu bảng, gồm: Logistic Regression (mô hình tuyến tính cơ sở), Random Forest (300 cây Bagging), XGBoost (200 cây Gradient Boosting) và Soft Voting Ensemble (mô hình hội đồng kết hợp cả 3 thuật toán trên với trọng số [1, 3, 2]).”",
            "“Như biểu đồ trên Slide 10 thể hiện, ở Kịch bản 1 (khi có đầy đủ đặc trưng bao gồm cả Burnout_Level), mô hình Random Forest vươn lên dẫn đầu với điểm ROC-AUC kiểm định chéo 5-Fold đạt 89.01%, theo sát là Soft Voting Ensemble đạt 88.89% và XGBoost đạt 88.64%.”",
            "“Đặc biệt, ở biểu đồ bên phải cho Kịch bản 2 (Cảnh báo sớm khi ẩn hoàn toàn biến Burnout_Level), hai mô hình XGBoost và Soft Voting Ensemble vẫn duy trì năng lực dự báo rất cao với ROC-AUC đạt 88.26%, chứng minh hệ thống hoàn toàn có thể cảnh báo sớm ngay từ đầu học kỳ.”",
        ],
        pointer_tip="Chỉ vào thanh cao nhất của biểu đồ bên trái (Random Forest - Kịch bản 1) và thanh cao nhất bên phải (XGBoost / Soft Voting - Kịch bản 2).",
        fig_name="07_model_performance_comparison.png",
    )

    # SLIDE 11
    add_slide_section(
        doc,
        slide_no=11,
        slide_title="BẢNG TỔNG HỢP SỐ LIỆU THỰC NGHIỆM (CV 5-FOLD & TẬP TEST ĐỘC LẬP)",
        speaker="Đỗ Quang Minh",
        duration="55 giây",
        visual_layout_points=[
            "Bảng số liệu tổng hợp đầy đủ 9 cột chỉ số (CV Accuracy, CV Recall, CV F1, CV AUC, Test Acc, Test F1, Test AUC, KB2 CV AUC) cho cả 4 mô hình.",
        ],
        spoken_paragraphs=[
            "“Slide 11 trình bày bảng số liệu chi tiết của cả 4 mô hình trên tập Kiểm định chéo 5-Fold (1,600 sinh viên) và tập Kiểm thử độc lập 20% (400 sinh viên hoàn toàn chưa từng xuất hiện lúc train).”",
            "“Có 2 điểm rất đáng chú ý trên bảng này: Thứ nhất, trên tập Test độc lập, mô hình Random Forest đạt Test Accuracy 81.50%, Test F1-Score 83.63% và độ nhạy Recall lên tới 86.38% — nghĩa là cứ 100 sinh viên thực sự có nguy cơ bỏ học thì mô hình phát hiện đúng hơn 86 em, giúp nhà trường giảm thiểu tối đa việc bỏ sót sinh viên cần cứu giúp. Thứ hai, độ chênh lệch giữa điểm CV và điểm Test chỉ khoảng 0.6%, chứng tỏ các mô hình hoàn toàn không bị học vẹt (Overfitting).”",
        ],
        pointer_tip="Chỉ vào dòng số 1 (Random Forest ★) được tô nền xanh nhạt trên bảng.",
    )

    doc.add_page_break()

    # SLIDE 12
    add_slide_section(
        doc,
        slide_no=12,
        slide_title="ĐƯỜNG CONG ROC / PRECISION-RECALL & ĐỘ QUAN TRỌNG ĐẶC TRƯNG",
        speaker="Đỗ Quang Minh",
        duration="60 giây",
        visual_layout_points=[
            "Bên trái (Hình 08): Đường cong ROC và Precision-Recall của 4 mô hình trên tập Test.",
            "Bên phải (Hình 10): Top 15 đặc trưng quan trọng nhất (Feature Importance) trên Random Forest và XGBoost.",
            "Hộp ngang ở dưới: Top 6 đặc trưng quyết định lớn nhất đến nguy cơ bỏ học.",
        ],
        spoken_paragraphs=[
            "“Để giải thích tại sao mô hình lại đưa ra dự đoán như vậy (Explainable AI), Slide 12 hiển thị đường cong ROC/PR ở bên trái và biểu đồ Mức độ quan trọng của các đặc trưng (Feature Importance) ở bên phải.”",
            "“Nhìn vào biểu đồ Feature Importance của Random Forest và XGBoost, chúng ta thấy một kết quả cực kỳ thuyết phục: Ngay sau biến Burnout_Ordinal, các chỉ số tổng hợp do nhóm tự thiết kế trong bước Feature Engineering là Stress_Engagement_Ratio (Tỷ số Căng thẳng trên Gắn kết), Psychological_Stress_Index và Academic_Engagement_Index đều vươn lên nằm trong Top 5 biến quan trọng nhất, vượt qua các biến đơn lẻ truyền thống. Điều này khẳng định 6 đặc trưng mới của nhóm thực sự mang lại giá trị cốt lõi cho mô hình.”",
        ],
        pointer_tip="Chỉ vào 4–5 thanh dài nhất ở trên cùng của biểu đồ Feature Importance bên phải.",
        fig_name="10_feature_importance_analysis.png",
    )

    # SLIDE 13
    add_slide_section(
        doc,
        slide_no=13,
        slide_title="KIẾN TRÚC HỆ THỐNG 3 TẦNG & ỨNG DỤNG WEB FASTAPI (APP.PY)",
        speaker="Vũ Quang Huy",
        duration="60 giây",
        visual_layout_points=[
            "Bên trái (Hình 12): Sơ đồ kiến trúc 3 tầng (Tầng Giao diện Web UI -> Tầng Backend FastAPI chạy 3 file .pkl -> Tầng Kết quả & Khuyến nghị).",
            "Bên phải: 3 tính năng nổi bật của trang Web Dashboard.",
        ],
        spoken_paragraphs=[
            "“Không dừng lại ở việc huấn luyện trên Notebook, nhóm đã đóng gói các mô hình tốt nhất thành ứng dụng Web cảnh báo sớm hoàn chỉnh theo kiến trúc 3 tầng trên Slide 13:”",
            "“Ở tầng Backend FastAPI, hệ thống nạp 3 file mô hình đã huấn luyện gồm: mô hình Soft Voting 3 lớp để dự báo mức độ kiệt sức (Low, Medium, High), mô hình Random Forest cho kịch bản đầy đủ và mô hình XGBoost cho kịch bản cảnh báo sớm. Khi cố vấn học tập nhập thông số sinh viên hoặc chọn 1 trong 3 mẫu hồ sơ có sẵn trên Web, chỉ trong chưa đầy 30 mili-giây, hệ thống sẽ trả về đồng thời xác suất bỏ học, 3 thanh xác suất Burnout, điểm chỉ số Gắn kết vs Căng thẳng và danh sách các khuyến nghị can thiệp cá nhân hóa (như kết nối phòng tham vấn tâm lý, phụ đạo học tập hay hỗ trợ học bổng tài chính).”",
        ],
        pointer_tip="Chỉ theo mũi tên từ Trái sang Phải trên sơ đồ Kiến trúc 3 tầng (Web Input -> FastAPI AI Engine -> Kết quả & Khuyến nghị).",
        fig_name="12_system_architecture_diagram.png",
    )

    # SLIDE 14
    add_slide_section(
        doc,
        slide_no=14,
        slide_title="KẾT LUẬN CHUNG, ĐỊNH HƯỚNG PHÁT TRIỂN & CHUYỂN SANG DEMO WEB",
        speaker="Vũ Quang Huy & Đỗ Quang Minh",
        duration="45 giây (+ 2 phút Demo Web trực tiếp)",
        visual_layout_points=[
            "Bên trái: 4 Đóng góp cốt lõi đã hoàn thành của đồ án.",
            "Bên phải: 4 Định hướng phát triển trong tương lai (Chuỗi thời gian LSTM, SHAP cục bộ, tích hợp LMS/Moodle, mở rộng đa trường).",
            "Banner cảm ơn & Q&A ở dưới đáy.",
        ],
        spoken_paragraphs=[
            "“Cuối cùng, trên Slide 14 là tổng kết 4 đóng góp cốt lõi của Nhóm 03 và định hướng phát triển tiếp theo. Trong tương lai, nhóm định hướng tích hợp giải thích cục bộ SHAP lên giao diện Web và kết nối trực tiếp với hệ thống LMS/Moodle của nhà trường để tự động sàng lọc dữ liệu theo từng tuần học.”",
            "“Ngay sau đây, em xin phép chuyển màn hình sang trình duyệt ở địa chỉ localhost:8000 trong khoảng 1 phút để Demo trực tiếp ứng dụng Web dự báo thời gian thực cho các thầy cô xem ạ. Em xin cảm ơn các thầy cô đã lắng nghe và rất mong nhận được các câu hỏi, góp ý từ Hội đồng ạ!”",
        ],
        pointer_tip="Nói xong câu cuối cùng là nhấn Alt + Tab mở ngay trình duyệt http://localhost:8000 đang chạy sẵn app.py để bấm thử 3 nút Xanh / Cam / Đỏ cho thầy cô xem!",
    )

    doc.save(OUT_ROOT)
    os.makedirs(os.path.dirname(OUT_SIC_VI), exist_ok=True)
    shutil.copy2(OUT_ROOT, OUT_SIC_VI)
    print(f"Successfully created Slide-by-Slide Script Word document at:\n  1. {OUT_ROOT}\n  2. {OUT_SIC_VI}")


if __name__ == "__main__":
    build_slide_script_doc()
