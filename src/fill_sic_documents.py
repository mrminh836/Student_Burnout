"""
Fills out all 5 Samsung Innovation Campus (SIC) AI Capstone Project files for Group 03:
1. SIC_AI_Capstone Project_Action Plan.docx
2. SIC_AI_Capstone Project_Evaluation Sheet.docx
3. SIC_AI_Capstone Project_Final Report.docx
4. SIC_AI_Capstone Project_Work Breakdown Structure.xlsx
5. SIC_AI_Capstone Project_Presentation Slide Template.pptx

Team: Nhóm 03 (Đỗ Quang Minh, Vũ Quang Huy)
Topic: Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học
       dựa trên các chỉ số về căng thẳng và sự tham gia học tập (Student Burnout & Dropout Risk Prediction)
"""

import os
import shutil
import zipfile
import datetime
import xml.etree.ElementTree as ET

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_FOLDER = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"
BACKUP_DIR = os.path.join(BASE_DIR, "templates_backup")
OUT_DIR = os.path.join(BASE_DIR, "SIC_Documents_Completed")
FIG_DIR = os.path.join(BASE_DIR, "figures")

os.makedirs(BACKUP_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

FILES = [
    "SIC_AI_Capstone Project_Action Plan.docx",
    "SIC_AI_Capstone Project_Evaluation Sheet.docx",
    "SIC_AI_Capstone Project_Final Report.docx",
    "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx",
    "SIC_AI_Capstone Project_Presentation Slide Template.pptx",
]

# Backup original files once if not already backed up
for f in FILES:
    src_p = os.path.join(SRC_FOLDER, f)
    bak_p = os.path.join(BACKUP_DIR, f)
    if os.path.exists(src_p) and not os.path.exists(bak_p):
        shutil.copy2(src_p, bak_p)


# ==============================================================================
# 1. FILL ACTION PLAN (.DOCX)
# ==============================================================================
def fill_action_plan():
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Action Plan.docx")
    doc = Document(in_path)

    # Table 0 (10x2)
    t0 = doc.tables[0]
    t0.cell(0, 1).text = "AI Course (Samsung Innovation Campus)"
    t0.cell(1, 1).text = "Nhóm 03 (Group 03 - Student Burnout AI)"
    t0.cell(2, 1).text = "Đỗ Quang Minh (Team Leader) / Vũ Quang Huy (Member)"
    t0.cell(3, 1).text = (
        "Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học "
        "dựa trên các chỉ số về căng thẳng và sự tham gia học tập\n"
        "(Building, Training, and Comparing Machine Learning Models for Student Dropout Risk Prediction "
        "Based on Stress and Study Engagement Indicators)"
    )

    t0.cell(5, 0).text = (
        "1. Phân tích khám phá dữ liệu (EDA) trên bộ dữ liệu 800 sinh viên với 25 thuộc tính về nhân khẩu học, "
        "sự tham gia học tập (Attendance, Study Hours, GPA, Motivation, Backlogs), lối sống và chỉ số căng thẳng tâm lý (Stress, Anxiety, Financial Stress, Burnout).\n"
        "2. Xây dựng quy trình tiền xử lý dữ liệu khuyết (KNNImputer + SimpleImputer) và tạo lập các đặc trưng chuyên sâu (Academic Engagement Index, Psychological Stress Index, Stress-to-Engagement Ratio).\n"
        "3. Huấn luyện, kiểm định chéo 5-Fold (5-Fold Stratified CV) và so sánh toàn diện 10 mô hình học máy (Logistic Regression, SVM, KNN, Decision Tree, Random Forest, Extra Trees, Gradient Boosting, XGBoost, MLP Neural Network, Soft Voting Ensemble) trên 2 kịch bản (Có và Không có biến trung gian Burnout_Level).\n"
        "4. Triển khai giao diện Web Dashboard (FastAPI UI) phục vụ cảnh báo sớm nguy cơ bỏ học và đề xuất giải pháp hỗ trợ tâm lý - học tập cá nhân hóa."
    )

    t0.cell(7, 0).text = (
        "Tình trạng kiệt sức học đường (Student Burnout) và bỏ học đại học (Student Dropout) đang là thách thức lớn đối với các cơ sở giáo dục. "
        "Đồ án này nghiên cứu mối liên hệ định lượng giữa các chỉ số căng thẳng tâm lý (Stress, Anxiety, Financial Stress, Peer Pressure) và sự tham gia học tập (Attendance_Percent, Study_Hours_Per_Day, Motivation_Score, Previous_GPA) nhằm dự báo sớm nguy cơ bỏ học (Dropout_Risk). "
        "Nhóm xây dựng Pipeline học máy khép kín kết hợp kỹ thuật tạo đặc trưng chuyên ngành và so sánh 10 thuật toán phân loại. "
        "Kết quả thực nghiệm cho thấy các mô hình tập hợp cây (Random Forest, Extra Trees) và Soft Voting Ensemble đạt hiệu năng cao nhất với ROC-AUC ~80.5%, F1-Score ~78.9% và Recall trên 81%, giúp nhà trường nhận diện chính xác nhóm sinh viên có nguy cơ bỏ học để can thiệp kịp thời."
    )

    t0.cell(9, 0).text = (
        "• Bước 1 - Tiền xử lý dữ liệu (Data Preprocessing): Xử lý 18/25 cột dữ liệu khuyết bằng KNNImputer (k=7) cho biến số và Mode Imputation cho biến phân loại; chuẩn hóa StandardScaler và mã hóa One-Hot / Ordinal.\n"
        "• Bước 2 - Kỹ thuật đặc trưng (Domain Feature Engineering): Xây dựng 6 chỉ số tổng hợp: Academic_Engagement_Index, Psychological_Stress_Index, Stress_Engagement_Ratio, Net_Stress_Support_Balance, Digital_Fatigue_Ratio, và Academic_Risk_Burden.\n"
        "• Bước 3 - Huấn luyện & So sánh mô hình (Model Training & Comparison): Đánh giá 10 mô hình học máy bằng 5-Fold Stratified Cross-Validation và Hold-out Test (80/20) theo các độ đo Accuracy, Precision, Recall, F1-Score, ROC-AUC và PR-AUC.\n"
        "• Bước 4 - Xây dựng ứng dụng (Web Deployment): Đóng gói Pipeline mô hình tốt nhất (.pkl) và tích hợp vào hệ thống Web Dashboard (FastAPI) dự báo trực tuyến."
    )

    # Table 1 (6x2)
    t1 = doc.tables[1]
    t1.cell(1, 0).text = (
        "• Nguồn dữ liệu: Bộ dữ liệu 'student_burnout_dropout_dataset_2.csv' do chương trình cung cấp (bài toán Student Burnout).\n"
        "• Quy mô & Cấu trúc: 800 bản ghi sinh viên x 25 thuộc tính, bao gồm:\n"
        "  - Nhân khẩu học & Bối cảnh: Age, Gender, Year_of_Study, Department, Residence_Type, Family_Income_Bracket.\n"
        "  - Sự tham gia học tập: Attendance_Percent, Study_Hours_Per_Day, Previous_GPA, Backlogs, Motivation_Score.\n"
        "  - Lối sống & Sinh hoạt: Sleep_Hours, Screen_Time_Hours, Exercise_Freq_Per_Week, Social_Activity_Score, Part_Time_Job.\n"
        "  - Căng thẳng & Tâm lý: Stress_Level, Anxiety_Score, Financial_Stress_Score, Peer_Pressure_Score, Family_Support_Score, Counseling_Access.\n"
        "  - Biến mục tiêu: Burnout_Level (Low: 320, Medium: 280, High: 200) và Dropout_Risk (Yes: 437 [54.6%], No: 363 [45.4%]).\n"
        "• Cách sử dụng: Chia tỷ lệ 80% Train (640 mẫu) / 20% Test (160 mẫu) có phân tầng (Stratified) kết hợp kiểm định chéo 5-Fold trên toàn tập."
    )

    t1.cell(3, 0).text = (
        "• Kết quả đạt được (Expected Outcomes):\n"
        "  1. Bộ báo cáo phân tích EDA trực quan (10 biểu đồ chuẩn xuất bản) làm rõ tác động của Burnout (nhóm High Burnout có tỷ lệ bỏ học tới 92.0%) và chuyên cần/căng thẳng lên nguy cơ bỏ học.\n"
        "  2. Bảng đối sánh thực nghiệm 10 mô hình học máy, chọn ra mô hình tối ưu (Random Forest / Extra Trees / Soft Voting Ensemble đạt ROC-AUC = 80.50%, F1-Score = 78.92%, Recall = 83.91%).\n"
        "  3. Ứng dụng Web Dashboard hoàn chỉnh cho phép nhập thông số sinh viên và dự báo tức thời cả Burnout_Level và Dropout_Risk kèm khuyến nghị can thiệp.\n"
        "• Lợi ích thực tiễn (Benefits): Giúp cố vấn học tập và phòng công tác sinh viên phát hiện sớm sinh viên có nguy cơ bỏ học ngay từ học kỳ đầu để hỗ trợ tâm lý, học bổng và phụ đạo học thuật, giảm tỷ lệ bỏ học tại trường đại học."
    )

    t1.cell(5, 0).text = (
        "1. Đỗ Quang Minh (Team Leader):\n"
        "   - Thiết kế kiến trúc Pipeline học máy và kỹ thuật tạo đặc trưng (Domain Feature Engineering).\n"
        "   - Xây dựng, huấn luyện và tinh chỉnh nhóm mô hình Ensemble & Boosting (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
        "   - Phát triển Backend FastAPI & tích hợp mô hình vào giao diện Web Demo.\n"
        "2. Vũ Quang Huy (Member):\n"
        "   - Khảo sát dữ liệu, xử lý dữ liệu khuyết và thực hiện Phân tích khám phá dữ liệu (EDA & Trực quan hóa dữ liệu).\n"
        "   - Huấn luyện và đánh giá nhóm mô hình cơ sở & phi tuyến (Logistic Regression, SVM, KNN, Decision Tree, MLP Neural Network).\n"
        "   - Thiết kế giao diện người dùng (Frontend Dashboard), tổng hợp kết quả thực nghiệm, biên soạn Final Report và Slide thuyết trình."
    )

    # Table 2 (4x2)
    t2 = doc.tables[2]
    t2.cell(1, 0).text = (
        "• Tuần 1 (Giai đoạn 1 - Khởi động & Khám phá dữ liệu): Xác định yêu cầu bài toán, phân tích thống kê mô tả, kiểm tra dữ liệu khuyết (18/25 cột) và vẽ 6 nhóm biểu đồ EDA.\n"
        "• Tuần 2 (Giai đoạn 2 - Tiền xử lý & Kỹ thuật đặc trưng): Xây dựng Pipeline KNNImputer, chuẩn hóa dữ liệu, thiết kế 6 đặc trưng kết hợp Stress & Engagement.\n"
        "• Tuần 3 (Giai đoạn 3 - Huấn luyện & So sánh mô hình): Huấn luyện 10 mô hình học máy trên 2 kịch bản, tối ưu siêu tham số, phân tích ROC-AUC, Confusion Matrix và Feature Importance.\n"
        "• Tuần 4 (Giai đoạn 4 - Xây dựng Web UI & Hoàn thiện Báo cáo): Phát triển ứng dụng Web Demo (FastAPI), kiểm thử hệ thống, hoàn thiện Final Report, WBS và Slide thuyết trình."
    )
    t2.cell(3, 0).text = "Đề tài có tính thực tiễn cao trong giáo dục đại học; phương pháp luận rõ ràng, so sánh đa dạng các họ mô hình học máy và có sản phẩm giao diện trực quan."

    for path in [
        os.path.join(OUT_DIR, "SIC_AI_Capstone Project_Action Plan.docx"),
        os.path.join(SRC_FOLDER, "SIC_AI_Capstone Project_Action Plan.docx"),
    ]:
        doc.save(path)
    print("Completed: SIC_AI_Capstone Project_Action Plan.docx")


# ==============================================================================
# 2. FILL EVALUATION SHEET (.DOCX)
# ==============================================================================
def fill_evaluation_sheet():
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Evaluation Sheet.docx")
    doc = Document(in_path)

    t0 = doc.tables[0]
    t0.cell(1, 2).text = "Dự đoán tỷ lệ bỏ học dựa trên chỉ số căng thẳng & tham gia học tập (Student Burnout)"
    t0.cell(1, 4).text = "Nhóm 03: Đỗ Quang Minh, Vũ Quang Huy"

    # Mark TRL Level 5/6 in Table 1
    t1 = doc.tables[1]
    t1.cell(5, 2).text = "✓ (Achieved - Prototype Integrated with Web UI)"
    t1.cell(6, 2).text = "✓ (Target - Tested on 800-Student Dataset)"

    for path in [
        os.path.join(OUT_DIR, "SIC_AI_Capstone Project_Evaluation Sheet.docx"),
        os.path.join(SRC_FOLDER, "SIC_AI_Capstone Project_Evaluation Sheet.docx"),
    ]:
        doc.save(path)
    print("Completed: SIC_AI_Capstone Project_Evaluation Sheet.docx")


# ==============================================================================
# 3. FILL WORK BREAKDOWN STRUCTURE (.XLSX)
# ==============================================================================
def fill_wbs_excel():
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx")
    wb = openpyxl.load_workbook(in_path)
    ws = wb["WBS"]

    ws["E7"] = datetime.date(2026, 11, 5)
    ws["E8"] = "Nhóm 03: Đỗ Quang Minh (Leader), Vũ Quang Huy"
    ws["E9"] = "Dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập (Student Burnout)"

    # Set timeline headers (14 working days across W1-W3)
    base_date = datetime.date(2026, 10, 19)
    for d_idx in range(14):
        col = 11 + d_idx
        ws.cell(13, col).value = base_date + datetime.timedelta(days=d_idx)

    # WBS rows (Rows 14 to 29)
    wbs_tasks = [
        # (row, level, text, start_offset, end_offset, owner, deliverable, status)
        (14, 1, "1. Khởi tạo Dự án & Phân tích Yêu cầu (Project Initiation)", 0, 1, "Đỗ Quang Minh, Vũ Quang Huy", "Action Plan & WBS", "Done"),
        (15, 2, "1.1. Xác định mục tiêu bài toán dự báo Burnout & Dropout Risk", 0, 0, "Đỗ Quang Minh", "Đặc tả bài toán", "Done"),
        (16, 2, "1.2. Khảo sát cấu trúc bộ dữ liệu (800 SV x 25 thuộc tính)", 0, 1, "Vũ Quang Huy", "Báo cáo cấu trúc dữ liệu", "Done"),
        (17, 1, "2. Phân tích Khám phá Dữ liệu (Exploratory Data Analysis - EDA)", 2, 4, "Vũ Quang Huy", "Bộ 6 biểu đồ EDA & Nhận xét", "Done"),
        (18, 2, "2.1. Thống kê tỷ lệ dữ liệu khuyết trên 18/25 cột & phân phối nhãn", 2, 2, "Vũ Quang Huy", "Biểu đồ Missing & Target", "Done"),
        (19, 2, "2.2. Phân tích tương quan Pearson & đối sánh Stress vs Engagement", 3, 4, "Vũ Quang Huy, Đỗ Quang Minh", "Heatmap & Violin/Boxplots", "Done"),
        (20, 1, "3. Tiền xử lý & Kỹ thuật Tạo Đặc trưng (Preprocessing & Feature Eng.)", 4, 6, "Đỗ Quang Minh", "Module DomainFeatureEngineer", "Done"),
        (21, 2, "3.1. Xây dựng Pipeline KNNImputer (k=7), StandardScaler, OneHotEncoder", 4, 5, "Vũ Quang Huy", "Preprocessing Pipeline", "Done"),
        (22, 2, "3.2. Thiết kế 6 chỉ số tổng hợp (Engagement Index, Stress Index, Ratio...)", 5, 6, "Đỗ Quang Minh", "6 Engineered Features", "Done"),
        (23, 1, "4. Huấn luyện, Tối ưu & So sánh 10 Mô hình Học máy (Modeling)", 7, 10, "Đỗ Quang Minh, Vũ Quang Huy", "10 Trained Models & Metrics CSV", "Done"),
        (24, 2, "4.1. Huấn luyện nhóm mô hình cơ sở (Logistic, SVM, KNN, DT, MLP)", 7, 8, "Vũ Quang Huy", "Kết quả 5 mô hình cơ sở", "Done"),
        (25, 2, "4.2. Huấn luyện nhóm Ensemble & Boosting (RF, ExtraTrees, GB, XGBoost, Voting)", 8, 9, "Đỗ Quang Minh", "Best Model (.pkl) AUC 80.5%", "Done"),
        (26, 2, "4.3. Đánh giá 2 kịch bản (Có & Không có Burnout_Level) và Feature Importance", 9, 10, "Đỗ Quang Minh, Vũ Quang Huy", "Biểu đồ ROC, PR, CM, Importance", "Done"),
        (27, 1, "5. Phát triển Web Dashboard Demo & Hoàn thiện Báo cáo Cuối kỳ", 11, 13, "Đỗ Quang Minh, Vũ Quang Huy", "FastAPI Web UI, Report & Slides", "Done"),
        (28, 2, "5.1. Xây dựng giao diện Web Demo dự báo trực tiếp Burnout & Dropout Risk", 11, 12, "Đỗ Quang Minh", "Ứng dụng app.py (FastAPI)", "Done"),
        (29, 2, "5.2. Biên soạn Final Report, Slide thuyết trình & Jupyter Notebook", 12, 13, "Vũ Quang Huy, Đỗ Quang Minh", "Final Report.docx & Slide.pptx", "Done"),
    ]

    fill_bar = PatternFill(start_color="2A9D8F", end_color="2A9D8F", fill_type="solid")
    fill_header_task = PatternFill(start_color="E9F5F3", end_color="E9F5F3", fill_type="solid")
    fill_clear = PatternFill(fill_type=None)

    for r, lvl, title, s_off, e_off, owner, deliv, status in wbs_tasks:
        # Clear columns B..E first
        for c in range(2, 6):
            ws.cell(r, c).value = None
        target_col = 2 if lvl == 1 else 3
        ws.cell(r, target_col).value = title
        ws.cell(r, target_col).font = Font(name="Arial", size=10, bold=(lvl == 1))

        s_date = base_date + datetime.timedelta(days=s_off)
        e_date = base_date + datetime.timedelta(days=e_off)
        ws.cell(r, 6).value = s_date
        ws.cell(r, 6).number_format = "yyyy-mm-dd"
        ws.cell(r, 7).value = e_date
        ws.cell(r, 7).number_format = "yyyy-mm-dd"
        ws.cell(r, 8).value = owner
        ws.cell(r, 9).value = deliv
        ws.cell(r, 10).value = status

        for d_idx in range(14):
            c_cell = ws.cell(r, 11 + d_idx)
            if s_off <= d_idx <= e_off:
                c_cell.fill = fill_bar
            else:
                c_cell.fill = fill_header_task if lvl == 1 else fill_clear

    ws.column_dimensions["C"].width = 48
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 13
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 28
    ws.column_dimensions["J"].width = 14

    for path in [
        os.path.join(OUT_DIR, "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx"),
        os.path.join(SRC_FOLDER, "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx"),
    ]:
        wb.save(path)
    print("Completed: SIC_AI_Capstone Project_Work Breakdown Structure.xlsx")


# ==============================================================================
# 4. FILL FINAL REPORT (.DOCX)
# ==============================================================================
def unwrap_sdt_in_docx(src_docx, dst_docx):
    """Unwraps <w:sdt> blocks in word/document.xml so python-docx can edit all headings and paragraphs."""
    with zipfile.ZipFile(src_docx, "r") as zin:
        xml_bytes = zin.read("word/document.xml")
        root = ET.fromstring(xml_bytes)

        def unwrap_element(elem):
            i = 0
            while i < len(elem):
                child = elem[i]
                tag = child.tag.split("}")[-1]
                if tag == "sdt":
                    sdt_content = None
                    for sub in child:
                        if sub.tag.split("}")[-1] == "sdtContent":
                            sdt_content = sub
                            break
                    elem.remove(child)
                    if sdt_content is not None:
                        for idx_s, item in enumerate(list(sdt_content)):
                            elem.insert(i + idx_s, item)
                        continue
                else:
                    unwrap_element(child)
                i += 1

        unwrap_element(root)
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))


def insert_paragraph_after(paragraph, text="", bold=False, italic=False, font_size=10.5, color_rgb=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = paragraph._parent.add_paragraph()
    new_para._p.getparent().remove(new_para._p)
    new_para._p = new_p
    if text:
        run = new_para.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(font_size)
        run.font.name = "Segoe UI"
        if color_rgb:
            run.font.color.rgb = RGBColor(*color_rgb)
    return new_para


def insert_image_after(paragraph, img_path, caption_text, width_inches=6.1):
    p_img = insert_paragraph_after(paragraph, "")
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_img = p_img.add_run()
    r_img.add_picture(img_path, width=Inches(width_inches))

    p_cap = insert_paragraph_after(
        p_img, caption_text, bold=True, italic=True, font_size=9.5, color_rgb=(40, 70, 95)
    )
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p_cap


def fill_final_report():
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Final Report.docx")
    temp_unwrapped = os.path.join(BACKUP_DIR, "_unwrapped_final_report.docx")
    unwrap_sdt_in_docx(in_path, temp_unwrapped)

    doc = Document(temp_unwrapped)

    # 1. Replace Cover Page placeholders
    replacements = {
        "< Project Title >": "DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CHỈ SỐ CĂNG THẲNG & THAM GIA HỌC TẬP\n(Student Burnout & Dropout Risk Prediction)",
        "< Date (DD/MM/YY) >": "05/11/2026",
        "Team Name": "Nhóm 03 - Student Burnout AI",
        "<Member 1>": "1. Đỗ Quang Minh (Team Leader)",
        "<Member 2>": "2. Vũ Quang Huy (Team Member)",
        "<Member 3>": "",
        "<Member 4>": "",
        "<Member 5>": "",
    }
    for p in doc.paragraphs:
        t = p.text.strip()
        if t in replacements:
            p.text = replacements[t]

    # Map section headings in body (only the second occurrence after Table of Contents)
    section_content = {
        "1.1. Background Information": [
            "Trong bối cảnh giáo dục đại học hiện đại, tình trạng sinh viên bỏ học giữa chừng (Student Dropout) và hội chứng kiệt sức học đường (Student Burnout) đang gia tăng mạnh mẽ. Nhiều nghiên cứu giáo dục chỉ ra rằng việc sinh viên rời bỏ giảng đường không chỉ xuất phát từ năng lực học thuật thuần túy mà là hệ quả tích lũy của hai nhóm yếu tố tương tác chặt chẽ: (1) Sự suy giảm mức độ tham gia học tập (Academic Engagement: tỷ lệ chuyên cần, thời lượng tự học, điểm động lực, kết quả GPA kỳ trước, số môn nợ) và (2) Áp lực căng thẳng tâm lý - xã hội kéo dài (Psychological & Financial Stress: căng thẳng học tập, lo âu, áp lực tài chính, áp lực đồng trang lứa, thiếu hụt thời gian ngủ và thiếu điểm tựa gia đình).",
            "Việc ứng dụng Học máy (Machine Learning) để phân tích đồng thời các chỉ số căng thẳng và hành vi tham gia học tập cho phép nhà trường chuyển đổi từ phương pháp xử lý bị động (khi sinh viên đã nộp đơn xin thôi học hoặc bị buộc thôi học) sang cơ chế Cảnh báo sớm chủ động (Proactive Early Warning System), giúp can thiệp đúng đối tượng và đúng thời điểm.",
        ],
        "1.2. Motivation and Objective": [
            "• Động lực nghiên cứu (Motivation): Tại nhiều trường đại học, dữ liệu điểm danh, điểm GPA và khảo sát tâm lý đầu kỳ của sinh viên thường được lưu trữ rời rạc, chưa được khai thác bằng các mô hình dự báo định lượng. Do đó, Nhóm 03 lựa chọn đề tài 'Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập' nhằm giải quyết bài toán thực tiễn này.",
            "• Mục tiêu cụ thể (Objectives):\n"
            "  1. Phân tích khám phá dữ liệu (EDA) trên tập dữ liệu 800 sinh viên x 25 thuộc tính để xác định các yếu tố rủi ro chính dẫn đến kiệt sức (Burnout_Level) và bỏ học (Dropout_Risk).\n"
            "  2. Xây dựng quy trình Tiền xử lý tự động (KNNImputer + StandardScaler + OneHotEncoder) và thiết kế 6 đặc trưng chuyên sâu (Domain Feature Engineering) đo lường tỷ số mất cân bằng giữa Căng thẳng và Gắn kết học tập.\n"
            "  3. Huấn luyện, kiểm định chéo 5-Fold (5-Fold Stratified Cross-Validation) và đối sánh toàn diện 10 thuật toán Học máy trên 2 kịch bản thực nghiệm (Kịch bản 1: Đầy đủ đặc trưng có Burnout_Level; Kịch bản 2: Cảnh báo sớm chỉ dùng chỉ số Stress & Engagement thô).\n"
            "  4. Phát triển giao diện Web Dashboard (FastAPI UI) phục vụ dự báo thời gian thực và đưa ra khuyến nghị can thiệp cá nhân hóa cho từng sinh viên.",
        ],
        "1.3. Members and Role Assignments": [
            "Dự án được thực hiện bởi Nhóm 03 với sự phân công công việc đồng đều, rõ ràng theo đúng thế mạnh của từng thành viên:\n"
            "• Đỗ Quang Minh (Team Leader - 50% khối lượng):\n"
            "  - Phụ trách thiết kế kiến trúc Pipeline học máy chống rò rỉ dữ liệu (No Data Leakage) và xây dựng lớp biến đổi đặc trưng DomainFeatureEngineer (6 chỉ số tổng hợp).\n"
            "  - Huấn luyện, tinh chỉnh siêu tham số và đánh giá nhóm mô hình Ensemble & Boosting (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
            "  - Xây dựng Backend FastAPI và tích hợp mô hình dự báo kép (Burnout_Level & Dropout_Risk) vào ứng dụng Web Demo.\n"
            "• Vũ Quang Huy (Team Member - 50% khối lượng):\n"
            "  - Phụ trách làm sạch dữ liệu, phân tích dữ liệu khuyết trên 18/25 cột và thiết kế bộ 6 biểu đồ Phân tích khám phá dữ liệu (EDA).\n"
            "  - Huấn luyện, tinh chỉnh và đánh giá nhóm mô hình tuyến tính, khoảng cách, cây đơn và mạng nơ-ron (Logistic Regression, SVM RBF, KNN, Decision Tree, MLP Neural Network).\n"
            "  - Thiết kế giao diện người dùng (Frontend Dashboard), tổng hợp bảng số liệu thực nghiệm, biên soạn Báo cáo cuối kỳ (Final Report), WBS và Slide thuyết trình.",
        ],
        "1.4. Schedule and Milestones": [
            "Lộ trình thực hiện dự án được chia thành 4 giai đoạn chính (tương ứng với bảng WBS chi tiết):\n"
            "• Giai đoạn 1 (19/10/2026 – 20/10/2026): Khởi tạo dự án, xác định yêu cầu nghiệp vụ, khảo sát cấu trúc 25 trường dữ liệu và hoàn thiện Action Plan & WBS. [Hoàn thành 100%]\n"
            "• Giai đoạn 2 (21/10/2026 – 25/10/2026): Phân tích khám phá dữ liệu (EDA), kiểm định tương quan Pearson, xử lý dữ liệu khuyết bằng KNNImputer (k=7) và xây dựng 6 đặc trưng chuyên ngành. [Hoàn thành 100%]\n"
            "• Giai đoạn 3 (26/10/2026 – 29/10/2026): Huấn luyện và kiểm định chéo 5-Fold trên 10 mô hình học máy ở 2 kịch bản; xuất các biểu đồ ROC, Precision-Recall, Confusion Matrix và Feature Importance. [Hoàn thành 100%]\n"
            "• Giai đoạn 4 (30/10/2026 – 05/11/2026): Phát triển Web Dashboard Demo (app.py), kiểm thử tích hợp, hoàn thiện Jupyter Notebook, Final Report và Slide thuyết trình. [Hoàn thành 100%]",
        ],
        "2.1. Data Acquisition": [
            "Bộ dữ liệu sử dụng trong đồ án là 'student_burnout_dropout_dataset_2.csv' gồm 800 hồ sơ sinh viên đại học và 25 trường thuộc tính, được chia thành 5 nhóm thông tin:\n"
            "1. Định danh & Nhân khẩu học (7 biến): Student_ID (STU00001–STU00800), Age (17–25 tuổi, trung bình 20.88), Gender (Female: 392, Male: 346, Other: 32, Khuyết: 30), Year_of_Study (Năm 1–4), Department (Engineering: 209, Business: 151, Science: 143, Arts: 136, Medicine: 77, Law: 61), Residence_Type (Hostel: 290, Day Scholar: 286, PG/Rented: 207), Family_Income_Bracket (Low, Lower-Middle, Middle, Upper-Middle, High).\n"
            "2. Chỉ số Tham gia Học tập - Academic Engagement (5 biến): Attendance_Percent (30%–100%, TB: 76.88%), Study_Hours_Per_Day (0.0–8.1 giờ/ngày, TB: 3.16h), Previous_GPA (2.42–10.0, TB: 6.86), Backlogs (0–6 môn nợ, TB: 0.82), Motivation_Score (0–10, TB: 5.47).\n"
            "3. Chỉ số Lối sống & Sinh hoạt (5 biến): Sleep_Hours (2.0–10.0h, TB: 6.20h), Screen_Time_Hours (0.5–12.6h, TB: 5.43h), Exercise_Freq_Per_Week (0–7 buổi/tuần), Social_Activity_Score (0–10), Part_Time_Job (No: 547, Yes: 228).\n"
            "4. Chỉ số Căng thẳng & Hỗ trợ Tâm lý (6 biến): Stress_Level (0.1–10.0, TB: 5.52), Anxiety_Score (0–10, TB: 4.98), Financial_Stress_Score (0–10, TB: 5.03), Peer_Pressure_Score (0–9.8, TB: 4.48), Family_Support_Score (0–10, TB: 6.42), Counseling_Access (No: 535, Yes: 265).\n"
            "5. Biến Kiệt sức & Biến Mục tiêu (2 biến): Burnout_Level (Low: 320 [40.0%], Medium: 280 [35.0%], High: 200 [25.0%]) và biến mục tiêu cần dự đoán Dropout_Risk (Yes: 437 sinh viên [54.62%], No: 363 sinh viên [45.38%]).",
        ],
        "2.2. Training Methodology": [
            "Nhóm thiết kế phương pháp huấn luyện nghiêm ngặt nhằm đảm bảo tính khách quan và khả năng tổng quát hóa cao:\n"
            "• Phân chia dữ liệu (Data Splitting): Tập dữ liệu 800 mẫu được chia theo tỷ lệ 80% Huấn luyện (640 mẫu) và 20% Kiểm thử độc lập (160 mẫu) có phân tầng theo nhãn (Stratified Train-Test Split, random_state=42) để giữ nguyên tỷ lệ 54.6% Yes / 45.4% No.\n"
            "• Kiểm định chéo phân tầng 5-Fold (5-Fold Stratified Cross-Validation): Toàn bộ quá trình tạo đặc trưng (Feature Engineering), điền khuyết (Imputation), chuẩn hóa (Scaling) và mã hóa (Encoding) đều được đóng gói bên trong đối tượng Pipeline của scikit-learn để tính toán tham số thống kê chỉ trên tập Train của từng Fold, triệt tiêu hoàn toàn hiện tượng rò rỉ dữ liệu (Data Leakage).\n"
            "• Thiết kế 2 kịch bản thực nghiệm đối chứng:\n"
            "  - Kịch bản 1 (Full Feature Set): Sử dụng toàn bộ đặc trưng đầu vào bao gồm cả trạng thái Burnout_Level và 6 chỉ số tổng hợp.\n"
            "  - Kịch bản 2 (Early Warning - Không dùng Burnout_Level): Dự báo Dropout_Risk trực tiếp từ các chỉ số hành vi học tập và căng thẳng tâm lý thô khi chưa có đánh giá Burnout_Level.\n"
            "• Hệ thống 10 mô hình học máy được huấn luyện và so sánh: (1) Logistic Regression, (2) Support Vector Machine (RBF Kernel + Probability Calibration), (3) K-Nearest Neighbors (KNN), (4) Decision Tree, (5) Random Forest, (6) Extra Trees, (7) Gradient Boosting, (8) XGBoost, (9) Multi-Layer Perceptron (MLP Neural Network), và (10) Soft Voting Ensemble.",
        ],
        "2.3. Workflow": [
            "Quy trình thực thi tổng thể của hệ thống học máy Nhóm 03 gồm 5 bước tuần tự khép kín từ thu thập dữ liệu, tiền xử lý & khám phá dữ liệu, kỹ thuật tạo đặc trưng chuyên ngành, huấn luyện & kiểm định chéo 10 mô hình đến triển khai ứng dụng Web cảnh báo sớm:",
        ],
        "2.4. System Design": [
            "Hệ thống phần mềm được thiết kế theo kiến trúc 3 tầng (3-Tier Architecture): (1) Tầng Giao diện Web Dashboard thu thập thông số học tập & tâm lý của sinh viên; (2) Tầng Xử lý AI Backend (FastAPI) thực thi DomainFeatureEngineer, KNNImputer Pipeline và bộ đôi mô hình dự báo Burnout_Level + Dropout_Risk; (3) Tầng Hiển thị Kết quả & Khuyến nghị can thiệp sớm.",
        ],
        "3.1. Data Preprocessing": [
            "1. Xử lý dữ liệu khuyết (Missing Data Imputation):\n"
            "   Qua khảo sát, có 18/25 cột chứa giá trị khuyết (Missing Values) với tỷ lệ từ 1.62% (Family_Income_Bracket - 13 mẫu) đến 6.38% (Attendance_Percent - 51 mẫu) và 6.12% (Stress_Level - 49 mẫu). Thay vì xóa bỏ các dòng khuyết (sẽ làm mất hơn 40% số bản ghi nếu dropna toàn bộ), nhóm áp dụng:\n"
            "   - Đối với các biến số liên tục (Numerical Features): Sử dụng thuật toán KNNImputer (n_neighbors=7) để ước lượng giá trị khuyết dựa trên 7 sinh viên có hồ sơ tương đồng nhất trong không gian đặc trưng chuẩn hóa.\n"
            "   - Đối với các biến phân loại (Categorical Features): Điền giá trị xuất hiện phổ biến nhất (Most Frequent / Mode Imputation).\n"
            "2. Kỹ thuật Tạo Đặc trưng Chuyên sâu (Domain Feature Engineering):\n"
            "   Nhóm xây dựng 6 đặc trưng mới phản ánh trực tiếp bản chất bài toán:\n"
            "   • Academic_Engagement_Index = 0.35*(Attendance_Percent/10) + 0.25*Study_Hours_Per_Day + 0.25*Motivation_Score + 0.15*Previous_GPA\n"
            "   • Psychological_Stress_Index = 0.35*Stress_Level + 0.30*Anxiety_Score + 0.20*Financial_Stress_Score + 0.15*Peer_Pressure_Score\n"
            "   • Stress_Engagement_Ratio = Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)\n"
            "   • Net_Stress_Support_Balance = Psychological_Stress_Index - (0.5*Family_Support_Score + 0.3*Social_Activity_Score + 0.2*Exercise_Freq_Per_Week)\n"
            "   • Digital_Fatigue_Ratio = Screen_Time_Hours / (Sleep_Hours + 1.0)\n"
            "   • Academic_Risk_Burden = 1.5*Backlogs + (100 - Attendance_Percent)/20.0\n"
            "3. Mã hóa & Chuẩn hóa (Encoding & Scaling): Mã hóa thứ tự (Ordinal Encoding) cho Family_Income_Bracket (0–4) và Burnout_Level (Low=0, Medium=1, High=2); mã hóa One-Hot (OneHotEncoder) cho các biến danh목 (Gender, Department, Residence_Type, Part_Time_Job, Counseling_Access); chuẩn hóa toàn bộ biến số bằng StandardScaler (μ=0, σ=1).",
        ],
        "3.2. Exploratory Data Analysis (EDA)": [
            "Quá trình phân tích khám phá dữ liệu (EDA) đã làm sáng tỏ các quy luật quan trọng trong bộ dữ liệu:\n"
            "• Phát hiện 1 - Mối quan hệ quyết định giữa Burnout_Level và Dropout_Risk: Trong số 200 sinh viên có mức độ kiệt sức cao (Burnout_Level = High), có tới 184 sinh viên bỏ học (chiếm 92.00%). Ở mức Medium, tỷ lệ bỏ học là 63.21% (177/280), trong khi ở mức Low chỉ là 23.75% (76/320). Hệ số tương quan Pearson giữa Burnout_Numeric và Dropout_Binary đạt r = +0.550.\n"
            "• Phát hiện 2 - Sự đối lập giữa nhóm Tham gia Học tập và nhóm Căng thẳng Tâm lý:\n"
            "  - Các yếu tố bảo vệ (giảm nguy cơ bỏ học - tương quan âm): Attendance_Percent (r = -0.244), Motivation_Score (r = -0.193), Study_Hours_Per_Day (r = -0.160), Family_Support_Score (r = -0.137), Previous_GPA (r = -0.106).\n"
            "  - Các yếu tố nguy cơ (tăng nguy cơ bỏ học - tương quan dương): Stress_Level (r = +0.164), Anxiety_Score (r = +0.145), Screen_Time_Hours (r = +0.127), Financial_Stress_Score (r = +0.114).\n"
            "• Phát hiện 3 - Phân tích theo nhóm nhân khẩu học: Sinh viên có thu nhập gia đình thấp (Low Income) có tỷ lệ bỏ học cao nhất (62.1%), giảm dần khi thu nhập tăng (High Income: 51.1%).",
        ],
        "3.3. Modeling": [
            "Nhóm tiến hành huấn luyện và so sánh 10 mô hình học máy trên cả 2 kịch bản. Dưới đây là tổng hợp kết quả thực nghiệm (5-Fold CV và Tập Test 20%):\n"
            "1. Kết quả Kịch bản 1 (Full Features - Có Burnout_Level & Chỉ số Tổng hợp):\n"
            "   • Random Forest: CV Accuracy = 74.38% | CV F1 = 77.46% | CV ROC-AUC = 80.50% (Cao nhất) | Test ROC-AUC = 80.24%\n"
            "   • Extra Trees: CV Accuracy = 75.00% (Cao nhất) | CV F1 = 78.10% | CV ROC-AUC = 79.22% | Test Accuracy = 75.63% | Test F1 = 78.92% | Test Recall = 83.91%\n"
            "   • Soft Voting Ensemble: CV Accuracy = 73.88% | CV F1 = 76.79% | CV ROC-AUC = 79.95% | Test Accuracy = 72.50% | Test F1 = 75.82% | Test ROC-AUC = 79.37%\n"
            "   • Logistic Regression: CV Accuracy = 72.25% | CV F1 = 75.19% | CV ROC-AUC = 79.82% | Test Accuracy = 73.75% | Test ROC-AUC = 79.86%\n"
            "   • SVM (RBF Kernel): CV Accuracy = 74.63% | CV F1 = 77.57% | CV ROC-AUC = 78.55% | Test Accuracy = 74.38% | Test F1 = 77.84%\n"
            "   • MLP Neural Network: CV Accuracy = 73.63% | CV F1 = 75.94% | CV ROC-AUC = 79.68%\n"
            "   • XGBoost: CV Accuracy = 71.75% | CV F1 = 74.28% | CV ROC-AUC = 79.25%\n"
            "   • Gradient Boosting: CV Accuracy = 71.75% | CV F1 = 74.44% | CV ROC-AUC = 79.08%\n"
            "   • K-Nearest Neighbors: CV Accuracy = 70.25% | CV F1 = 72.59% | CV ROC-AUC = 77.14%\n"
            "   • Decision Tree: CV Accuracy = 66.88% | CV F1 = 67.75% | CV ROC-AUC = 73.88%\n\n"
            "2. Kết quả Kịch bản 2 (Early Warning - Không dùng Burnout_Level, chỉ dùng chỉ số Stress & Engagement):\n"
            "   • Soft Voting Ensemble dẫn đầu với CV Accuracy = 70.63%, CV F1 = 73.89%, CV ROC-AUC = 77.10% (Test Accuracy = 70.63%, Test ROC-AUC = 77.63%).\n"
            "   • Logistic Regression đạt CV ROC-AUC = 76.82% và Test ROC-AUC = 79.00%.\n"
            "   • XGBoost và Random Forest đạt Test Accuracy = 70.00%, chứng tỏ các đặc trưng tổng hợp (Engagement Index, Stress Index, Stress-Engagement Ratio) giúp mô hình duy trì khả năng dự báo mạnh mẽ ngay cả khi ẩn biến Burnout_Level.",
        ],
        "3.4. User Interface": [
            "Để đưa mô hình vào ứng dụng thực tế, Nhóm 03 đã xây dựng hệ thống Web Dashboard tương tác bằng FastAPI (tệp app.py). Giao diện cho phép cố vấn học tập hoặc sinh viên:\n"
            "• Điều chỉnh trực quan 18 thanh trượt/danh sách chọn về Chuyên cần (Attendance %), Giờ tự học (Study Hours), GPA, Động lực, Số môn nợ, Căng thẳng, Lo âu, Áp lực tài chính, Hỗ trợ gia đình, Giờ ngủ và Giờ màn hình.\n"
            "• Chọn nhanh 3 mẫu hồ sơ mẫu (Mẫu SV Tích cực - Nguy cơ Thấp; Mẫu SV Căng thẳng TB; Mẫu SV Kiệt sức - Nguy cơ Cao) để kiểm thử tức thời.\n"
            "• Hiển thị kết quả dự báo kép theo thời gian thực: (1) Xác suất nguy cơ bỏ học (%) kèm phân tầng màu sắc cảnh báo; (2) Phân bố xác suất 3 mức độ kiệt sức (Low / Medium / High Burnout); (3) Ba chỉ số giải thích mô hình (Academic Engagement Index, Psychological Stress Index, Stress-to-Engagement Ratio) và (4) Danh sách khuyến nghị can thiệp cá nhân hóa.",
        ],
        "3.5. Testing and Improvements": [
            "• Quá trình Kiểm thử & Cải tiến chất lượng mô hình:\n"
            "  1. Cải tiến xử lý dữ liệu khuyết: Thay thế SimpleImputer (median) bằng KNNImputer (k=7) kết hợp chuẩn hóa giúp giữ vững mối liên hệ đa biến giữa Attendance, Study Hours và Stress.\n"
            "  2. Cải tiến Feature Engineering: Việc bổ sung 6 đặc trưng tương tác (đặc biệt là Stress_Engagement_Ratio và Academic_Engagement_Index) đã nâng điểm ROC-AUC của Random Forest từ 80.02% lên 80.50% và giúp Extra Trees đạt độ chính xác Test 75.63% (F1-Score đạt 78.92%, Recall đạt 83.91%).\n"
            "  3. Kiểm soát Overfitting trên tập dữ liệu 800 mẫu: Giới hạn độ sâu cây (max_depth=3..8), tăng min_samples_leaf và sử dụng điều chuẩn L2 (reg_lambda=2.0 cho XGBoost, C=0.5 cho Logistic Regression), giúp sai lệch giữa CV và Test luôn dưới 1.5%.",
        ],
        "4.1. Accomplishments and Benefits": [
            "• Thành tựu đạt được (Accomplishments):\n"
            "  - Hoàn thiện Pipeline Học máy chuẩn công nghiệp từ EDA, Imputation, Domain Feature Engineering đến huấn luyện & so sánh 10 thuật toán học máy trên 2 kịch bản.\n"
            "  - Xác lập được bộ chỉ số định lượng đánh giá mức độ gắn kết học tập (Academic Engagement Index) và căng thẳng tâm lý (Psychological Stress Index), giải thích rõ ràng nguyên nhân gốc rễ dẫn đến bỏ học.\n"
            "  - Xây dựng thành công Web Dashboard Demo hoạt động mượt mà với thời gian phản hồi < 30ms/yêu cầu.\n"
            "• Lợi ích mang lại (Benefits):\n"
            "  - Đối với Nhà trường & Cố vấn học tập: Công cụ sàng lọc tự động hàng nghìn sinh viên mỗi học kỳ, ưu tiên nguồn lực tư vấn cho nhóm sinh viên có xác suất bỏ học > 65%.\n"
            "  - Đối với Sinh viên: Giúp sinh viên tự đánh giá mức độ cân bằng giữa học tập - sinh hoạt - tâm lý và sớm tiếp cận các dịch vụ tham vấn tâm lý học đường.",
        ],
        "4.2. Future Improvements": [
            "Trong tương lai, nhóm định hướng mở rộng đề tài theo 4 hướng:\n"
            "1. Thu thập dữ liệu chuỗi thời gian (Longitudinal / Time-Series Data) theo từng tuần học trong học kỳ để dự báo xu hướng biến đổi của Stress và Attendance bằng mô hình LSTM / Transformer.\n"
            "2. Tích hợp thư viện SHAP (SHapley Additive exPlanations) trực tiếp lên giao diện Web để giải thích đóng góp của từng biến đối với từng cá nhân sinh viên.\n"
            "3. Kết nối trực tiếp với Hệ thống Quản lý Học tập (LMS / Moodle / Canvas) và hệ thống điểm danh thông minh của nhà trường.\n"
            "4. Mở rộng quy mô dữ liệu lên 5,000 – 10,000 sinh viên từ nhiều trường đại học tại Việt Nam để hiệu chỉnh mô hình sát với đặc thù sinh viên trong nước.",
        ],
    }

    # Image attachments per section
    section_images = {
        "2.3. Workflow": [
            (os.path.join(FIG_DIR, "11_system_workflow_diagram.png"), "Hình 2.1: Quy trình thực hiện dự án Học máy (ML Project Workflow - Nhóm 03)"),
        ],
        "2.4. System Design": [
            (os.path.join(FIG_DIR, "12_system_architecture_diagram.png"), "Hình 2.2: Kiến trúc hệ thống Cảnh báo sớm Nguy cơ Bỏ học & Giao diện Web UI"),
        ],
        "3.1. Data Preprocessing": [
            (os.path.join(FIG_DIR, "01_missing_values_analysis.png"), "Hình 3.1: Phân tích tỷ lệ dữ liệu khuyết trên 18/25 thuộc tính trước khi qua KNNImputer"),
        ],
        "3.2. Exploratory Data Analysis (EDA)": [
            (os.path.join(FIG_DIR, "02_target_and_burnout_distribution.png"), "Hình 3.2: Phân phối biến mục tiêu Dropout_Risk và tỷ lệ bỏ học theo từng mức độ Burnout_Level"),
            (os.path.join(FIG_DIR, "03_stress_vs_engagement_distributions.png"), "Hình 3.3: So sánh phân phối các chỉ số Tham gia Học tập và Căng thẳng Tâm lý theo Dropout_Risk"),
            (os.path.join(FIG_DIR, "04_correlation_matrix_heatmap.png"), "Hình 3.4: Ma trận tương quan Pearson giữa các chỉ số Căng thẳng, Tham gia Học tập, Burnout và Dropout"),
            (os.path.join(FIG_DIR, "05_categorical_factors_analysis.png"), "Hình 3.5: Phân tích tỷ lệ bỏ học theo các yếu tố nhân khẩu học và môi trường học tập"),
            (os.path.join(FIG_DIR, "06_multivariate_burnout_engagement_scatter.png"), "Hình 3.6: Phân tích đa biến giữa Tỷ lệ chuyên cần, Căng thẳng tâm lý và Hồ sơ Burnout"),
        ],
        "3.3. Modeling": [
            (os.path.join(FIG_DIR, "07_model_performance_comparison.png"), "Hình 3.7: So sánh hiệu năng 10 mô hình học máy trên Kịch bản 1 (Full) và Kịch bản 2 (Early Warning)"),
            (os.path.join(FIG_DIR, "08_roc_and_precision_recall_curves.png"), "Hình 3.8: Đường cong ROC và Precision-Recall của 10 mô hình trên tập kiểm thử (Test Set)"),
            (os.path.join(FIG_DIR, "09_confusion_matrices_top_models.png"), "Hình 3.9: Ma trận nhầm lẫn (Confusion Matrix) của 4 mô hình tiêu biểu trên tập kiểm thử"),
            (os.path.join(FIG_DIR, "10_feature_importance_analysis.png"), "Hình 3.10: Phân tích mức độ quan trọng của các đặc trưng (Feature Importance) trên Random Forest & XGBoost"),
        ],
    }

    # Find the second occurrence of each heading (the body section, not TOC)
    seen_counts = {}
    for p in list(doc.paragraphs):
        txt = p.text.strip()
        if txt in section_content:
            seen_counts[txt] = seen_counts.get(txt, 0) + 1
            if seen_counts[txt] >= 2 or txt in ["2.4. System Design"]:
                curr_p = p
                for para_text in section_content[txt]:
                    curr_p = insert_paragraph_after(curr_p, para_text, font_size=10.5)
                if txt in section_images:
                    for img_file, cap in section_images[txt]:
                        if os.path.exists(img_file):
                            curr_p = insert_image_after(curr_p, img_file, cap, width_inches=6.0)

    # Fill Team Member Review Table dynamically (6x2 table with NAME | REVIEW and COMMENT)
    for t_mem in doc.tables:
        if len(t_mem.rows) == 6 and len(t_mem.columns) == 2:
            t_mem.cell(0, 0).text = "NAME"
            t_mem.cell(0, 1).text = "REVIEW and COMMENT"
            t_mem.cell(1, 0).text = "Đỗ Quang Minh\n(Team Leader)"
            t_mem.cell(1, 1).text = (
                "Qua dự án Capstone của khóa học Samsung Innovation Campus (SIC AI), tôi đã nắm vững cách xây dựng một Pipeline Học máy hoàn chỉnh và chống rò rỉ dữ liệu. "
                "Việc thiết kế các chỉ số kết hợp giữa Căng thẳng tâm lý (Stress Index) và Sự tham gia học tập (Engagement Index) giúp nâng cao rõ rệt hiệu năng của các mô hình Ensemble (Random Forest, Extra Trees, Soft Voting) và mang lại ý nghĩa thực tiễn cao trong giáo dục."
            )
            t_mem.cell(2, 0).text = "Vũ Quang Huy\n(Team Member)"
            t_mem.cell(2, 1).text = (
                "Đồ án giúp tôi rèn luyện kỹ năng phân tích khám phá dữ liệu (EDA), xử lý dữ liệu khuyết bằng KNNImputer và đối sánh thực nghiệm 10 thuật toán phân loại khác nhau. "
                "Sự phối hợp chặt chẽ trong nhóm từ khâu phân tích dữ liệu, huấn luyện mô hình đến triển khai giao diện Web Demo đã giúp nhóm hoàn thành tốt mọi mục tiêu đề ra."
            )
            break

    for path in [
        os.path.join(OUT_DIR, "SIC_AI_Capstone Project_Final Report.docx"),
        os.path.join(SRC_FOLDER, "SIC_AI_Capstone Project_Final Report.docx"),
    ]:
        doc.save(path)

    if os.path.exists(temp_unwrapped):
        os.remove(temp_unwrapped)
    print("Completed: SIC_AI_Capstone Project_Final Report.docx")


# ==============================================================================
# 5. FILL PRESENTATION SLIDE TEMPLATE (.PPTX)
# ==============================================================================
def add_content_slide(prs, layout, unit_num, unit_title, sec_num, slide_title, subtitle_text, bullets=None, img_path=None, table_data=None):
    slide = prs.slides.add_slide(layout)

    # Populate placeholders by idx or name
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        nm = shape.name
        txt = shape.text_frame.text.strip()
        if nm == "제목 1" or txt == "Slide Title":
            shape.text_frame.text = slide_title
        elif nm == "텍스트 개체 틀 2" or txt == "1.1":
            shape.text_frame.text = sec_num
        elif nm == "텍스트 개체 틀 3" or txt == "Subtitle":
            shape.text_frame.text = subtitle_text
        elif nm == "텍스트 개체 틀 4" or txt == "01":
            shape.text_frame.text = unit_num
        elif nm == "텍스트 개체 틀 5" or txt == "UNIT":
            shape.text_frame.text = unit_title
        elif nm == "텍스트 개체 틀 6" or "Level-1" in txt:
            if bullets:
                tf = shape.text_frame
                tf.clear()
                for i, b in enumerate(bullets):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = b
                    p.font.size = PptPt(13.5 if img_path or table_data else 15)
                    p.space_after = PptPt(6)
                if img_path:
                    shape.width = PptInches(4.6)
            else:
                shape.text_frame.text = ""

    # Remove yellow instruction callout if present
    for shape in list(slide.shapes):
        if shape.has_text_frame and "Please edit project name" in shape.text_frame.text:
            sp = shape._element
            sp.getparent().remove(sp)

    if img_path and os.path.exists(img_path):
        left = PptInches(5.35) if bullets else PptInches(1.2)
        top = PptInches(1.85)
        width = PptInches(7.5) if bullets else PptInches(10.8)
        slide.shapes.add_picture(img_path, left, top, width=width)

    if table_data:
        rows = len(table_data)
        cols = len(table_data[0])
        left = PptInches(0.65)
        top = PptInches(2.55) if bullets else PptInches(1.85)
        width = PptInches(12.0)
        height = PptInches(0.38 * rows)
        t_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = t_shape.table
        for r_i, row in enumerate(table_data):
            for c_i, val in enumerate(row):
                cell = table.cell(r_i, c_i)
                cell.text = str(val)
                for p in cell.text_frame.paragraphs:
                    p.font.size = PptPt(10.5 if rows > 8 else 11.5)
                    p.font.bold = (r_i == 0 or r_i == 1)
                    if r_i == 0:
                        p.font.color.rgb = PptRGBColor(255, 255, 255)
                if r_i == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = PptRGBColor(29, 53, 87)

    return slide


def fill_presentation_pptx():
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx")
    prs = Presentation(in_path)

    # Slide 1: Cover Slide
    s1 = prs.slides[0]
    for shape in s1.shapes:
        if shape.has_text_frame:
            if shape.name == "텍스트 개체 틀 4":
                shape.text_frame.text = "Nhóm 03: Đỗ Quang Minh • Vũ Quang Huy"
            elif shape.name == "제목 3":
                shape.text_frame.text = (
                    "DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CHỈ SỐ CĂNG THẲNG & SỰ THAM GIA HỌC TẬP\n"
                    "(Student Burnout & Dropout Risk Prediction)"
                )
                for p in shape.text_frame.paragraphs:
                    p.font.size = PptPt(24)

    # Slide 2: Agenda / Table of Contents
    s2 = prs.slides[1]
    for shape in s2.shapes:
        if shape.has_text_frame and "Project Name" in shape.text_frame.text:
            shape.text_frame.text = "Student Burnout & Dropout Risk Prediction — Nhóm 03"

    agenda_box = s2.shapes.add_textbox(PptInches(1.2), PptInches(2.1), PptInches(10.8), PptInches(4.5))
    tf_ag = agenda_box.text_frame
    tf_ag.word_wrap = True
    agenda_items = [
        "01. GIỚI THIỆU ĐỀ TÀI (INTRODUCTION): Bối cảnh, Động lực, Mục tiêu & Phân công Nhóm 03",
        "02. DỮ LIỆU & PHƯƠNG PHÁP LUẬN (METHODOLOGY & WORKFLOW): Bộ dữ liệu 800 SV x 25 biến & Kiến trúc Pipeline",
        "03. TIỀN XỬ LÝ & KỸ THUẬT ĐẶC TRƯNG (PREPROCESSING & FEATURE ENGINEERING): KNNImputer & 6 Chỉ số Stress-Engagement",
        "04. PHÂN TÍCH KHÁM PHÁ DỮ LIỆU (EXPLORATORY DATA ANALYSIS): Tác động của Burnout, Chuyên cần & Căng thẳng tâm lý",
        "05. HUẤN LUYỆN & SO SÁNH 10 MÔ HÌNH HỌC MÁY (MODEL COMPARISON): Kết quả 2 kịch bản, ROC-AUC & Feature Importance",
        "06. GIAO DIỆN WEB DEMO & KẾT LUẬN (WEB UI & IMPACT): Hệ thống Cảnh báo sớm FastAPI & Định hướng phát triển",
    ]
    for i, item in enumerate(agenda_items):
        p = tf_ag.paragraphs[0] if i == 0 else tf_ag.add_paragraph()
        p.text = item
        p.font.size = PptPt(16)
        p.font.bold = True
        p.font.color.rgb = PptRGBColor(29, 53, 87)
        p.space_after = PptPt(14)

    # Slide 3: Use existing Slide 3 as Section 1.1 Introduction & Objectives
    s3 = prs.slides[2]
    content_layout = s3.slide_layout
    for shape in list(s3.shapes):
        if shape.has_text_frame and "Please edit project name" in shape.text_frame.text:
            shape._element.getparent().remove(shape._element)
        elif shape.has_text_frame:
            nm = shape.name
            if nm == "제목 1":
                shape.text_frame.text = "Bối cảnh, Mục tiêu & Phân công Nhóm 03"
            elif nm == "텍스트 개체 틀 2":
                shape.text_frame.text = "1.1"
            elif nm == "텍스트 개체 틀 3":
                shape.text_frame.text = "Dự báo sớm nguy cơ bỏ học đại học từ chỉ số Căng thẳng (Stress) và Sự tham gia học tập (Engagement)"
            elif nm == "텍스트 개체 틀 4":
                shape.text_frame.text = "01"
            elif nm == "텍스트 개체 틀 5":
                shape.text_frame.text = "INTRODUCTION"
            elif nm == "텍스트 개체 틀 6":
                tf = shape.text_frame
                tf.clear()
                b_list = [
                    "• Bài toán thực tiễn: Sinh viên bỏ học (Dropout) là hệ quả tích lũy của Kiệt sức học đường (Burnout), áp lực tâm lý - tài chính kéo dài và sự suy giảm gắn kết học tập.",
                    "• Mục tiêu đồ án:\n  1. Khám phá quy luật dữ liệu (EDA) trên 800 sinh viên x 25 thuộc tính.\n  2. Xây dựng Pipeline xử lý khuyết (KNNImputer) & tạo 6 đặc trưng chuyên sâu (Stress vs. Engagement).\n  3. Huấn luyện & so sánh 10 mô hình học máy trên 2 kịch bản (Có và Không có biến Burnout_Level).\n  4. Triển khai Web Dashboard cảnh báo sớm trực quan.",
                    "• Thành viên Nhóm 03:\n  - Đỗ Quang Minh (Leader): Thiết kế Feature Engineering, huấn luyện nhóm Ensemble/Boosting (RF, ExtraTrees, XGBoost, Voting) & phát triển Backend FastAPI.\n  - Vũ Quang Huy (Member): Phân tích EDA & xử lý khuyết, huấn luyện nhóm mô hình cơ sở (LR, SVM, KNN, DT, MLP), thiết kế UI & biên soạn báo cáo.",
                ]
                for idx_b, line in enumerate(b_list):
                    p = tf.paragraphs[0] if idx_b == 0 else tf.add_paragraph()
                    p.text = line
                    p.font.size = PptPt(14)
                    p.space_after = PptPt(8)

    # Remove the Closing Slide temporarily so we can append it at the very end
    closing_layout = prs.slides[3].slide_layout
    rId = prs.slides._sldIdLst[3].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[3]

    # Add Slide 4: Dataset Overview & Missing Data Analysis
    add_content_slide(
        prs, content_layout, "02", "DATA & PREPROCESSING", "2.1",
        "Tổng quan Bộ dữ liệu & Xử lý Dữ liệu Khuyết",
        "Bộ dữ liệu student_burnout_dropout_dataset_2.csv gồm 800 sinh viên x 25 thuộc tính (18 cột có Missing Values)",
        bullets=[
            "• Quy mô: 800 sinh viên, 25 cột chia thành 5 nhóm:",
            "  1. Nhân khẩu học (Age, Gender, Year, Dept, Residence, Income)",
            "  2. Tham gia học tập (Attendance %, Study Hours, GPA, Backlogs, Motivation)",
            "  3. Lối sống (Sleep, Screen Time, Exercise, Social, Part-time Job)",
            "  4. Căng thẳng & Tâm lý (Stress, Anxiety, Financial Stress, Peer Pressure, Support)",
            "  5. Biến mục tiêu: Burnout_Level (Low/Med/High) & Dropout_Risk (Yes: 54.6%, No: 45.4%)",
            "• Giải pháp xử lý 18 cột khuyết (1.6% - 6.4%):",
            "  - Biến liên tục: KNNImputer (k=7) trong Pipeline",
            "  - Biến phân loại: Most Frequent Imputation + OneHot/Ordinal Encoding",
        ],
        img_path=os.path.join(FIG_DIR, "01_missing_values_analysis.png"),
    )

    # Add Slide 5: Workflow & Feature Engineering
    add_content_slide(
        prs, content_layout, "02", "METHODOLOGY", "2.2",
        "Quy trình Học máy & Kỹ thuật Tạo Đặc trưng (Feature Engineering)",
        "Thiết kế 6 chỉ số tổng hợp đo lường sự tương tác giữa Căng thẳng Tâm lý và Sự Tham gia Học tập",
        bullets=[
            "• 6 Đặc trưng chuyên sâu (DomainFeatureEngineer):",
            "  1. Academic_Engagement_Index: Tổ hợp Attendance, Study Hours, Motivation & GPA.",
            "  2. Psychological_Stress_Index: Tổ hợp Stress, Anxiety, Financial & Peer Pressure.",
            "  3. Stress_Engagement_Ratio: Tỷ số mất cân bằng giữa Áp lực và Gắn kết học tập.",
            "  4. Net_Stress_Support_Balance: Áp lực trừ điểm tựa Gia đình, Xã hội & Thể chất.",
            "  5. Digital_Fatigue_Ratio: Tỷ lệ Screen_Time_Hours / Sleep_Hours.",
            "  6. Academic_Risk_Burden: Gánh nặng nợ môn (Backlogs) & thiếu hụt chuyên cần.",
        ],
        img_path=os.path.join(FIG_DIR, "11_system_workflow_diagram.png"),
    )

    # Add Slide 6: EDA 1 - Target & Burnout Relationship
    add_content_slide(
        prs, content_layout, "03", "EDA RESULTS", "3.1",
        "Phân tích Biến mục tiêu: Tác động của Burnout lên Nguy cơ Bỏ học",
        "Mối tương quan mạnh mẽ giữa Mức độ Kiệt sức (Burnout_Level) và Tỷ lệ Bỏ học (r = +0.550)",
        bullets=[
            "• Phân bố nhãn cân bằng: Yes = 437 SV (54.6%), No = 363 SV (45.4%).",
            "• Tỷ lệ bỏ học tăng vọt theo mức độ Burnout:",
            "  - Low Burnout (320 SV): Chỉ 23.8% bỏ học (76 SV).",
            "  - Medium Burnout (280 SV): 63.2% bỏ học (177 SV).",
            "  - High Burnout (200 SV): Lên tới 92.0% bỏ học (184 SV)!",
            "• Kết luận: Kiệt sức học đường là trạng thái trung gian trực tiếp dẫn tới quyết định bỏ học của sinh viên.",
        ],
        img_path=os.path.join(FIG_DIR, "02_target_and_burnout_distribution.png"),
    )

    # Add Slide 7: EDA 2 - Stress vs Engagement Distributions
    add_content_slide(
        prs, content_layout, "03", "EDA RESULTS", "3.2",
        "Đối sánh Chỉ số Tham gia Học tập & Căng thẳng Tâm lý",
        "Sự khác biệt có ý nghĩa thống kê giữa nhóm Sinh viên Bỏ học (Yes) và Không bỏ học (No)",
        bullets=[
            "• Nhóm Bỏ học (Dropout = Yes) có sự suy giảm rõ rệt về gắn kết học tập:",
            "  - Chuyên cần TB thấp hơn: 73.7% so với 80.7% (nhóm No).",
            "  - Giờ tự học ít hơn: 2.93h/ngày so với 3.43h/ngày.",
            "  - Điểm động lực thấp hơn: 5.11 so với 5.90.",
            "• Đồng thời chịu áp lực tâm lý cao hơn:",
            "  - Stress_Level TB: 5.81 (Yes) vs 5.17 (No).",
            "  - Anxiety_Score TB: 5.26 (Yes) vs 4.65 (No).",
        ],
        img_path=os.path.join(FIG_DIR, "03_stress_vs_engagement_distributions.png"),
    )

    # Add Slide 8: EDA 3 - Correlation Heatmap & Multivariate Analysis
    add_content_slide(
        prs, content_layout, "03", "EDA RESULTS", "3.3",
        "Ma trận Tương quan & Tương tác Đa biến",
        "Phân tích hệ số tương quan Pearson giữa 18 biến định lượng trong bộ dữ liệu",
        bullets=[
            "• Tương quan âm mạnh nhất với Dropout_Risk (Yếu tố bảo vệ):",
            "  - Attendance_Percent (r = -0.244)",
            "  - Motivation_Score (r = -0.193)",
            "  - Study_Hours_Per_Day (r = -0.160)",
            "  - Family_Support_Score (r = -0.137)",
            "• Tương quan dương mạnh nhất với Dropout_Risk (Yếu tố rủi ro):",
            "  - Burnout_Numeric (r = +0.550)",
            "  - Stress_Level (r = +0.164) & Anxiety_Score (r = +0.145)",
            "  - Screen_Time_Hours (r = +0.127) & Financial_Stress (r = +0.114)",
        ],
        img_path=os.path.join(FIG_DIR, "04_correlation_matrix_heatmap.png"),
    )

    # Add Slide 9: Model Comparison Chart (10 Models x 2 Scenarios)
    add_content_slide(
        prs, content_layout, "04", "MODELING", "3.4",
        "So sánh Hiệu năng 10 Mô hình Học máy trên 2 Kịch bản",
        "Đánh giá bằng 5-Fold Stratified Cross-Validation & Tập Kiểm thử Độc lập (Test 20%)",
        bullets=[
            "• Kịch bản 1 (Đầy đủ đặc trưng - Có Burnout_Level):",
            "  - Random Forest đạt ROC-AUC cao nhất: 80.50% (CV) | 80.24% (Test).",
            "  - Extra Trees đạt Accuracy & F1 cao nhất: Acc 75.63%, F1 78.92%, Recall 83.91%.",
            "  - Soft Voting Ensemble & Logistic Regression đạt ROC-AUC ~79.8% - 80.0%.",
            "• Kịch bản 2 (Cảnh báo sớm - Ẩn hoàn toàn Burnout_Level):",
            "  - Soft Voting Ensemble dẫn đầu với CV Acc 70.63%, F1 73.89%, ROC-AUC 77.10%.",
            "  - Chứng minh các chỉ số Stress & Engagement dự báo tốt ngay từ giai đoạn sớm.",
        ],
        img_path=os.path.join(FIG_DIR, "07_model_performance_comparison.png"),
    )

    # Add Slide 10: Detailed Metrics Table
    table_metrics = [
        ["Mô hình (Model)", "KB1: CV Acc", "KB1: CV F1", "KB1: CV AUC", "KB1: Test Acc", "KB1: Test F1", "KB1: Test AUC", "KB2: CV AUC (Sớm)"],
        ["1. Random Forest", "74.38%", "77.46%", "80.50%", "71.88%", "75.14%", "80.24%", "76.14%"],
        ["2. Extra Trees", "75.00%", "78.10%", "79.22%", "75.63%", "78.92%", "78.81%", "76.19%"],
        ["3. Soft Voting Ensemble", "73.88%", "76.79%", "79.95%", "72.50%", "75.82%", "79.37%", "77.10%"],
        ["4. Logistic Regression", "72.25%", "75.19%", "79.82%", "73.75%", "76.14%", "79.86%", "76.82%"],
        ["5. SVM (RBF Kernel)", "74.63%", "77.57%", "78.55%", "74.38%", "77.84%", "79.15%", "76.13%"],
        ["6. MLP Neural Network", "73.63%", "75.94%", "79.68%", "70.63%", "72.83%", "77.75%", "75.02%"],
        ["7. XGBoost", "71.75%", "74.28%", "79.25%", "71.25%", "73.86%", "77.77%", "76.45%"],
        ["8. Gradient Boosting", "71.75%", "74.44%", "79.08%", "73.13%", "75.14%", "78.52%", "75.68%"],
        ["9. K-Nearest Neighbors", "70.25%", "72.59%", "77.14%", "71.88%", "74.86%", "76.37%", "71.99%"],
        ["10. Decision Tree", "66.88%", "67.75%", "73.88%", "68.75%", "71.59%", "73.82%", "67.12%"],
    ]
    add_content_slide(
        prs, content_layout, "04", "MODELING", "3.5",
        "Bảng Tổng hợp Kết quả Thực nghiệm 10 Mô hình Học máy",
        "So sánh chi tiết các độ đo Accuracy, F1-Score và ROC-AUC trên cả 2 kịch bản thực nghiệm",
        bullets=None,
        img_path=None,
        table_data=table_metrics,
    )

    # Add Slide 11: ROC Curves, Confusion Matrices & Feature Importance
    add_content_slide(
        prs, content_layout, "04", "EVALUATION", "3.6",
        "Phân tích Đường cong ROC, Precision-Recall & Feature Importance",
        "Các đặc trưng tổng hợp Stress & Engagement giữ vai trò quyết định thứ 2 ngay sau Burnout_Level",
        bullets=[
            "• Đường cong ROC & PR: Random Forest (AUC=0.802, AP=0.824) và Logistic Regression (AUC=0.799, AP=0.832) duy trì độ nhạy (Recall) cao ổn định.",
            "• Top Đặc trưng quan trọng nhất (Random Forest & XGBoost):",
            "  1. Burnout_Ordinal / Burnout_Level (Mức độ kiệt sức)",
            "  2. Stress_Engagement_Ratio (Tỷ số Căng thẳng / Gắn kết)",
            "  3. Academic_Engagement_Index (Chỉ số Tham gia học tập)",
            "  4. Psychological_Stress_Index (Chỉ số Căng thẳng tâm lý)",
            "  5. Attendance_Percent (Tỷ lệ chuyên cần)",
            "  6. Net_Stress_Support_Balance & Academic_Risk_Burden",
        ],
        img_path=os.path.join(FIG_DIR, "10_feature_importance_analysis.png"),
    )

    # Add Slide 12: System Architecture & Web UI Demo
    add_content_slide(
        prs, content_layout, "05", "WEB APPLICATION", "4.1",
        "Kiến trúc Hệ thống & Giao diện Web Dashboard Cảnh báo Sớm",
        "Ứng dụng FastAPI (app.py) dự báo thời gian thực Burnout_Level & Dropout_Risk kèm khuyến nghị cá nhân hóa",
        bullets=[
            "• Tầng Giao diện (Web UI): Nhập 18 thông số học tập, tâm lý, lối sống hoặc chọn nhanh 3 mẫu hồ sơ sinh viên (Low / Medium / High Risk).",
            "• Tầng AI Engine: Tự động tính 6 chỉ số chuyên ngành và chạy song song mô hình dự báo Burnout_Level (3 lớp) + Dropout_Risk.",
            "• Đầu ra thông minh: Hiển thị xác suất bỏ học (%), mức độ kiệt sức và danh sách giải pháp hỗ trợ học tập - tâm lý cho cố vấn học tập.",
        ],
        img_path=os.path.join(FIG_DIR, "12_system_architecture_diagram.png"),
    )

    # Re-add Closing Slide at the end
    end_slide = prs.slides.add_slide(closing_layout)
    tb = end_slide.shapes.add_textbox(PptInches(2.5), PptInches(4.3), PptInches(8.5), PptInches(1.5))
    p = tb.text_frame.paragraphs[0]
    p.text = "Cảm ơn Giảng viên và Hội đồng Chấm điểm!\nNhóm 03: Đỗ Quang Minh • Vũ Quang Huy"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = PptPt(20)
    p.font.bold = True
    p.font.color.rgb = PptRGBColor(29, 53, 87)

    for path in [
        os.path.join(OUT_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx"),
        os.path.join(SRC_FOLDER, "SIC_AI_Capstone Project_Presentation Slide Template.pptx"),
    ]:
        prs.save(path)
    print("Completed: SIC_AI_Capstone Project_Presentation Slide Template.pptx")


if __name__ == "__main__":
    fill_action_plan()
    fill_evaluation_sheet()
    fill_wbs_excel()
    fill_final_report()
    fill_presentation_pptx()
    print("\nALL 5 SIC CAPSTONE PROJECT FILES SUCCESSFULLY FILLED AND SAVED!")
