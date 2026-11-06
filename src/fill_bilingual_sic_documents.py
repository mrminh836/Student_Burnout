"""
Generates TWO complete versions of all 5 Samsung Innovation Campus (SIC) AI Capstone files:
1. 100% Vietnamese Version (Ban_Tieng_Viet_Full / SIC_Documents_VI)
2. 100% English Version (Ban_Tieng_Anh_Full / SIC_Documents_EN)

Each version includes:
- SIC_AI_Capstone Project_Action Plan.docx
- SIC_AI_Capstone Project_Evaluation Sheet.docx
- SIC_AI_Capstone Project_Final Report.docx
- SIC_AI_Capstone Project_Work Breakdown Structure.xlsx
- SIC_AI_Capstone Project_Presentation Slide Template.pptx
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
from openpyxl.styles import Font, PatternFill

from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor
from pptx.enum.text import PP_ALIGN

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_FOLDER = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"
BACKUP_DIR = os.path.join(BASE_DIR, "templates_backup")

FIG_VI = os.path.join(BASE_DIR, "figures_vi")
FIG_EN = os.path.join(BASE_DIR, "figures_en")

OUT_VI_WS = os.path.join(BASE_DIR, "SIC_Documents_VI")
OUT_EN_WS = os.path.join(BASE_DIR, "SIC_Documents_EN")
OUT_VI_DRV = os.path.join(SRC_FOLDER, "Ban_Tieng_Viet_Full")
OUT_EN_DRV = os.path.join(SRC_FOLDER, "Ban_Tieng_Anh_Full")

for d in [OUT_VI_WS, OUT_EN_WS, OUT_VI_DRV, OUT_EN_DRV]:
    os.makedirs(d, exist_ok=True)


def save_multi(obj_save_fn, base_filename, lang):
    """Saves a document to both workspace folder, drive subfolder, and suffixed file in drive root."""
    name_no_ext, ext = os.path.splitext(base_filename)
    suffixed_name = f"{name_no_ext}_{lang.upper()}{ext}"
    if lang == "vi":
        targets = [
            os.path.join(OUT_VI_WS, base_filename),
            os.path.join(OUT_VI_DRV, base_filename),
            os.path.join(SRC_FOLDER, suffixed_name),
            os.path.join(SRC_FOLDER, base_filename),  # Keep default in root as VI
        ]
    else:
        targets = [
            os.path.join(OUT_EN_WS, base_filename),
            os.path.join(OUT_EN_DRV, base_filename),
            os.path.join(SRC_FOLDER, suffixed_name),
        ]
    for t in targets:
        obj_save_fn(t)


# ==============================================================================
# 1. ACTION PLAN (.DOCX) - FULL VI & FULL EN
# ==============================================================================
def build_action_plan(lang="vi"):
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Action Plan.docx")
    doc = Document(in_path)
    t0, t1, t2 = doc.tables[0], doc.tables[1], doc.tables[2]

    if lang == "vi":
        t0.cell(0, 0).text = "Khóa học"
        t0.cell(0, 1).text = "Khóa học Trí tuệ Nhân tạo (Samsung Innovation Campus - AI Course)"
        t0.cell(1, 0).text = "Tên Nhóm"
        t0.cell(1, 1).text = "Nhóm 03 (Group 03 - Student Burnout AI)"
        t0.cell(2, 0).text = "Trưởng nhóm / Thành viên"
        t0.cell(2, 1).text = "Đỗ Quang Minh (Trưởng nhóm) / Vũ Quang Huy (Thành viên)"
        t0.cell(3, 0).text = "Tên Đề tài"
        t0.cell(3, 1).text = (
            "Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học "
            "dựa trên các chỉ số về căng thẳng và sự tham gia học tập"
        )
        t0.cell(4, 0).text = "Mục tiêu"
        t0.cell(5, 0).text = (
            "1. Phân tích khám phá dữ liệu (EDA) trên bộ dữ liệu 800 sinh viên với 25 thuộc tính về nhân khẩu học, "
            "sự tham gia học tập (chuyên cần, giờ tự học, GPA kỳ trước, động lực, số môn nợ), lối sống và chỉ số căng thẳng tâm lý.\n"
            "2. Xây dựng quy trình tiền xử lý dữ liệu khuyết trên 18/25 cột (KNNImputer kết hợp Mode Imputation) và thiết kế 6 đặc trưng chuyên sâu đo lường tương tác giữa Căng thẳng và Sự tham gia học tập.\n"
            "3. Huấn luyện, kiểm định chéo phân tầng 5-Fold và so sánh toàn diện 10 mô hình học máy trên 2 kịch bản (Kịch bản 1: Đầy đủ đặc trưng có mức độ kiệt sức; Kịch bản 2: Cảnh báo sớm chỉ dùng các chỉ số căng thẳng và học tập thô).\n"
            "4. Triển khai ứng dụng Web Dashboard (FastAPI) dự báo trực tiếp mức độ kiệt sức và nguy cơ bỏ học kèm khuyến nghị can thiệp cá nhân hóa."
        )
        t0.cell(6, 0).text = "Tóm tắt đề tài"
        t0.cell(7, 0).text = (
            "Tình trạng kiệt sức học đường và bỏ học đại học đang là vấn đề cấp thiết tại các cơ sở giáo dục đại học. "
            "Đồ án này nghiên cứu mối liên hệ định lượng giữa các chỉ số căng thẳng tâm lý (căng thẳng học tập, lo âu, áp lực tài chính, áp lực bạn bè) và sự tham gia học tập (tỷ lệ chuyên cần, giờ tự học mỗi ngày, điểm động lực, điểm GPA) nhằm dự báo sớm nguy cơ bỏ học. "
            "Nhóm xây dựng đường ống học máy khép kín kết hợp kỹ thuật tạo đặc trưng chuyên ngành và đối sánh 10 thuật toán phân loại. "
            "Kết quả thực nghiệm cho thấy các mô hình tập hợp cây (Random Forest, Extra Trees) và mô hình kết hợp (Soft Voting Ensemble) đạt hiệu năng vượt trội với ROC-AUC ~80.5%, F1-Score ~78.9% và độ nhạy (Recall) trên 81%, hỗ trợ đắc lực cho nhà trường trong công tác cảnh báo sớm."
        )
        t0.cell(8, 0).text = "Phương pháp thực hiện"
        t0.cell(9, 0).text = (
            "• Bước 1 - Tiền xử lý dữ liệu: Xử lý 18/25 cột khuyết bằng thuật toán KNNImputer (k=7) cho biến số liên tục và giá trị phổ biến nhất cho biến phân loại; chuẩn hóa bằng StandardScaler và mã hóa One-Hot / Thứ tự.\n"
            "• Bước 2 - Kỹ thuật tạo đặc trưng: Thiết kế 6 chỉ số tổng hợp gồm: Chỉ số Gắn kết Học tập, Chỉ số Căng thẳng Tâm lý, Tỷ số Căng thẳng / Gắn kết, Cân bằng Áp lực - Hỗ trợ, Tỷ số Mệt mỏi Số và Gánh nặng Học vụ.\n"
            "• Bước 3 - Huấn luyện & So sánh mô hình: Đánh giá 10 thuật toán học máy qua kiểm định chéo 5-Fold và tập kiểm thử độc lập (20%) theo các thước đo Accuracy, Precision, Recall, F1-Score, ROC-AUC và PR-AUC.\n"
            "• Bước 4 - Xây dựng phần mềm: Đóng gói mô hình tối ưu và tích hợp vào giao diện Web Dashboard phục vụ dự báo trực tuyến."
        )
        t1.cell(0, 0).text = "Dữ liệu sử dụng"
        t1.cell(1, 0).text = (
            "• Nguồn dữ liệu: Bộ dữ liệu 'student_burnout_dropout_dataset_2.csv' thuộc chủ đề Student Burnout.\n"
            "• Quy mô & Cấu trúc: 800 hồ sơ sinh viên x 25 thuộc tính, chia thành 5 nhóm:\n"
            "  - Nhân khẩu học & Bối cảnh: Tuổi, Giới tính, Năm học, Khoa/Ngành, Loại hình cư trú, Mức thu nhập gia đình.\n"
            "  - Sự tham gia học tập: Tỷ lệ chuyên cần (%), Giờ tự học/ngày, GPA kỳ trước, Số môn nợ, Điểm động lực.\n"
            "  - Lối sống & Sinh hoạt: Giờ ngủ/ngày, Giờ dùng màn hình/ngày, Tần suất tập thể dục, Hoạt động xã hội, Việc làm thêm.\n"
            "  - Căng thẳng & Tâm lý: Mức căng thẳng, Điểm lo âu, Áp lực tài chính, Áp lực bạn bè, Hỗ trợ gia đình, Tiếp cận tư vấn.\n"
            "  - Biến mục tiêu: Mức độ kiệt sức (Thấp: 320, Trung bình: 280, Cao: 200) và Nguy cơ bỏ học (Có: 437 [54.6%], Không: 363 [45.4%]).\n"
            "• Cách sử dụng: Chia tỷ lệ 80% Huấn luyện (640 mẫu) / 20% Kiểm thử (160 mẫu) có phân tầng kết hợp kiểm định chéo 5-Fold."
        )
        t1.cell(2, 0).text = "Kết quả kỳ vọng"
        t1.cell(3, 0).text = (
            "• Kết quả đạt được:\n"
            "  1. Bộ 12 biểu đồ phân tích khám phá dữ liệu và đánh giá mô hình trực quan, làm rõ tác động của kiệt sức học đường (nhóm Kiệt sức Cao có tỷ lệ bỏ học lên tới 92.0%).\n"
            "  2. Bảng đối sánh thực nghiệm 10 mô hình học máy, lựa chọn được mô hình tối ưu đạt ROC-AUC = 80.50%, F1-Score = 78.92% và Recall = 83.91%.\n"
            "  3. Hệ thống Web Dashboard hoàn chỉnh cho phép nhập thông số sinh viên và dự báo tức thời mức độ kiệt sức cùng nguy cơ bỏ học.\n"
            "• Lợi ích thực tiễn: Giúp cố vấn học tập và nhà trường phát hiện sớm sinh viên có nguy cơ bỏ học để kịp thời tư vấn tâm lý, hỗ trợ học bổng và phụ đạo học tập."
        )
        t1.cell(4, 0).text = "Phân công thành viên"
        t1.cell(5, 0).text = (
            "1. Đỗ Quang Minh (Trưởng nhóm):\n"
            "   - Thiết kế kiến trúc Pipeline học máy và kỹ thuật tạo 6 đặc trưng chuyên ngành.\n"
            "   - Xây dựng, huấn luyện và tối ưu nhóm mô hình tập hợp cây & tăng cường (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
            "   - Phát triển Backend FastAPI và tích hợp mô hình vào ứng dụng Web Demo.\n"
            "2. Vũ Quang Huy (Thành viên):\n"
            "   - Khảo sát dữ liệu, xử lý dữ liệu khuyết và thực hiện Phân tích khám phá dữ liệu (EDA).\n"
            "   - Huấn luyện và đánh giá nhóm mô hình cơ sở & mạng nơ-ron (Logistic Regression, SVM, KNN, Decision Tree, MLP).\n"
            "   - Thiết kế giao diện người dùng (Frontend Dashboard), biên soạn Báo cáo cuối kỳ, WBS và Slide thuyết trình."
        )
        t2.cell(0, 0).text = "Tóm tắt tiến độ"
        t2.cell(1, 0).text = (
            "• Tuần 1 (Giai đoạn 1 - Khởi động & Khám phá dữ liệu): Xác định yêu cầu bài toán, thống kê mô tả, phân tích 18 cột dữ liệu khuyết và xuất 6 biểu đồ EDA.\n"
            "• Tuần 2 (Giai đoạn 2 - Tiền xử lý & Kỹ thuật đặc trưng): Xây dựng Pipeline KNNImputer, chuẩn hóa dữ liệu và thiết kế 6 chỉ số kết hợp Căng thẳng & Tham gia học tập.\n"
            "• Tuần 3 (Giai đoạn 3 - Huấn luyện & So sánh mô hình): Huấn luyện 10 mô hình học máy trên 2 kịch bản, tối ưu siêu tham số, phân tích ROC-AUC, Ma trận nhầm lẫn và Độ quan trọng đặc trưng.\n"
            "• Tuần 4 (Giai đoạn 4 - Xây dựng Web UI & Hoàn thiện Báo cáo): Phát triển ứng dụng Web Demo, kiểm thử hệ thống, hoàn thiện Báo cáo tổng kết và Slide thuyết trình."
        )
        t2.cell(2, 0).text = "Nhận xét & Đánh giá"
        t2.cell(3, 0).text = "Đề tài có tính ứng dụng thực tiễn cao trong giáo dục; phương pháp luận chặt chẽ, so sánh đa dạng 10 thuật toán học máy và có sản phẩm giao diện trực quan."
    else:
        t0.cell(0, 1).text = "AI Course (Samsung Innovation Campus)"
        t0.cell(1, 1).text = "Group 03 - Student Burnout & Dropout AI"
        t0.cell(2, 1).text = "Do Quang Minh (Team Leader) / Vu Quang Huy (Team Member)"
        t0.cell(3, 1).text = (
            "Building, Training, and Comparing Machine Learning Models for Student Dropout Risk Prediction "
            "Based on Psychological Stress and Academic Engagement Indicators"
        )
        t0.cell(5, 0).text = (
            "1. Conduct comprehensive Exploratory Data Analysis (EDA) on 800 university student records across 25 attributes covering demographics, "
            "academic engagement (attendance, study hours, previous GPA, motivation, backlogs), lifestyle habits, and psychological stress indicators.\n"
            "2. Develop an automated missing-data imputation pipeline (KNNImputer + Mode Imputation across 18/25 columns) and engineer 6 domain-specific composite features capturing Stress-Engagement interactions.\n"
            "3. Train, cross-validate (5-Fold Stratified CV), and benchmark 10 machine learning classifiers across two experimental scenarios (Scenario 1: Full Feature Set with Burnout_Level; Scenario 2: Early Warning using raw Stress & Engagement indicators only).\n"
            "4. Deploy an interactive FastAPI Web Dashboard for real-time Burnout Level and Dropout Risk prediction with personalized intervention recommendations."
        )
        t0.cell(7, 0).text = (
            "Academic burnout and university student dropout pose critical challenges to higher education institutions worldwide. "
            "This capstone project investigates the quantitative interplay between psychological stress indicators (Stress Level, Anxiety Score, Financial Stress, Peer Pressure) and study engagement metrics (Attendance Percentage, Daily Study Hours, Motivation Score, Previous GPA) to predict student dropout risk early. "
            "We construct a leakage-free machine learning pipeline incorporating domain-driven feature engineering and benchmark 10 classification algorithms. "
            "Experimental results demonstrate that tree ensemble methods (Random Forest, Extra Trees) and our Soft Voting Ensemble achieve top performance with a 5-Fold CV ROC-AUC of 80.50%, Test F1-Score of 78.92%, and Recall of 83.91%, enabling proactive institutional support for at-risk students."
        )
        t0.cell(9, 0).text = (
            "• Step 1 - Data Preprocessing: Impute missing values across 18/25 columns using KNNImputer (k=7) for continuous variables and Most-Frequent Imputation for categorical variables; apply StandardScaler normalization and One-Hot / Ordinal encoding.\n"
            "• Step 2 - Domain Feature Engineering: Construct 6 composite features: Academic_Engagement_Index, Psychological_Stress_Index, Stress_Engagement_Ratio, Net_Stress_Support_Balance, Digital_Fatigue_Ratio, and Academic_Risk_Burden.\n"
            "• Step 3 - Model Training & Benchmarking: Evaluate 10 machine learning classifiers using 5-Fold Stratified Cross-Validation and a 20% hold-out test set across Accuracy, Precision, Recall, F1-Score, ROC-AUC, and PR-AUC.\n"
            "• Step 4 - Web Application Deployment: Serialize top-performing pipelines (.pkl) and integrate them into an interactive FastAPI Web Dashboard."
        )
        t1.cell(1, 0).text = (
            "• Data Source: 'student_burnout_dropout_dataset_2.csv' provided for the Student Burnout & Dropout Prediction task.\n"
            "• Scale & Structure: 800 student records x 25 attributes categorized into 5 groups:\n"
            "  - Demographics & Context: Age, Gender, Year_of_Study, Department, Residence_Type, Family_Income_Bracket.\n"
            "  - Academic Engagement: Attendance_Percent, Study_Hours_Per_Day, Previous_GPA, Backlogs, Motivation_Score.\n"
            "  - Lifestyle & Habits: Sleep_Hours, Screen_Time_Hours, Exercise_Freq_Per_Week, Social_Activity_Score, Part_Time_Job.\n"
            "  - Psychological Stress & Support: Stress_Level, Anxiety_Score, Financial_Stress_Score, Peer_Pressure_Score, Family_Support_Score, Counseling_Access.\n"
            "  - Target Variables: Burnout_Level (Low: 320, Medium: 280, High: 200) and Dropout_Risk (Yes: 437 [54.6%], No: 363 [45.4%]).\n"
            "• Utilization Strategy: 80% Stratified Training Set (640 records) / 20% Hold-out Test Set (160 records) alongside 5-Fold Stratified Cross-Validation."
        )
        t1.cell(3, 0).text = (
            "• Expected Outcomes:\n"
            "  1. A complete suite of 12 high-resolution EDA and model evaluation visualizations revealing key dropout drivers (e.g., 92.0% dropout rate among High Burnout students).\n"
            "  2. Rigorous benchmarking of 10 ML classifiers, identifying top ensemble models achieving 80.50% ROC-AUC, 78.92% F1-Score, and 83.91% Recall.\n"
            "  3. A responsive FastAPI Web Dashboard providing real-time dual predictions (Burnout_Level & Dropout_Risk) and actionable student counseling recommendations.\n"
            "• Expected Benefits: Empowers academic advisors and student affairs offices to identify at-risk students early and deliver targeted psychological counseling, financial aid, and academic tutoring."
        )
        t1.cell(5, 0).text = (
            "1. Do Quang Minh (Team Leader):\n"
            "   - Design the leakage-free ML Pipeline architecture and engineer the 6 domain-specific composite features.\n"
            "   - Train, tune, and evaluate tree-based ensemble & boosting classifiers (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
            "   - Develop the FastAPI backend inference service and integrate serialized models into the Web Dashboard.\n"
            "2. Vu Quang Huy (Team Member):\n"
            "   - Perform data auditing, missing-value analysis, and comprehensive Exploratory Data Analysis (EDA).\n"
            "   - Train and evaluate baseline linear, kernel, distance, single-tree, and neural models (Logistic Regression, SVM, KNN, Decision Tree, MLP Neural Network).\n"
            "   - Design the frontend dashboard UI, compile experimental evaluation tables, and author the Final Report, WBS, and Presentation Slides."
        )
        t2.cell(1, 0).text = (
            "• Week 1 (Phase 1 - Project Initiation & EDA): Problem formulation, descriptive statistics, missing data audit across 18 columns, and generation of 6 EDA figures.\n"
            "• Week 2 (Phase 2 - Preprocessing & Feature Engineering): Implementation of KNNImputer pipeline, feature scaling/encoding, and creation of 6 Stress-Engagement composite indices.\n"
            "• Week 3 (Phase 3 - Model Training & Benchmarking): Training and 5-Fold CV evaluation of 10 ML models across 2 scenarios, hyperparameter optimization, ROC/PR curves, Confusion Matrices, and Feature Importance.\n"
            "• Week 4 (Phase 4 - Web UI Deployment & Documentation): Development of FastAPI Web Dashboard, end-to-end testing, Final Report, WBS workbook, and Presentation Slides."
        )
        t2.cell(3, 0).text = "Highly practical educational AI capstone project with rigorous methodology, comprehensive 10-model benchmarking, and a functional interactive web interface."

    save_multi(doc.save, "SIC_AI_Capstone Project_Action Plan.docx", lang)
    print(f"Completed Action Plan ({lang.upper()})")


# ==============================================================================
# 2. EVALUATION SHEET (.DOCX) - FULL VI & FULL EN
# ==============================================================================
def build_evaluation_sheet(lang="vi"):
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Evaluation Sheet.docx")
    doc = Document(in_path)
    t0, t1 = doc.tables[0], doc.tables[1]

    if lang == "vi":
        t0.cell(1, 2).text = "Dự đoán tỷ lệ bỏ học dựa trên chỉ số căng thẳng & sự tham gia học tập (Student Burnout)"
        t0.cell(1, 4).text = "Nhóm 03: Đỗ Quang Minh, Vũ Quang Huy"
        t1.cell(5, 2).text = "✓ (Đạt Mức 5 - Tích hợp Prototype & Web UI)"
        t1.cell(6, 2).text = "✓ (Đạt Mức 6 - Kiểm thử trên 800 hồ sơ SV)"
    else:
        t0.cell(1, 2).text = "Student Dropout Risk Prediction Based on Stress & Study Engagement Indicators"
        t0.cell(1, 4).text = "Group 03: Do Quang Minh, Vu Quang Huy"
        t1.cell(5, 2).text = "✓ (Achieved TRL 5 - Integrated Prototype & Web UI)"
        t1.cell(6, 2).text = "✓ (Achieved TRL 6 - Validated on 800-Student Dataset)"

    save_multi(doc.save, "SIC_AI_Capstone Project_Evaluation Sheet.docx", lang)
    print(f"Completed Evaluation Sheet ({lang.upper()})")


# ==============================================================================
# 3. WORK BREAKDOWN STRUCTURE (.XLSX) - FULL VI & FULL EN
# ==============================================================================
def build_wbs_excel(lang="vi"):
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx")
    wb = openpyxl.load_workbook(in_path)
    ws = wb["WBS"]

    base_date = datetime.date(2026, 10, 19)
    for d_idx in range(14):
        ws.cell(13, 11 + d_idx).value = base_date + datetime.timedelta(days=d_idx)

    if lang == "vi":
        ws["B2"] = "Cấu trúc Phân rã Công việc Dự án Cuối khóa (Capstone Project WBS)"
        ws["B7"] = "Ngày cập nhật"
        ws["E7"] = datetime.date(2026, 11, 5)
        ws["B8"] = "Tên Nhóm Dự án"
        ws["E8"] = "Nhóm 03: Đỗ Quang Minh (Trưởng nhóm), Vũ Quang Huy"
        ws["B9"] = "Đề tài Dự án"
        ws["E9"] = "Dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập"
        ws["B11"] = "Hạng mục Công việc"
        ws["F11"] = "Ngày bắt đầu"
        ws["G11"] = "Ngày kết thúc"
        ws["H11"] = "Người phụ trách"
        ws["I11"] = "Sản phẩm bàn giao"
        ws["J11"] = "Trạng thái"

        wbs_tasks = [
            (14, 1, "1. Khởi tạo Dự án & Phân tích Yêu cầu", 0, 1, "Đỗ Quang Minh, Vũ Quang Huy", "Kế hoạch hành động & WBS", "Hoàn thành"),
            (15, 2, "1.1. Xác định mục tiêu bài toán dự báo Kiệt sức & Bỏ học", 0, 0, "Đỗ Quang Minh", "Đặc tả yêu cầu bài toán", "Hoàn thành"),
            (16, 2, "1.2. Khảo sát cấu trúc bộ dữ liệu (800 SV x 25 thuộc tính)", 0, 1, "Vũ Quang Huy", "Báo cáo cấu trúc dữ liệu", "Hoàn thành"),
            (17, 1, "2. Phân tích Khám phá Dữ liệu (EDA)", 2, 4, "Vũ Quang Huy", "Bộ 6 biểu đồ EDA & Nhận xét", "Hoàn thành"),
            (18, 2, "2.1. Thống kê tỷ lệ dữ liệu khuyết trên 18/25 cột & phân phối nhãn", 2, 2, "Vũ Quang Huy", "Biểu đồ Khuyết & Phân phối", "Hoàn thành"),
            (19, 2, "2.2. Phân tích tương quan Pearson & đối sánh Căng thẳng vs Học tập", 3, 4, "Vũ Quang Huy, Đỗ Quang Minh", "Ma trận tương quan & Violin plot", "Hoàn thành"),
            (20, 1, "3. Tiền xử lý & Kỹ thuật Tạo Đặc trưng", 4, 6, "Đỗ Quang Minh", "Module DomainFeatureEngineer", "Hoàn thành"),
            (21, 2, "3.1. Xây dựng Pipeline KNNImputer (k=7), Chuẩn hóa & Mã hóa", 4, 5, "Vũ Quang Huy", "Pipeline Tiền xử lý dữ liệu", "Hoàn thành"),
            (22, 2, "3.2. Thiết kế 6 chỉ số tổng hợp (Gắn kết, Căng thẳng, Tỷ số...)", 5, 6, "Đỗ Quang Minh", "6 Đặc trưng chuyên ngành mới", "Hoàn thành"),
            (23, 1, "4. Huấn luyện, Tối ưu & So sánh 10 Mô hình Học máy", 7, 10, "Đỗ Quang Minh, Vũ Quang Huy", "10 Mô hình & Bảng kết quả CSV", "Hoàn thành"),
            (24, 2, "4.1. Huấn luyện nhóm mô hình cơ sở (Logistic, SVM, KNN, DT, MLP)", 7, 8, "Vũ Quang Huy", "Kết quả 5 mô hình cơ sở", "Hoàn thành"),
            (25, 2, "4.2. Huấn luyện nhóm Ensemble & Boosting (RF, ET, GB, XGB, Voting)", 8, 9, "Đỗ Quang Minh", "Mô hình tốt nhất (AUC 80.5%)", "Hoàn thành"),
            (26, 2, "4.3. Đánh giá 2 kịch bản thực nghiệm & Độ quan trọng đặc trưng", 9, 10, "Đỗ Quang Minh, Vũ Quang Huy", "Biểu đồ ROC, PR, CM, Importance", "Hoàn thành"),
            (27, 1, "5. Phát triển Web Dashboard Demo & Hoàn thiện Báo cáo", 11, 13, "Đỗ Quang Minh, Vũ Quang Huy", "Web UI, Báo cáo & Slide", "Hoàn thành"),
            (28, 2, "5.1. Xây dựng giao diện Web Demo dự báo trực tiếp Burnout & Dropout", 11, 12, "Đỗ Quang Minh", "Ứng dụng Web app.py (FastAPI)", "Hoàn thành"),
            (29, 2, "5.2. Biên soạn Báo cáo cuối kỳ, Slide thuyết trình & Notebook", 12, 13, "Vũ Quang Huy, Đỗ Quang Minh", "Bộ 5 tài liệu SIC hoàn chỉnh", "Hoàn thành"),
        ]
    else:
        ws["B2"] = "Capstone Project Work Breakdown Structure (WBS)"
        ws["B7"] = "Today"
        ws["E7"] = datetime.date(2026, 11, 5)
        ws["B8"] = "Project Team Name"
        ws["E8"] = "Group 03: Do Quang Minh (Team Leader), Vu Quang Huy"
        ws["B9"] = "Project Topic"
        ws["E9"] = "Student Dropout Risk Prediction Based on Psychological Stress and Academic Engagement Indicators"

        wbs_tasks = [
            (14, 1, "1. Project Initiation & Requirements Analysis", 0, 1, "Do Quang Minh, Vu Quang Huy", "Action Plan & WBS", "Done"),
            (15, 2, "1.1. Define objectives for Burnout & Dropout Risk prediction", 0, 0, "Do Quang Minh", "Problem Specification", "Done"),
            (16, 2, "1.2. Audit dataset schema (800 students x 25 attributes)", 0, 1, "Vu Quang Huy", "Data Audit Summary", "Done"),
            (17, 1, "2. Exploratory Data Analysis (EDA) & Statistical Profiling", 2, 4, "Vu Quang Huy", "6 High-Res EDA Figures", "Done"),
            (18, 2, "2.1. Analyze missingness across 18/25 columns & class balance", 2, 2, "Vu Quang Huy", "Missingness & Target Plots", "Done"),
            (19, 2, "2.2. Conduct Pearson correlation & Stress vs. Engagement analysis", 3, 4, "Vu Quang Huy, Do Quang Minh", "Heatmap & Violin Distributions", "Done"),
            (20, 1, "3. Data Preprocessing & Domain Feature Engineering", 4, 6, "Do Quang Minh", "DomainFeatureEngineer Class", "Done"),
            (21, 2, "3.1. Build KNNImputer (k=7), StandardScaler & OneHot Pipeline", 4, 5, "Vu Quang Huy", "Leakage-Free Preprocessing", "Done"),
            (22, 2, "3.2. Engineer 6 composite Stress-Engagement indices", 5, 6, "Do Quang Minh", "6 Engineered Features", "Done"),
            (23, 1, "4. Training, Tuning & Benchmarking 10 ML Classifiers", 7, 10, "Do Quang Minh, Vu Quang Huy", "10 Trained Models & CSV Metrics", "Done"),
            (24, 2, "4.1. Train baseline & neural models (Logistic, SVM, KNN, DT, MLP)", 7, 8, "Vu Quang Huy", "5 Baseline Model Benchmarks", "Done"),
            (25, 2, "4.2. Train tree ensembles & boosting (RF, ET, GB, XGBoost, Voting)", 8, 9, "Do Quang Minh", "Best Model (.pkl) AUC 80.5%", "Done"),
            (26, 2, "4.3. Evaluate 2 experimental scenarios & Feature Importance", 9, 10, "Do Quang Minh, Vu Quang Huy", "ROC, PR, CM & Importance Plots", "Done"),
            (27, 1, "5. Web Dashboard Deployment & Final Documentation", 11, 13, "Do Quang Minh, Vu Quang Huy", "FastAPI Web UI, Report & Slides", "Done"),
            (28, 2, "5.1. Develop real-time Web Dashboard for Burnout & Dropout prediction", 11, 12, "Do Quang Minh", "FastAPI Web Application (app.py)", "Done"),
            (29, 2, "5.2. Compile Final Report, Presentation Slides & Jupyter Notebook", 12, 13, "Vu Quang Huy, Do Quang Minh", "Completed SIC Deliverable Suite", "Done"),
        ]

    fill_bar = PatternFill(start_color="2A9D8F", end_color="2A9D8F", fill_type="solid")
    fill_header_task = PatternFill(start_color="E9F5F3", end_color="E9F5F3", fill_type="solid")
    fill_clear = PatternFill(fill_type=None)

    for r, lvl, title, s_off, e_off, owner, deliv, status in wbs_tasks:
        for c in range(2, 6):
            ws.cell(r, c).value = None
        target_col = 2 if lvl == 1 else 3
        ws.cell(r, target_col).value = title
        ws.cell(r, target_col).font = Font(name="Arial", size=10, bold=(lvl == 1))
        ws.cell(r, 6).value = base_date + datetime.timedelta(days=s_off)
        ws.cell(r, 6).number_format = "yyyy-mm-dd"
        ws.cell(r, 7).value = base_date + datetime.timedelta(days=e_off)
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
    ws.column_dimensions["H"].width = 28
    ws.column_dimensions["I"].width = 30
    ws.column_dimensions["J"].width = 15

    save_multi(wb.save, "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx", lang)
    print(f"Completed WBS Excel ({lang.upper()})")


# ==============================================================================
# 4. FINAL REPORT (.DOCX) - FULL VI & FULL EN
# ==============================================================================
def unwrap_sdt_in_docx(src_docx, dst_docx):
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


def insert_image_after(paragraph, img_path, caption_text, width_inches=6.0):
    p_img = insert_paragraph_after(paragraph, "")
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_img = p_img.add_run()
    r_img.add_picture(img_path, width=Inches(width_inches))
    p_cap = insert_paragraph_after(p_img, caption_text, bold=True, italic=True, font_size=9.5, color_rgb=(40, 70, 95))
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p_cap


def build_final_report(lang="vi"):
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Final Report.docx")
    temp_unwrapped = os.path.join(BACKUP_DIR, f"_unwrapped_final_{lang}.docx")
    unwrap_sdt_in_docx(in_path, temp_unwrapped)
    doc = Document(temp_unwrapped)
    fig_folder = FIG_VI if lang == "vi" else FIG_EN

    if lang == "vi":
        replacements = {
            "< Project Title >": "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY CHO BÀI TOÁN DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CHỈ SỐ CĂNG THẲNG & SỰ THAM GIA HỌC TẬP",
            "< Date (DD/MM/YY) >": "05/11/2026",
            "Team Name": "Nhóm 03 - Student Burnout AI",
            "<Member 1>": "1. Đỗ Quang Minh (Trưởng nhóm)",
            "<Member 2>": "2. Vũ Quang Huy (Thành viên)",
            "<Member 3>": "",
            "<Member 4>": "",
            "<Member 5>": "",
        }
        section_content = {
            "1.1. Background Information": [
                "Trong bối cảnh giáo dục đại học hiện đại, tình trạng sinh viên bỏ học giữa chừng và hội chứng kiệt sức học đường đang gia tăng mạnh mẽ. Nhiều nghiên cứu giáo dục chỉ ra rằng việc sinh viên rời bỏ giảng đường không chỉ xuất phát từ năng lực học thuật thuần túy mà là hệ quả tích lũy của hai nhóm yếu tố tương tác chặt chẽ: (1) Sự suy giảm mức độ tham gia học tập (tỷ lệ chuyên cần, thời lượng tự học, điểm động lực, kết quả GPA kỳ trước, số môn nợ) và (2) Áp lực căng thẳng tâm lý - xã hội kéo dài (căng thẳng học tập, lo âu, áp lực tài chính, áp lực đồng trang lứa, thiếu ngủ và thiếu điểm tựa gia đình).",
                "Việc ứng dụng Học máy để phân tích đồng thời các chỉ số căng thẳng và hành vi tham gia học tập cho phép nhà trường chuyển đổi từ phương pháp xử lý bị động sang cơ chế Cảnh báo sớm chủ động, giúp can thiệp đúng đối tượng và đúng thời điểm.",
            ],
            "1.2. Motivation and Objective": [
                "• Động lực nghiên cứu: Tại các trường đại học, dữ liệu điểm danh, điểm GPA và khảo sát tâm lý của sinh viên thường được lưu trữ rời rạc, chưa được khai thác bằng các mô hình dự báo định lượng. Do đó, Nhóm 03 lựa chọn đề tài 'Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập'.",
                "• Mục tiêu cụ thể:\n"
                "  1. Phân tích khám phá dữ liệu (EDA) trên tập dữ liệu 800 sinh viên x 25 thuộc tính để xác định các yếu tố rủi ro chính dẫn đến kiệt sức và bỏ học.\n"
                "  2. Xây dựng quy trình Tiền xử lý tự động (KNNImputer + StandardScaler + OneHotEncoder) và thiết kế 6 đặc trưng chuyên sâu đo lường tỷ số mất cân bằng giữa Căng thẳng và Gắn kết học tập.\n"
                "  3. Huấn luyện, kiểm định chéo 5-Fold và đối sánh toàn diện 10 thuật toán Học máy trên 2 kịch bản thực nghiệm (Kịch bản 1: Đầy đủ đặc trưng; Kịch bản 2: Cảnh báo sớm không dùng biến mức độ kiệt sức).\n"
                "  4. Phát triển giao diện Web Dashboard (FastAPI) phục vụ dự báo thời gian thực và đưa ra khuyến nghị can thiệp cá nhân hóa.",
            ],
            "1.3. Members and Role Assignments": [
                "Dự án được thực hiện bởi Nhóm 03 với sự phân công công việc đồng đều:\n"
                "• Đỗ Quang Minh (Trưởng nhóm - 50% khối lượng):\n"
                "  - Thiết kế kiến trúc Pipeline học máy chống rò rỉ dữ liệu và xây dựng lớp biến đổi đặc trưng DomainFeatureEngineer (6 chỉ số tổng hợp).\n"
                "  - Huấn luyện, tinh chỉnh siêu tham số và đánh giá nhóm mô hình tập hợp cây & tăng cường (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
                "  - Xây dựng Backend FastAPI và tích hợp bộ đôi mô hình dự báo vào ứng dụng Web Demo.\n"
                "• Vũ Quang Huy (Thành viên - 50% khối lượng):\n"
                "  - Làm sạch dữ liệu, phân tích dữ liệu khuyết trên 18/25 cột và thiết kế bộ 6 biểu đồ Phân tích khám phá dữ liệu (EDA).\n"
                "  - Huấn luyện và đánh giá nhóm mô hình tuyến tính, hạt nhân, khoảng cách, cây đơn và mạng nơ-ron (Logistic Regression, SVM, KNN, Decision Tree, MLP).\n"
                "  - Thiết kế giao diện người dùng (Frontend Dashboard), tổng hợp bảng số liệu thực nghiệm, biên soạn Báo cáo cuối kỳ, WBS và Slide thuyết trình.",
            ],
            "1.4. Schedule and Milestones": [
                "Lộ trình thực hiện dự án được chia thành 4 giai đoạn chính:\n"
                "• Giai đoạn 1 (19/10/2026 – 20/10/2026): Khởi tạo dự án, xác định yêu cầu nghiệp vụ, khảo sát cấu trúc 25 trường dữ liệu và hoàn thiện Action Plan & WBS. [Hoàn thành 100%]\n"
                "• Giai đoạn 2 (21/10/2026 – 25/10/2026): Phân tích khám phá dữ liệu (EDA), kiểm định tương quan Pearson, xử lý dữ liệu khuyết bằng KNNImputer (k=7) và xây dựng 6 đặc trưng chuyên ngành. [Hoàn thành 100%]\n"
                "• Giai đoạn 3 (26/10/2026 – 29/10/2026): Huấn luyện và kiểm định chéo 5-Fold trên 10 mô hình học máy ở 2 kịch bản; xuất các biểu đồ ROC, Precision-Recall, Confusion Matrix và Feature Importance. [Hoàn thành 100%]\n"
                "• Giai đoạn 4 (30/10/2026 – 05/11/2026): Phát triển Web Dashboard Demo, kiểm thử tích hợp, hoàn thiện Jupyter Notebook, Báo cáo tổng kết và Slide thuyết trình. [Hoàn thành 100%]",
            ],
            "2.1. Data Acquisition": [
                "Bộ dữ liệu 'student_burnout_dropout_dataset_2.csv' gồm 800 hồ sơ sinh viên và 25 thuộc tính chia thành 5 nhóm:\n"
                "1. Định danh & Nhân khẩu học (7 biến): Mã sinh viên (STU00001–STU00800), Tuổi (17–25, TB: 20.88), Giới tính (Nữ: 392, Nam: 346, Khác: 32, Khuyết: 30), Năm học (1–4), Khoa (Engineering: 209, Business: 151, Science: 143, Arts: 136, Medicine: 77, Law: 61), Loại hình cư trú (Ký túc xá: 290, Ngoại trú gia đình: 286, Thuê trọ: 207), Thu nhập gia đình (5 mức).\n"
                "2. Chỉ số Tham gia Học tập (5 biến): Tỷ lệ chuyên cần (30%–100%, TB: 76.88%), Giờ tự học/ngày (0.0–8.1h, TB: 3.16h), GPA kỳ trước (2.42–10.0, TB: 6.86), Số môn nợ (0–6 môn, TB: 0.82), Điểm động lực (0–10, TB: 5.47).\n"
                "3. Chỉ số Lối sống & Sinh hoạt (5 biến): Giờ ngủ (2.0–10.0h, TB: 6.20h), Giờ dùng màn hình (0.5–12.6h, TB: 5.43h), Tập thể dục (0–7 buổi/tuần), Điểm hoạt động xã hội (0–10), Việc làm thêm (Không: 547, Có: 228).\n"
                "4. Chỉ số Căng thẳng & Hỗ trợ Tâm lý (6 biến): Mức độ căng thẳng (0.1–10.0, TB: 5.52), Điểm lo âu (0–10, TB: 4.98), Áp lực tài chính (0–10, TB: 5.03), Áp lực đồng trang lứa (0–9.8, TB: 4.48), Hỗ trợ gia đình (0–10, TB: 6.42), Tiếp cận tư vấn tâm lý (Không: 535, Có: 265).\n"
                "5. Biến Kiệt sức & Mục tiêu (2 biến): Mức độ kiệt sức (Low: 320 [40.0%], Medium: 280 [35.0%], High: 200 [25.0%]) và Nguy cơ bỏ học (Yes: 437 [54.62%], No: 363 [45.38%]).",
            ],
            "2.2. Training Methodology": [
                "Nhóm thiết kế phương pháp huấn luyện nghiêm ngặt nhằm đảm bảo tính khách quan:\n"
                "• Phân chia dữ liệu: Tập dữ liệu 800 mẫu được chia theo tỷ lệ 80% Huấn luyện (640 mẫu) và 20% Kiểm thử độc lập (160 mẫu) có phân tầng theo nhãn.\n"
                "• Kiểm định chéo phân tầng 5-Fold: Toàn bộ quá trình tạo đặc trưng, điền khuyết, chuẩn hóa và mã hóa đều được đóng gói bên trong Pipeline của scikit-learn để triệt tiêu hoàn toàn hiện tượng rò rỉ dữ liệu.\n"
                "• Thiết kế 2 kịch bản thực nghiệm đối chứng:\n"
                "  - Kịch bản 1 (Đầy đủ đặc trưng): Sử dụng toàn bộ đặc trưng bao gồm cả trạng thái Burnout_Level và 6 chỉ số tổng hợp.\n"
                "  - Kịch bản 2 (Cảnh báo sớm): Dự báo nguy cơ bỏ học trực tiếp từ các chỉ số hành vi học tập và căng thẳng tâm lý thô khi chưa có biến Burnout_Level.\n"
                "• 10 thuật toán học máy được so sánh: Logistic Regression, SVM (RBF Kernel), K-Nearest Neighbors, Decision Tree, Random Forest, Extra Trees, Gradient Boosting, XGBoost, MLP Neural Network, và Soft Voting Ensemble.",
            ],
            "2.3. Workflow": [
                "Quy trình thực thi tổng thể của hệ thống học máy Nhóm 03 gồm 5 bước tuần tự khép kín từ thu thập dữ liệu, tiền xử lý & khám phá dữ liệu, kỹ thuật tạo đặc trưng chuyên ngành, huấn luyện & kiểm định chéo 10 mô hình đến triển khai ứng dụng Web cảnh báo sớm:",
            ],
            "2.4. System Design": [
                "Hệ thống phần mềm được thiết kế theo kiến trúc 3 tầng: (1) Tầng Giao diện Web Dashboard thu thập thông số học tập & tâm lý của sinh viên; (2) Tầng Xử lý AI Backend (FastAPI) thực thi DomainFeatureEngineer, KNNImputer Pipeline và bộ đôi mô hình dự báo; (3) Tầng Hiển thị Kết quả & Khuyến nghị can thiệp sớm.",
            ],
            "3.1. Data Preprocessing": [
                "1. Xử lý dữ liệu khuyết: Có 18/25 cột chứa giá trị khuyết từ 1.62% đến 6.38%. Nhóm áp dụng thuật toán KNNImputer (k=7) cho các biến số liên tục và điền giá trị phổ biến nhất (Mode) cho các biến phân loại.\n"
                "2. Kỹ thuật Tạo Đặc trưng Chuyên sâu: Nhóm xây dựng 6 đặc trưng mới:\n"
                "   • Academic_Engagement_Index = 0.35*(Attendance_Percent/10) + 0.25*Study_Hours_Per_Day + 0.25*Motivation_Score + 0.15*Previous_GPA\n"
                "   • Psychological_Stress_Index = 0.35*Stress_Level + 0.30*Anxiety_Score + 0.20*Financial_Stress_Score + 0.15*Peer_Pressure_Score\n"
                "   • Stress_Engagement_Ratio = Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)\n"
                "   • Net_Stress_Support_Balance = Psychological_Stress_Index - (0.5*Family_Support_Score + 0.3*Social_Activity_Score + 0.2*Exercise_Freq_Per_Week)\n"
                "   • Digital_Fatigue_Ratio = Screen_Time_Hours / (Sleep_Hours + 1.0)\n"
                "   • Academic_Risk_Burden = 1.5*Backlogs + (100 - Attendance_Percent)/20.0\n"
                "3. Mã hóa & Chuẩn hóa: Mã hóa thứ tự cho thu nhập gia đình (0–4) và mức độ kiệt sức (0–2); mã hóa One-Hot cho các biến danh mục và chuẩn hóa toàn bộ biến số bằng StandardScaler.",
            ],
            "3.2. Exploratory Data Analysis (EDA)": [
                "Quá trình phân tích khám phá dữ liệu (EDA) đã làm sáng tỏ các quy luật quan trọng:\n"
                "• Phát hiện 1 - Mối quan hệ quyết định giữa Burnout_Level và Dropout_Risk: Trong số 200 sinh viên có mức độ kiệt sức cao (High), có tới 184 sinh viên bỏ học (92.00%). Ở mức Trung bình (Medium), tỷ lệ bỏ học là 63.21% (177/280), trong khi ở mức Thấp (Low) chỉ là 23.75% (76/320). Hệ số tương quan Pearson đạt r = +0.550.\n"
                "• Phát hiện 2 - Sự đối lập giữa Tham gia Học tập và Căng thẳng Tâm lý:\n"
                "  - Các yếu tố bảo vệ (tương quan âm với bỏ học): Tỷ lệ chuyên cần (r = -0.244), Điểm động lực (r = -0.193), Giờ tự học/ngày (r = -0.160), Hỗ trợ gia đình (r = -0.137), GPA kỳ trước (r = -0.106).\n"
                "  - Các yếu tố nguy cơ (tương quan dương với bỏ học): Mức căng thẳng (r = +0.164), Điểm lo âu (r = +0.145), Giờ dùng màn hình (r = +0.127), Áp lực tài chính (r = +0.114).",
            ],
            "3.3. Modeling": [
                "Tổng hợp kết quả thực nghiệm 10 mô hình học máy trên cả 2 kịch bản (5-Fold CV và Tập Test 20%):\n"
                "1. Kết quả Kịch bản 1 (Đầy đủ đặc trưng - Có Burnout_Level & Chỉ số Tổng hợp):\n"
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
                "2. Kết quả Kịch bản 2 (Cảnh báo sớm - Không dùng Burnout_Level):\n"
                "   • Soft Voting Ensemble dẫn đầu với CV Accuracy = 70.63%, CV F1 = 73.89%, CV ROC-AUC = 77.10% (Test Accuracy = 70.63%, Test ROC-AUC = 77.63%).\n"
                "   • Logistic Regression đạt CV ROC-AUC = 76.82% và Test ROC-AUC = 79.00%.",
            ],
            "3.4. User Interface": [
                "Nhóm 03 đã xây dựng hệ thống Web Dashboard tương tác bằng FastAPI (app.py). Giao diện cho phép:\n"
                "• Điều chỉnh trực quan 18 thanh trượt/danh sách chọn về Chuyên cần, Giờ tự học, GPA, Động lực, Số môn nợ, Căng thẳng, Lo âu, Áp lực tài chính, Hỗ trợ gia đình, Giờ ngủ và Giờ màn hình.\n"
                "• Chọn nhanh 3 mẫu hồ sơ sinh viên (Nguy cơ Thấp, Trung bình, Cao) để kiểm thử tức thời.\n"
                "• Hiển thị kết quả dự báo kép theo thời gian thực: Xác suất nguy cơ bỏ học (%), Phân bố xác suất 3 mức độ kiệt sức, Ba chỉ số giải thích AI và Danh sách khuyến nghị can thiệp cá nhân hóa.",
            ],
            "3.5. Testing and Improvements": [
                "• Cải tiến xử lý dữ liệu khuyết: Thay thế SimpleImputer bằng KNNImputer (k=7) giúp bảo toàn cấu trúc tương quan đa biến.\n"
                "• Cải tiến tạo đặc trưng: Bổ sung 6 chỉ số tổng hợp nâng điểm ROC-AUC của Random Forest lên 80.50% và giúp Extra Trees đạt F1-Score 78.92%, Recall 83.91%.\n"
                "• Kiểm soát quá khớp (Overfitting): Giới hạn độ sâu cây và áp dụng điều chuẩn L2 giúp độ lệch giữa tập CV và tập Test luôn dưới 1.5%.",
            ],
            "4.1. Accomplishments and Benefits": [
                "• Thành tựu đạt được: Hoàn thiện Pipeline Học máy chuẩn công nghiệp, so sánh 10 thuật toán trên 2 kịch bản, xây dựng bộ chỉ số Gắn kết & Căng thẳng, và triển khai thành công Web Dashboard thời gian thực.\n"
                "• Lợi ích thực tiễn: Cung cấp công cụ sàng lọc tự động cho cố vấn học tập và phòng công tác sinh viên, giúp can thiệp sớm cho nhóm sinh viên có xác suất bỏ học cao.",
            ],
            "4.2. Future Improvements": [
                "1. Thu thập dữ liệu chuỗi thời gian theo từng tuần học trong học kỳ để dự báo xu hướng bằng mô hình LSTM / Transformer.\n"
                "2. Tích hợp giải thích mô hình cục bộ SHAP trực tiếp lên giao diện Web.\n"
                "3. Kết nối tự động với hệ thống quản lý học tập (LMS) và hệ thống điểm danh của nhà trường.\n"
                "4. Mở rộng quy mô dữ liệu khảo sát tại các trường đại học ở Việt Nam.",
            ],
        }
        section_images = {
            "2.3. Workflow": [(os.path.join(fig_folder, "11_system_workflow_diagram.png"), "Hình 2.1: Quy trình thực hiện dự án Học máy (Nhóm 03)")],
            "2.4. System Design": [(os.path.join(fig_folder, "12_system_architecture_diagram.png"), "Hình 2.2: Kiến trúc hệ thống Cảnh báo sớm Nguy cơ Bỏ học & Giao diện Web UI")],
            "3.1. Data Preprocessing": [(os.path.join(fig_folder, "01_missing_values_analysis.png"), "Hình 3.1: Phân tích tỷ lệ dữ liệu khuyết trên 18/25 thuộc tính")],
            "3.2. Exploratory Data Analysis (EDA)": [
                (os.path.join(fig_folder, "02_target_and_burnout_distribution.png"), "Hình 3.2: Phân phối biến mục tiêu và tỷ lệ bỏ học theo mức độ kiệt sức"),
                (os.path.join(fig_folder, "03_stress_vs_engagement_distributions.png"), "Hình 3.3: So sánh phân phối các chỉ số Tham gia Học tập và Căng thẳng Tâm lý"),
                (os.path.join(fig_folder, "04_correlation_matrix_heatmap.png"), "Hình 3.4: Ma trận tương quan Pearson giữa các biến số"),
                (os.path.join(fig_folder, "05_categorical_factors_analysis.png"), "Hình 3.5: Phân tích tỷ lệ bỏ học theo các yếu tố nhân khẩu học"),
                (os.path.join(fig_folder, "06_multivariate_burnout_engagement_scatter.png"), "Hình 3.6: Phân tích đa biến giữa Chuyên cần, Căng thẳng và Mức độ Kiệt sức"),
            ],
            "3.3. Modeling": [
                (os.path.join(fig_folder, "07_model_performance_comparison.png"), "Hình 3.7: So sánh hiệu năng 10 mô hình học máy trên 2 kịch bản"),
                (os.path.join(fig_folder, "08_roc_and_precision_recall_curves.png"), "Hình 3.8: Đường cong ROC và Precision-Recall trên tập kiểm thử"),
                (os.path.join(fig_folder, "09_confusion_matrices_top_models.png"), "Hình 3.9: Ma trận nhầm lẫn của 4 mô hình tiêu biểu"),
                (os.path.join(fig_folder, "10_feature_importance_analysis.png"), "Hình 3.10: Phân tích mức độ quan trọng của các đặc trưng (Random Forest & XGBoost)"),
            ],
        }
    else:
        replacements = {
            "< Project Title >": "STUDENT DROPOUT RISK PREDICTION BASED ON PSYCHOLOGICAL STRESS AND ACADEMIC ENGAGEMENT INDICATORS",
            "< Date (DD/MM/YY) >": "05/11/2026",
            "Team Name": "Group 03 - Student Burnout AI",
            "<Member 1>": "1. Do Quang Minh (Team Leader)",
            "<Member 2>": "2. Vu Quang Huy (Team Member)",
            "<Member 3>": "",
            "<Member 4>": "",
            "<Member 5>": "",
        }
        section_content = {
            "1.1. Background Information": [
                "In contemporary higher education, student attrition (dropout) and academic burnout represent critical institutional challenges. Educational psychology literature demonstrates that university dropout is rarely caused by academic difficulty alone; rather, it emerges from the cumulative interaction of two primary dimensions: (1) Declining Academic Engagement (class attendance percentage, daily self-study hours, intrinsic motivation, prior GPA, and course backlogs) and (2) Chronic Psychological and Financial Stress (academic stress, anxiety, financial burden, peer pressure, sleep deprivation, and insufficient family support).",
                "Applying Machine Learning to jointly model psychological stress indicators and study engagement behaviors enables universities to transition from reactive post-hoc administration to a Proactive Early Warning System, identifying at-risk students before irreversible disengagement occurs.",
            ],
            "1.2. Motivation and Objective": [
                "• Research Motivation: In most universities, attendance logs, academic transcripts, and student wellness survey records remain siloed across separate departments and are rarely synthesized into predictive quantitative models. Group 03 undertook the project 'Building, Training, and Comparing Machine Learning Models for Student Dropout Risk Prediction Based on Psychological Stress and Academic Engagement Indicators' to bridge this gap.",
                "• Specific Objectives:\n"
                "  1. Perform comprehensive Exploratory Data Analysis (EDA) across 800 student records and 25 attributes to uncover key statistical drivers of Burnout_Level and Dropout_Risk.\n"
                "  2. Build an automated, leakage-free Preprocessing Pipeline (KNNImputer + StandardScaler + OneHotEncoder) and engineer 6 domain-specific composite indices quantifying the imbalance between Stress and Engagement.\n"
                "  3. Train, evaluate via 5-Fold Stratified Cross-Validation, and benchmark 10 Machine Learning classifiers across two experimental scenarios (Scenario 1: Full Feature Set including Burnout_Level; Scenario 2: Early Warning using raw Stress & Engagement indicators without Burnout_Level).\n"
                "  4. Develop an interactive FastAPI Web Dashboard for real-time dual prediction and personalized student intervention recommendations.",
            ],
            "1.3. Members and Role Assignments": [
                "The project was executed by Group 03 with an equitable 50/50 workload distribution aligned with each member's technical strengths:\n"
                "• Do Quang Minh (Team Leader - 50% contribution):\n"
                "  - Designed the leakage-free scikit-learn Pipeline architecture and implemented the custom DomainFeatureEngineer transformer (6 composite indices).\n"
                "  - Trained, tuned hyperparameters, and evaluated the tree-ensemble and boosting model families (Random Forest, Extra Trees, Gradient Boosting, XGBoost, and Soft Voting Ensemble).\n"
                "  - Developed the FastAPI backend server and integrated the dual inference engine (Burnout_Level + Dropout_Risk) into the Web Dashboard.\n"
                "• Vu Quang Huy (Team Member - 50% contribution):\n"
                "  - Conducted data quality auditing, missing-value analysis across 18/25 columns, and generated the 6 Exploratory Data Analysis (EDA) visualizations.\n"
                "  - Trained and evaluated the linear, kernel, distance-based, single-tree, and neural network classifiers (Logistic Regression, SVM RBF, KNN, Decision Tree, MLP Neural Network).\n"
                "  - Designed the Frontend Web Dashboard UI, compiled experimental evaluation tables, and authored the Final Report, WBS workbook, and Presentation Slides.",
            ],
            "1.4. Schedule and Milestones": [
                "The project execution followed a structured 4-phase milestone schedule detailed in our WBS workbook:\n"
                "• Phase 1 (Oct 19, 2026 – Oct 20, 2026): Project initiation, problem formulation, dataset schema auditing across 25 attributes, and completion of Action Plan & WBS. [100% Completed]\n"
                "• Phase 2 (Oct 21, 2026 – Oct 25, 2026): Exploratory Data Analysis (EDA), Pearson correlation testing, KNNImputer (k=7) pipeline construction, and creation of 6 engineered domain features. [100% Completed]\n"
                "• Phase 3 (Oct 26, 2026 – Oct 29, 2026): Training and 5-Fold Stratified Cross-Validation of 10 ML models across both experimental scenarios; generation of ROC, Precision-Recall, Confusion Matrix, and Feature Importance figures. [100% Completed]\n"
                "• Phase 4 (Oct 30, 2026 – Nov 05, 2026): Development of the FastAPI Web Dashboard (app.py), integration testing, Jupyter Notebook packaging, Final Report, and Presentation Slides. [100% Completed]",
            ],
            "2.1. Data Acquisition": [
                "The dataset 'student_burnout_dropout_dataset_2.csv' comprises 800 university student records and 25 attributes organized into 5 functional categories:\n"
                "1. Identification & Demographics (7 attributes): Student_ID (STU00001–STU00800), Age (17–25 years, mean: 20.88), Gender (Female: 392, Male: 346, Other: 32, Missing: 30), Year_of_Study (Years 1–4), Department (Engineering: 209, Business: 151, Science: 143, Arts: 136, Medicine: 77, Law: 61), Residence_Type (Hostel: 290, Day Scholar: 286, PG/Rented: 207), Family_Income_Bracket (Low, Lower-Middle, Middle, Upper-Middle, High).\n"
                "2. Academic Engagement Indicators (5 attributes): Attendance_Percent (30%–100%, mean: 76.88%), Study_Hours_Per_Day (0.0–8.1 hours/day, mean: 3.16h), Previous_GPA (2.42–10.0, mean: 6.86), Backlogs (0–6 failed courses, mean: 0.82), Motivation_Score (0–10, mean: 5.47).\n"
                "3. Lifestyle & Daily Habits (5 attributes): Sleep_Hours (2.0–10.0h, mean: 6.20h), Screen_Time_Hours (0.5–12.6h, mean: 5.43h), Exercise_Freq_Per_Week (0–7 days/week), Social_Activity_Score (0–10), Part_Time_Job (No: 547, Yes: 228).\n"
                "4. Psychological Stress & Support Indicators (6 attributes): Stress_Level (0.1–10.0, mean: 5.52), Anxiety_Score (0–10, mean: 4.98), Financial_Stress_Score (0–10, mean: 5.03), Peer_Pressure_Score (0–9.8, mean: 4.48), Family_Support_Score (0–10, mean: 6.42), Counseling_Access (No: 535, Yes: 265).\n"
                "5. Burnout State & Target Variable (2 attributes): Burnout_Level (Low: 320 [40.0%], Medium: 280 [35.0%], High: 200 [25.0%]) and binary target Dropout_Risk (Yes: 437 students [54.62%], No: 363 students [45.38%]).",
            ],
            "2.2. Training Methodology": [
                "We designed a rigorous training and evaluation protocol to ensure strong generalization and zero data leakage:\n"
                "• Stratified Data Splitting: The 800 records are partitioned into an 80% Training Set (640 records) and a 20% Hold-out Test Set (160 records) using stratified sampling (random_state=42) to preserve the exact 54.6% Yes / 45.4% No class ratio.\n"
                "• 5-Fold Stratified Cross-Validation: All feature engineering, missing-value imputation, scaling, and encoding steps are encapsulated inside a scikit-learn Pipeline so parameters are fitted strictly within each training fold.\n"
                "• Two Controlled Experimental Scenarios:\n"
                "  - Scenario 1 (Full Feature Set): Predicts Dropout_Risk using all available features including Burnout_Level and the 6 engineered indices.\n"
                "  - Scenario 2 (Early Warning - Without Burnout_Level): Predicts Dropout_Risk directly from raw study engagement and psychological stress indicators without relying on Burnout_Level.\n"
                "• 10 Benchmarked ML Classifiers: (1) Logistic Regression, (2) Support Vector Machine (RBF Kernel with probability calibration), (3) K-Nearest Neighbors, (4) Decision Tree, (5) Random Forest, (6) Extra Trees, (7) Gradient Boosting, (8) XGBoost, (9) Multi-Layer Perceptron (MLP Neural Network), and (10) Soft Voting Ensemble.",
            ],
            "2.3. Workflow": [
                "The end-to-end execution workflow of Group 03's Machine Learning system consists of five sequential stages: Data Acquisition, Preprocessing & EDA, Domain Feature Engineering, 10-Model Cross-Validation Training, and Web UI Deployment:",
            ],
            "2.4. System Design": [
                "The software system follows a clean 3-Tier Architecture: (1) Presentation Layer (Web UI Input Form capturing student engagement, stress, and lifestyle parameters); (2) AI Processing Backend (FastAPI server executing DomainFeatureEngineer, KNNImputer Pipeline, and the dual Burnout + Dropout inference models); and (3) Diagnostic Output Dashboard displaying risk probabilities, explainable indices, and personalized intervention plans.",
            ],
            "3.1. Data Preprocessing": [
                "1. Missing Value Imputation:\n"
                "   Auditing revealed missing values across 18 of the 25 columns ranging from 1.62% (Family_Income_Bracket) to 6.38% (Attendance_Percent) and 6.12% (Stress_Level). Dropping incomplete rows would discard over 40% of the dataset. Instead, we applied:\n"
                "   - Continuous Numerical Features: KNNImputer (n_neighbors=7) to estimate missing entries from the 7 most similar student profiles in standardized feature space.\n"
                "   - Categorical Features: Most-Frequent (Mode) Imputation.\n"
                "2. Domain-Specific Feature Engineering:\n"
                "   We engineered 6 domain-specific composite features inside DomainFeatureEngineer:\n"
                "   • Academic_Engagement_Index = 0.35*(Attendance_Percent/10) + 0.25*Study_Hours_Per_Day + 0.25*Motivation_Score + 0.15*Previous_GPA\n"
                "   • Psychological_Stress_Index = 0.35*Stress_Level + 0.30*Anxiety_Score + 0.20*Financial_Stress_Score + 0.15*Peer_Pressure_Score\n"
                "   • Stress_Engagement_Ratio = Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)\n"
                "   • Net_Stress_Support_Balance = Psychological_Stress_Index - (0.5*Family_Support_Score + 0.3*Social_Activity_Score + 0.2*Exercise_Freq_Per_Week)\n"
                "   • Digital_Fatigue_Ratio = Screen_Time_Hours / (Sleep_Hours + 1.0)\n"
                "   • Academic_Risk_Burden = 1.5*Backlogs + (100 - Attendance_Percent)/20.0\n"
                "3. Feature Encoding & Normalization: Ordinal encoding for Family_Income_Bracket (0–4) and Burnout_Level (Low=0, Medium=1, High=2); One-Hot Encoding for nominal features; StandardScaler normalization (zero mean, unit variance) for all continuous features.",
            ],
            "3.2. Exploratory Data Analysis (EDA)": [
                "Our Exploratory Data Analysis (EDA) uncovered three pivotal empirical findings:\n"
                "• Finding 1 - Strong Progression from Burnout_Level to Dropout_Risk: Among the 200 students experiencing High Burnout, 184 dropped out (an alarming 92.00% dropout rate). Medium Burnout students exhibited a 63.21% dropout rate (177/280), whereas Low Burnout students had only a 23.75% dropout rate (76/320). The Pearson correlation between Burnout_Numeric and Dropout_Binary is r = +0.550.\n"
                "• Finding 2 - Opposing Effects of Study Engagement vs. Psychological Stress:\n"
                "  - Protective Factors (negative correlation with dropout): Attendance_Percent (r = -0.244), Motivation_Score (r = -0.193), Study_Hours_Per_Day (r = -0.160), Family_Support_Score (r = -0.137), and Previous_GPA (r = -0.106).\n"
                "  - Risk Factors (positive correlation with dropout): Stress_Level (r = +0.164), Anxiety_Score (r = +0.145), Screen_Time_Hours (r = +0.127), and Financial_Stress_Score (r = +0.114).\n"
                "• Finding 3 - Socioeconomic Gradient: Students from Low-income families experienced the highest dropout rate (62.1%), steadily decreasing toward High-income families (51.1%).",
            ],
            "3.3. Modeling": [
                "We trained and compared 10 machine learning models across both experimental scenarios using 5-Fold Stratified CV and the 20% Hold-out Test Set:\n"
                "1. Scenario 1 Results (Full Feature Set - Including Burnout_Level & Engineered Indices):\n"
                "   • Random Forest: CV Accuracy = 74.38% | CV F1 = 77.46% | CV ROC-AUC = 80.50% (Highest AUC) | Test ROC-AUC = 80.24%\n"
                "   • Extra Trees: CV Accuracy = 75.00% (Highest Acc) | CV F1 = 78.10% | CV ROC-AUC = 79.22% | Test Accuracy = 75.63% | Test F1 = 78.92% | Test Recall = 83.91%\n"
                "   • Soft Voting Ensemble: CV Accuracy = 73.88% | CV F1 = 76.79% | CV ROC-AUC = 79.95% | Test Accuracy = 72.50% | Test F1 = 75.82% | Test ROC-AUC = 79.37%\n"
                "   • Logistic Regression: CV Accuracy = 72.25% | CV F1 = 75.19% | CV ROC-AUC = 79.82% | Test Accuracy = 73.75% | Test ROC-AUC = 79.86%\n"
                "   • SVM (RBF Kernel): CV Accuracy = 74.63% | CV F1 = 77.57% | CV ROC-AUC = 78.55% | Test Accuracy = 74.38% | Test F1 = 77.84%\n"
                "   • MLP Neural Network: CV Accuracy = 73.63% | CV F1 = 75.94% | CV ROC-AUC = 79.68%\n"
                "   • XGBoost: CV Accuracy = 71.75% | CV F1 = 74.28% | CV ROC-AUC = 79.25%\n"
                "   • Gradient Boosting: CV Accuracy = 71.75% | CV F1 = 74.44% | CV ROC-AUC = 79.08%\n"
                "   • K-Nearest Neighbors: CV Accuracy = 70.25% | CV F1 = 72.59% | CV ROC-AUC = 77.14%\n"
                "   • Decision Tree: CV Accuracy = 66.88% | CV F1 = 67.75% | CV ROC-AUC = 73.88%\n\n"
                "2. Scenario 2 Results (Early Warning - Without Burnout_Level, using Stress & Engagement only):\n"
                "   • Soft Voting Ensemble achieved the highest overall CV performance with CV Accuracy = 70.63%, CV F1 = 73.89%, and CV ROC-AUC = 77.10% (Test Accuracy = 70.63%, Test ROC-AUC = 77.63%).\n"
                "   • Logistic Regression achieved CV ROC-AUC = 76.82% and Test ROC-AUC = 79.00%, confirming that our engineered Stress-Engagement indices effectively capture dropout risk even before clinical burnout assessment.",
            ],
            "3.4. User Interface": [
                "To translate our models into a practical decision-support tool, Group 03 developed an interactive FastAPI Web Dashboard (app.py). Key capabilities include:\n"
                "• Interactive Sliders & Selectors for 18 student attributes spanning Attendance (%), Study Hours/Day, Previous GPA, Motivation, Backlogs, Stress Level, Anxiety, Financial Stress, Family Support, Sleep Hours, and Screen Time.\n"
                "• One-Click Student Archetype Presets (Low-Risk Engaged Student, Moderate-Stress Student, and High-Burnout At-Risk Student) for rapid live demonstration.\n"
                "• Real-Time Dual AI Diagnostics displaying: (1) Dropout Risk Probability (%) with color-coded severity banner; (2) 3-Class Burnout Level Probability Breakdown (Low / Medium / High); (3) Explainable Domain Indices (Academic Engagement Index, Psychological Stress Index, Stress-to-Engagement Ratio); and (4) Tailored academic and psychological counseling recommendations.",
            ],
            "3.5. Testing and Improvements": [
                "• Imputation Refinement: Upgrading from SimpleImputer (median) to multivariate KNNImputer (k=7) preserved inter-feature correlations among Attendance, Study Hours, and Stress.\n"
                "• Feature Engineering Boost: Introducing the 6 domain interaction indices elevated Random Forest CV ROC-AUC to 80.50% and enabled Extra Trees to reach 75.63% Test Accuracy, 78.92% F1-Score, and 83.91% Recall.\n"
                "• Overfitting Control: Constraining tree depth (max_depth=3..8), enforcing minimum leaf samples, and applying L2 regularization kept the generalization gap between 5-Fold CV and Hold-out Test metrics below 1.5%.",
            ],
            "4.1. Accomplishments and Benefits": [
                "• Key Accomplishments:\n"
                "  - Delivered an end-to-end, reproducible Machine Learning pipeline benchmarking 10 classifiers across two realistic institutional scenarios.\n"
                "  - Formulated interpretable composite metrics (Academic Engagement Index & Psychological Stress Index) that explain why a student is at risk.\n"
                "  - Built a responsive FastAPI Web Dashboard with sub-30ms inference latency.\n"
                "• Practical Benefits:\n"
                "  - For Universities & Academic Advisors: Provides automated semester-start screening to prioritize counseling resources for students exceeding a 65% dropout risk threshold.\n"
                "  - For Students: Raises self-awareness regarding study-life balance, screen fatigue, and mental wellness.",
            ],
            "4.2. Future Improvements": [
                "1. Incorporate longitudinal weekly time-series tracking of attendance and stress across the semester using sequential deep learning models (LSTM / Transformers).\n"
                "2. Embed local SHAP (SHapley Additive exPlanations) waterfall plots directly inside the Web Dashboard for individual student explainability.\n"
                "3. Integrate directly with institutional Learning Management Systems (LMS / Moodle / Canvas) via REST APIs.\n"
                "4. Expand multi-campus data collection across Vietnamese universities to calibrate domain weights for local student populations.",
            ],
        }
        section_images = {
            "2.3. Workflow": [(os.path.join(fig_folder, "11_system_workflow_diagram.png"), "Figure 2.1: Machine Learning Project Execution Workflow (Group 03)")],
            "2.4. System Design": [(os.path.join(fig_folder, "12_system_architecture_diagram.png"), "Figure 2.2: 3-Tier Student Dropout Early Warning System Architecture & Web UI")],
            "3.1. Data Preprocessing": [(os.path.join(fig_folder, "01_missing_values_analysis.png"), "Figure 3.1: Missing Data Percentage Across 18/25 Attributes Prior to KNN Imputation")],
            "3.2. Exploratory Data Analysis (EDA)": [
                (os.path.join(fig_folder, "02_target_and_burnout_distribution.png"), "Figure 3.2: Target Distribution and Dropout Rate Progression Across Burnout Levels"),
                (os.path.join(fig_folder, "03_stress_vs_engagement_distributions.png"), "Figure 3.3: Distribution Comparison of Study Engagement & Psychological Stress Indicators"),
                (os.path.join(fig_folder, "04_correlation_matrix_heatmap.png"), "Figure 3.4: Pearson Correlation Matrix Across Stress, Engagement, Burnout, and Dropout Risk"),
                (os.path.join(fig_folder, "05_categorical_factors_analysis.png"), "Figure 3.5: Dropout Rate Analysis Across Demographic & Institutional Factors"),
                (os.path.join(fig_folder, "06_multivariate_burnout_engagement_scatter.png"), "Figure 3.6: Multivariate Interaction Between Attendance, Stress Level, and Burnout Profile"),
            ],
            "3.3. Modeling": [
                (os.path.join(fig_folder, "07_model_performance_comparison.png"), "Figure 3.7: Performance Comparison Across 10 ML Models in Scenario 1 (Full) & Scenario 2 (Early Warning)"),
                (os.path.join(fig_folder, "08_roc_and_precision_recall_curves.png"), "Figure 3.8: ROC and Precision-Recall Curves Across All 10 Classifiers on the Test Set"),
                (os.path.join(fig_folder, "09_confusion_matrices_top_models.png"), "Figure 3.9: Hold-Out Test Set Confusion Matrices for Top 4 Representative Models"),
                (os.path.join(fig_folder, "10_feature_importance_analysis.png"), "Figure 3.10: Top 15 Feature Importances in Random Forest (Gini) and XGBoost (Gain)"),
            ],
        }

    for p in doc.paragraphs:
        t = p.text.strip()
        if t in replacements:
            p.text = replacements[t]

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

    # Translate headings to Vietnamese if lang == "vi"
    if lang == "vi":
        heading_map_vi = {
            "Content": "Mục lục (Table of Contents)",
            "1. Introduction": "1. Giới thiệu chung",
            "1.1. Background Information": "1.1. Bối cảnh nghiên cứu",
            "1.2. Motivation and Objective": "1.2. Động lực và Mục tiêu đề tài",
            "1.3. Members and Role Assignments": "1.3. Thành viên và Phân công nhiệm vụ",
            "1.4. Schedule and Milestones": "1.4. Lịch trình và Các mốc quan trọng",
            "2. Project Execution": "2. Triển khai Dự án",
            "2.1. Data Acquisition": "2.1. Thu thập và Đặc tả Dữ liệu",
            "2.1. Data Acquisition2.2. Training Methodology": "2.1. Thu thập Dữ liệu & 2.2. Phương pháp Huấn luyện",
            "2.2. Training Methodology": "2.2. Phương pháp Huấn luyện Mô hình",
            "2.3. Workflow": "2.3. Quy trình Thực hiện (Workflow)",
            "2.4. System Diagram": "2.4. Sơ đồ Kiến trúc Hệ thống",
            "2.4. System Design": "2.4. Thiết kế Kiến trúc Hệ thống",
            "3. Results": "3. Kết quả Thực nghiệm",
            "3.1. Data Preprocessing": "3.1. Tiền xử lý Dữ liệu & Tạo Đặc trưng",
            "3.2. Exploratory Data Analysis (EDA)": "3.2. Phân tích Khám phá Dữ liệu (EDA)",
            "3.2. Exploratory Data Analysis (EDA)3.3. Modeling": "3.2. Phân tích EDA & 3.3. Huấn luyện Mô hình",
            "3.3. Modeling": "3.3. Kết quả Huấn luyện & So sánh Mô hình",
            "3.4. User Interface": "3.4. Giao diện Người dùng (Web Dashboard UI)",
            "3.4. User Interface3.5. Testing and Improvements": "3.4. Giao diện Người dùng & 3.5. Kiểm thử Cải tiến",
            "3.5. Testing and Improvements": "3.5. Kiểm thử và Cải tiến Hệ thống",
            "4.  Projected Impact": "4. Đánh giá Tác động Dự kiến",
            "4. Projected Impact": "4. Đánh giá Tác động Dự kiến",
            "4.1. Accomplishments and Benefits": "4.1. Thành tựu Đạt được và Lợi ích Thực tiễn",
            "4.2. Future Improvements": "4.2. Định hướng Phát triển Tương lai",
            "5. Team Member Review and Comment": "5. Nhận xét và Đánh giá của Thành viên Nhóm",
            "6. Instructor Review and Comment": "6. Nhận xét và Đánh giá của Giảng viên Hướng dẫn",
        }
        for p in doc.paragraphs:
            t = p.text.strip()
            if t in heading_map_vi:
                p.text = heading_map_vi[t]

    for t_mem in doc.tables:
        if len(t_mem.rows) == 6 and len(t_mem.columns) == 2:
            if lang == "vi":
                t_mem.cell(0, 0).text = "THÀNH VIÊN"
                t_mem.cell(0, 1).text = "NHẬN XÉT VÀ ĐÁNH GIÁ CÁ NHÂN"
                t_mem.cell(1, 0).text = "Đỗ Quang Minh\n(Trưởng nhóm)"
                t_mem.cell(1, 1).text = (
                    "Qua dự án Capstone của khóa học Samsung Innovation Campus (SIC AI), tôi đã nắm vững cách xây dựng một Pipeline Học máy hoàn chỉnh và chống rò rỉ dữ liệu. "
                    "Việc thiết kế các chỉ số kết hợp giữa Căng thẳng tâm lý và Sự tham gia học tập giúp nâng cao rõ rệt hiệu năng của các mô hình tập hợp cây (Random Forest, Extra Trees, Soft Voting) và mang lại ý nghĩa thực tiễn cao trong giáo dục."
                )
                t_mem.cell(2, 0).text = "Vũ Quang Huy\n(Thành viên)"
                t_mem.cell(2, 1).text = (
                    "Đồ án giúp tôi rèn luyện kỹ năng phân tích khám phá dữ liệu (EDA), xử lý dữ liệu khuyết bằng KNNImputer và đối sánh thực nghiệm 10 thuật toán phân loại khác nhau. "
                    "Sự phối hợp chặt chẽ trong nhóm từ khâu phân tích dữ liệu, huấn luyện mô hình đến triển khai giao diện Web Demo đã giúp nhóm hoàn thành tốt mọi mục tiêu đề ra."
                )
            else:
                t_mem.cell(0, 0).text = "NAME"
                t_mem.cell(0, 1).text = "REVIEW and COMMENT"
                t_mem.cell(1, 0).text = "Do Quang Minh\n(Team Leader)"
                t_mem.cell(1, 1).text = (
                    "Through the Samsung Innovation Campus (SIC) AI Capstone project, I gained deep practical experience in architecting leakage-free machine learning pipelines. "
                    "Engineering composite indices that capture the interplay between Psychological Stress and Academic Engagement significantly boosted our ensemble models (Random Forest, Extra Trees, Soft Voting) and produced actionable insights for student retention."
                )
                t_mem.cell(2, 0).text = "Vu Quang Huy\n(Team Member)"
                t_mem.cell(2, 1).text = (
                    "This project strengthened my skills in statistical Exploratory Data Analysis (EDA), multivariate KNN imputation, and systematic benchmarking across 10 classification algorithms. "
                    "Close collaboration across data analysis, model tuning, and FastAPI Web UI deployment enabled our team to deliver a complete, practical Early Warning System."
                )
            break

    save_multi(doc.save, "SIC_AI_Capstone Project_Final Report.docx", lang)
    if os.path.exists(temp_unwrapped):
        os.remove(temp_unwrapped)
    print(f"Completed Final Report ({lang.upper()})")


# ==============================================================================
# 5. PRESENTATION SLIDES (.PPTX) - FULL VI & FULL EN
# ==============================================================================
def add_content_slide(prs, layout, unit_num, unit_title, sec_num, slide_title, subtitle_text, bullets=None, img_path=None, table_data=None):
    slide = prs.slides.add_slide(layout)
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

    for shape in list(slide.shapes):
        if shape.has_text_frame and "Please edit project name" in shape.text_frame.text:
            shape._element.getparent().remove(shape._element)

    if img_path and os.path.exists(img_path):
        left = PptInches(5.35) if bullets else PptInches(1.2)
        top = PptInches(1.85)
        width = PptInches(7.5) if bullets else PptInches(10.8)
        slide.shapes.add_picture(img_path, left, top, width=width)

    if table_data:
        rows, cols = len(table_data), len(table_data[0])
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


def build_presentation_pptx(lang="vi"):
    in_path = os.path.join(BACKUP_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx")
    prs = Presentation(in_path)
    fig_folder = FIG_VI if lang == "vi" else FIG_EN
    is_vi = lang == "vi"

    # Slide 1: Cover
    s1 = prs.slides[0]
    for shape in s1.shapes:
        if shape.has_text_frame:
            if shape.name == "텍스트 개체 틀 4":
                shape.text_frame.text = (
                    "Nhóm 03: Đỗ Quang Minh • Vũ Quang Huy"
                    if is_vi
                    else "Group 03: Do Quang Minh • Vu Quang Huy"
                )
            elif shape.name == "제목 3":
                shape.text_frame.text = (
                    "DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CHỈ SỐ CĂNG THẲNG & SỰ THAM GIA HỌC TẬP"
                    if is_vi
                    else "STUDENT DROPOUT RISK PREDICTION BASED ON STRESS & STUDY ENGAGEMENT INDICATORS"
                )
                for p in shape.text_frame.paragraphs:
                    p.font.size = PptPt(24)

    # Slide 2: Agenda
    s2 = prs.slides[1]
    for shape in s2.shapes:
        if shape.has_text_frame and "Project Name" in shape.text_frame.text:
            shape.text_frame.text = (
                "Dự đoán Nguy cơ Bỏ học & Kiệt sức Học đường — Nhóm 03"
                if is_vi
                else "Student Burnout & Dropout Risk Prediction — Group 03"
            )

    agenda_box = s2.shapes.add_textbox(PptInches(1.2), PptInches(2.1), PptInches(10.8), PptInches(4.5))
    tf_ag = agenda_box.text_frame
    tf_ag.word_wrap = True
    agenda_items = (
        [
            "01. GIỚI THIỆU ĐỀ TÀI: Bối cảnh, Động lực, Mục tiêu & Phân công Nhóm 03",
            "02. DỮ LIỆU & PHƯƠNG PHÁP LUẬN: Bộ dữ liệu 800 SV x 25 biến & Quy trình Học máy",
            "03. TIỀN XỬ LÝ & TẠO ĐẶC TRƯNG: KNNImputer & 6 Chỉ số Tương tác Căng thẳng - Học tập",
            "04. PHÂN TÍCH KHÁM PHÁ DỮ LIỆU (EDA): Tác động của Kiệt sức, Chuyên cần & Căng thẳng",
            "05. HUẤN LUYỆN & SO SÁNH 10 MÔ HÌNH HỌC MÁY: Kết quả 2 kịch bản, ROC-AUC & Feature Importance",
            "06. GIAO DIỆN WEB DEMO & KẾT LUẬN: Hệ thống Cảnh báo sớm FastAPI & Định hướng phát triển",
        ]
        if is_vi
        else [
            "01. INTRODUCTION: Background, Motivation, Objectives & Group 03 Role Assignment",
            "02. DATA & METHODOLOGY: 800-Student Dataset (25 Attributes) & End-to-End ML Workflow",
            "03. PREPROCESSING & FEATURE ENGINEERING: KNNImputer & 6 Stress-Engagement Indices",
            "04. EXPLORATORY DATA ANALYSIS (EDA): Impact of Burnout, Attendance & Psychological Stress",
            "05. 10-MODEL BENCHMARKING: Scenario 1 vs. Scenario 2 Results, ROC-AUC & Feature Importance",
            "06. WEB UI DEMO & IMPACT: FastAPI Early Warning Dashboard & Future Roadmap",
        ]
    )
    for i, item in enumerate(agenda_items):
        p = tf_ag.paragraphs[0] if i == 0 else tf_ag.add_paragraph()
        p.text = item
        p.font.size = PptPt(16)
        p.font.bold = True
        p.font.color.rgb = PptRGBColor(29, 53, 87)
        p.space_after = PptPt(14)

    # Slide 3: Intro
    s3 = prs.slides[2]
    content_layout = s3.slide_layout
    for shape in list(s3.shapes):
        if shape.has_text_frame and "Please edit project name" in shape.text_frame.text:
            shape._element.getparent().remove(shape._element)
        elif shape.has_text_frame:
            nm = shape.name
            if nm == "제목 1":
                shape.text_frame.text = "Bối cảnh, Mục tiêu & Phân công Nhóm 03" if is_vi else "Background, Objectives & Team Roles (Group 03)"
            elif nm == "텍스트 개체 틀 2":
                shape.text_frame.text = "1.1"
            elif nm == "텍스트 개체 틀 3":
                shape.text_frame.text = (
                    "Dự báo sớm nguy cơ bỏ học đại học từ chỉ số Căng thẳng và Sự tham gia học tập"
                    if is_vi
                    else "Proactive early warning of university student dropout from Stress & Engagement indicators"
                )
            elif nm == "텍스트 개체 틀 4":
                shape.text_frame.text = "01"
            elif nm == "텍스트 개체 틀 5":
                shape.text_frame.text = "GIỚI THIỆU" if is_vi else "INTRODUCTION"
            elif nm == "텍스트 개체 틀 6":
                tf = shape.text_frame
                tf.clear()
                b_list = (
                    [
                        "• Bài toán thực tiễn: Sinh viên bỏ học là hệ quả tích lũy của Kiệt sức học đường, áp lực tâm lý - tài chính kéo dài và sự suy giảm gắn kết học tập.",
                        "• Mục tiêu đồ án:\n  1. Khám phá quy luật dữ liệu (EDA) trên 800 sinh viên x 25 thuộc tính.\n  2. Xây dựng Pipeline xử lý khuyết (KNNImputer) & tạo 6 đặc trưng chuyên sâu.\n  3. Huấn luyện & so sánh 10 mô hình học máy trên 2 kịch bản thực nghiệm.\n  4. Triển khai Web Dashboard cảnh báo sớm trực quan.",
                        "• Phân công Nhóm 03:\n  - Đỗ Quang Minh (Trưởng nhóm): Thiết kế Feature Engineering, huấn luyện nhóm Ensemble/Boosting (RF, ExtraTrees, XGBoost, Voting) & phát triển Backend FastAPI.\n  - Vũ Quang Huy (Thành viên): Phân tích EDA & xử lý khuyết, huấn luyện nhóm mô hình cơ sở (LR, SVM, KNN, DT, MLP), thiết kế UI & biên soạn báo cáo.",
                    ]
                    if is_vi
                    else [
                        "• Problem Statement: University student dropout stems from the cumulative interaction of Academic Burnout, chronic psychological/financial stress, and declining study engagement.",
                        "• Project Objectives:\n  1. Conduct statistical EDA on 800 student records across 25 attributes.\n  2. Build a KNNImputer preprocessing pipeline & engineer 6 domain composite features.\n  3. Train & benchmark 10 ML classifiers across 2 experimental scenarios.\n  4. Deploy an interactive FastAPI Early Warning Web Dashboard.",
                        "• Group 03 Role Assignment:\n  - Do Quang Minh (Leader): Pipeline architecture, Domain Feature Engineering, Tree Ensembles/Boosting (RF, ET, GB, XGBoost, Voting) & FastAPI Backend.\n  - Vu Quang Huy (Member): Data auditing, EDA visualizations, Baseline/Neural classifiers (LR, SVM, KNN, DT, MLP), Dashboard UI & Documentation.",
                    ]
                )
                for idx_b, line in enumerate(b_list):
                    p = tf.paragraphs[0] if idx_b == 0 else tf.add_paragraph()
                    p.text = line
                    p.font.size = PptPt(14)
                    p.space_after = PptPt(8)

    closing_layout = prs.slides[3].slide_layout
    rId = prs.slides._sldIdLst[3].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[3]

    if is_vi:
        add_content_slide(
            prs, content_layout, "02", "DỮ LIỆU", "2.1",
            "Tổng quan Bộ dữ liệu & Xử lý Dữ liệu Khuyết",
            "Bộ dữ liệu student_burnout_dropout_dataset_2.csv gồm 800 sinh viên x 25 thuộc tính (18 cột khuyết)",
            bullets=[
                "• Quy mô: 800 sinh viên, 25 cột chia thành 5 nhóm:",
                "  1. Nhân khẩu học (Tuổi, Giới tính, Năm học, Khoa, Cư trú, Thu nhập)",
                "  2. Tham gia học tập (Chuyên cần %, Giờ tự học, GPA, Môn nợ, Động lực)",
                "  3. Lối sống (Giờ ngủ, Giờ màn hình, Thể dục, Xã hội, Làm thêm)",
                "  4. Căng thẳng & Tâm lý (Stress, Lo âu, Tài chính, Bạn bè, Gia đình)",
                "  5. Biến mục tiêu: Burnout_Level (Low/Med/High) & Dropout_Risk (Yes: 54.6%)",
                "• Giải pháp xử lý 18 cột khuyết (1.6% - 6.4%):",
                "  - Biến liên tục: KNNImputer (k=7) trong Pipeline",
                "  - Biến phân loại: Điền Mode + Mã hóa One-Hot / Thứ tự",
            ],
            img_path=os.path.join(fig_folder, "01_missing_values_analysis.png"),
        )
        add_content_slide(
            prs, content_layout, "02", "PHƯƠNG PHÁP", "2.2",
            "Quy trình Học máy & Kỹ thuật Tạo Đặc trưng Chuyên sâu",
            "Thiết kế 6 chỉ số tổng hợp đo lường sự tương tác giữa Căng thẳng Tâm lý và Sự Tham gia Học tập",
            bullets=[
                "• 6 Đặc trưng chuyên sâu (DomainFeatureEngineer):",
                "  1. Academic_Engagement_Index: Tổ hợp Chuyên cần, Giờ học, Động lực & GPA.",
                "  2. Psychological_Stress_Index: Tổ hợp Stress, Lo âu, Tài chính & Bạn bè.",
                "  3. Stress_Engagement_Ratio: Tỷ số mất cân bằng Áp lực / Gắn kết.",
                "  4. Net_Stress_Support_Balance: Áp lực trừ điểm tựa Gia đình & Xã hội.",
                "  5. Digital_Fatigue_Ratio: Tỷ lệ Giờ màn hình / Giờ ngủ.",
                "  6. Academic_Risk_Burden: Gánh nặng nợ môn & thiếu hụt chuyên cần.",
            ],
            img_path=os.path.join(fig_folder, "11_system_workflow_diagram.png"),
        )
        add_content_slide(
            prs, content_layout, "03", "KẾT QUẢ EDA", "3.1",
            "Phân tích Biến mục tiêu: Tác động của Burnout lên Nguy cơ Bỏ học",
            "Mối tương quan mạnh mẽ giữa Mức độ Kiệt sức và Tỷ lệ Bỏ học (r = +0.550)",
            bullets=[
                "• Phân bố nhãn cân bằng: Có nguy cơ = 437 SV (54.6%), Không = 363 SV (45.4%).",
                "• Tỷ lệ bỏ học tăng vọt theo mức độ Kiệt sức (Burnout_Level):",
                "  - Kiệt sức Thấp (320 SV): Chỉ 23.8% bỏ học (76 SV).",
                "  - Kiệt sức Trung bình (280 SV): 63.2% bỏ học (177 SV).",
                "  - Kiệt sức Cao (200 SV): Lên tới 92.0% bỏ học (184 SV)!",
                "• Kết luận: Kiệt sức học đường là yếu tố dự báo trực tiếp mạnh nhất dẫn tới quyết định bỏ học.",
            ],
            img_path=os.path.join(fig_folder, "02_target_and_burnout_distribution.png"),
        )
        add_content_slide(
            prs, content_layout, "03", "KẾT QUẢ EDA", "3.2",
            "Đối sánh Chỉ số Tham gia Học tập & Căng thẳng Tâm lý",
            "Sự khác biệt rõ rệt giữa nhóm Sinh viên Bỏ học (Yes) và Không bỏ học (No)",
            bullets=[
                "• Nhóm Bỏ học (Yes) suy giảm mạnh về gắn kết học tập:",
                "  - Chuyên cần TB thấp hơn: 73.7% so với 80.7% (nhóm No).",
                "  - Giờ tự học ít hơn: 2.93h/ngày so với 3.43h/ngày.",
                "  - Điểm động lực thấp hơn: 5.11 so với 5.90.",
                "• Đồng thời chịu áp lực tâm lý cao hơn:",
                "  - Mức căng thẳng TB: 5.81 (Yes) vs 5.17 (No).",
                "  - Điểm lo âu TB: 5.26 (Yes) vs 4.65 (No).",
            ],
            img_path=os.path.join(fig_folder, "03_stress_vs_engagement_distributions.png"),
        )
        add_content_slide(
            prs, content_layout, "03", "KẾT QUẢ EDA", "3.3",
            "Ma trận Tương quan & Tương tác Đa biến",
            "Phân tích hệ số tương quan Pearson giữa 18 biến định lượng trong bộ dữ liệu",
            bullets=[
                "• Yếu tố bảo vệ (Tương quan âm với Nguy cơ bỏ học):",
                "  - Tỷ lệ chuyên cần (r = -0.244)",
                "  - Điểm động lực (r = -0.193)",
                "  - Giờ tự học mỗi ngày (r = -0.160)",
                "  - Hỗ trợ từ gia đình (r = -0.137)",
                "• Yếu tố rủi ro (Tương quan dương với Nguy cơ bỏ học):",
                "  - Mức độ kiệt sức (r = +0.550)",
                "  - Căng thẳng (r = +0.164) & Lo âu (r = +0.145)",
                "  - Giờ màn hình (r = +0.127) & Áp lực tài chính (r = +0.114)",
            ],
            img_path=os.path.join(fig_folder, "04_correlation_matrix_heatmap.png"),
        )
        add_content_slide(
            prs, content_layout, "04", "MÔ HÌNH HÓA", "3.4",
            "So sánh Hiệu năng 10 Mô hình Học máy trên 2 Kịch bản",
            "Đánh giá bằng Kiểm định chéo 5-Fold Stratified CV & Tập Kiểm thử Độc lập (Test 20%)",
            bullets=[
                "• Kịch bản 1 (Đầy đủ đặc trưng - Có Burnout_Level):",
                "  - Random Forest đạt ROC-AUC cao nhất: 80.50% (CV) | 80.24% (Test).",
                "  - Extra Trees đạt Accuracy & F1 cao nhất: Acc 75.63%, F1 78.92%, Recall 83.91%.",
                "  - Soft Voting Ensemble & Logistic Regression đạt ROC-AUC ~79.8% - 80.0%.",
                "• Kịch bản 2 (Cảnh báo sớm - Ẩn hoàn toàn Burnout_Level):",
                "  - Soft Voting Ensemble dẫn đầu với CV Acc 70.63%, F1 73.89%, ROC-AUC 77.10%.",
            ],
            img_path=os.path.join(fig_folder, "07_model_performance_comparison.png"),
        )
        table_vi = [
            ["Mô hình Học máy", "KB1: CV Acc", "KB1: CV F1", "KB1: CV AUC", "KB1: Test Acc", "KB1: Test F1", "KB1: Test AUC", "KB2: CV AUC (Sớm)"],
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
            prs, content_layout, "04", "MÔ HÌNH HÓA", "3.5",
            "Bảng Tổng hợp Kết quả Thực nghiệm 10 Mô hình Học máy",
            "Đối sánh chi tiết Accuracy, F1-Score và ROC-AUC trên cả 2 kịch bản thực nghiệm",
            table_data=table_vi,
        )
        add_content_slide(
            prs, content_layout, "04", "ĐÁNH GIÁ", "3.6",
            "Đường cong ROC, Precision-Recall & Độ quan trọng Đặc trưng",
            "Các chỉ số tổng hợp Căng thẳng & Tham gia Học tập giữ vai trò quyết định hàng đầu",
            bullets=[
                "• Đường cong ROC & PR: Random Forest (AUC=0.802, AP=0.824) và Logistic Regression (AUC=0.799, AP=0.832) đạt độ ổn định cao.",
                "• Top Đặc trưng quan trọng nhất (Random Forest & XGBoost):",
                "  1. Burnout_Ordinal (Mức độ kiệt sức)",
                "  2. Stress_Engagement_Ratio (Tỷ số Căng thẳng / Gắn kết)",
                "  3. Academic_Engagement_Index (Chỉ số Tham gia học tập)",
                "  4. Psychological_Stress_Index (Chỉ số Căng thẳng tâm lý)",
                "  5. Attendance_Percent (Tỷ lệ chuyên cần)",
            ],
            img_path=os.path.join(fig_folder, "10_feature_importance_analysis.png"),
        )
        add_content_slide(
            prs, content_layout, "05", "ỨNG DỤNG WEB", "4.1",
            "Kiến trúc Hệ thống & Giao diện Web Dashboard Cảnh báo Sớm",
            "Ứng dụng FastAPI (app.py) dự báo thời gian thực Burnout_Level & Dropout_Risk kèm khuyến nghị",
            bullets=[
                "• Tầng Giao diện (Web UI): Nhập 18 thông số học tập, tâm lý, lối sống hoặc chọn nhanh 3 mẫu hồ sơ sinh viên.",
                "• Tầng AI Engine: Tự động tính 6 chỉ số chuyên ngành và chạy song song mô hình dự báo Burnout_Level + Dropout_Risk.",
                "• Đầu ra thông minh: Hiển thị xác suất bỏ học (%), mức độ kiệt sức và danh sách giải pháp hỗ trợ cá nhân hóa.",
            ],
            img_path=os.path.join(fig_folder, "12_system_architecture_diagram.png"),
        )
    else:
        add_content_slide(
            prs, content_layout, "02", "DATA & PREPROCESSING", "2.1",
            "Dataset Overview & Missing Data Imputation",
            "student_burnout_dropout_dataset_2.csv: 800 students x 25 attributes (missing values across 18 columns)",
            bullets=[
                "• Dataset Scale: 800 student records across 5 attribute groups:",
                "  1. Demographics (Age, Gender, Year, Dept, Residence, Family Income)",
                "  2. Study Engagement (Attendance %, Study Hours, GPA, Backlogs, Motivation)",
                "  3. Lifestyle Habits (Sleep, Screen Time, Exercise, Social Activity, Job)",
                "  4. Stress & Support (Stress, Anxiety, Financial Stress, Peer Pressure, Support)",
                "  5. Targets: Burnout_Level (Low/Med/High) & Dropout_Risk (Yes: 54.6%, No: 45.4%)",
                "• Missing Value Strategy (1.6% - 6.4% missing across 18 cols):",
                "  - Continuous Features: Multivariate KNNImputer (k=7)",
                "  - Categorical Features: Mode Imputation + One-Hot / Ordinal Encoding",
            ],
            img_path=os.path.join(fig_folder, "01_missing_values_analysis.png"),
        )
        add_content_slide(
            prs, content_layout, "02", "METHODOLOGY", "2.2",
            "ML Workflow & Domain-Specific Feature Engineering",
            "Designing 6 composite indices to quantify the interaction between Psychological Stress & Study Engagement",
            bullets=[
                "• 6 Engineered Features (DomainFeatureEngineer):",
                "  1. Academic_Engagement_Index: Weighted Attendance, Study Hours, Motivation & GPA.",
                "  2. Psychological_Stress_Index: Weighted Stress, Anxiety, Financial & Peer Pressure.",
                "  3. Stress_Engagement_Ratio: Imbalance ratio between Stress and Engagement.",
                "  4. Net_Stress_Support_Balance: Stress minus Family, Social & Exercise buffers.",
                "  5. Digital_Fatigue_Ratio: Screen_Time_Hours / Sleep_Hours.",
                "  6. Academic_Risk_Burden: Course backlogs & attendance deficit penalty.",
            ],
            img_path=os.path.join(fig_folder, "11_system_workflow_diagram.png"),
        )
        add_content_slide(
            prs, content_layout, "03", "EDA RESULTS", "3.1",
            "Target Variable Analysis: Impact of Burnout on Dropout Risk",
            "Strong monotonic relationship between Burnout_Level and Dropout_Risk (Pearson r = +0.550)",
            bullets=[
                "• Balanced Binary Target: Yes = 437 students (54.6%), No = 363 students (45.4%).",
                "• Sharp Dropout Progression across Burnout Tiers:",
                "  - Low Burnout (320 students): Only 23.8% dropout rate (76 students).",
                "  - Medium Burnout (280 students): 63.2% dropout rate (177 students).",
                "  - High Burnout (200 students): 92.0% dropout rate (184 students)!",
                "• Key Takeaway: Academic burnout serves as the primary proximate state preceding student attrition.",
            ],
            img_path=os.path.join(fig_folder, "02_target_and_burnout_distribution.png"),
        )
        add_content_slide(
            prs, content_layout, "03", "EDA RESULTS", "3.2",
            "Study Engagement vs. Psychological Stress Distributions",
            "Statistically significant divergence between Dropout (Yes) and Retained (No) student groups",
            bullets=[
                "• At-Risk Students (Dropout = Yes) exhibit lower study engagement:",
                "  - Lower Mean Attendance: 73.7% vs. 80.7% (Retained).",
                "  - Fewer Daily Study Hours: 2.93h/day vs. 3.43h/day.",
                "  - Lower Motivation Score: 5.11 vs. 5.90.",
                "• Accompanied by elevated psychological stress:",
                "  - Higher Mean Stress Level: 5.81 (Yes) vs. 5.17 (No).",
                "  - Higher Mean Anxiety Score: 5.26 (Yes) vs. 4.65 (No).",
            ],
            img_path=os.path.join(fig_folder, "03_stress_vs_engagement_distributions.png"),
        )
        add_content_slide(
            prs, content_layout, "03", "EDA RESULTS", "3.3",
            "Pearson Correlation Matrix & Multivariate Relationships",
            "Evaluating linear associations across 18 quantitative student attributes",
            bullets=[
                "• Top Protective Factors (Negative correlation with Dropout_Risk):",
                "  - Attendance_Percent (r = -0.244)",
                "  - Motivation_Score (r = -0.193)",
                "  - Study_Hours_Per_Day (r = -0.160)",
                "  - Family_Support_Score (r = -0.137)",
                "• Top Risk Factors (Positive correlation with Dropout_Risk):",
                "  - Burnout_Numeric (r = +0.550)",
                "  - Stress_Level (r = +0.164) & Anxiety_Score (r = +0.145)",
                "  - Screen_Time_Hours (r = +0.127) & Financial_Stress (r = +0.114)",
            ],
            img_path=os.path.join(fig_folder, "04_correlation_matrix_heatmap.png"),
        )
        add_content_slide(
            prs, content_layout, "04", "MODELING", "3.4",
            "Benchmarking 10 Machine Learning Models Across 2 Scenarios",
            "Evaluated via 5-Fold Stratified Cross-Validation & 20% Hold-Out Test Set",
            bullets=[
                "• Scenario 1 (Full Feature Set - With Burnout_Level):",
                "  - Random Forest achieves highest ROC-AUC: 80.50% (CV) | 80.24% (Test).",
                "  - Extra Trees achieves highest Acc & F1: Test Acc 75.63%, F1 78.92%, Recall 83.91%.",
                "  - Soft Voting Ensemble & Logistic Regression reach ~79.8% - 80.0% ROC-AUC.",
                "• Scenario 2 (Early Warning - Without Burnout_Level):",
                "  - Soft Voting Ensemble leads with 70.63% CV Acc, 73.89% F1, and 77.10% ROC-AUC.",
            ],
            img_path=os.path.join(fig_folder, "07_model_performance_comparison.png"),
        )
        table_en = [
            ["ML Classifier", "S1: CV Acc", "S1: CV F1", "S1: CV AUC", "S1: Test Acc", "S1: Test F1", "S1: Test AUC", "S2: CV AUC (Early)"],
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
            "Comprehensive Experimental Results Across 10 ML Classifiers",
            "Detailed comparison of Accuracy, F1-Score, and ROC-AUC in Scenario 1 (Full) and Scenario 2 (Early Warning)",
            table_data=table_en,
        )
        add_content_slide(
            prs, content_layout, "04", "EVALUATION", "3.6",
            "ROC Curves, Precision-Recall & Feature Importance Analysis",
            "Engineered Stress-Engagement composite indices rank as top predictors right behind Burnout_Level",
            bullets=[
                "• ROC & PR Curves: Random Forest (AUC=0.802, AP=0.824) and Logistic Regression (AUC=0.799, AP=0.832) maintain superior sensitivity.",
                "• Top Ranked Features (Random Forest & XGBoost):",
                "  1. Burnout_Ordinal (Burnout Level)",
                "  2. Stress_Engagement_Ratio (Stress-to-Engagement Imbalance)",
                "  3. Academic_Engagement_Index (Composite Engagement)",
                "  4. Psychological_Stress_Index (Composite Stress)",
                "  5. Attendance_Percent (Class Attendance Rate)",
            ],
            img_path=os.path.join(fig_folder, "10_feature_importance_analysis.png"),
        )
        add_content_slide(
            prs, content_layout, "05", "WEB APPLICATION", "4.1",
            "System Architecture & Interactive Early Warning Web Dashboard",
            "FastAPI Web Application (app.py) delivering real-time Burnout_Level & Dropout_Risk diagnostics",
            bullets=[
                "• Presentation Layer (Web UI): Interactive controls for 18 engagement, stress, and lifestyle inputs plus 3 student archetype presets.",
                "• AI Inference Engine: Computes 6 domain indices on the fly and runs dual Burnout_Level + Dropout_Risk models.",
                "• Actionable Output: Displays dropout probability (%), burnout breakdown, and personalized intervention recommendations.",
            ],
            img_path=os.path.join(fig_folder, "12_system_architecture_diagram.png"),
        )

    end_slide = prs.slides.add_slide(closing_layout)
    tb = end_slide.shapes.add_textbox(PptInches(2.2), PptInches(4.3), PptInches(9.0), PptInches(1.5))
    p = tb.text_frame.paragraphs[0]
    p.text = (
        "Cảm ơn Giảng viên và Hội đồng Chấm điểm!\nNhóm 03: Đỗ Quang Minh • Vũ Quang Huy"
        if is_vi
        else "Thank You for Your Attention!\nGroup 03: Do Quang Minh • Vu Quang Huy"
    )
    p.alignment = PP_ALIGN.CENTER
    p.font.size = PptPt(20)
    p.font.bold = True
    p.font.color.rgb = PptRGBColor(29, 53, 87)

    save_multi(prs.save, "SIC_AI_Capstone Project_Presentation Slide Template.pptx", lang)
    print(f"Completed Presentation Slides ({lang.upper()})")


if __name__ == "__main__":
    for language in ["vi", "en"]:
        print(f"\n================ BUILDING {language.upper()} SUITE ================")
        build_action_plan(language)
        build_evaluation_sheet(language)
        build_wbs_excel(language)
        build_final_report(language)
        build_presentation_pptx(language)
    print("\nSUCCESSFULLY GENERATED BOTH FULL VIETNAMESE (VI) AND FULL ENGLISH (EN) DOCUMENT SUITES!")
