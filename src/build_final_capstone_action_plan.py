"""
Builds FINAL_CAPSTONE/01_Final_Action_Plan.docx
Strictly conforms to official Samsung Innovation Campus (SIC) template,
cleans student self-evaluation, standardizes all metrics and dates.
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = r"c:\Users\mrmin\Desktop\Sudent_burnout"
SRC_TEMPLATE = os.path.join(BASE_DIR, "templates_backup", "SIC_AI_Capstone Project_Action Plan.docx")
OUT_PATH = os.path.join(BASE_DIR, "FINAL_CAPSTONE", "01_Final_Action_Plan.docx")


def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'))


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def build_action_plan():
    doc = Document(SRC_TEMPLATE)

    # Table 0: General Information, Goal, Abstract, Method
    t0 = doc.tables[0]
    t0.cell(0, 1).text = "Khóa học Trí tuệ Nhân tạo (Samsung Innovation Campus - AI Course)"
    t0.cell(1, 1).text = "Nhóm 03 (Group 03 - Student Burnout AI)"
    t0.cell(2, 1).text = "Đỗ Quang Minh (Trưởng nhóm / Team Leader) / Vũ Quang Huy (Thành viên / Team Member)"
    t0.cell(3, 1).text = (
        "Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học "
        "dựa trên các chỉ số về căng thẳng và sự tham gia học tập\n"
        "(Building, Training, and Comparing Machine Learning Models for Student Dropout Risk Prediction "
        "Based on Stress and Academic Engagement Indicators)"
    )

    t0.cell(5, 0).text = (
        "1. Khảo sát và phân tích khám phá dữ liệu (Comprehensive EDA) trên bộ dữ liệu 800 sinh viên với 25 thuộc tính "
        "về nhân khẩu học, sự tham gia học tập (Attendance, Study Hours, GPA, Motivation, Backlogs), lối sống và căng thẳng tâm lý.\n"
        "2. Xây dựng quy trình tiền xử lý 6 bước chống rò rỉ dữ liệu (KNNImputer k=7 trên 18 cột khuyết) kết hợp kỹ thuật "
        "Domain Feature Engineering tạo 6 chỉ số chuyên ngành đo lường cán cân Căng thẳng Tâm lý vs. Gắn kết Học tập.\n"
        "3. Huấn luyện, kiểm định chéo phân tầng 5-Fold (Stratified 5-Fold CV) và đối sánh toàn diện các thuật toán phân loại "
        "trên 2 kịch bản (Kịch bản 1: Đầy đủ đặc trưng chẩn đoán; Kịch bản 2: Cảnh báo sớm trực tiếp từ hành vi thô).\n"
        "4. Mở rộng mẫu thống kê (2.000 hồ sơ sạch) phục vụ thử nghiệm mở rộng quy mô kỹ thuật và triển khai ứng dụng Web Dashboard "
        "(FastAPI) dự báo thời gian thực kèm khuyến nghị can thiệp sư phạm cá nhân hóa."
    )

    t0.cell(7, 0).text = (
        "Tình trạng kiệt sức học đường (Academic Burnout) và bỏ học đại học (Student Dropout) đang là thách thức nghiêm trọng tại các cơ sở giáo dục. "
        "Đồ án này nghiên cứu mối liên hệ định lượng giữa áp lực tâm lý - xã hội (căng thẳng học vụ, lo âu, áp lực tài chính, đồng trang lứa) "
        "và sự tham gia học tập (chuyên cần, giờ tự học, động lực, GPA) nhằm xây dựng hệ thống cảnh báo sớm chủ động. "
        "Nhóm thực hiện quy trình học máy chuẩn mực: khảo sát trên 800 hồ sơ sinh viên gốc (25 thuộc tính, 612 ô khuyết trên 18 cột), thiết kế 6 đặc trưng "
        "chuyên ngành, đối sánh 10 mô hình học máy trên dữ liệu gốc (Random Forest đạt CV ROC-AUC 80.50% ± 3.83%, Extra Trees đạt Test Recall 83.91%), "
        "sau đó mở rộng mẫu quy mô kỹ thuật lên 2.000 hồ sơ sạch để kiểm thử 4 mô hình cốt lõi (Random Forest đạt CV ROC-AUC 89.01% ở KB1, XGBoost đạt 88.26% ở KB2). "
        "Toàn bộ mô hình được đóng gói và tích hợp vào Web Dashboard tương tác (FastAPI) hỗ trợ đắc lực cho cố vấn học tập."
    )

    t0.cell(9, 0).text = (
        "• Bước 1 - Tiền xử lý dữ liệu (Data Preprocessing): Loại bỏ trùng lặp Student_ID, điền khuyết đa biến bằng KNNImputer (k=7) cho biến số liên tục "
        "và Mode Imputation cho biến phân loại trong Pipeline scikit-learn; lọc ngoại lai bằng IQR kết hợp chặn biên nghiệp vụ giáo dục.\n"
        "• Bước 2 - Kỹ thuật tạo đặc trưng (Domain Feature Engineering): Thiết kế 6 chỉ số tương tác chuyên ngành: "
        "(1) Academic_Engagement_Index, (2) Psychological_Stress_Index, (3) Stress_Engagement_Ratio, (4) Net_Stress_Support_Balance, "
        "(5) Digital_Fatigue_Ratio, và (6) Academic_Risk_Burden; bổ sung tích số Chuyên cần × Giờ học.\n"
        "• Bước 3 - Huấn luyện & Đối sánh mô hình (Model Benchmarking): Đánh giá khách quan qua kiểm định chéo phân tầng 5-Fold (Stratified 5-Fold CV) "
        "và tập kiểm thử độc lập (Hold-out Test 20%) trên cả 2 kịch bản; đo lường đa chỉ số: Accuracy, Precision, Recall, F1-Score, ROC-AUC và PR-AUC.\n"
        "• Bước 4 - Triển khai ứng dụng (Web Deployment): Đóng gói mô hình tối ưu (.pkl) và xây dựng ứng dụng Web Dashboard tương tác bằng FastAPI "
        "với giao diện nhập 18 tham số, 3 bộ hồ sơ mẫu (Presets) và module sinh khuyến nghị can thiệp tự động."
    )

    # Table 1: Data, Expected Outcome, Roles
    t1 = doc.tables[1]
    t1.cell(1, 0).text = (
        "• Nguồn dữ liệu: Bộ dữ liệu nghiên cứu 'student_burnout_dropout_dataset_2.csv' (Kaggle / SIC AI Capstone).\n"
        "• Quy mô & Đặc điểm dữ liệu gốc: 800 bản ghi sinh viên × 25 thuộc tính (7 biến Nhân khẩu học, 5 biến Tham gia học tập, 5 biến Lối sống, 6 biến Căng thẳng & Hỗ trợ, 2 biến Mục tiêu). "
        "Có 612 ô khuyết trên 18 cột (tỷ lệ 1.62%–6.38%; 440/800 dòng chứa ít nhất một ô khuyết [55.0%]).\n"
        "• Quy mô mở rộng kỹ thuật: 2.000 hồ sơ sạch (0 NaN) gồm 800 bản ghi gốc đã làm sạch kết hợp 1.200 bản ghi tổng hợp phân tầng (Multivariate Gaussian) "
        "phục vụ huấn luyện mô hình mở rộng quy mô và kiểm thử giao diện Web.\n"
        "• Giao thức phân chia: Tỷ lệ phân tầng 80% Huấn luyện (Train) / 20% Kiểm thử độc lập (Test Set) bảo toàn tỷ lệ nhãn (54.6% Có nguy cơ / 45.4% An toàn)."
    )

    t1.cell(3, 0).text = (
        "• Kết quả đạt được:\n"
        "  1. Bộ 12 biểu đồ phân tích khám phá dữ liệu (EDA) và kiểm định mô hình trực quan, làm sáng tỏ mối liên hệ nhân quả thống kê: "
        "nhóm Kiệt sức Thấp (Low Burnout) chỉ có 23.8% bỏ học, trong khi nhóm Kiệt sức Cao (High Burnout) có tỷ lệ bỏ học lên tới 92.0% (tăng gấp gần 3.9 lần).\n"
        "  2. Bảng đối sánh thực nghiệm khoa học, minh bạch: Trên dữ liệu gốc 800 SV, Random Forest đạt CV ROC-AUC = 80.50% ± 3.83% (Extra Trees đạt Test Recall = 83.91%); "
        "trên tập mở rộng 2.000 hồ sơ, Random Forest đạt CV ROC-AUC = 89.01% ở Kịch bản 1 và XGBoost đạt CV ROC-AUC = 88.26% ở Kịch bản 2 (Cảnh báo sớm).\n"
        "  3. Ứng dụng Web Dashboard hoàn chỉnh (FastAPI) phản hồi tức thời (~145–153ms), hiển thị trực quan xác suất rủi ro và đề xuất giải pháp can thiệp.\n"
        "• Lợi ích thực tiễn: Giúp nhà trường và cố vấn học tập chủ động phát hiện sinh viên suy giảm gắn kết để kịp thời tư vấn tâm lý, giãn tiến độ học tập hoặc cấp học bổng hỗ trợ."
    )

    t1.cell(5, 0).text = (
        "1. Đỗ Quang Minh (Trưởng nhóm / Team Leader — 50% khối lượng):\n"
        "   - Thiết kế kiến trúc Pipeline học máy chống rò rỉ dữ liệu và xây dựng module DomainFeatureEngineer (6 chỉ số tương tác).\n"
        "   - Huấn luyện, kiểm định chéo và tối ưu nhóm mô hình tập hợp cây & tăng cường (Random Forest, Extra Trees, Gradient Boosting, XGBoost, Soft Voting Ensemble).\n"
        "   - Phát triển Backend FastAPI (app.py), đóng gói mô hình (.pkl) và tích hợp hệ thống Web Dashboard Demo.\n"
        "2. Vũ Quang Huy (Thành viên / Team Member — 50% khối lượng):\n"
        "   - Khảo sát chất lượng dữ liệu, thống kê 612 ô khuyết trên 18 cột, thiết lập KNNImputer (k=7) và thực hiện Phân tích EDA.\n"
        "   - Huấn luyện và đánh giá nhóm mô hình phân loại cơ sở & mạng nơ-ron (Logistic Regression, SVM RBF, KNN, Decision Tree, MLP).\n"
        "   - Thiết kế giao diện Frontend Web Dashboard, tổng hợp số liệu thực nghiệm, biên soạn Báo cáo cuối kỳ, WBS và Slide thuyết trình."
    )

    # Table 2: Schedule Summary, Comment & Assessment
    t2 = doc.tables[2]
    t2.cell(1, 0).text = (
        "• Giai đoạn 1 (19/10 – 20/10/2026): Khởi tạo dự án, phân tích yêu cầu bài toán, thống kê mô tả 800 bản ghi gốc, phân tích 18 cột dữ liệu khuyết.\n"
        "• Giai đoạn 2 (21/10 – 25/10/2026): Tiền xử lý dữ liệu (KNNImputer k=7), thiết kế 6 đặc trưng chuyên ngành, xuất bộ 6 biểu đồ phân tích EDA.\n"
        "• Giai đoạn 3 (26/10 – 29/10/2026): Xây dựng Pipeline học máy, huấn luyện 10 mô hình qua 5-Fold Stratified CV trên 2 kịch bản, phân tích ROC/PR và Feature Importance.\n"
        "• Giai đoạn 4 (30/10 – 05/11/2026): Mở rộng mẫu quy mô kỹ thuật (2.000 hồ sơ), huấn luyện 4 mô hình cốt lõi, phát triển Web Dashboard FastAPI, hoàn thiện Báo cáo, WBS và Slide."
    )

    # Rule C29: Leave Instructor's Comment & Assessment completely blank for instructor!
    t2.cell(3, 0).text = ""

    # Polish font styling
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        run.font.name = "Arial"
                        if run.font.size is None or run.font.size > Pt(11):
                            run.font.size = Pt(10)

    doc.save(OUT_PATH)
    print(f"Generated successfully: {OUT_PATH}")


if __name__ == "__main__":
    build_action_plan()

