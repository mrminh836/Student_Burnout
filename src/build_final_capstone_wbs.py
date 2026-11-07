"""
Builds FINAL_CAPSTONE/05_Final_WBS.xlsx
Strictly conforms to official Samsung Innovation Campus (SIC) template,
corrects the 2021 date error on timeline to continuous 2026 dates,
standardizes task status to 'Done' matching legend, aligns all milestones.
"""
import os
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = r"c:\Users\mrmin\Desktop\Sudent_burnout"
SRC_TEMPLATE = os.path.join(BASE_DIR, "templates_backup", "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx")
OUT_PATH = os.path.join(BASE_DIR, "FINAL_CAPSTONE", "05_Final_WBS.xlsx")


def build_wbs():
    wb = openpyxl.load_workbook(SRC_TEMPLATE)
    ws = wb["WBS"]

    # Basic project info
    ws["E7"] = datetime.date(2026, 11, 6)
    ws["E8"] = "Nhóm 03: Đỗ Quang Minh (Trưởng nhóm), Vũ Quang Huy"
    ws["E9"] = (
        "Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học "
        "dựa trên các chỉ số về căng thẳng và sự tham gia học tập"
    )

    # Styling constants
    thin = Side(style="thin", color="CCCCCC")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    font_bold = Font(name="Arial", size=10, bold=True)
    font_norm = Font(name="Arial", size=9.5)
    font_sub = Font(name="Arial", size=9.0)

    fill_l1 = PatternFill("solid", fgColor="0B162C")      # Navy Header
    fill_l2 = PatternFill("solid", fgColor="F1F5F9")      # Light Slate Sub-row
    fill_done = PatternFill("solid", fgColor="D9F2E6")    # Soft Green for Done
    font_done = Font(name="Arial", size=9.5, bold=True, color="0D6832")
    font_white = Font(name="Arial", size=10, bold=True, color="FFFFFF")

    fill_bar_l1 = PatternFill("solid", fgColor="0066FF")  # Electric Blue for Phase Gantt
    fill_bar_l2 = PatternFill("solid", fgColor="00B4D8")  # Cyan for Subtask Gantt

    # Fix Timeline Dates (D1 to D20: 19/10/2026 to 07/11/2026 continuous)
    # Columns K (col 11) to AD (col 30)
    base_date = datetime.date(2026, 10, 19)
    for d_idx in range(20):
        c = ws.cell(row=13, column=11 + d_idx)
        c.value = base_date + datetime.timedelta(days=d_idx)
        c.number_format = "yyyy-mm-dd"
        c.font = Font(name="Arial", size=8.5, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)

    # WBS Task Definitions
    # (row, level, task_name, start_d_idx, end_d_idx, start_date, end_date, owner, deliverable, status)
    tasks = [
        # Phase 1
        (14, 1, "1. Khởi tạo Dự án & Phân tích Yêu cầu", 0, 1, datetime.date(2026, 10, 19), datetime.date(2026, 10, 20),
         "Đỗ Quang Minh, Vũ Quang Huy", "Kế hoạch hành động & WBS", "Done"),
        (15, 2, "1.1. Xác định mục tiêu bài toán dự báo Burnout & Dropout", 0, 0, datetime.date(2026, 10, 19), datetime.date(2026, 10, 19),
         "Đỗ Quang Minh", "Đặc tả bài toán & 4 mục tiêu", "Done"),
        (16, 2, "1.2. Khảo sát cấu trúc bộ dữ liệu (800 SV × 25 thuộc tính)", 0, 1, datetime.date(2026, 10, 19), datetime.date(2026, 10, 20),
         "Vũ Quang Huy", "Báo cáo đặc tả dữ liệu", "Done"),

        # Phase 2
        (17, 1, "2. Phân tích Khám phá Dữ liệu (EDA) & Tiền xử lý", 2, 6, datetime.date(2026, 10, 21), datetime.date(2026, 10, 25),
         "Vũ Quang Huy, Đỗ Quang Minh", "Bộ 6 biểu đồ EDA & Pipeline tiền xử lý", "Done"),
        (18, 2, "2.1. Thống kê 612 ô khuyết trên 18/25 cột & phân phối mục tiêu", 2, 2, datetime.date(2026, 10, 21), datetime.date(2026, 10, 21),
         "Vũ Quang Huy", "Biểu đồ khuyết & nhãn", "Done"),
        (19, 2, "2.2. Phân tích tương quan Pearson & phân bố Căng thẳng vs Gắn kết", 3, 4, datetime.date(2026, 10, 22), datetime.date(2026, 10, 23),
         "Vũ Quang Huy", "Ma trận tương quan & Violin plots", "Done"),
        (20, 2, "2.3. Xây dựng KNNImputer (k=7) & lọc ngoại lai IQR", 4, 6, datetime.date(2026, 10, 23), datetime.date(2026, 10, 25),
         "Vũ Quang Huy", "Module tiền xử lý dữ liệu sạch", "Done"),

        # Phase 3
        (21, 1, "3. Thiết kế Đặc trưng Chuyên ngành & Kiến trúc Pipeline", 4, 6, datetime.date(2026, 10, 23), datetime.date(2026, 10, 25),
         "Đỗ Quang Minh", "Lớp DomainFeatureEngineer (6 chỉ số)", "Done"),
        (22, 2, "3.1. Thiết kế 6 chỉ số tương tác (Gắn kết, Căng thẳng, Tỷ số...)", 5, 6, datetime.date(2026, 10, 24), datetime.date(2026, 10, 25),
         "Đỗ Quang Minh", "6 Đặc trưng chuyên ngành", "Done"),
        (23, 2, "3.2. Đóng gói Pipeline ColumnTransformer chống rò rỉ dữ liệu", 5, 6, datetime.date(2026, 10, 24), datetime.date(2026, 10, 25),
         "Đỗ Quang Minh", "Pipeline scikit-learn chuẩn hóa", "Done"),

        # Phase 4
        (24, 1, "4. Huấn luyện, Kiểm định chéo & Đối sánh Mô hình Học máy", 7, 10, datetime.date(2026, 10, 26), datetime.date(2026, 10, 29),
         "Đỗ Quang Minh, Vũ Quang Huy", "Bảng kết quả CSV & Biểu đồ kiểm định", "Done"),
        (25, 2, "4.1. Huấn luyện 10 mô hình trên 800 SV gốc (5-Fold Stratified CV)", 7, 8, datetime.date(2026, 10, 26), datetime.date(2026, 10, 27),
         "Vũ Quang Huy", "Kết quả 10 mô hình cơ sở & Ensemble", "Done"),
        (26, 2, "4.2. Huấn luyện 4 mô hình cốt lõi trên 2.000 hồ sơ mở rộng", 8, 9, datetime.date(2026, 10, 27), datetime.date(2026, 10, 28),
         "Đỗ Quang Minh", "Mô hình tốt nhất (RF 89.01%, XGB 88.26%)", "Done"),
        (27, 2, "4.3. Đánh giá đa chiều: ROC/PR, Confusion Matrix & Feature Importance", 9, 10, datetime.date(2026, 10, 28), datetime.date(2026, 10, 29),
         "Đỗ Quang Minh, Vũ Quang Huy", "Biểu đồ ROC, PR, CM, Feature Importance", "Done"),

        # Phase 5
        (28, 1, "5. Phát triển Web Dashboard Demo, Mở rộng Mẫu & Hoàn thiện Hồ sơ SIC", 11, 18, datetime.date(2026, 10, 30), datetime.date(2026, 11, 6),
         "Đỗ Quang Minh, Vũ Quang Huy", "Web UI Demo & Bộ 5 tài liệu hoàn chỉnh", "Done"),
        (29, 2, "5.1. Xây dựng giao diện Web Dashboard FastAPI (app.py)", 11, 13, datetime.date(2026, 10, 30), datetime.date(2026, 11, 1),
         "Đỗ Quang Minh", "Ứng dụng Web dự báo thời gian thực", "Done"),
        (30, 2, "5.2. Mở rộng mẫu kỹ thuật 2.000 hồ sơ sạch (Multivariate Gaussian)", 13, 15, datetime.date(2026, 11, 1), datetime.date(2026, 11, 3),
         "Đỗ Quang Minh, Vũ Quang Huy", "Bộ dữ liệu mở rộng 0 NaN", "Done"),
        (31, 2, "5.3. Kiểm toán tính nhất quán, biên soạn Báo cáo, Slide & Hoàn thiện", 15, 18, datetime.date(2026, 11, 3), datetime.date(2026, 11, 6),
         "Vũ Quang Huy, Đỗ Quang Minh", "Bộ 5 tài liệu SIC chuẩn nghiệm thu", "Done"),
    ]

    # Clear old gantt cells first
    for r in range(14, 40):
        for col_idx in range(11, 31):
            cell = ws.cell(row=r, column=col_idx)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)

    # Write each task row
    for r_idx, lvl, name, s_idx, e_idx, s_date, e_date, owner, deliv, st in tasks:
        # Task name
        c_name = ws.cell(row=r_idx, column=2 if lvl == 1 else 3)
        c_name.value = name
        c_name.font = font_bold if lvl == 1 else font_sub

        # Dates & Info
        ws.cell(row=r_idx, column=6, value=s_date).number_format = "yyyy-mm-dd"
        ws.cell(row=r_idx, column=7, value=e_date).number_format = "yyyy-mm-dd"
        ws.cell(row=r_idx, column=8, value=owner)
        ws.cell(row=r_idx, column=9, value=deliv)

        # Status
        c_st = ws.cell(row=r_idx, column=10, value=st)
        c_st.fill = fill_done
        c_st.font = font_done
        c_st.alignment = Alignment(horizontal="center", vertical="center")

        # Row styling
        for col_i in range(2, 11):
            cell = ws.cell(row=r_idx, column=col_i)
            cell.border = border_all
            if lvl == 1:
                cell.font = font_bold

        # Gantt bar
        fill_bar = fill_bar_l1 if lvl == 1 else fill_bar_l2
        for d in range(s_idx, e_idx + 1):
            g_cell = ws.cell(row=r_idx, column=11 + d)
            g_cell.fill = fill_bar
            g_cell.border = border_all

    # Auto-adjust column widths
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 13
    ws.column_dimensions["H"].width = 28
    ws.column_dimensions["I"].width = 32
    ws.column_dimensions["J"].width = 14
    for d in range(20):
        col_letter = get_column_letter(11 + d)
        ws.column_dimensions[col_letter].width = 4.2

    wb.save(OUT_PATH)
    print(f"Generated successfully: {OUT_PATH}")


if __name__ == "__main__":
    build_wbs()
