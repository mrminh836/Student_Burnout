"""
MASTER PROMPT – PROFESSIONAL POWERPOINT DESIGNER (16:9 WIDESCREEN)
Generates modern, grid-aligned, visually rich PowerPoint presentations (.pptx)
in both 100% Vietnamese (VI) and 100% English (EN) with zero text-overflow or overlap.
"""

import os
from PIL import Image as PilImage
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_FOLDER = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"

FIG_VI = os.path.join(BASE_DIR, "figures_vi")
FIG_EN = os.path.join(BASE_DIR, "figures_en")

OUT_VI_WS = os.path.join(BASE_DIR, "SIC_Documents_VI")
OUT_EN_WS = os.path.join(BASE_DIR, "SIC_Documents_EN")
OUT_VI_DRV = os.path.join(SRC_FOLDER, "Ban_Tieng_Viet_Full")
OUT_EN_DRV = os.path.join(SRC_FOLDER, "Ban_Tieng_Anh_Full")

# Design System Colors
C_NAVY = RGBColor(29, 53, 87)       # #1D3557 Primary Brand
C_DARK_BG = RGBColor(15, 29, 50)    # #0F1D32 Cover Background
C_SLATE = RGBColor(38, 70, 83)      # #264653 Secondary
C_BLUE = RGBColor(69, 123, 157)     # #457B9D Accent Blue
C_TEAL = RGBColor(42, 157, 143)     # #2A9D8F Success / Protective
C_AMBER = RGBColor(244, 162, 97)    # #F4A261 Warning / Medium
C_CORAL = RGBColor(230, 57, 70)     # #E63946 Alert / High Risk
C_BG_LIGHT = RGBColor(244, 247, 251)# #F4F7FB Slide Canvas
C_WHITE = RGBColor(255, 255, 255)
C_TEXT_DARK = RGBColor(30, 41, 59)
C_BORDER = RGBColor(203, 213, 225)
C_SOFT_TEAL = RGBColor(230, 244, 241)


def add_rect(slide, left, top, width, height, fill_rgb, border_rgb=None, border_pt=1.0, rounded=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_rgb
    if border_rgb:
        shp.line.color.rgb = border_rgb
        shp.line.width = Pt(border_pt)
    else:
        shp.line.fill.background()
    if rounded and hasattr(shp, "adjustments") and len(shp.adjustments) > 0:
        shp.adjustments[0] = 0.04
    return shp


def add_picture_fitted(slide, img_path, box_left, box_top, box_width, box_height, pad=0.10):
    """Draws a white rounded container card and centers the image inside it without overflowing."""
    add_rect(slide, box_left, box_top, box_width, box_height, C_WHITE, border_rgb=C_BORDER, rounded=True)
    if not os.path.exists(img_path):
        return
    with PilImage.open(img_path) as im:
        w_px, h_px = im.size
    aspect = w_px / float(h_px)

    avail_w = box_width - 2 * pad
    avail_h = box_height - 2 * pad

    if avail_w / avail_h > aspect:
        # Height-constrained
        final_h = avail_h
        final_w = avail_h * aspect
    else:
        # Width-constrained
        final_w = avail_w
        final_h = avail_w / aspect

    x = box_left + (box_width - final_w) / 2.0
    y = box_top + (box_height - final_h) / 2.0
    slide.shapes.add_picture(img_path, Inches(x), Inches(y), width=Inches(final_w), height=Inches(final_h))


def add_text(slide, left, top, width, height, text, font_size=13, bold=False, italic=False, color=C_TEXT_DARK, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = "Segoe UI"
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def add_card_with_header(slide, left, top, width, height, header_text, body_lines, header_bg=C_NAVY, card_bg=C_WHITE, body_font=11.5):
    add_rect(slide, left, top, width, height, card_bg, border_rgb=C_BORDER, border_pt=1.0, rounded=True)
    add_rect(slide, left, top, width, 0.42, header_bg, rounded=False)
    add_text(slide, left + 0.12, top + 0.06, width - 0.24, 0.32, header_text, font_size=12, bold=True, color=C_WHITE)

    tb = slide.shapes.add_textbox(Inches(left + 0.12), Inches(top + 0.48), Inches(width - 0.24), Inches(height - 0.54))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(body_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(4)
        r = p.add_run()
        r.text = line
        r.font.name = "Segoe UI"
        r.font.size = Pt(body_font)
        r.font.bold = line.startswith("★") or line.startswith("• [")
        r.font.color.rgb = C_TEXT_DARK


def create_master_slide(prs, slide_num, total_slides, section_badge, title_text, takeaway_text, lang="vi"):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_rect(slide, 0, 0, 13.333, 7.5, C_BG_LIGHT)
    add_rect(slide, 0, 0, 13.333, 1.15, C_WHITE)
    add_rect(slide, 0, 1.13, 13.333, 0.04, C_NAVY)
    add_rect(slide, 0, 0, 0.14, 1.15, C_TEAL)

    # Wider Section Pill Badge (3.15 in) so text never wraps to 2 lines
    add_rect(slide, 0.38, 0.13, 3.15, 0.27, C_TEAL, rounded=True)
    add_text(slide, 0.42, 0.13, 3.07, 0.26, section_badge, font_size=9.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    add_text(slide, 0.36, 0.42, 12.6, 0.40, title_text, font_size=20, bold=True, color=C_NAVY)
    add_text(slide, 0.36, 0.80, 12.6, 0.30, takeaway_text, font_size=11.5, italic=True, color=C_SLATE)

    add_rect(slide, 0, 7.12, 13.333, 0.38, C_NAVY)
    footer_left = (
        "CHƯƠNG TRÌNH SAMSUNG INNOVATION CAMPUS (SIC) — ĐỒ ÁN AI CAPSTONE  |  NHÓM 03: ĐỖ QUANG MINH • VŨ QUANG HUY"
        if lang == "vi"
        else "SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE PROJECT  |  GROUP 03: DO QUANG MINH • VU QUANG HUY"
    )
    add_text(slide, 0.38, 7.17, 10.2, 0.28, footer_left, font_size=9.5, bold=True, color=C_WHITE)
    add_text(slide, 11.2, 7.17, 1.8, 0.28, f"Slide {slide_num:02d} / {total_slides:02d}", font_size=9.5, bold=True, color=C_AMBER, align=PP_ALIGN.RIGHT)

    return slide


def build_deck(lang="vi"):
    is_vi = lang == "vi"
    fig_dir = FIG_VI if is_vi else FIG_EN
    total_slides = 14

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ==========================================================================
    # SLIDE 01: COVER SLIDE
    # ==========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_rect(s1, 0, 0, 13.333, 7.5, C_DARK_BG)
    add_rect(s1, 0, 0, 0.28, 7.5, C_TEAL)
    add_rect(s1, 0.28, 0, 0.10, 7.5, C_AMBER)

    add_rect(s1, 0.85, 0.65, 5.4, 0.38, C_TEAL, rounded=True)
    badge_cov = "SAMSUNG INNOVATION CAMPUS — ĐỒ ÁN TỐT NGHIỆP AI" if is_vi else "SAMSUNG INNOVATION CAMPUS — AI CAPSTONE PROJECT"
    add_text(s1, 0.90, 0.68, 5.3, 0.32, badge_cov, font_size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    title_main = (
        "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY\nDỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG & THAM GIA HỌC TẬP"
        if is_vi
        else "BUILDING, TRAINING & COMPARING MACHINE LEARNING MODELS\nFOR STUDENT DROPOUT PREDICTION FROM STRESS & ENGAGEMENT"
    )
    add_text(s1, 0.85, 1.25, 11.8, 1.45, title_main, font_size=26, bold=True, color=C_WHITE)

    sub_main = (
        "Đề tài: Dự báo Kiệt sức Học đường & Nguy cơ Bỏ học — Hệ thống Cảnh báo Sớm Học đường Thông minh"
        if is_vi
        else "Topic: Student Burnout & Dropout Risk Prediction — Proactive Academic Early Warning System"
    )
    add_text(s1, 0.85, 2.75, 11.5, 0.45, sub_main, font_size=15, italic=True, color=C_AMBER)

    add_rect(s1, 0.85, 3.45, 11.6, 1.35, C_NAVY, border_rgb=C_TEAL, border_pt=1.5, rounded=True)
    team_header = "THỰC HIỆN BỞI NHÓM 03:" if is_vi else "PRESENTED BY GROUP 03:"
    team_body = (
        "• Đỗ Quang Minh (Trưởng nhóm) — Thiết kế Pipeline, Kỹ thuật Tạo Đặc trưng, Mô hình Cây Tập hợp & Backend FastAPI\n"
        "• Vũ Quang Huy (Thành viên) — Phân tích Khám phá Dữ liệu (EDA), Điền khuyết KNN, Mô hình Cơ sở & Giao diện Web UI"
        if is_vi
        else "• Do Quang Minh (Team Leader) — Leakage-Free Pipeline, Domain Feature Engineering, Tree Ensembles & FastAPI Backend\n"
        "• Vu Quang Huy (Team Member) — Statistical EDA, Multivariate KNN Imputation, Baseline/Neural Models & Web UI"
    )
    add_text(s1, 1.1, 3.55, 11.1, 0.35, team_header, font_size=13, bold=True, color=C_AMBER)
    add_text(s1, 1.1, 3.92, 11.1, 0.78, team_body, font_size=12.5, color=C_WHITE)

    cover_kpis = (
        [
            ("2,000", "Hồ sơ Sinh viên (Sạch 100%)", C_TEAL),
            ("25 → 34", "Thuộc tính & 6 Chỉ số Mới", C_BLUE),
            ("4 Mô hình", "LR, RF, XGBoost, Soft Voting", C_AMBER),
            ("89.01%", "Đỉnh ROC-AUC (5-Fold CV)", C_CORAL),
        ]
        if is_vi
        else [
            ("2,000", "Cleaned Student Profiles", C_TEAL),
            ("25 → 34", "Raw & Engineered Features", C_BLUE),
            ("4 Models", "LR, RF, XGBoost, Soft Voting", C_AMBER),
            ("89.01%", "Peak 5-Fold CV ROC-AUC", C_CORAL),
        ]
    )
    for idx, (val, lbl, col) in enumerate(cover_kpis):
        x = 0.85 + idx * 2.95
        add_rect(s1, x, 5.15, 2.75, 1.45, C_SLATE, border_rgb=col, border_pt=2.0, rounded=True)
        add_text(s1, x, 5.30, 2.75, 0.65, val, font_size=26, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s1, x + 0.1, 6.00, 2.55, 0.45, lbl, font_size=11.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 02: AGENDA
    # ==========================================================================
    s2 = create_master_slide(
        prs, 2, total_slides,
        "00 NỘI DUNG TRÌNH BÀY" if is_vi else "00 PRESENTATION AGENDA",
        "Cấu trúc Bài Thuyết trình Bảo vệ Đồ án" if is_vi else "Presentation Agenda & Executive Structure",
        "Hành trình khép kín từ Đặt vấn đề, Khám phá dữ liệu, Kỹ thuật Đặc trưng, Huấn luyện 10 Mô hình đến Ứng dụng Web Thực tế"
        if is_vi else "A structured walkthrough from problem formulation and statistical EDA to 10-model benchmarking and live Web UI deployment",
        lang=lang,
    )
    agenda_cards = (
        [
            ("01. GIỚI THIỆU & BỐI CẢNH", ["• Thực trạng bỏ học đại học & kiệt sức học đường.", "• Mục tiêu nghiên cứu & tính cấp thiết của Cảnh báo sớm.", "• Phân công nhiệm vụ Nhóm 03 & Lộ trình 4 giai đoạn."], C_NAVY),
            ("02. DỮ LIỆU & TIỀN XỬ LÝ", ["• Khảo sát 800 sinh viên × 25 thuộc tính (5 nhóm chức năng).", "• Phân tích 18 cột khuyết dữ liệu (1.6% – 6.4%).", "• Quy trình điền khuyết đa biến KNNImputer (k=7) chống rò rỉ."], C_BLUE),
            ("03. KỸ THUẬT TẠO ĐẶC TRƯNG", ["• Quy trình học máy 5 giai đoạn khép kín.", "• Thiết kế 6 chỉ số chuyên ngành trong DomainFeatureEngineer.", "• Đo lường cán cân Căng thẳng vs. Gắn kết học tập."], C_TEAL),
            ("04. PHÂN TÍCH KHÁM PHÁ (EDA)", ["• Tác động quyết định của mức độ kiệt sức lên bỏ học (92%).", "• Đối sánh phân phối 8 biến Tham gia học tập & Tâm lý.", "• Ma trận tương quan Pearson & Phân tích nhân khẩu học."], C_SLATE),
            ("05. ĐỐI SÁNH 10 MÔ HÌNH HỌC MÁY", ["• Thực nghiệm Kịch bản 1 (Đầy đủ) & Kịch bản 2 (Cảnh báo sớm).", "• Phân tích ROC-AUC (80.5%), F1-Score (78.9%) & Recall (83.9%).", "• Giải thích mô hình qua Ma trận nhầm lẫn & Độ quan trọng biến."], C_CORAL),
            ("06. HỆ THỐNG WEB UI & KẾT LUẬN", ["• Kiến trúc phần mềm 3 tầng & Ứng dụng Web FastAPI (app.py).", "• Dự báo kép Kiệt sức + Bỏ học kèm khuyến nghị cá nhân hóa.", "• Tổng kết đóng góp, lộ trình nâng cấp & Hỏi đáp."], C_AMBER),
        ]
        if is_vi
        else [
            ("01. INTRODUCTION & BACKGROUND", ["• University student attrition & academic burnout dynamics.", "• Research motivation & need for Proactive Early Warning.", "• Group 03 role distribution & 4-phase milestone schedule."], C_NAVY),
            ("02. DATASET & PREPROCESSING", ["• 800 student records × 25 attributes across 5 domains.", "• Auditing missingness across 18/25 columns (1.6% – 6.4%).", "• Leakage-free multivariate KNNImputer (k=7) pipeline."], C_BLUE),
            ("03. DOMAIN FEATURE ENGINEERING", ["• 5-stage end-to-end Machine Learning execution workflow.", "• Formulating 6 composite indices in DomainFeatureEngineer.", "• Quantifying the Stress vs. Academic Engagement balance."], C_TEAL),
            ("04. EXPLORATORY DATA ANALYSIS", ["• Monotonic progression from Burnout to Dropout (up to 92%).", "• Violin/Boxplot comparisons of 8 Stress & Engagement metrics.", "• Pearson correlation matrix & socioeconomic stratification."], C_SLATE),
            ("05. 10-MODEL BENCHMARKING", ["• Scenario 1 (Full Features) vs. Scenario 2 (Early Warning).", "• Evaluating ROC-AUC (80.5%), F1-Score (78.9%) & Recall (83.9%).", "• Confusion matrices, ROC/PR curves & Feature Importance."], C_CORAL),
            ("06. WEB DASHBOARD & CONCLUSION", ["• 3-Tier System Architecture & FastAPI Web UI (app.py).", "• Real-time dual predictions & prescriptive interventions.", "• Key accomplishments, future roadmap (LSTM/SHAP) & Q&A."], C_AMBER),
        ]
    )
    for idx, (hdr, lines, col) in enumerate(agenda_cards):
        r_i, c_i = divmod(idx, 3)
        x = 0.40 + c_i * 4.22
        y = 1.40 + r_i * 2.72
        add_card_with_header(s2, x, y, 4.05, 2.52, hdr, lines, header_bg=col, body_font=11.5)

    # ==========================================================================
    # SLIDE 03: INTRODUCTION & CONCEPTUAL MECHANISM
    # ==========================================================================
    s3 = create_master_slide(
        prs, 3, total_slides,
        "01 GIỚI THIỆU ĐỀ TÀI" if is_vi else "01 INTRODUCTION",
        "Đặt Vấn đề & Cơ chế Hình thành Nguy cơ Bỏ học Đại học" if is_vi else "Problem Statement & Conceptual Mechanism of Student Dropout",
        "Bỏ học không phải sự kiện ngẫu nhiên mà là hệ quả tích lũy từ sự suy giảm Gắn kết Học tập kết hợp với Căng thẳng Tâm lý kéo dài"
        if is_vi else "Student dropout is not an isolated event but the cumulative result of declining Academic Engagement and chronic Psychological Stress",
        lang=lang,
    )
    add_card_with_header(
        s3, 0.40, 1.40, 3.85, 2.60,
        "1. SỰ THAM GIA HỌC TẬP SUY GIẢM" if is_vi else "1. DECLINING STUDY ENGAGEMENT",
        [
            "• Tỷ lệ chuyên cần thấp (Attendance_Percent < 70%).",
            "• Giờ tự học mỗi ngày sụt giảm (Study_Hours_Per_Day).",
            "• Mất động lực học tập nội tại (Motivation_Score).",
            "• Điểm GPA kỳ trước thấp & tích lũy môn nợ (Backlogs).",
        ] if is_vi else [
            "• Deteriorating class attendance (< 70%).",
            "• Reduced daily self-study hours (Study_Hours_Per_Day).",
            "• Eroding intrinsic academic motivation.",
            "• Lower prior GPA & accumulating course Backlogs.",
        ],
        header_bg=C_TEAL,
    )
    add_card_with_header(
        s3, 0.40, 4.20, 3.85, 2.65,
        "2. ÁP LỰC CĂNG THẲNG TÂM LÝ - XÃ HỘI" if is_vi else "2. CHRONIC PSYCHOLOGICAL STRESS",
        [
            "• Điểm căng thẳng (Stress_Level) & lo âu (Anxiety) cao.",
            "• Áp lực tài chính gia đình (Financial_Stress_Score).",
            "• Áp lực đồng trang lứa (Peer_Pressure_Score).",
            "• Thiếu ngủ, dùng màn hình nhiều & thiếu hỗ trợ gia đình.",
        ] if is_vi else [
            "• Elevated academic Stress_Level & Anxiety_Score.",
            "• Severe financial burden (Financial_Stress_Score).",
            "• Peer comparison pressure (Peer_Pressure_Score).",
            "• Sleep deprivation, screen fatigue & low family support.",
        ],
        header_bg=C_CORAL,
    )
    add_card_with_header(
        s3, 4.65, 2.25, 3.80, 3.60,
        "TRẠNG THÁI TRUNG GIAN: KIỆT SỨC" if is_vi else "MEDIATING STATE: ACADEMIC BURNOUT",
        [
            "★ Hội chứng Kiệt sức Học đường (Burnout_Level):",
            "• Khi Áp lực Tâm lý vượt quá Điểm tựa Hỗ trợ và Động lực Học tập.",
            "• Phân tầng 3 cấp độ rõ rệt trong dữ liệu:",
            "  - Mức Thấp (Low): 320 SV (40.0%)",
            "  - Mức Vừa (Medium): 280 SV (35.0%)",
            "  - Mức Cao (High): 200 SV (25.0%)",
            "• Đóng vai trò cầu nối trực tiếp dẫn tới quyết định bỏ học.",
        ] if is_vi else [
            "★ Academic Burnout Syndrome (Burnout_Level):",
            "• Triggered when chronic stress exceeds support buffers and motivation.",
            "• Stratified into 3 distinct severity tiers:",
            "  - Low Burnout: 320 students (40.0%)",
            "  - Medium Burnout: 280 students (35.0%)",
            "  - High Burnout: 200 students (25.0%)",
            "• Acts as the proximate mediator of student attrition.",
        ],
        header_bg=C_AMBER, card_bg=C_WHITE,
    )
    add_card_with_header(
        s3, 8.85, 1.40, 4.08, 5.45,
        "3. MỤC TIÊU DỰ BÁO & GIẢI PHÁP AI" if is_vi else "3. TARGET OUTCOME & AI SOLUTION",
        [
            "★ Biến mục tiêu: Dropout_Risk (Yes / No)",
            "• Yes (Có nguy cơ bỏ học): 437 SV (54.62%)",
            "• No (An toàn / Duy trì): 363 SV (45.38%)",
            "",
            "★ Giải pháp của Nhóm 03:",
            "• Xây dựng hệ thống Học máy dự báo trên cả 2 kịch bản:",
            "  1. Chẩn đoán chính xác khi biết mức độ kiệt sức (ROC-AUC 80.50%).",
            "  2. Cảnh báo sớm ngay từ chỉ số Căng thẳng & Học tập thô (ROC-AUC 77.10%).",
            "• Cung cấp Dashboard khuyến nghị can thiệp sớm cho cố vấn học tập.",
        ] if is_vi else [
            "★ Target Variable: Dropout_Risk (Yes / No)",
            "• Yes (At-Risk of Dropout): 437 students (54.62%)",
            "• No (Retained / Safe): 363 students (45.38%)",
            "",
            "★ Group 03 AI Solution:",
            "• Dual-scenario Machine Learning architecture:",
            "  1. Full diagnostic prediction with Burnout_Level (80.50% ROC-AUC).",
            "  2. Early warning directly from raw Stress & Engagement (77.10% ROC-AUC).",
            "• Interactive Web Dashboard for proactive advising.",
        ],
        header_bg=C_NAVY, card_bg=C_SOFT_TEAL,
    )

    # ==========================================================================
    # SLIDE 04: OBJECTIVES, TEAM ROLES & TIMELINE
    # ==========================================================================
    s4 = create_master_slide(
        prs, 4, total_slides,
        "01 MỤC TIÊU & PHÂN CÔNG" if is_vi else "01 OBJECTIVES & TEAM",
        "Mục tiêu Đề tài, Phân công Nhóm 03 & Lộ trình Thực hiện" if is_vi else "Project Objectives, Group 03 Role Assignment & Schedule",
        "Phân công khối lượng công việc cân bằng (50% - 50%) và hoàn thành 100% các mốc bàn giao theo chuẩn WBS"
        if is_vi else "Balanced 50/50 technical workload distribution with 100% milestone completion across 4 project phases",
        lang=lang,
    )
    add_card_with_header(
        s4, 0.40, 1.40, 6.10, 3.30,
        "4 MỤC TIÊU CỐT LÕI CỦA ĐỒ ÁN" if is_vi else "4 CORE TECHNICAL OBJECTIVES",
        [
            "• [Mục tiêu 1 - EDA]: Khám phá quy luật thống kê trên 800 sinh viên × 25 biến; lượng hóa tác động của Kiệt sức, Chuyên cần và Căng thẳng.",
            "• [Mục tiêu 2 - Tạo Đặc trưng]: Xử lý 18 cột khuyết bằng KNNImputer (k=7) và thiết kế 6 chỉ số tương tác Căng thẳng – Gắn kết.",
            "• [Mục tiêu 3 - So sánh Mô hình]: Huấn luyện & kiểm định chéo 5-Fold trên 10 mô hình Học máy ở 2 kịch bản đối chứng.",
            "• [Mục tiêu 4 - Ứng dụng Web]: Triển khai ứng dụng FastAPI dự báo thời gian thực cả Burnout_Level và Dropout_Risk.",
        ] if is_vi else [
            "• [Objective 1 - Statistical EDA]: Profile 800 students × 25 attributes to quantify Burnout, Engagement, and Stress effects.",
            "• [Objective 2 - Feature Engineering]: Impute 18 incomplete columns via KNNImputer (k=7) and build 6 composite indices.",
            "• [Objective 3 - 10-Model Benchmarking]: Train & cross-validate 10 ML classifiers across 2 experimental scenarios.",
            "• [Objective 4 - Web Dashboard UI]: Deploy a real-time FastAPI Web App predicting both Burnout_Level and Dropout_Risk.",
        ],
        header_bg=C_NAVY,
    )
    add_card_with_header(
        s4, 6.80, 1.40, 6.10, 1.58,
        "ĐỖ QUANG MINH — TRƯỞNG NHÓM (50% KHỐI LƯỢNG)" if is_vi else "DO QUANG MINH — TEAM LEADER (50% SHARE)",
        [
            "• Thiết kế kiến trúc Pipeline chống rò rỉ dữ liệu & lớp DomainFeatureEngineer (6 chỉ số).",
            "• Huấn luyện nhóm Cây tập hợp & Boosting (Random Forest, Extra Trees, GB, XGBoost, Voting).",
            "• Phát triển Backend FastAPI (app.py) và đóng gói mô hình dự báo (.pkl).",
        ] if is_vi else [
            "• Architected leakage-free ML Pipeline & DomainFeatureEngineer (6 composite indices).",
            "• Trained Tree Ensembles & Boosting (Random Forest, Extra Trees, GB, XGBoost, Voting).",
            "• Built FastAPI Backend service (app.py) and serialized production models (.pkl).",
        ],
        header_bg=C_TEAL,
    )
    add_card_with_header(
        s4, 6.80, 3.12, 6.10, 1.58,
        "VŨ QUANG HUY — THÀNH VIÊN (50% KHỐI LƯỢNG)" if is_vi else "VU QUANG HUY — TEAM MEMBER (50% SHARE)",
        [
            "• Kiểm toán dữ liệu khuyết (18/25 cột), phân tích thống kê & xuất 12 biểu đồ chuẩn 200 DPI.",
            "• Huấn luyện nhóm mô hình Cơ sở & Mạng nơ-ron (Logistic, SVM RBF, KNN, Decision Tree, MLP).",
            "• Thiết kế giao diện Frontend Web Dashboard, biên soạn Báo cáo, WBS & Slide.",
        ] if is_vi else [
            "• Audited missing data (18/25 cols), performed statistical EDA & created 12 200-DPI figures.",
            "• Trained Baseline, Kernel & Neural models (Logistic, SVM RBF, KNN, Decision Tree, MLP).",
            "• Designed Frontend Dashboard UI and authored Final Report, WBS & Presentation Deck.",
        ],
        header_bg=C_BLUE,
    )

    phases = (
        [
            ("GĐ 1 (19–20/10)", "Khởi tạo & Đặc tả Dữ liệu\n(Action Plan & WBS)", C_NAVY),
            ("GĐ 2 (21–25/10)", "Phân tích EDA & KNNImputer\n+ Tạo 6 Chỉ số Mới", C_BLUE),
            ("GĐ 3 (26–29/10)", "Huấn luyện & Kiểm định 5-Fold\n10 Mô hình × 2 Kịch bản", C_TEAL),
            ("GĐ 4 (30/10–05/11)", "Web Dashboard FastAPI\n& Hoàn thiện Báo cáo SIC", C_CORAL),
        ]
        if is_vi
        else [
            ("Phase 1 (Oct 19–20)", "Initiation & Schema Audit\n(Action Plan & WBS)", C_NAVY),
            ("Phase 2 (Oct 21–25)", "Statistical EDA, KNNImputer\n& 6 Engineered Indices", C_BLUE),
            ("Phase 3 (Oct 26–29)", "5-Fold CV Benchmarking\n10 Models × 2 Scenarios", C_TEAL),
            ("Phase 4 (Oct 30–Nov 05)", "FastAPI Web Dashboard\n& Final SIC Deliverables", C_CORAL),
        ]
    )
    for idx, (p_title, p_desc, col) in enumerate(phases):
        x = 0.40 + idx * 3.18
        add_rect(s4, x, 4.95, 2.98, 1.85, C_WHITE, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s4, x, 4.95, 2.98, 0.42, col)
        add_text(s4, x + 0.08, 5.00, 2.82, 0.32, p_title, font_size=11.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        done_lbl = "\n[✓ Hoàn thành 100%]" if is_vi else "\n[✓ 100% Completed]"
        add_text(s4, x + 0.10, 5.50, 2.78, 1.15, p_desc + done_lbl, font_size=11, bold=True, color=C_TEXT_DARK, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 05: DATASET SCHEMA & MISSING DATA ANALYSIS
    # ==========================================================================
    s5 = create_master_slide(
        prs, 5, total_slides,
        "02 DỮ LIỆU & TIỀN XỬ LÝ" if is_vi else "02 DATA & PREPROCESSING",
        "Kiến trúc Bộ Dữ liệu (800 SV × 25 Biến) & Xử lý Dữ liệu Khuyết" if is_vi else "Dataset Architecture (800 × 25) & Multivariate Missing Data Imputation",
        "18/25 thuộc tính có dữ liệu khuyết (1.6% – 6.4%) được xử lý bằng KNNImputer (k=7) trong Pipeline để bảo toàn cấu trúc tương quan"
        if is_vi else "18/25 attributes contain missing values (1.6% – 6.4%), imputed via multivariate KNNImputer (k=7) inside the CV Pipeline",
        lang=lang,
    )
    add_card_with_header(
        s5, 0.40, 1.40, 4.85, 5.45,
        "5 NHÓM THUỘC TÍNH & CHIẾN LƯỢC TIỀN XỬ LÝ" if is_vi else "5 ATTRIBUTE DOMAINS & PREPROCESSING STRATEGY",
        [
            "• [1. Nhân khẩu học - 7 biến]: Student_ID, Age (17–25), Gender, Year_of_Study (1–4), Department (6 khoa), Residence_Type, Family_Income_Bracket.",
            "• [2. Tham gia học tập - 5 biến]: Attendance_Percent (30–100%), Study_Hours_Per_Day (0–8.1h), Previous_GPA, Backlogs (0–6), Motivation_Score.",
            "• [3. Lối sống - 5 biến]: Sleep_Hours, Screen_Time_Hours, Exercise_Freq, Social_Activity, Part_Time_Job.",
            "• [4. Căng thẳng & Hỗ trợ - 6 biến]: Stress_Level, Anxiety_Score, Financial_Stress, Peer_Pressure, Family_Support, Counseling_Access.",
            "• [5. Mục tiêu - 2 biến]: Burnout_Level (Low/Med/High) & Dropout_Risk (Yes: 437 | No: 363).",
            "★ Tại sao dùng KNNImputer (k=7)?",
            "  Xóa dòng khuyết sẽ làm mất >40% dữ liệu. KNNImputer ước lượng từ 7 sinh viên tương đồng nhất.",
        ] if is_vi else [
            "• [1. Demographics - 7 cols]: Student_ID, Age (17–25), Gender, Year (1–4), Department (6 depts), Residence_Type, Family_Income_Bracket.",
            "• [2. Study Engagement - 5 cols]: Attendance_Percent (30–100%), Study_Hours (0–8.1h), Previous_GPA, Backlogs (0–6), Motivation_Score.",
            "• [3. Lifestyle - 5 cols]: Sleep_Hours, Screen_Time, Exercise_Freq, Social_Activity, Part_Time_Job.",
            "• [4. Stress & Support - 6 cols]: Stress_Level, Anxiety_Score, Financial_Stress, Peer_Pressure, Family_Support, Counseling_Access.",
            "• [5. Targets - 2 cols]: Burnout_Level (Low/Med/High) & Dropout_Risk (Yes: 437 | No: 363).",
            "★ Why Multivariate KNNImputer (k=7)?",
            "  Listwise deletion loses >40% of rows; KNN preserves multi-feature correlations.",
        ],
        header_bg=C_NAVY, body_font=11.0,
    )
    add_picture_fitted(s5, os.path.join(fig_dir, "01_missing_values_analysis.png"), 5.45, 1.40, 7.48, 5.45, pad=0.10)

    # ==========================================================================
    # SLIDE 06: WORKFLOW & DOMAIN FEATURE ENGINEERING
    # ==========================================================================
    s6 = create_master_slide(
        prs, 6, total_slides,
        "02 PHƯƠNG PHÁP LUẬN" if is_vi else "02 METHODOLOGY & FEATURES",
        "Quy trình Học máy 5 Bước & Kỹ thuật Tạo 6 Đặc trưng Chuyên sâu" if is_vi else "5-Stage ML Workflow & Domain-Specific Feature Engineering",
        "Xây dựng lớp biến đổi DomainFeatureEngineer lượng hóa trực tiếp sự mất cân bằng giữa Căng thẳng và Gắn kết Học tập"
        if is_vi else "Custom DomainFeatureEngineer transformer quantifying the structural interplay between Stress and Study Engagement",
        lang=lang,
    )
    add_picture_fitted(s6, os.path.join(fig_dir, "11_system_workflow_diagram.png"), 0.40, 1.35, 12.53, 2.75, pad=0.06)

    fe_cards = [
        ("1. Academic_Engagement_Index", "0.35*(Attendance/10) + 0.25*Study\n+ 0.25*Motivation + 0.15*GPA", C_TEAL),
        ("2. Psychological_Stress_Index", "0.35*Stress + 0.30*Anxiety\n+ 0.20*Financial + 0.15*Peer", C_CORAL),
        ("3. Stress_Engagement_Ratio", "Psychological_Stress_Index /\n(Academic_Engagement_Index + 1)", C_NAVY),
        ("4. Net_Stress_Support_Balance", "Stress_Index - (0.5*Family\n+ 0.3*Social + 0.2*Exercise)", C_BLUE),
        ("5. Digital_Fatigue_Ratio", "Screen_Time_Hours /\n(Sleep_Hours + 1.0)", C_AMBER),
        ("6. Academic_Risk_Burden", "1.5 * Backlogs +\n(100 - Attendance) / 20.0", C_SLATE),
    ]
    for idx, (f_name, f_eq, col) in enumerate(fe_cards):
        r_i, c_i = divmod(idx, 3)
        x = 0.40 + c_i * 4.22
        y = 4.25 + r_i * 1.35
        add_rect(s6, x, y, 4.05, 1.22, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s6, x, y, 4.05, 0.34, col)
        add_text(s6, x + 0.10, y + 0.03, 3.85, 0.28, f_name, font_size=10.5, bold=True, color=C_WHITE)
        add_text(s6, x + 0.10, y + 0.40, 3.85, 0.75, f_eq, font_size=10.5, bold=True, color=C_TEXT_DARK, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 07: EDA 1 - TARGET & BURNOUT PROGRESSION
    # ==========================================================================
    s7 = create_master_slide(
        prs, 7, total_slides,
        "03 PHÂN TÍCH KHÁM PHÁ (EDA)" if is_vi else "03 EDA INSIGHTS",
        "Phát hiện EDA 1: Tác động Quyết định của Burnout lên Nguy cơ Bỏ học" if is_vi else "EDA Insight 1: Burnout Level as the Primary Driver of Dropout Risk",
        "Tỷ lệ bỏ học tăng gấp gần 4 lần khi sinh viên chuyển từ trạng thái Kiệt sức Thấp (23.8%) sang Kiệt sức Cao (92.0%)"
        if is_vi else "Student dropout rate surges nearly 4-fold from Low Burnout (23.8%) to High Burnout (92.0%, Pearson r = +0.550)",
        lang=lang,
    )
    burn_kpis = (
        [
            ("KIỆT SỨC THẤP - LOW (n=320)", "23.8% Bỏ học", "244 SV An toàn vs. 76 SV Bỏ học\nNhóm duy trì cân bằng học tập tốt", C_TEAL),
            ("KIỆT SỨC VỪA - MEDIUM (n=280)", "63.2% Bỏ học", "103 SV An toàn vs. 177 SV Bỏ học\nNgưỡng chuyển giao cần can thiệp sớm", C_AMBER),
            ("KIỆT SỨC CAO - HIGH (n=200)", "92.0% Bỏ học", "Chỉ 16 SV An toàn vs. 184 SV Bỏ học!\nBáo động đỏ: Nguy cơ bỏ học gần như tuyệt đối", C_CORAL),
        ]
        if is_vi
        else [
            ("LOW BURNOUT (n=320)", "23.8% Dropout", "244 Retained vs. 76 Dropped Out\nHealthy study-life equilibrium", C_TEAL),
            ("MEDIUM BURNOUT (n=280)", "63.2% Dropout", "103 Retained vs. 177 Dropped Out\nCritical transition zone for early advising", C_AMBER),
            ("HIGH BURNOUT (n=200)", "92.0% Dropout", "Only 16 Retained vs. 184 Dropped Out!\nSevere attrition risk requiring urgent action", C_CORAL),
        ]
    )
    for idx, (b_hdr, b_big, b_sub, col) in enumerate(burn_kpis):
        x = 0.40 + idx * 4.22
        add_rect(s7, x, 1.35, 4.05, 1.55, C_WHITE, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s7, x, 1.35, 4.05, 0.34, col)
        add_text(s7, x + 0.1, 1.38, 3.85, 0.28, b_hdr, font_size=10.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s7, x + 0.1, 1.72, 3.85, 0.48, b_big, font_size=20, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s7, x + 0.1, 2.22, 3.85, 0.62, b_sub, font_size=10, color=C_TEXT_DARK, align=PP_ALIGN.CENTER)

    add_picture_fitted(s7, os.path.join(fig_dir, "02_target_and_burnout_distribution.png"), 0.40, 3.05, 12.53, 3.85, pad=0.06)

    # ==========================================================================
    # SLIDE 08: EDA 2 - STUDY ENGAGEMENT VS. PSYCHOLOGICAL STRESS
    # ==========================================================================
    s8 = create_master_slide(
        prs, 8, total_slides,
        "03 PHÂN TÍCH KHÁM PHÁ (EDA)" if is_vi else "03 EDA INSIGHTS",
        "Phát hiện EDA 2: Đối sánh Phân phối Tham gia Học tập & Căng thẳng Tâm lý" if is_vi else "EDA Insight 2: Study Engagement vs. Psychological Stress Distributions",
        "Nhóm bỏ học (Yes) có chuyên cần thấp hơn 7% (73.7% vs 80.7%), tự học ít hơn, động lực thấp hơn và căng thẳng/lo âu cao hơn rõ rệt"
        if is_vi else "Dropout students (Yes) show a 7% attendance deficit (73.7% vs 80.7%), fewer study hours, lower motivation, and elevated stress/anxiety",
        lang=lang,
    )
    add_picture_fitted(s8, os.path.join(fig_dir, "03_stress_vs_engagement_distributions.png"), 0.40, 1.38, 9.15, 5.50, pad=0.08)

    add_card_with_header(
        s8, 9.72, 1.38, 3.21, 5.50,
        "CHÊNH LỆCH TRUNG BÌNH (μ)" if is_vi else "MEAN COMPARISON (No vs Yes)",
        [
            "★ Nhóm Gắn kết Học tập:" if is_vi else "★ Academic Engagement:",
            "• Attendance_Percent:",
            "  No: 80.68%  |  Yes: 73.70%",
            "• Study_Hours_Per_Day:",
            "  No: 3.43h   |  Yes: 2.93h",
            "• Motivation_Score:",
            "  No: 5.90    |  Yes: 5.11",
            "• Previous_GPA:",
            "  No: 7.02    |  Yes: 6.72",
            "",
            "★ Nhóm Căng thẳng Tâm lý:" if is_vi else "★ Psychological Stress:",
            "• Stress_Level:",
            "  No: 5.17    |  Yes: 5.81",
            "• Anxiety_Score:",
            "  No: 4.65    |  Yes: 5.26",
            "• Financial_Stress:",
            "  No: 4.75    |  Yes: 5.26",
            "• Family_Support:",
            "  No: 6.73    |  Yes: 6.16",
        ],
        header_bg=C_NAVY, body_font=10.5,
    )

    # ==========================================================================
    # SLIDE 09: EDA 3 - CORRELATION MATRIX & PROTECTIVE VS RISK FACTORS
    # ==========================================================================
    s9 = create_master_slide(
        prs, 9, total_slides,
        "03 PHÂN TÍCH KHÁM PHÁ (EDA)" if is_vi else "03 EDA INSIGHTS",
        "Phát hiện EDA 3: Ma trận Tương quan Pearson & Các Yếu tố Nhân khẩu học" if is_vi else "EDA Insight 3: Pearson Correlation Matrix & Socioeconomic Factors",
        "Xác định rõ 5 yếu tố bảo vệ (tương quan âm) và 5 yếu tố rủi ro (tương quan dương) tác động trực tiếp lên cả Burnout và Dropout"
        if is_vi else "Identifying top 5 protective factors (negative correlation) and top 5 risk factors (positive correlation) driving student attrition",
        lang=lang,
    )
    add_picture_fitted(s9, os.path.join(fig_dir, "04_correlation_matrix_heatmap.png"), 0.40, 1.38, 7.35, 5.50, pad=0.08)

    add_card_with_header(
        s9, 7.95, 1.38, 4.98, 2.65,
        "YẾU TỐ BẢO VỆ (GIẢM NGUY CƠ BỎ HỌC)" if is_vi else "PROTECTIVE FACTORS (NEGATIVE CORRELATION)",
        [
            "• 1. Attendance_Percent: r = -0.244 (với Dropout) | -0.325 (với Burnout)",
            "• 2. Motivation_Score: r = -0.193 (với Dropout) | -0.249 (với Burnout)",
            "• 3. Study_Hours_Per_Day: r = -0.160 (với Dropout) | -0.257 (với Burnout)",
            "• 4. Family_Support_Score: r = -0.137 (với Dropout) | -0.205 (với Burnout)",
            "• 5. Previous_GPA & Exercise: r = -0.106 (Cải thiện sức bền học tập)",
        ] if is_vi else [
            "• 1. Attendance_Percent: r = -0.244 (Dropout) | -0.325 (Burnout)",
            "• 2. Motivation_Score: r = -0.193 (Dropout) | -0.249 (Burnout)",
            "• 3. Study_Hours_Per_Day: r = -0.160 (Dropout) | -0.257 (Burnout)",
            "• 4. Family_Support_Score: r = -0.137 (Dropout) | -0.205 (Burnout)",
            "• 5. Previous_GPA & Exercise: r = -0.106 (Academic resilience)",
        ],
        header_bg=C_TEAL, body_font=10.8,
    )
    add_card_with_header(
        s9, 7.95, 4.20, 4.98, 2.68,
        "YẾU TỐ NGUY CƠ (TĂNG NGUY CƠ BỎ HỌC)" if is_vi else "RISK FACTORS (POSITIVE CORRELATION)",
        [
            "• 1. Burnout_Numeric: r = +0.550 (Yếu tố dự báo trung gian mạnh nhất)",
            "• 2. Stress_Level: r = +0.164 (với Dropout) | +0.270 (với Burnout)",
            "• 3. Anxiety_Score: r = +0.145 (với Dropout) | +0.247 (với Burnout)",
            "• 4. Screen_Time_Hours: r = +0.127 (với Dropout) | +0.212 (với Burnout)",
            "• 5. Financial_Stress & Low Income: Nhóm thu nhập thấp bỏ học tới 62.1%",
        ] if is_vi else [
            "• 1. Burnout_Numeric: r = +0.550 (Strongest proximate predictor)",
            "• 2. Stress_Level: r = +0.164 (Dropout) | +0.270 (Burnout)",
            "• 3. Anxiety_Score: r = +0.145 (Dropout) | +0.247 (Burnout)",
            "• 4. Screen_Time_Hours: r = +0.127 (Dropout) | +0.212 (Burnout)",
            "• 5. Financial_Stress & Low Income: 62.1% dropout in Low-Income tier",
        ],
        header_bg=C_CORAL, body_font=10.8,
    )

    # ==========================================================================
    # SLIDE 10: MODEL BENCHMARKING VISUALIZATION (10 MODELS x 2 SCENARIOS)
    # ==========================================================================
    s10 = create_master_slide(
        prs, 10, total_slides,
        "04 ĐỐI SÁNH 10 MÔ HÌNH ML" if is_vi else "04 MODEL BENCHMARKING",
        "So sánh Hiệu năng 10 Mô hình Học máy trên 2 Kịch bản Thực nghiệm" if is_vi else "Benchmarking 10 Machine Learning Models Across 2 Experimental Scenarios",
        "Đánh giá khách quan bằng 5-Fold Stratified Cross-Validation: Random Forest & Extra Trees dẫn đầu Kịch bản 1; Voting dẫn đầu Kịch bản 2"
        if is_vi else "Rigorous 5-Fold Stratified CV evaluation: Random Forest & Extra Trees lead Scenario 1; Soft Voting leads Scenario 2",
        lang=lang,
    )
    add_picture_fitted(s10, os.path.join(fig_dir, "07_model_performance_comparison.png"), 0.40, 1.36, 12.53, 4.12, pad=0.06)

    add_card_with_header(
        s10, 0.40, 5.62, 6.15, 1.30,
        "KỊCH BẢN 1: ĐẦY ĐỦ ĐẶC TRƯNG (CÓ BURNOUT_LEVEL)" if is_vi else "SCENARIO 1: FULL FEATURES (WITH BURNOUT_LEVEL)",
        [
            "• Random Forest đạt đỉnh ROC-AUC: 80.50% (CV) | 80.24% (Test).",
            "• Extra Trees đạt đỉnh Accuracy & F1: Test Acc 75.63%, F1 78.92%, Recall 83.91%.",
        ] if is_vi else [
            "• Random Forest achieves peak ROC-AUC: 80.50% (CV) | 80.24% (Test).",
            "• Extra Trees achieves peak Acc & F1: Test Acc 75.63%, F1 78.92%, Recall 83.91%.",
        ],
        header_bg=C_NAVY, body_font=10.5,
    )
    add_card_with_header(
        s10, 6.78, 5.62, 6.15, 1.30,
        "KỊCH BẢN 2: CẢNH BÁO SỚM (ẨN BURNOUT_LEVEL)" if is_vi else "SCENARIO 2: EARLY WARNING (WITHOUT BURNOUT_LEVEL)",
        [
            "• Soft Voting Ensemble dẫn đầu CV: Acc 70.63%, F1 73.89%, ROC-AUC 77.10%.",
            "• Logistic Regression đạt Test ROC-AUC 79.00% chỉ từ chỉ số Stress & Engagement.",
        ] if is_vi else [
            "• Soft Voting Ensemble leads 5-Fold CV: Acc 70.63%, F1 73.89%, ROC-AUC 77.10%.",
            "• Logistic Regression reaches 79.00% Test ROC-AUC from Stress & Engagement alone.",
        ],
        header_bg=C_TEAL, body_font=10.5,
    )

    # ==========================================================================
    # SLIDE 11: QUANTITATIVE COMPARISON MATRIX (TABLE)
    # ==========================================================================
    s11 = create_master_slide(
        prs, 11, total_slides,
        "04 BẢNG SỐ LIỆU THỰC NGHIỆM" if is_vi else "04 QUANTITATIVE METRICS",
        "Bảng Tổng hợp Kết quả Thực nghiệm 10 Mô hình Học máy" if is_vi else "Comprehensive Experimental Evaluation Matrix Across 10 Classifiers",
        "Độ lệch giữa Kiểm định chéo 5-Fold (CV) và Tập kiểm thử độc lập (Test 20%) luôn dưới 1.5%, khẳng định mô hình không bị Overfitting"
        if is_vi else "Generalization gap between 5-Fold CV and 20% Hold-Out Test remains under 1.5%, confirming strong regularization",
        lang=lang,
    )
    table_data = [
        [
            "Hạng & Mô hình Học máy" if is_vi else "Rank & ML Classifier",
            "KB1: CV Acc" if is_vi else "S1: CV Acc",
            "KB1: CV Recall" if is_vi else "S1: CV Recall",
            "KB1: CV F1" if is_vi else "S1: CV F1",
            "KB1: CV AUC" if is_vi else "S1: CV AUC",
            "KB1: Test Acc" if is_vi else "S1: Test Acc",
            "KB1: Test F1" if is_vi else "S1: Test F1",
            "KB1: Test AUC" if is_vi else "S1: Test AUC",
            "KB2: CV AUC (Sớm)" if is_vi else "S2: Early CV AUC",
        ],
        ["1. Random Forest ★", "80.25%", "86.38%", "82.98%", "89.01%", "81.50%", "83.63%", "88.42%", "87.97%"],
        ["2. Soft Voting Ensemble ★", "79.70%", "84.74%", "82.27%", "88.89%", "80.25%", "82.17%", "88.32%", "88.25%"],
        ["3. XGBoost ★", "79.95%", "84.56%", "82.39%", "88.64%", "78.50%", "80.28%", "87.99%", "88.26%"],
        ["4. Logistic Regression", "78.45%", "81.64%", "80.62%", "87.61%", "78.25%", "80.18%", "87.60%", "87.25%"],
    ]
    t_shp = s11.shapes.add_table(len(table_data), len(table_data[0]), Inches(0.40), Inches(1.42), Inches(12.53), Inches(5.35))
    tbl = t_shp.table
    col_widths = [2.65, 1.18, 1.25, 1.18, 1.22, 1.25, 1.20, 1.25, 1.35]
    for c_i, w in enumerate(col_widths):
        tbl.columns[c_i].width = Inches(w)

    for r_i, row in enumerate(table_data):
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text = str(val)
            if r_i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_NAVY
            elif r_i in [1, 2, 3]:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_SOFT_TEAL
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_WHITE if r_i % 2 == 1 else C_BG_LIGHT

            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.CENTER
                p.font.name = "Segoe UI"
                p.font.size = Pt(10.5)
                p.font.bold = (r_i in [0, 1, 2, 3] or c_i == 0)
                p.font.color.rgb = C_WHITE if r_i == 0 else C_TEXT_DARK

    # ==========================================================================
    # SLIDE 12: ROC/PR CURVES & FEATURE IMPORTANCE
    # ==========================================================================
    s12 = create_master_slide(
        prs, 12, total_slides,
        "04 GIẢI THÍCH MÔ HÌNH" if is_vi else "04 MODEL EXPLAINABILITY",
        "Đường cong ROC / Precision-Recall & Độ quan trọng Đặc trưng" if is_vi else "ROC / Precision-Recall Curves & Feature Importance Analysis",
        "Các chỉ số tương tác Stress_Engagement_Ratio và Academic_Engagement_Index đứng đầu bảng xếp hạng ngay sau Burnout_Ordinal"
        if is_vi else "Engineered indices (Stress_Engagement_Ratio & Academic_Engagement_Index) rank immediately behind Burnout_Ordinal",
        lang=lang,
    )
    add_picture_fitted(s12, os.path.join(fig_dir, "08_roc_and_precision_recall_curves.png"), 0.40, 1.36, 6.15, 4.10, pad=0.06)
    add_picture_fitted(s12, os.path.join(fig_dir, "10_feature_importance_analysis.png"), 6.78, 1.36, 6.15, 4.10, pad=0.06)

    add_card_with_header(
        s12, 0.40, 5.60, 12.53, 1.35,
        "Ý NGHĨA HỌC THUẬT & THỰC TIỄN TỪ PHÂN TÍCH ĐỘ QUAN TRỌNG ĐẶC TRƯNG" if is_vi else "KEY INSIGHTS FROM ROC/PR CURVES & FEATURE IMPORTANCE",
        [
            "• Top 6 Đặc trưng Quyết định Nguy cơ Bỏ học: (1) Burnout_Ordinal, (2) Stress_Engagement_Ratio, (3) Academic_Engagement_Index, (4) Psychological_Stress_Index, (5) Attendance_Percent, (6) Net_Stress_Support_Balance.",
            "• Độ nhạy cao (Recall = 80.3% – 83.9%): Đảm bảo hệ thống phát hiện được hơn 4/5 số sinh viên thực sự có nguy cơ bỏ học để can thiệp sớm.",
        ] if is_vi else [
            "• Top 6 Predictors of Student Dropout: (1) Burnout_Ordinal, (2) Stress_Engagement_Ratio, (3) Academic_Engagement_Index, (4) Psychological_Stress_Index, (5) Attendance_Percent, (6) Net_Stress_Support_Balance.",
            "• High Recall (80.3% – 83.9%): Captures over 4 out of 5 truly at-risk students, minimizing costly false-negative omissions in university advising.",
        ],
        header_bg=C_SLATE, body_font=10.8,
    )

    # ==========================================================================
    # SLIDE 13: SYSTEM ARCHITECTURE & FASTAPI WEB DASHBOARD UI
    # ==========================================================================
    s13 = create_master_slide(
        prs, 13, total_slides,
        "05 ỨNG DỤNG WEB THỰC TẾ" if is_vi else "05 WEB APPLICATION",
        "Kiến trúc Hệ thống 3 Tầng & Giao diện Web Dashboard Cảnh báo Sớm" if is_vi else "3-Tier System Architecture & Interactive FastAPI Web Dashboard",
        "Ứng dụng Web (app.py) cho phép nhập 18 chỉ số hoặc chọn 3 mẫu Hồ sơ Sinh viên, trả về dự báo kép và khuyến nghị trong < 30ms"
        if is_vi else "FastAPI Web App (app.py) with 18 interactive sliders and 3 student presets delivering dual AI diagnostics in < 30ms",
        lang=lang,
    )
    add_picture_fitted(s13, os.path.join(fig_dir, "12_system_architecture_diagram.png"), 0.40, 1.38, 7.85, 5.50, pad=0.08)

    add_card_with_header(
        s13, 8.45, 1.38, 4.48, 5.50,
        "TÍNH NĂNG GIAO DIỆN WEB (APP.PY)" if is_vi else "WEB DASHBOARD CAPABILITIES",
        [
            "★ 1. Ba Mẫu Hồ sơ Kiểm thử Nhanh:",
            "  • [An toàn] Mẫu SV Tích cực: Nguy cơ 8.0%",
            "  • [Cảnh báo] Mẫu SV Căng thẳng TB: ~58%",
            "  • [Báo động] Mẫu SV Kiệt sức Cao: 95.8%!",
            "",
            "★ 2. Dự báo Kép Thời gian thực:",
            "  • Dự báo tự động Burnout_Level (3 mức).",
            "  • Dự báo xác suất Dropout_Risk (%).",
            "",
            "★ 3. Giải thích AI & Khuyến nghị:",
            "  • Hiển thị Chỉ số Gắn kết vs. Căng thẳng.",
            "  • Đề xuất cá nhân hóa: Phụ đạo học tập, tham vấn tâm lý, hỗ trợ học bổng.",
        ] if is_vi else [
            "★ 1. Three One-Click Student Presets:",
            "  • [Safe Tier] Engaged Student: 8.0% Risk",
            "  • [Warning Tier] Moderate Stress: ~58% Risk",
            "  • [Critical Tier] High Burnout: 95.8% Risk!",
            "",
            "★ 2. Real-Time Dual AI Inference:",
            "  • 3-Class Burnout_Level probabilities.",
            "  • Calibrated Dropout_Risk probability (%).",
            "",
            "★ 3. Explainable Indices & Actions:",
            "  • Displays Engagement vs. Stress scores.",
            "  • Generates targeted tutoring, counseling & financial aid referrals.",
        ],
        header_bg=C_NAVY, body_font=11.0,
    )

    # ==========================================================================
    # SLIDE 14: CONCLUSION, FUTURE ROADMAP & Q&A
    # ==========================================================================
    s14 = create_master_slide(
        prs, 14, total_slides,
        "06 KẾT LUẬN & HỎI ĐÁP" if is_vi else "06 CONCLUSION & Q&A",
        "Kết luận Chung, Định hướng Phát triển & Thảo luận (Q&A)" if is_vi else "Conclusions, Future Research Roadmap & Q&A",
        "Hoàn thành xuất sắc toàn bộ mục tiêu đề tài — Sẵn sàng tích hợp vào Hệ thống Quản lý Học tập (LMS) của các trường đại học"
        if is_vi else "Successfully achieved all capstone objectives — Ready for integration with institutional Learning Management Systems (LMS)",
        lang=lang,
    )
    add_card_with_header(
        s14, 0.40, 1.40, 6.15, 3.80,
        "TỔNG KẾT ĐÓNG GÓP CỐT LÕI" if is_vi else "KEY PROJECT ACCOMPLISHMENTS",
        [
            "• [1. Khám phá Dữ liệu]: Lượng hóa bước nhảy vọt tỷ lệ bỏ học theo mức độ Kiệt sức (Thấp 23.8% → Vừa 63.2% → Cao 92.0%).",
            "• [2. Kỹ thuật Đặc trưng]: Thiết kế 6 chỉ số chuyên ngành giúp mô hình nắm bắt chính xác cán cân Căng thẳng – Gắn kết học tập.",
            "• [3. Hệ thống 10 Mô hình ML]: Random Forest & Extra Trees đạt ROC-AUC 80.50%, F1-Score 78.92% và Recall 83.91%.",
            "• [4. Sản phẩm Thực tiễn]: Ứng dụng Web FastAPI hoàn chỉnh hỗ trợ cố vấn học tập sàng lọc và can thiệp sớm.",
        ] if is_vi else [
            "• [1. Empirical Discovery]: Quantified the sharp dropout surge across Burnout tiers (Low 23.8% → Medium 63.2% → High 92.0%).",
            "• [2. Domain Engineering]: Formulated 6 composite indices capturing the balance between Psychological Stress and Engagement.",
            "• [3. 10-Model Benchmarking]: Achieved 80.50% CV ROC-AUC, 78.92% Test F1-Score, and 83.91% Recall with Tree Ensembles.",
            "• [4. Practical Software]: Delivered a real-time FastAPI Early Warning Dashboard for university advisors.",
        ],
        header_bg=C_TEAL, body_font=11.5,
    )
    add_card_with_header(
        s14, 6.78, 1.40, 6.15, 3.80,
        "ĐỊNH HƯỚNG PHÁT TRIỂN TƯƠNG LAI" if is_vi else "FUTURE RESEARCH & DEPLOYMENT ROADMAP",
        [
            "• [Hướng 1 - Dữ liệu Chuỗi thời gian]: Thu thập chuyên cần và căng thẳng theo từng tuần học để huấn luyện mô hình LSTM / Transformer.",
            "• [Hướng 2 - Giải thích Cục bộ SHAP]: Tích hợp biểu đồ SHAP Waterfall trực tiếp lên Web UI cho từng hồ sơ sinh viên.",
            "• [Hướng 3 - Kết nối LMS / Moodle API]: Tự động hóa quy trình cảnh báo sớm định kỳ hàng tháng cho Phòng Công tác Sinh viên.",
            "• [Hướng 4 - Mở rộng Đa trường]: Thu thập 5,000–10,000 mẫu tại các trường đại học Việt Nam.",
        ] if is_vi else [
            "• [Direction 1 - Longitudinal Time-Series]: Track weekly attendance and stress logs using sequential LSTM / Transformer models.",
            "• [Direction 2 - Local Explainability (SHAP)]: Embed per-student SHAP waterfall plots directly inside the Web Dashboard.",
            "• [Direction 3 - LMS API Integration]: Connect with Moodle / Canvas APIs for automated monthly cohort screening.",
            "• [Direction 4 - Multi-Campus Scaling]: Expand dataset to 5,000–10,000 students across Vietnamese universities.",
        ],
        header_bg=C_NAVY, body_font=11.5,
    )

    add_rect(s14, 0.40, 5.40, 12.53, 1.50, C_DARK_BG, border_rgb=C_AMBER, border_pt=2.0, rounded=True)
    qa_title = (
        "CẢM ƠN GIẢNG VIÊN VÀ HỘI ĐỒNG CHẤM ĐIỂM ĐÃ LẮNG NGHE! — THẢO LUẬN & HỎI ĐÁP"
        if is_vi
        else "THANK YOU FOR YOUR ATTENTION! — QUESTIONS & DISCUSSION (Q&A)"
    )
    qa_sub = (
        "Nhóm 03: Đỗ Quang Minh (Trưởng nhóm) • Vũ Quang Huy (Thành viên)  |  Web Demo: http://127.0.0.1:8000"
        if is_vi
        else "Group 03: Do Quang Minh (Team Leader) • Vu Quang Huy (Team Member)  |  Live Web Demo: http://127.0.0.1:8000"
    )
    add_text(s14, 0.60, 5.58, 12.13, 0.55, qa_title, font_size=20, bold=True, color=C_AMBER, align=PP_ALIGN.CENTER)
    add_text(s14, 0.60, 6.18, 12.13, 0.45, qa_sub, font_size=13.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    base_name = "SIC_AI_Capstone Project_Presentation Slide Template.pptx"
    suffixed = f"SIC_AI_Capstone Project_Presentation Slide Template_{lang.upper()}.pptx"
    targets = [
        os.path.join(OUT_VI_WS if is_vi else OUT_EN_WS, base_name),
        os.path.join(OUT_VI_DRV if is_vi else OUT_EN_DRV, base_name),
        os.path.join(SRC_FOLDER, suffixed),
    ]
    if is_vi:
        targets.append(os.path.join(SRC_FOLDER, base_name))

    for t in targets:
        prs.save(t)
    print(f"Saved Master 14-Slide Presentation Deck ({lang.upper()}) -> {len(targets)} locations")


if __name__ == "__main__":
    for language in ["vi", "en"]:
        build_deck(language)
    print("ALL MASTER POWERPOINT PRESENTATIONS UPDATED SUCCESSFULLY!")
