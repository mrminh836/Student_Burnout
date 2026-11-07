"""
Builds FINAL_CAPSTONE/04_Final_Presentation.pptx (9 Slides, 5-8 Minute Defense Deck)
Strictly conforms to master prompt requirements:
- 9 Slides: Project Overview, Problem Statement, Dataset & Understanding,
  Preprocessing & Feature Engineering, Exploratory Data Analysis, AI/ML Methodology,
  Model Performance & Evaluation, Demo / System Architecture, Conclusion & Future Work.
- 16:9 Widescreen, Navy/Cyan/White palette, Zero Emojis, Native Vector Shapes.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = r"c:\Users\mrmin\Desktop\Sudent_burnout"
FIG_DIR = os.path.join(BASE_DIR, "figures_vi")
OUT_PATH = os.path.join(BASE_DIR, "FINAL_CAPSTONE", "04_Final_Presentation.pptx")

# Color Palette
C_NAVY_DARK = RGBColor(11, 22, 44)       # #0B162C Primary Dark
C_NAVY_CARD = RGBColor(18, 32, 60)       # #12203C Surface Dark
C_BLUE_ELEC = RGBColor(0, 102, 255)      # #0066FF Primary Brand
C_CYAN = RGBColor(0, 180, 216)           # #00B4D8 Bright Cyan
C_TEAL = RGBColor(13, 148, 136)          # #0D9488 Emerald Teal
C_CORAL = RGBColor(225, 29, 72)          # #E11D48 Warning Coral
C_AMBER = RGBColor(217, 119, 6)          # #D97706 Warning Amber
C_AMBER_BRIGHT = RGBColor(245, 158, 11)  # #F59E0B Accent
C_SLATE_900 = RGBColor(15, 23, 42)       # #0F172A Text Dark
C_SLATE_700 = RGBColor(51, 65, 85)       # #334155 Text Slate
C_SLATE_500 = RGBColor(100, 116, 139)    # #64748B Secondary Text
C_SLATE_300 = RGBColor(203, 213, 225)    # #CBD5E1 Border Light
C_SLATE_100 = RGBColor(241, 245, 249)    # #F1F5F9 Card Background
C_BG_LIGHT = RGBColor(248, 250, 252)     # #F8FAFC Page Background
C_WHITE = RGBColor(255, 255, 255)        # #FFFFFF Pure White


def add_rect(slide, left, top, width, height, fill_rgb, border_rgb=None, border_pt=1.0, rounded=False):
    # Enforce sharp square/rectangle boxes (no rounded corners) per user request
    st = MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(st, Inches(left), Inches(top), Inches(width), Inches(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_rgb
    if border_rgb:
        shp.line.color.rgb = border_rgb
        shp.line.width = Pt(border_pt)
    else:
        shp.line.fill.background()
    return shp


def add_text(slide, left, top, width, height, text, font_size=12, bold=False, italic=False, color=C_SLATE_900,
             align=PP_ALIGN.LEFT, font_name="Segoe UI"):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = str(text)
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    p.alignment = align
    return tb


def apply_chrome(slide, cur_slide, total_slides, eyebrow, title, subtitle="", dark_mode=False):
    bg_col = C_NAVY_DARK if dark_mode else C_BG_LIGHT
    add_rect(slide, 0, 0, 13.333, 7.50, bg_col)

    if not dark_mode:
        add_rect(slide, 0, 0, 13.333, 0.08, C_BLUE_ELEC)

    badge_bg = C_NAVY_CARD if dark_mode else C_SLATE_100
    badge_txt = C_CYAN if dark_mode else C_BLUE_ELEC
    badge_w = 0.16 + len(eyebrow) * 0.088
    add_rect(slide, 0.50, 0.32, max(badge_w, 2.2), 0.32, badge_bg, border_rgb=badge_txt, border_pt=1.0, rounded=True)
    add_text(slide, 0.58, 0.36, max(badge_w - 0.16, 2.0), 0.24, eyebrow.upper(), font_size=9.5, bold=True, color=badge_txt)

    title_col = C_WHITE if dark_mode else C_SLATE_900
    add_text(slide, 0.50, 0.72, 12.33, 0.44, title, font_size=20.0, bold=True, color=title_col)

    if subtitle:
        sub_col = C_SLATE_300 if dark_mode else C_SLATE_500
        add_text(slide, 0.50, 1.18, 12.33, 0.28, subtitle, font_size=11.0, color=sub_col)

    # Footer
    foot_div = C_NAVY_CARD if dark_mode else C_SLATE_300
    foot_txt = C_SLATE_500 if dark_mode else C_SLATE_700
    add_rect(slide, 0.50, 7.05, 12.33, 0.02, foot_div)
    add_text(slide, 0.50, 7.12, 9.5, 0.26,
             "SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE PROJECT  |  NHÓM 03: ĐỖ QUANG MINH • VŨ QUANG HUY",
             font_size=9.0, color=foot_txt)
    page_str = f"{cur_slide:02d} / {total_slides:02d}"
    add_text(slide, 11.33, 7.12, 1.50, 0.26, page_str, font_size=9.5, bold=True, color=C_CYAN, align=PP_ALIGN.RIGHT)


def add_fitted_image(slide, img_path, left, top, width, height, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.05):
    add_rect(slide, left, top, width, height, bg_rgb, border_rgb=border_rgb, border_pt=1.0, rounded=True)
    if os.path.exists(img_path):
        slide.shapes.add_picture(
            img_path,
            Inches(left + pad),
            Inches(top + pad),
            width=Inches(width - 2 * pad),
            height=Inches(height - 2 * pad)
        )


def build_9slide_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.50)
    blank = prs.slide_layouts[6]
    total_slides = 9

    # =========================================================================
    # SLIDE 1: PROJECT OVERVIEW (HERO COVER LAYOUT - DARK MODE)
    # =========================================================================
    s1 = prs.slides.add_slide(blank)
    add_rect(s1, 0, 0, 13.333, 7.50, C_NAVY_DARK)
    add_rect(s1, 0, 0, 13.333, 0.10, C_BLUE_ELEC)

    # Eyebrow Pill
    add_rect(s1, 0.80, 0.65, 3.80, 0.36, C_NAVY_CARD, border_rgb=C_CYAN, border_pt=1.0, rounded=True)
    add_text(s1, 0.95, 0.70, 3.50, 0.26, "SAMSUNG INNOVATION CAMPUS — AI CAPSTONE", font_size=10.0, bold=True, color=C_CYAN)

    # Main Title
    add_text(s1, 0.80, 1.25, 7.30, 1.70,
             "XÂY DỰNG, HUẤN LUYỆN & SO SÁNH\nCÁC MÔ HÌNH HỌC MÁY DỰ ĐOÁN\nTỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG\nVÀ THAM GIA HỌC TẬP",
             font_size=24.0, bold=True, color=C_WHITE)

    # Subtitle
    add_text(s1, 0.80, 3.15, 7.30, 0.75,
             "Hệ thống Cảnh báo Sớm Học đường Thông minh — Phân tích Tương tác giữa Kiệt sức (Academic Burnout) "
             "và Nguy cơ Bỏ học Đại học (Student Dropout Early Warning System)",
             font_size=12.5, color=C_SLATE_300)

    # Team Members Box
    add_rect(s1, 0.80, 4.25, 7.30, 1.75, C_NAVY_CARD, border_rgb=C_BLUE_ELEC, border_pt=1.5, rounded=True)
    add_text(s1, 1.10, 4.45, 6.70, 0.32, "NHÓM 03 — THÀNH VIÊN THỰC HIỆN ĐỒ ÁN (50% / 50%)", font_size=11.5, bold=True, color=C_CYAN)

    # Member 1
    add_text(s1, 1.10, 4.88, 3.30, 0.30, "ĐỖ QUANG MINH — Trưởng nhóm", font_size=12.0, bold=True, color=C_WHITE)
    add_text(s1, 1.10, 5.20, 3.30, 0.65, "Kiến trúc Pipeline • Tạo 6 Đặc trưng\nMô hình Cây tập hợp & Boosting • FastAPI Backend", font_size=9.5, color=C_SLATE_300)

    # Member 2
    add_text(s1, 4.60, 4.88, 3.30, 0.30, "VŨ QUANG HUY — Thành viên", font_size=12.0, bold=True, color=C_WHITE)
    add_text(s1, 4.60, 5.20, 3.30, 0.65, "Kiểm toán Dữ liệu • Phân tích Khám phá EDA\nKNNImputer (k=7) • Mô hình Cơ sở • Web Frontend", font_size=9.5, color=C_SLATE_300)

    # Right Hero KPI Cards (4 Cards)
    kpis = [
        ("800 → 2.000", "Quy mô Hồ sơ Sinh viên", "800 gốc (612 ô khuyết) mở rộng 2.000 sạch (0 NaN)", C_CYAN),
        ("25 → 34", "Thuộc tính Dữ liệu", "Tích hợp 6 chỉ số tương tác chuyên ngành mới", C_TEAL),
        ("89.01%", "Đỉnh CV ROC-AUC (KB1)", "Random Forest (Kịch bản chẩn đoán đầy đủ)", C_AMBER_BRIGHT),
        ("88.26%", "Cảnh báo Sớm AUC (KB2)", "XGBoost (Kịch bản hành vi thô, không dùng Burnout)", C_BLUE_ELEC),
    ]
    for idx, (val, title, desc, col) in enumerate(kpis):
        y = 1.25 + idx * 1.38
        add_rect(s1, 8.50, y, 4.05, 1.22, C_NAVY_CARD, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s1, 8.50, y, 0.10, 1.22, col)
        add_text(s1, 8.75, y + 0.14, 1.65, 0.50, val, font_size=23.0, bold=True, color=col)
        add_text(s1, 10.45, y + 0.16, 2.00, 0.30, title, font_size=11.5, bold=True, color=C_WHITE)
        add_text(s1, 10.45, y + 0.52, 2.00, 0.55, desc, font_size=9.0, color=C_SLATE_300)

    add_text(s1, 11.33, 7.12, 1.50, 0.26, "01 / 09", font_size=9.5, bold=True, color=C_CYAN, align=PP_ALIGN.RIGHT)

    # =========================================================================
    # SLIDE 2: PROBLEM STATEMENT (BÀI TOÁN CHÚNG TÔI GIẢI QUYẾT)
    # =========================================================================
    s2 = prs.slides.add_slide(blank)
    apply_chrome(s2, 2, total_slides,
                 "01 ĐẶT VẤN ĐỀ NGHIÊN CỨU",
                 "Bối Cảnh Bài Toán & Cơ Chế Hình Thành Nguy Cơ Bỏ Học",
                 "Bỏ học không diễn ra đột ngột mà là kết quả tích lũy từ suy giảm Gắn kết Học tập và Căng thẳng Tâm lý kéo dài")

    # Left Context Card (Target Distribution & Critical Insight)
    add_rect(s2, 0.50, 1.65, 3.80, 5.15, C_WHITE, border_rgb=C_BLUE_ELEC, border_pt=1.5, rounded=True)
    add_rect(s2, 0.50, 1.65, 3.80, 0.40, C_BLUE_ELEC)
    add_text(s2, 0.70, 1.72, 3.40, 0.28, "BIẾN MỤC TIÊU: DROPOUT_RISK", font_size=11.5, bold=True, color=C_WHITE)

    add_text(s2, 0.70, 2.25, 3.40, 0.55, "54.62%", font_size=32.0, bold=True, color=C_CORAL, align=PP_ALIGN.CENTER)
    add_text(s2, 0.70, 2.85, 3.40, 0.45, "Sinh viên trong mẫu thuộc nhóm Có Nguy cơ Bỏ học (437/800 sinh viên)",
             font_size=11.0, bold=True, color=C_SLATE_900, align=PP_ALIGN.CENTER)

    bullet_p2 = [
        "• Thách thức thực tiễn: Nhà trường thường chỉ phát hiện khi sinh viên đã tích lũy nợ môn hoặc nộp đơn xin thôi học.",
        "• Mô hình tích hợp: Kết hợp dữ liệu điểm danh, học lực, tâm lý và áp lực tài chính.",
        "• Mục tiêu giải pháp: Xây dựng hệ thống học máy chuyển từ can thiệp bị động sang Cảnh báo Sớm Chủ động (Proactive Early Warning)."
    ]
    for i, b in enumerate(bullet_p2):
        add_text(s2, 0.70, 3.50 + i * 1.05, 3.40, 0.95, b, font_size=10.0, color=C_SLATE_700)

    # Right Side: 3 Sequential Causal Stages
    stages = [
        ("GIAI ĐOẠN 1: NGUYÊN NHÂN GỐC", "Suy Giảm Học Tập & Áp Lực Tâm Lý Kéo Dài",
         "• Chuyên cần sụt giảm (Attendance < 70%), giờ tự học ít (<2h/ngày).\n"
         "• Áp lực đa chiều: Căng thẳng học vụ, lo âu, gánh nặng tài chính gia đình, mất ngủ do dùng màn hình quá mức.\n"
         "• Điểm tựa suy yếu: Thiếu sự đồng hành của gia đình và bạn bè.",
         C_SLATE_700),
        ("GIAI ĐOẠN 2: BIẾN TRUNG GIAN", "Kiệt Sức Học Đường (Academic Burnout_Level)",
         "• Cầu nối tâm lý trực tiếp: Khi áp lực vượt quá ngưỡng chịu đựng, sinh viên rơi vào trạng thái kiệt sức.\n"
         "• Phân tầng 3 mức: Thấp (Low: 40%), Vừa (Medium: 35%), Cao (High: 25%).\n"
         "• Tương quan cực mạnh: Hệ số Pearson r = +0.550 với nguy cơ bỏ học.",
         C_AMBER),
        ("GIAI ĐOẠN 3: CAN THIỆP SỚM", "Hệ Thống Cảnh Báo Sớm Bằng Học Máy (AI Solution)",
         "• Kịch bản 1 (Đầy đủ): Chẩn đoán chính xác khi có thông tin kiệt sức (CV ROC-AUC 89.01%).\n"
         "• Kịch bản 2 (Cảnh báo sớm): Dự báo trực tiếp từ hành vi và áp lực thô (CV ROC-AUC 88.26%).\n"
         "• Can thiệp kịp thời: Tư vấn tâm lý, phụ đạo chuyên môn, giãn tiến độ, cấp học bổng hỗ trợ.",
         C_TEAL),
    ]

    for idx, (st_eyebrow, st_title, st_body, col) in enumerate(stages):
        y = 1.65 + idx * 1.76
        add_rect(s2, 4.60, y, 8.23, 1.62, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s2, 4.60, y, 0.12, 1.62, col)
        add_text(s2, 4.88, y + 0.12, 3.00, 0.24, st_eyebrow, font_size=9.5, bold=True, color=col)
        add_text(s2, 4.88, y + 0.38, 7.80, 0.30, st_title, font_size=12.0, bold=True, color=C_SLATE_900)
        add_text(s2, 4.88, y + 0.72, 7.80, 0.82, st_body, font_size=9.5, color=C_SLATE_700)

    # =========================================================================
    # SLIDE 3: DATASET & DATA UNDERSTANDING (DỮ LIỆU ĐẦU VÀO)
    # =========================================================================
    s3 = prs.slides.add_slide(blank)
    apply_chrome(s3, 3, total_slides,
                 "02 DỮ LIỆU ĐẦU VÀO",
                 "Kiến Trúc Bộ Dữ Liệu (800 → 2.000 Hồ Sơ) & Phân Bố Thuộc Tính",
                 "Minh bạch quy mô: 800 bản ghi gốc ban đầu, xử lý 612 ô khuyết trên 18 cột và mở rộng mẫu kỹ thuật 2.000 hồ sơ sạch")

    # Top 5 KPI Cards Strip
    top_metrics = [
        ("800", "Bản ghi Gốc ban đầu", "Khảo sát 25 thuộc tính", C_BLUE_ELEC),
        ("612", "Ô dữ liệu bị khuyết", "Nằm trên 18/25 cột (1.6%–6.4%)", C_CORAL),
        ("k = 7", "KNNImputer Đa biến", "Nội suy từ 7 SV láng giềng", C_TEAL),
        ("2.000", "Hồ sơ Sau Mở rộng", "800 gốc sạch + 1.200 tổng hợp", C_AMBER),
        ("34", "Thuộc tính Hoàn chỉnh", "Bao gồm 6 đặc trưng mới (0 NaN)", C_NAVY_DARK),
    ]
    for idx, (m_val, m_t, m_s, col) in enumerate(top_metrics):
        x = 0.50 + idx * 2.50
        add_rect(s3, x, 1.65, 2.35, 1.25, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s3, x, 1.65, 2.35, 0.08, col)
        add_text(s3, x + 0.08, 1.75, 2.18, 0.45, m_val, font_size=23.0, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s3, x + 0.08, 2.22, 2.18, 0.28, m_t, font_size=10.5, bold=True, color=C_SLATE_900, align=PP_ALIGN.CENTER)
        add_text(s3, x + 0.08, 2.50, 2.18, 0.32, m_s, font_size=9.0, color=C_SLATE_500, align=PP_ALIGN.CENTER)

    # Bottom Left: 5 Attribute Domains & Provenance
    add_rect(s3, 0.50, 3.10, 5.80, 3.70, C_WHITE, border_rgb=C_SLATE_300, border_pt=1.0, rounded=True)
    add_rect(s3, 0.50, 3.10, 5.80, 0.38, C_NAVY_DARK)
    add_text(s3, 0.70, 3.16, 5.40, 0.26, "5 NHÓM DỮ LIỆU & NGUỒN GỐC MẪU MỞ RỘNG", font_size=11.0, bold=True, color=C_CYAN)

    domain_bullets = [
        "1. Nhân khẩu học (7 biến): Tuổi (17–25), Giới tính, Năm học, Khoa đào tạo, Loại hình cư trú, Mức thu nhập gia đình.",
        "2. Tham gia học tập (5 biến): Chuyên cần (%), Giờ tự học/ngày, GPA kỳ trước, Số môn nợ, Điểm động lực nội tại.",
        "3. Lối sống & Sinh hoạt (5 biến): Giờ ngủ, Giờ màn hình, Tập thể dục, Hoạt động xã hội, Làm thêm.",
        "4. Căng thẳng & Hỗ trợ (6 biến): Căng thẳng học vụ, Lo âu, Áp lực tài chính, Bạn bè, Gia đình, Tiếp cận tư vấn.",
        "5. Biến mục tiêu (2 biến): Mức kiệt sức (Low/Med/High) & Nguy cơ bỏ học (Dropout_Risk: Yes/No).",
        "Minh bạch nguồn gốc 2.000 hồ sơ: 800 bản ghi gốc được làm sạch 6 bước, kết hợp 1.200 bản ghi sinh từ phân phối Gaussian đa biến phân tầng nhằm kiểm thử khả năng mở rộng quy mô kỹ thuật và giao diện Web."
    ]
    for i, b in enumerate(domain_bullets):
        add_text(s3, 0.70, 3.55 + i * 0.52, 5.40, 0.50, b, font_size=9.2, color=C_SLATE_700)

    # Bottom Right: Missing Value Analysis Image
    add_fitted_image(s3, os.path.join(FIG_DIR, "01_missing_values_analysis.png"), 6.50, 3.10, 6.33, 3.70, pad=0.06)

    # =========================================================================
    # SLIDE 4: PREPROCESSING & FEATURE ENGINEERING (TIỀN XỬ LÝ & ĐẶC TRƯNG)
    # =========================================================================
    s4 = prs.slides.add_slide(blank)
    apply_chrome(s4, 4, total_slides,
                 "03 TIỀN XỬ LÝ & KỸ THUẬT ĐẶC TRƯNG",
                 "Quy Trình Học Máy 5 Bước & Thiết Kế 6 Đặc Trưng Chuyên Ngành",
                 "Biến đổi thuộc tính rời rạc thành 6 chỉ số tổng hợp lượng hóa trực tiếp cán cân Căng thẳng Tâm lý vs. Gắn kết Học tập")

    # Top 5-Step Process Ribbon
    steps = [
        ("BƯỚC 1: LÀM SẠCH", "Xóa trùng ID & Lọc ngoại lai IQR", C_SLATE_700),
        ("BƯỚC 2: ĐIỀN KHUYẾT", "KNNImputer (k=7) trong Pipeline", C_BLUE_ELEC),
        ("BƯỚC 3: TẠO ĐẶC TRƯNG", "DomainFeatureEngineer (6 chỉ số)", C_TEAL),
        ("BƯỚC 4: CHUẨN HÓA", "OneHotEncoder & StandardScaler", C_CYAN),
        ("BƯỚC 5: KIỂM ĐỊNH", "Stratified 5-Fold CV & Test 20%", C_NAVY_DARK),
    ]
    for idx, (st_t, st_s, col) in enumerate(steps):
        x = 0.50 + idx * 2.49
        add_rect(s4, x, 1.60, 2.36, 0.82, col, rounded=True)
        add_text(s4, x + 0.08, 1.66, 2.20, 0.28, st_t, font_size=10.0, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.08, 1.96, 2.20, 0.40, st_s, font_size=9.0, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Bottom 2x3 Grid of 6 Engineered Features
    features = [
        ("F1", "Academic_Engagement_Index", "0.35×(Attendance/10) + 0.25×Study + 0.25×Motivation + 0.15×GPA",
         "Đo lường tổng hợp mức độ chuyên cần, thời lượng tự học, động lực và kết quả học tập.", "Top 3 Đặc trưng Quan trọng", C_TEAL),
        ("F2", "Psychological_Stress_Index", "0.35×Stress + 0.30×Anxiety + 0.20×Financial + 0.15×Peer_Pressure",
         "Đo lường tổng hợp cường độ áp lực tâm lý, lo âu, gánh nặng tài chính và bạn bè.", "Top 4 Đặc trưng Quan trọng", C_CORAL),
        ("F3", "Stress_Engagement_Ratio", "Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)",
         "Tỷ số mất cân bằng khi áp lực tâm lý vượt quá độ gắn kết học tập (+1.0 tránh chia 0).", "Top 2 Đặc trưng Quan trọng", C_BLUE_ELEC),
        ("F4", "Net_Stress_Support_Balance", "Stress_Index − (0.50×Family + 0.30×Social + 0.20×Exercise)",
         "Mức căng thẳng ròng còn lại sau khi đã khấu trừ các điểm tựa hỗ trợ gia đình & thể chất.", "Top 6 Đặc trưng Quan trọng", C_CYAN),
        ("F5", "Digital_Fatigue_Ratio", "Screen_Time_Hours / (Sleep_Hours + 1.0)",
         "Đánh giá hội chứng mệt mỏi số hóa và rối loạn nhịp sinh học do dùng màn hình quá mức.", "Chỉ báo rủi ro lối sống", C_AMBER),
        ("F6", "Academic_Risk_Burden", "1.50 × Backlogs + (100.0 − Attendance_Percent) / 20.0",
         "Lượng hóa gánh nặng học vụ tích lũy từ số môn nợ và tỷ lệ vắng học trên lớp.", "Chỉ báo cảnh báo học vụ", C_NAVY_DARK),
    ]
    for idx, (code, fname, formula, desc, rank, col) in enumerate(features):
        r_i, c_i = divmod(idx, 3)
        x = 0.50 + c_i * 4.15
        y = 2.60 + r_i * 2.10
        add_rect(s4, x, y, 4.02, 1.95, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s4, x, y, 4.02, 0.08, col)
        add_text(s4, x + 0.15, y + 0.14, 0.45, 0.30, code, font_size=12.0, bold=True, color=col)
        add_text(s4, x + 0.65, y + 0.14, 3.20, 0.30, fname, font_size=11.0, bold=True, color=C_SLATE_900)

        # Formula strip
        add_rect(s4, x + 0.15, y + 0.46, 3.72, 0.44, C_NAVY_CARD, rounded=True)
        add_text(s4, x + 0.20, y + 0.50, 3.62, 0.36, formula, font_size=8.8, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

        add_text(s4, x + 0.15, y + 0.96, 3.72, 0.55, desc, font_size=9.0, color=C_SLATE_700)

        add_rect(s4, x + 0.15, y + 1.55, 3.72, 0.28, C_SLATE_100, border_rgb=col, border_pt=1.0, rounded=True)
        add_text(s4, x + 0.20, y + 1.58, 3.62, 0.22, rank, font_size=8.8, bold=True, color=col, align=PP_ALIGN.CENTER)

    # =========================================================================
    # SLIDE 5: EXPLORATORY DATA ANALYSIS (PHÂN TÍCH KHÁM PHÁ DỮ LIỆU)
    # =========================================================================
    s5 = prs.slides.add_slide(blank)
    apply_chrome(s5, 5, total_slides,
                 "04 PHÂN TÍCH KHÁM PHÁ (EDA)",
                 "Phát Hiện Trọng Tâm: Tỷ Lệ Bỏ Học Theo Mức Kiệt Sức & Tương Quan",
                 "Mức độ kiệt sức học đường càng cao thì tỷ lệ bỏ học quan sát được trong bộ dữ liệu càng lớn (tăng từ 23.8% lên 92.0%)")

    # Left: Burnout Progression Chart
    add_fitted_image(s5, os.path.join(FIG_DIR, "art_burnout_progression.png"), 0.50, 1.65, 5.80, 5.15, pad=0.08)

    # Right: 3 Big Numbers + Correlation Highlights
    b_rates = [
        ("23.8%", "MỨC THẤP — LOW BURNOUT", "76 / 320 sinh viên bỏ học | 244 sinh viên duy trì việc học an toàn", C_TEAL),
        ("63.2%", "MỨC VỪA — MEDIUM BURNOUT", "177 / 280 sinh viên bỏ học | Ngưỡng chuyển giao báo động cần can thiệp", C_AMBER),
        ("92.0%", "MỨC CAO — HIGH BURNOUT", "184 / 200 sinh viên bỏ học | Nguy cơ bỏ học gần như tuyệt đối (gấp 3.87 lần)", C_CORAL),
    ]
    for idx, (pct_val, title, sub, col) in enumerate(b_rates):
        y = 1.65 + idx * 1.25
        add_rect(s5, 6.50, y, 6.33, 1.12, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s5, 6.50, y, 0.12, 1.12, col)
        add_text(s5, 6.75, y + 0.14, 1.80, 0.50, pct_val, font_size=28.0, bold=True, color=col)
        add_text(s5, 8.65, y + 0.18, 4.05, 0.30, title, font_size=11.5, bold=True, color=C_SLATE_900)
        add_text(s5, 8.65, y + 0.52, 4.05, 0.50, sub, font_size=9.2, color=C_SLATE_700)

    # Bottom Right Box: Correlation Matrix Highlights
    add_rect(s5, 6.50, 5.50, 6.33, 1.30, C_NAVY_CARD, border_rgb=C_CYAN, border_pt=1.5, rounded=True)
    add_text(s5, 6.75, 5.58, 5.90, 0.28, "ĐỐI LẬP HAI CỰC: YẾU TỐ BẢO VỆ VS. YẾU TỐ NGUY CƠ (PEARSON r TRÊN 800 MẪU)",
             font_size=10.0, bold=True, color=C_CYAN)
    corr_summary = (
        "• Nhóm yếu tố bảo vệ (r < 0): Attendance_Percent (r = -0.244, chênh lệch 80.7% vs 73.7%), Motivation (r = -0.193), "
        "Study_Hours (-0.160), Family_Support & GPA (-0.137).\n"
        "• Nhóm yếu tố nguy cơ (r > 0): Burnout_Numeric (r = +0.550, mạnh nhất), Stress_Level (+0.164), Anxiety (+0.145), "
        "Financial_Stress (+0.114). Sinh viên thu nhập thấp có tỷ lệ bỏ học lên tới 68.4%."
    )
    add_text(s5, 6.75, 5.90, 5.90, 0.82, corr_summary, font_size=9.0, color=C_WHITE)

    # =========================================================================
    # SLIDE 6: AI/ML METHODOLOGY (PHƯƠNG PHÁP LUẬN HỌC MÁY)
    # =========================================================================
    s6 = prs.slides.add_slide(blank)
    apply_chrome(s6, 6, total_slides,
                 "05 PHƯƠNG PHÁP LUẬN HỌC MÁY",
                 "Thiết Kế Thực Nghiệm Đối Chứng 2 Kịch Bản & Chống Rò Rỉ Dữ Liệu",
                 "Đánh giá khách quan qua 5-Fold Stratified Cross-Validation và kiểm thử độc lập 20% Hold-out Test")

    # 2 Comparison Columns (Scenario 1 vs Scenario 2)
    scenarios = [
        ("KỊCH BẢN 1: CHẨN ĐOÁN TOÀN DIỆN (FULL FEATURES)",
         "Sử dụng toàn bộ đặc trưng bao gồm cả trạng thái Burnout_Level đã xác định",
         [
             "• Mục đích: Đánh giá độ chính xác chẩn đoán tối đa khi có đầy đủ dữ liệu khảo sát tâm lý chuyên sâu.",
             "• Tập đặc trưng: 32 biến đầu vào (bao gồm Burnout_Ordinal, 6 đặc trưng tương tác và mã hóa One-Hot).",
             "• Mô hình tối ưu: Random Forest đạt CV ROC-AUC 89.01% ± 0.80% (Test ROC-AUC: 88.42%, Test F1: 83.63%).",
             "• Ý nghĩa nghiệp vụ: Dùng để sàng lọc đợt cao điểm giữa kỳ khi đã có kết quả đánh giá kiệt sức học đường."
         ],
         C_AMBER_BRIGHT),
        ("KỊCH BẢN 2: CẢNH BÁO SỚM HỌC ĐƯỜNG (EARLY WARNING)",
         "Loại bỏ hoàn toàn biến Burnout_Level, chỉ dùng các chỉ số hành vi học tập & áp lực thô",
         [
             "• Mục đích: Cảnh báo sớm chủ động ngay từ đầu kỳ học khi sinh viên chưa bộc lộ hoặc chưa khảo sát kiệt sức.",
             "• Tập đặc trưng: 30 biến đầu vào (chỉ dùng chuyên cần, giờ học, GPA, stress, lo âu, tài chính thô).",
             "• Mô hình tối ưu: XGBoost đạt CV ROC-AUC 88.26% ± 1.03% (Test ROC-AUC: 87.24%, Test F1: 79.82%).",
             "• Ý nghĩa nghiệp vụ: Trọng tâm ứng dụng thực tiễn của hệ thống cảnh báo sớm học đường."
         ],
         C_BLUE_ELEC),
    ]

    for idx, (title, sub, bullets, col) in enumerate(scenarios):
        x = 0.50 + idx * 6.25
        add_rect(s6, x, 1.65, 6.08, 3.45, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s6, x, 1.65, 6.08, 0.40, col)
        add_text(s6, x + 0.20, 1.72, 5.68, 0.28, title, font_size=11.0, bold=True, color=C_WHITE)
        add_text(s6, x + 0.20, 2.15, 5.68, 0.32, sub, font_size=10.0, italic=True, color=C_SLATE_500)
        for b_i, b in enumerate(bullets):
            add_text(s6, x + 0.20, 2.55 + b_i * 0.58, 5.68, 0.55, b, font_size=9.5, color=C_SLATE_700)

    # Bottom Full-Width Strip: Scientific Rigor & Pipeline Architecture
    add_rect(s6, 0.50, 5.25, 12.33, 1.55, C_NAVY_CARD, border_rgb=C_CYAN, border_pt=1.5, rounded=True)
    add_text(s6, 0.75, 5.35, 11.83, 0.26, "NGUYÊN TẮC BẢO TOÀN TÍNH KHOA HỌC & CHỐNG RÒ RỈ DỮ LIỆU (NO DATA LEAKAGE)",
             font_size=11.0, bold=True, color=C_CYAN)

    rigor_points = (
        "• Pipeline Khép Kín: Toàn bộ quá trình điền khuyết (KNNImputer k=7), chuẩn hóa (StandardScaler) và mã hóa (OneHotEncoder) "
        "được đóng gói trong đối tượng Pipeline của scikit-learn, chỉ khớp (fit) trên fold huấn luyện.\n"
        "• Phân chia Phân tầng (Stratified Split): Bảo toàn tuyệt đối tỷ lệ nhãn Dropout_Risk (54.6% Yes / 45.4% No) trên cả 5 folds của Cross-Validation "
        "và tập kiểm thử độc lập 20% (Hold-out Test Set).\n"
        "• Lưu ý học thuật: Việc mở rộng 1.200 bản ghi tổng hợp phân tầng giúp kiểm thử khả năng mở rộng kỹ thuật của hệ thống, nhưng đánh giá hiệu năng "
        "thực chất nhất của mô hình trên sinh viên thực tế nằm trong khoảng ROC-AUC 80.50% (trên 800 mẫu gốc)."
    )
    add_text(s6, 0.75, 5.68, 11.83, 1.05, rigor_points, font_size=9.5, color=C_WHITE)

    # =========================================================================
    # SLIDE 7: MODEL PERFORMANCE & EVALUATION (KẾT QUẢ THỰC NGHIỆM)
    # =========================================================================
    s7 = prs.slides.add_slide(blank)
    apply_chrome(s7, 7, total_slides,
                 "06 KẾT QUẢ THỰC NGHIỆM & ĐỐI SÁNH",
                 "Đối Sánh 4 Mô Hình Cốt Lõi Trên 2 Kịch Bản & Độ Quan Trọng Đặc Trưng",
                 "Bảng số liệu đối chứng minh bạch giữa 5-Fold Cross-Validation và Tập kiểm thử độc lập 20%")

    # Left: Benchmark Chart
    add_fitted_image(s7, os.path.join(FIG_DIR, "art_model_benchmark.png"), 0.50, 1.65, 5.80, 2.70, pad=0.06)

    # Right: ROC/PR Curves
    add_fitted_image(s7, os.path.join(FIG_DIR, "08_roc_and_precision_recall_curves.png"), 6.50, 1.65, 6.33, 2.70, pad=0.06)

    # Bottom Table: 4 Flagship Models Benchmarking
    table_shp = s7.shapes.add_table(5, 9, Inches(0.50), Inches(4.50), Inches(12.33), Inches(1.85))
    tbl = table_shp.table
    tbl.columns[0].width = Inches(2.65)
    for c_i in range(1, 9):
        tbl.columns[c_i].width = Inches(1.21)

    headers = ["Mô hình Học máy", "KB1 CV Acc", "KB1 CV Recall", "KB1 CV F1", "KB1 CV AUC",
               "KB1 Test Acc", "KB1 Test F1", "KB1 Test AUC", "KB2 CV AUC"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Segoe UI"
        p.font.size = Pt(9.2)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    data_rows = [
        ["1. Random Forest (Top 1 KB1)", "80.25% ± 0.65%", "87.32%", "82.98%", "89.01% ± 0.80%", "81.50%", "83.63%", "88.42%", "87.97%"],
        ["2. Soft Voting Ensemble", "79.70% ± 1.27%", "85.33%", "82.27%", "88.89% ± 0.84%", "80.25%", "82.17%", "88.32%", "88.25%"],
        ["3. XGBoost (Top 1 KB2 Sớm)", "79.95% ± 1.11%", "84.96%", "82.39%", "88.64% ± 0.84%", "78.50%", "80.28%", "87.99%", "88.26% ± 1.03%"],
        ["4. Logistic Regression", "78.45% ± 0.68%", "81.25%", "80.62%", "87.61% ± 1.02%", "78.25%", "80.18%", "87.60%", "87.25%"],
    ]
    for i, row in enumerate(data_rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_SLATE_100 if i % 2 == 1 else C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Segoe UI"
            p.font.size = Pt(9.0)
            p.font.color.rgb = C_SLATE_900
            if j == 0:
                p.alignment = PP_ALIGN.LEFT
                p.font.bold = True
            else:
                p.alignment = PP_ALIGN.CENTER
                if "89.01%" in val or "88.26%" in val:
                    p.font.bold = True
                    p.font.color.rgb = C_BLUE_ELEC

    # Note below table
    add_text(s7, 0.50, 6.45, 12.33, 0.50,
             "Ma trận Nhầm lẫn (Confusion Matrix n=400): Random Forest KB1 đạt TP=189, TN=137, FP=42, FN=32 (Độ nhạy Recall = 85.52%). "
             "Top Đặc trưng quyết định: (1) Stress_Engagement_Ratio → (2) Academic_Engagement_Index → (3) Attendance_Percent → (4) Burnout_Ordinal.",
             font_size=9.0, color=C_SLATE_700)

    # =========================================================================
    # SLIDE 8: DEMO / SYSTEM ARCHITECTURE (KIẾN TRÚC & WEB DASHBOARD)
    # =========================================================================
    s8 = prs.slides.add_slide(blank)
    apply_chrome(s8, 8, total_slides,
                 "07 TRIỂN KHAI HỆ THỐNG",
                 "Kiến Trúc 3 Tầng & Ứng Dụng Web Dashboard Cảnh Báo Sớm (FastAPI)",
                 "Chuyển hóa mô hình học máy thành công cụ hỗ trợ ra quyết định thời gian thực cho cố vấn học tập")

    # Left: 3-Tier Architecture Flow Cards
    tiers = [
        ("TẦNG 1: GIAO DIỆN NHẬP LIỆU (INPUT LAYER)",
         "• Thu thập 18 chỉ số học tập, lối sống và áp lực tâm lý qua thanh trượt trực quan.\n"
         "• Tích hợp 3 nút chọn nhanh hồ sơ mẫu (Presets): SV Tích cực, SV Căng thẳng TB, SV Kiệt sức Cao.",
         C_SLATE_700),
        ("TẦNG 2: MÁY HỌC DỰ BÁO (AI BACKEND FASTAPI)",
         "• Vận hành song song mô hình dự báo Burnout (3 mức) và mô hình Dropout (Random Forest / XGBoost).\n"
         "• Tính toán tự động 6 chỉ số đặc trưng chuyên ngành; độ trễ phản hồi đo nội bộ: ≈145–153ms.",
         C_BLUE_ELEC),
        ("TẦNG 3: KẾT QUẢ & KHUYẾN NGHỊ (OUTPUT & ACTION)",
         "• Hiển thị trực quan xác suất rủi ro bỏ học (%) và đồng hồ cảnh báo phân tầng màu sắc.\n"
         "• Module khuyến nghị tự động: Đề xuất nhắc nhở chuyên cần, học nhóm, tư vấn tâm lý hoặc quỹ học bổng.",
         C_TEAL),
    ]
    for idx, (t_title, t_body, col) in enumerate(tiers):
        y = 1.65 + idx * 1.48
        add_rect(s8, 0.50, y, 5.50, 1.36, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s8, 0.50, y, 0.10, 1.36, col)
        add_text(s8, 0.72, y + 0.10, 5.15, 0.26, t_title, font_size=10.0, bold=True, color=col)
        add_text(s8, 0.72, y + 0.38, 5.15, 0.90, t_body, font_size=8.8, color=C_SLATE_700)

    # Right: Live Web Screenshot
    add_fitted_image(s8, os.path.join(FIG_DIR, "13_live_web_dashboard_screenshot.png"), 6.20, 1.65, 6.63, 4.25, pad=0.06)

    # Bottom Right Box: Verified Presets Output
    add_rect(s8, 6.20, 6.00, 6.63, 0.95, C_NAVY_CARD, border_rgb=C_CYAN, border_pt=1.5, rounded=True)
    add_text(s8, 6.35, 6.06, 6.30, 0.24, "KẾT QUẢ ĐO LƯỜNG THỰC TẾ TRÊN 3 HỒ SƠ MẪU (PRESETS)", font_size=9.5, bold=True, color=C_CYAN)
    preset_txt = (
        "• Mẫu 1 (SV Tích cực): Bỏ học 1.7% (Auto: 2.2% - An toàn)  |  "
        "• Mẫu 2 (SV Căng thẳng TB): Bỏ học 60.5% (Auto: 59.3% - Cảnh báo)  |  "
        "• Mẫu 3 (SV Kiệt sức Cao): Bỏ học 98.6% (Auto: 98.9% - Báo động đỏ)"
    )
    add_text(s8, 6.35, 6.34, 6.30, 0.55, preset_txt, font_size=8.8, color=C_WHITE)

    # =========================================================================
    # SLIDE 9: CONCLUSION & FUTURE WORK (KẾT LUẬN & HƯỚNG PHÁT TRIỂN)
    # =========================================================================
    s9 = prs.slides.add_slide(blank)
    apply_chrome(s9, 9, total_slides,
                 "08 TỔNG KẾT & ĐỊNH HƯỚNG",
                 "Đóng Góp Đồ Án, Nhìn Nhận Hạn Chế Khoa Học & Hướng Phát Triển",
                 "Hoàn thiện chuỗi giá trị từ dữ liệu, kỹ thuật tạo đặc trưng đến ứng dụng Web phục vụ cộng đồng giáo dục",
                 dark_mode=True)

    # 3 Pillar Cards
    pillars = [
        ("01 ĐÓNG GÓP KHOA HỌC & KỸ THUẬT",
         "• Xây dựng thành công quy trình học máy chuẩn mực: điền khuyết KNNImputer (k=7) trong Pipeline chống rò rỉ dữ liệu.\n"
         "• Thiết kế 6 đặc trưng chuyên ngành định lượng cán cân Căng thẳng vs Gắn kết; chứng minh Stress_Engagement_Ratio và Academic_Engagement_Index giữ vai trò quyết định.\n"
         "• Đối sánh khách quan 10 mô hình trên 800 mẫu gốc và 4 mô hình cốt lõi trên 2.000 hồ sơ; xây dựng ứng dụng Web hoàn chỉnh.",
         C_CYAN),
        ("02 NHÌN NHẬN HẠN CHẾ KHOA HỌC",
         "• Dữ liệu cắt ngang (Cross-sectional): Chưa theo dõi được diễn tiến tâm lý sinh viên theo từng tuần trong kỳ học.\n"
         "• Thử nghiệm dữ liệu tổng hợp: Mở rộng 1.200 bản ghi Gaussian phân tầng giúp kiểm thử tải hệ thống, nhưng độ phân tách trên mẫu sinh cao hơn thực tế (AUC phần thật đạt ~77–80%).\n"
         "• Feature Importance: Phản ánh mức đóng góp phân tách của thuật toán cây, không đồng nhất với quan hệ nhân quả một chiều.",
         C_CORAL),
        ("03 ĐỊNH HƯỚNG PHÁT TRIỂN TƯƠNG LAI",
         "• Thu thập chuỗi thời gian (Longitudinal Data): Ghi nhận chuyên cần và điểm kiểm tra định kỳ hàng tuần để huấn luyện mô hình LSTM / Transformer.\n"
         "• Giải thích cục bộ (Explainable AI): Tích hợp biểu đồ SHAP Waterfall trực tiếp lên Web UI giải thích từng trường hợp sinh viên.\n"
         "• Tích hợp hệ thống LMS: Kết nối tự động qua API với Moodle / Canvas của nhà trường để đồng bộ hóa dữ liệu điểm danh.",
         C_TEAL),
    ]

    for idx, (p_title, p_body, col) in enumerate(pillars):
        x = 0.50 + idx * 4.15
        add_rect(s9, x, 1.65, 4.02, 4.25, C_NAVY_CARD, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s9, x, 1.65, 4.02, 0.40, col)
        add_text(s9, x + 0.15, 1.72, 3.72, 0.28, p_title, font_size=10.5, bold=True, color=C_WHITE)
        add_text(s9, x + 0.15, 2.20, 3.72, 3.55, p_body, font_size=9.2, color=C_SLATE_300)

    # Bottom Thank You & Q&A Bar
    add_rect(s9, 0.50, 6.05, 12.33, 0.85, C_NAVY_CARD, border_rgb=C_BLUE_ELEC, border_pt=1.5, rounded=True)
    add_text(s9, 0.75, 6.15, 8.50, 0.30, "TRÂN TRỌNG CẢM ƠN QUÝ THẦY CÔ VÀ HỘI ĐỒNG CHẤM ĐỒ ÁN! — Q & A",
             font_size=12.5, bold=True, color=C_WHITE)
    add_text(s9, 0.75, 6.48, 8.50, 0.30,
             "Nhóm 03: Đỗ Quang Minh & Vũ Quang Huy  |  Samsung Innovation Campus (SIC)  |  Web App: http://127.0.0.1:8000",
             font_size=9.5, color=C_CYAN)

    prs.save(OUT_PATH)
    print(f"Generated successfully: {OUT_PATH}")


if __name__ == "__main__":
    build_9slide_presentation()
