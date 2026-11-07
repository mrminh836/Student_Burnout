"""
Builds the upgraded, high-tech, academic Dark Navy master PowerPoint deck for
Samsung Innovation Campus (SIC) AI Capstone Project:
"XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG VÀ THAM GIA HỌC TẬP"
Group 03: Đỗ Quang Minh (Trưởng nhóm) & Vũ Quang Huy (Thành viên)

Design System:
- Canvas Background: Dark Navy #0B1220
- Card Background: #111C2E
- Borders: #1E293B / #2A3B53
- Primary Accent: Electric Cyan #00D4FF
- Secondary Accent: Vibrant Teal #00E5B0
- Amber Accent: #FFB703
- Coral / Alert: #FF6B6B
- Text Main: White #F8FAFC
- Text Secondary: Muted Silver-Blue #AAB8CC
- 16:9 Widescreen (13.333 x 7.5 inches)
"""

import os
from PIL import Image as PilImage
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG_VI = os.path.join(BASE_DIR, "figures_vi")
FINAL_DIR = os.path.join(BASE_DIR, "FINAL_PRESENTATION")
os.makedirs(FINAL_DIR, exist_ok=True)
PPTX_OUT = os.path.join(FINAL_DIR, "SIC_AI_Capstone_Final.pptx")

# Palette definitions
C_DARK_BG = RGBColor(11, 18, 32)      # #0B1220 Primary Canvas
C_CARD_BG = RGBColor(17, 28, 46)      # #111C2E Secondary Background / Card
C_CARD_INNER = RGBColor(22, 36, 58)   # #16243A Slightly lighter card header
C_BORDER = RGBColor(30, 41, 59)       # #1E293B Muted Slate Border
C_CYAN = RGBColor(0, 212, 255)        # #00D4FF Primary Accent Cyan
C_TEAL = RGBColor(0, 229, 176)        # #00E5B0 Secondary Accent Teal
C_AMBER = RGBColor(255, 183, 3)       # #FFB703 Warm Gold/Amber
C_CORAL = RGBColor(255, 107, 107)     # #FF6B6B Alert Coral/Red
C_BLUE = RGBColor(56, 189, 248)       # #38BDF8 Sky Blue
C_WHITE = RGBColor(248, 250, 252)     # #F8FAFC Main Text White
C_MUTED = RGBColor(170, 184, 204)     # #AAB8CC Secondary Muted Text
C_CONTAINER_BG = RGBColor(255, 255, 255) # Clean White for Figure cards

FONT_FAMILY = "Segoe UI"


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


def add_text(slide, left, top, width, height, text, font_size=13, bold=False, italic=False, color=C_WHITE, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT_FAMILY
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def add_dark_card(slide, left, top, width, height, header_text, body_lines, border_color=C_BORDER, header_color=C_CYAN, body_font=11.0):
    """Draws a premium dark tech card with a subtle top accent bar and high-contrast text."""
    # Outer card
    add_rect(slide, left, top, width, height, C_CARD_BG, border_rgb=border_color, border_pt=1.2, rounded=True)
    # Header bar
    add_rect(slide, left, top, width, 0.40, C_CARD_INNER, rounded=False)
    # Thin accent highlight strip under header
    add_rect(slide, left, top + 0.38, width, 0.02, header_color, rounded=False)
    # Header text
    add_text(slide, left + 0.12, top + 0.05, width - 0.24, 0.30, header_text, font_size=11.5, bold=True, color=header_color)

    # Body text
    tb = slide.shapes.add_textbox(Inches(left + 0.12), Inches(top + 0.45), Inches(width - 0.24), Inches(height - 0.50))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(body_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(3)
        r = p.add_run()
        r.text = line
        r.font.name = FONT_FAMILY
        r.font.size = Pt(body_font)
        
        # Color coding bullets
        if line.startswith("★") or line.startswith("• ["):
            r.font.bold = True
            r.font.color.rgb = C_WHITE
        elif "89.01%" in line or "88.26%" in line or "92.0%" in line:
            r.font.bold = True
            r.font.color.rgb = C_CYAN
        elif line.startswith("  -") or line.startswith("  •"):
            r.font.color.rgb = C_MUTED
        else:
            r.font.color.rgb = C_WHITE


def add_figure_card(slide, img_path, box_left, box_top, box_width, box_height, border_color=C_BORDER, pad=0.08):
    """Embeds an image inside a clean, rounded container card with dark border."""
    add_rect(slide, box_left, box_top, box_width, box_height, C_CONTAINER_BG, border_rgb=border_color, border_pt=1.5, rounded=True)
    if not os.path.exists(img_path):
        return
    with PilImage.open(img_path) as im:
        w_px, h_px = im.size
    aspect = w_px / float(h_px)

    avail_w = box_width - 2 * pad
    avail_h = box_height - 2 * pad

    if avail_w / avail_h > aspect:
        final_h = avail_h
        final_w = avail_h * aspect
    else:
        final_w = avail_w
        final_h = avail_w / aspect

    x = box_left + (box_width - final_w) / 2.0
    y = box_top + (box_height - final_h) / 2.0
    slide.shapes.add_picture(img_path, Inches(x), Inches(y), width=Inches(final_w), height=Inches(final_h))


def create_master_slide(prs, slide_num, total_slides, section_badge, title_text, takeaway_text):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Dark Navy Canvas
    add_rect(slide, 0, 0, 13.333, 7.5, C_DARK_BG)
    # Top Header Band
    add_rect(slide, 0, 0, 13.333, 1.15, C_CARD_BG)
    # Cyan accent line at bottom of header
    add_rect(slide, 0, 1.13, 13.333, 0.03, C_CYAN)
    # Teal left vertical accent
    add_rect(slide, 0, 0, 0.12, 1.15, C_TEAL)

    # Section Pill Badge
    add_rect(slide, 0.36, 0.12, 3.20, 0.26, C_CARD_INNER, border_rgb=C_TEAL, border_pt=1.0, rounded=True)
    add_text(slide, 0.38, 0.13, 3.16, 0.24, section_badge, font_size=9.5, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    # Title & Takeaway
    add_text(slide, 0.36, 0.40, 12.6, 0.38, title_text, font_size=20, bold=True, color=C_WHITE)
    add_text(slide, 0.36, 0.78, 12.6, 0.30, takeaway_text, font_size=11.5, italic=True, color=C_CYAN)

    # Bottom Footer
    add_rect(slide, 0, 7.12, 13.333, 0.38, C_CARD_BG)
    add_rect(slide, 0, 7.12, 13.333, 0.02, C_BORDER)
    footer_left = "CHƯƠNG TRÌNH SAMSUNG INNOVATION CAMPUS (SIC) — ĐỒ ÁN AI CAPSTONE  |  NHÓM 03: ĐỖ QUANG MINH • VŨ QUANG HUY"
    add_text(slide, 0.38, 7.17, 10.2, 0.28, footer_left, font_size=9.5, bold=True, color=C_MUTED)
    add_text(slide, 11.2, 7.17, 1.8, 0.28, f"Slide {slide_num:02d} / {total_slides:02d}", font_size=9.5, bold=True, color=C_AMBER, align=PP_ALIGN.RIGHT)

    return slide


def build_deck():
    total_slides = 14
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ==========================================================================
    # SLIDE 01: COVER SLIDE (Premium Dark AI Tech)
    # ==========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_rect(s1, 0, 0, 13.333, 7.5, C_DARK_BG)
    # Dual tech vertical accents
    add_rect(s1, 0, 0, 0.20, 7.5, C_CYAN)
    add_rect(s1, 0.20, 0, 0.10, 7.5, C_TEAL)

    # SIC Badge
    add_rect(s1, 0.85, 0.60, 5.60, 0.38, C_CARD_BG, border_rgb=C_TEAL, border_pt=1.5, rounded=True)
    add_text(s1, 0.90, 0.63, 5.50, 0.32, "SAMSUNG INNOVATION CAMPUS — ĐỒ ÁN TỐT NGHIỆP AI", font_size=11, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    # Main Project Title
    title_main = (
        "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY\n"
        "DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG & THAM GIA HỌC TẬP"
    )
    add_text(s1, 0.85, 1.18, 11.8, 1.45, title_main, font_size=26, bold=True, color=C_WHITE)

    # Subtitle
    sub_main = "Hệ thống Cảnh báo Sớm Học đường Thông minh — Dự báo Kiệt sức (Burnout) & Nguy cơ Bỏ học Sinh viên"
    add_text(s1, 0.85, 2.70, 11.5, 0.40, sub_main, font_size=14.5, italic=True, color=C_CYAN)

    # Team Card
    add_rect(s1, 0.85, 3.35, 11.6, 1.45, C_CARD_BG, border_rgb=C_CYAN, border_pt=1.5, rounded=True)
    add_rect(s1, 0.85, 3.35, 11.6, 0.36, C_CARD_INNER, rounded=False)
    add_text(s1, 1.05, 3.40, 11.2, 0.30, "NHÓM THỰC HIỆN: NHÓM 03 (SIC AI CAPSTONE)", font_size=12, bold=True, color=C_CYAN)
    team_body = (
        "• Đỗ Quang Minh (Trưởng nhóm) — Kiến trúc Pipeline Chống rò rỉ, Kỹ thuật Tạo Đặc trưng (6 chỉ số), Mô hình Cây & Backend FastAPI\n"
        "• Vũ Quang Huy (Thành viên) — Phân tích Khám phá Dữ liệu (EDA), Điền khuyết Đa biến KNN (k=7), Mô hình Cơ sở & Giao diện Web UI"
    )
    add_text(s1, 1.05, 3.82, 11.2, 0.88, team_body, font_size=12, color=C_WHITE)

    # 4 KPI Cards
    kpis = [
        ("2,000", "Hồ sơ Sinh viên\n(Sạch chuẩn 100%, 0 NaN)", C_TEAL),
        ("25 → 34", "Thuộc tính Dữ liệu\n(Bổ sung 6 Chỉ số Mới)", C_CYAN),
        ("4 Mô hình", "So sánh Chuyên sâu\n(LR, RF, XGB, Voting)", C_AMBER),
        ("89.01%", "Đỉnh ROC-AUC\n(5-Fold Stratified CV)", C_CORAL),
    ]
    for idx, (val, lbl, col) in enumerate(kpis):
        x = 0.85 + idx * 2.95
        add_rect(s1, x, 5.15, 2.75, 1.50, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_text(s1, x, 5.28, 2.75, 0.65, val, font_size=26, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s1, x + 0.10, 6.02, 2.55, 0.52, lbl, font_size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 02: AGENDA (6 Sections)
    # ==========================================================================
    s2 = create_master_slide(
        prs, 2, total_slides,
        "00 NỘI DUNG TRÌNH BÀY",
        "Cấu trúc Bài Thuyết trình Bảo vệ Đồ án",
        "Hành trình khép kín từ Đặt vấn đề, Làm sạch 6 bước, Kỹ thuật Đặc trưng, Đối sánh 4 Mô hình đến Triển khai Web UI"
    )
    agenda_cards = [
        ("01. GIỚI THIỆU & BỐI CẢNH", [
            "• Thực trạng bỏ học đại học & kiệt sức học đường.",
            "• Mục tiêu nghiên cứu & tính cấp thiết của Cảnh báo sớm.",
            "• Phân công nhiệm vụ Nhóm 03 & Lộ trình WBS 4 giai đoạn."
        ], C_CYAN),
        ("02. DỮ LIỆU & LÀM SẠCH 6 BƯỚC", [
            "• Khảo sát 800 SV gốc mở rộng lên 2,000 SV sạch (0 NaN).",
            "• Xử lý 18 cột khuyết bằng KNNImputer (k=7) + lọc IQR.",
            "• Chuẩn hóa 100% dữ liệu không rò rỉ thông tin (no leakage)."
        ], C_TEAL),
        ("03. KỸ THUẬT TẠO ĐẶC TRƯNG", [
            "• Quy trình học máy 5 giai đoạn khép kín từ thô đến suy luận.",
            "• Thiết kế 6 chỉ số chuyên ngành trong DomainFeatureEngineer.",
            "• Lượng hóa cán cân Căng thẳng vs. Gắn kết học tập."
        ], C_BLUE),
        ("04. PHÂN TÍCH KHÁM PHÁ (EDA)", [
            "• Tác động quyết định của mức độ kiệt sức lên bỏ học (92%).",
            "• Đối sánh phân phối 8 biến Tham gia học tập & Tâm lý.",
            "• Ma trận tương quan Pearson & Phân tích nhân khẩu học."
        ], C_AMBER),
        ("05. ĐỐI SÁNH 4 MÔ HÌNH HỌC MÁY", [
            "• Thực nghiệm Kịch bản 1 (Đầy đủ) & Kịch bản 2 (Cảnh báo sớm).",
            "• So sánh LR, Random Forest, XGBoost & Soft Voting Ensemble.",
            "• Đạt đỉnh ROC-AUC 89.01%, Test F1 83.63%, Test Acc 81.50%."
        ], C_CORAL),
        ("06. HỆ THỐNG WEB UI & KẾT LUẬN", [
            "• Kiến trúc phần mềm 3 tầng & Ứng dụng Web FastAPI (app.py).",
            "• Dự báo kép Kiệt sức + Bỏ học kèm gợi ý can thiệp (<30ms).",
            "• Tổng kết đóng góp, 4 hướng mở rộng (LSTM, SHAP, LMS) & Q&A."
        ], C_TEAL),
    ]
    for idx, (hdr, lines, col) in enumerate(agenda_cards):
        r_i, c_i = divmod(idx, 3)
        x = 0.40 + c_i * 4.22
        y = 1.40 + r_i * 2.72
        add_dark_card(s2, x, y, 4.05, 2.52, hdr, lines, border_color=col, header_color=col, body_font=11.2)

    # ==========================================================================
    # SLIDE 03: INTRODUCTION & CONCEPTUAL MECHANISM
    # ==========================================================================
    s3 = create_master_slide(
        prs, 3, total_slides,
        "01 GIỚI THIỆU ĐỀ TÀI",
        "Bối cảnh Bài toán & Cơ chế Hình thành Nguy cơ Bỏ học Đại học",
        "Bỏ học không phải sự kiện ngẫu nhiên đơn lẻ mà là hệ quả tích lũy từ suy giảm Gắn kết Học tập và Căng thẳng kéo dài"
    )
    add_dark_card(
        s3, 0.40, 1.40, 3.85, 2.65,
        "1. SỰ THAM GIA HỌC TẬP SUY GIẢM",
        [
            "• Tỷ lệ chuyên cần thấp (Attendance_Percent < 70%).",
            "• Giờ tự học mỗi ngày sụt giảm (Study_Hours_Per_Day).",
            "• Mất động lực học tập nội tại (Motivation_Score).",
            "• Điểm GPA kỳ trước thấp & tích lũy môn nợ (Backlogs ≥ 2).",
        ],
        border_color=C_TEAL, header_color=C_TEAL, body_font=11.0
    )
    add_dark_card(
        s3, 0.40, 4.25, 3.85, 2.65,
        "2. ÁP LỰC CĂNG THẲNG TÂM LÝ - XÃ HỘI",
        [
            "• Điểm căng thẳng (Stress_Level) & lo âu (Anxiety) cao.",
            "• Áp lực tài chính gia đình (Financial_Stress_Score).",
            "• Áp lực đồng trang lứa (Peer_Pressure_Score).",
            "• Thiếu ngủ (<6h), dùng màn hình nhiều (>6h) gây mệt mỏi.",
        ],
        border_color=C_CORAL, header_color=C_CORAL, body_font=11.0
    )
    add_dark_card(
        s3, 4.60, 2.25, 3.90, 3.70,
        "TRẠNG THÁI TRUNG GIAN: KIỆT SỨC (BURNOUT)",
        [
            "★ Hội chứng Kiệt sức Học đường (Burnout_Level):",
            "• Cầu nối trực tiếp giữa áp lực và quyết định bỏ học.",
            "• Phân tầng 3 cấp độ rõ rệt trong dữ liệu nghiên cứu:",
            "  - Mức Thấp (Low): 320 SV (40.0%) -> Tỷ lệ an toàn cao",
            "  - Mức Vừa (Medium): 280 SV (35.0%) -> Ngưỡng chuyển tiếp",
            "  - Mức Cao (High): 200 SV (25.0%) -> Nguy cơ bỏ học 92%!",
            "• Tương quan Pearson rất mạnh với bỏ học (r = +0.550).",
        ],
        border_color=C_AMBER, header_color=C_AMBER, body_font=11.0
    )
    add_dark_card(
        s3, 8.85, 1.40, 4.08, 5.50,
        "3. BIẾN MỤC TIÊU & GIẢI PHÁP AI NHÓM 03",
        [
            "★ Biến mục tiêu: Dropout_Risk (Yes / No)",
            "• Yes (Có nguy cơ bỏ học): 437 SV (54.62%)",
            "• No (An toàn / Duy trì): 363 SV (45.38%)",
            "",
            "★ Giải pháp AI Toàn diện của Nhóm 03:",
            "• Xây dựng hệ thống Học máy dự báo trên cả 2 kịch bản:",
            "  1. Kịch bản 1 (Chẩn đoán lâm sàng): Kết hợp Burnout_Level để đạt độ chính xác tối đa (ROC-AUC 89.01%).",
            "  2. Kịch bản 2 (Cảnh báo sớm): Ẩn Burnout_Level, phát hiện sớm từ chỉ số Căng thẳng & Học tập thô (ROC-AUC 88.26%).",
            "• Cung cấp Web Dashboard can thiệp sớm cho Cố vấn học tập.",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=11.0
    )

    # ==========================================================================
    # SLIDE 04: OBJECTIVES, TEAM 50/50 & TIMELINE
    # ==========================================================================
    s4 = create_master_slide(
        prs, 4, total_slides,
        "01 MỤC TIÊU & PHÂN CÔNG",
        "Mục tiêu Kỹ thuật, Phân công Nhóm 03 & Lộ trình Thực hiện",
        "Phân chia khối lượng công việc cân bằng tuyệt đối (50% - 50%), hoàn thành đúng hạn 100% các mốc bàn giao WBS"
    )
    add_dark_card(
        s4, 0.40, 1.40, 6.10, 3.35,
        "4 MỤC TIÊU CỐT LÕI CỦA ĐỒ ÁN",
        [
            "• [Mục tiêu 1 - EDA & Mở rộng]: Làm sạch 6 bước trên 800 SV gốc, mở rộng lên 2,000 SV sạch chuẩn (0 NaN); lượng hóa tác động của Burnout.",
            "• [Mục tiêu 2 - Tạo Đặc trưng]: Điền khuyết đa biến KNNImputer (k=7), lọc ngoại lai IQR và thiết kế 6 chỉ số tương tác Căng thẳng – Gắn kết.",
            "• [Mục tiêu 3 - So sánh 4 Mô hình]: Huấn luyện & kiểm định chéo 5-Fold trên 4 mô hình tiêu biểu (LR, RF, XGBoost, Soft Voting) ở 2 kịch bản.",
            "• [Mục tiêu 4 - Ứng dụng Web]: Triển khai ứng dụng FastAPI dự báo thời gian thực kép cả Burnout_Level và Dropout_Risk (<30ms).",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=11.2
    )
    add_dark_card(
        s4, 6.80, 1.40, 6.10, 1.60,
        "ĐỖ QUANG MINH — TRƯỞNG NHÓM (50% KHỐI LƯỢNG)",
        [
            "• Thiết kế kiến trúc Pipeline chống rò rỉ dữ liệu & lớp DomainFeatureEngineer (6 chỉ số).",
            "• Huấn luyện nhóm Cây tập hợp & Boosting (Random Forest, XGBoost, Soft Voting Ensemble).",
            "• Phát triển Backend FastAPI (app.py) và đóng gói mô hình suy luận (.pkl).",
        ],
        border_color=C_TEAL, header_color=C_TEAL, body_font=11.0
    )
    add_dark_card(
        s4, 6.80, 3.15, 6.10, 1.60,
        "VŨ QUANG HUY — THÀNH VIÊN (50% KHỐI LƯỢNG)",
        [
            "• Kiểm toán dữ liệu khuyết (18/25 cột), tiền xử lý EDA & xuất bản 12 biểu đồ chuẩn 200 DPI.",
            "• Huấn luyện nhóm mô hình Cơ sở (Logistic Regression, KNN, SVM, Decision Tree, MLP).",
            "• Thiết kế giao diện Frontend Web Dashboard, biên soạn WBS, Báo cáo SIC & Slide thuyết trình.",
        ],
        border_color=C_BLUE, header_color=C_BLUE, body_font=11.0
    )
    timeline = [
        ("GĐ 1 (19–20/10)", "Khởi tạo & Đặc tả Dữ liệu\n(Action Plan & WBS)", C_CYAN),
        ("GĐ 2 (21–25/10)", "Phân tích EDA & KNNImputer\n+ Tạo 6 Chỉ số Mới", C_TEAL),
        ("GĐ 3 (26–29/10)", "Huấn luyện & Kiểm định 5-Fold\n4 Mô hình × 2 Kịch bản", C_BLUE),
        ("GĐ 4 (30/10–05/11)", "Web Dashboard FastAPI\n& Hoàn thiện Báo cáo SIC", C_AMBER),
    ]
    for idx, (p_t, p_d, col) in enumerate(timeline):
        x = 0.40 + idx * 3.18
        add_rect(s4, x, 4.95, 2.98, 1.85, C_CARD_BG, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s4, x, 4.95, 2.98, 0.38, C_CARD_INNER, rounded=False)
        add_text(s4, x + 0.08, 4.98, 2.82, 0.30, p_t, font_size=11, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.10, 5.45, 2.78, 1.25, p_d + "\n[✓ Hoàn thành 100%]", font_size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 05: DATASET ARCHITECTURE & PREPROCESSING
    # ==========================================================================
    s5 = create_master_slide(
        prs, 5, total_slides,
        "02 DỮ LIỆU & TIỀN XỬ LÝ",
        "Kiến trúc Bộ Dữ liệu (800 → 2,000 SV) & Quy trình Điền khuyết Đa biến",
        "18/25 thuộc tính ban đầu có dữ liệu khuyết (1.6% – 6.4%), được xử lý triệt để bằng KNNImputer (k=7) trong Pipeline"
    )
    add_dark_card(
        s5, 0.40, 1.40, 4.90, 5.50,
        "5 NHÓM THUỘC TÍNH & CHIẾN LƯỢC TIỀN XỬ LÝ",
        [
            "• [1. Nhân khẩu học - 7 biến]: Student_ID, Age (17–25), Gender, Year_of_Study (1–4), Department (6 khoa), Residence_Type, Family_Income_Bracket.",
            "• [2. Tham gia học tập - 5 biến]: Attendance_Percent (30–100%), Study_Hours_Per_Day (0–8.1h), Previous_GPA, Backlogs (0–6), Motivation_Score.",
            "• [3. Lối sống - 5 biến]: Sleep_Hours, Screen_Time_Hours, Exercise_Freq, Social_Activity, Part_Time_Job.",
            "• [4. Căng thẳng & Hỗ trợ - 6 biến]: Stress_Level, Anxiety_Score, Financial_Stress, Peer_Pressure, Family_Support, Counseling_Access.",
            "• [5. Biến mục tiêu - 2 biến]: Burnout_Level (Low/Med/High) & Dropout_Risk (Yes: 437 | No: 363).",
            "★ Sơ đồ 6 bước làm sạch & mở rộng:",
            "  1. Khử trùng lặp ID -> 2. KNNImputer (k=7) số -> 3. Mode imputer phân loại -> 4. Cắt tỉa IQR bounds -> 5. Mở rộng đa biến 1,200 mẫu -> 6. Dataset sạch 2,000 SV (0 NaN).",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=10.8
    )
    add_figure_card(s5, os.path.join(FIG_VI, "01_missing_values_analysis.png"), 5.50, 1.40, 7.43, 5.50, border_color=C_CYAN)

    # ==========================================================================
    # SLIDE 06: METHODOLOGY & 6 ENGINEERED INDICES
    # ==========================================================================
    s6 = create_master_slide(
        prs, 6, total_slides,
        "02 PHƯƠNG PHÁP LUẬN",
        "Quy trình Học máy 5 Bước & Kỹ thuật Thiết kế 6 Đặc trưng Chuyên ngành",
        "Lớp biến đổi DomainFeatureEngineer định lượng hóa sự mất cân bằng giữa Căng thẳng và Gắn kết, nâng vượt trội khả năng phân tách"
    )
    add_figure_card(s6, os.path.join(FIG_VI, "11_system_workflow_diagram.png"), 0.40, 1.35, 12.53, 2.75, border_color=C_TEAL, pad=0.05)

    fe_cards = [
        ("1. Academic_Engagement_Index", "0.35*(Attendance/10) + 0.25*Study\n+ 0.25*Motivation + 0.15*GPA", "Điểm tổng hợp gắn kết học vụ (0-10)", C_TEAL),
        ("2. Psychological_Stress_Index", "0.35*Stress + 0.30*Anxiety\n+ 0.20*Financial + 0.15*Peer", "Điểm tổng hợp áp lực tâm lý đa chiều", C_CORAL),
        ("3. Stress_Engagement_Ratio", "Psychological_Stress_Index /\n(Academic_Engagement_Index + 1.0)", "Tỷ số Căng thẳng/Gắn kết (Top 2 Quan trọng)", C_CYAN),
        ("4. Net_Stress_Support_Balance", "Stress_Index - (0.5*Family\n+ 0.3*Social + 0.2*Exercise)", "Cán cân áp lực ròng sau khi bù trừ hỗ trợ", C_BLUE),
        ("5. Digital_Fatigue_Ratio", "Screen_Time_Hours /\n(Sleep_Hours + 1.0)", "Tỷ số mệt mỏi số hóa (màn hình vs giấc ngủ)", C_AMBER),
        ("6. Academic_Risk_Burden", "1.5 * Backlogs +\n(100 - Attendance) / 20.0", "Gánh nặng rủi ro từ môn nợ & vắng học", C_TEAL),
    ]
    for idx, (f_name, f_eq, f_desc, col) in enumerate(fe_cards):
        r_i, c_i = divmod(idx, 3)
        x = 0.40 + c_i * 4.22
        y = 4.25 + r_i * 1.35
        add_rect(s6, x, y, 4.05, 1.25, C_CARD_BG, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s6, x, y, 4.05, 0.32, C_CARD_INNER, rounded=False)
        add_text(s6, x + 0.10, y + 0.03, 3.85, 0.26, f_name, font_size=10.5, bold=True, color=col)
        add_text(s6, x + 0.10, y + 0.36, 3.85, 0.50, f_eq, font_size=10.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s6, x + 0.10, y + 0.90, 3.85, 0.28, f_desc, font_size=9.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 07: EDA 1 - TARGET & BURNOUT PROGRESSION
    # ==========================================================================
    s7 = create_master_slide(
        prs, 7, total_slides,
        "03 PHÂN TÍCH KHÁM PHÁ (EDA)",
        "Phát hiện EDA 1: Tác động Quyết định của Burnout lên Nguy cơ Bỏ học",
        "Tỷ lệ bỏ học tăng gấp gần 4 lần từ mức Kiệt sức Thấp (23.8%) lên Kiệt sức Cao (92.0%, Pearson r = +0.550)"
    )
    burn_kpis = [
        ("KIỆT SỨC THẤP - LOW (n=320)", "23.8% Bỏ học", "244 SV An toàn vs. 76 SV Bỏ học\nNhóm duy trì nhịp học cân bằng tốt", C_TEAL),
        ("KIỆT SỨC VỪA - MEDIUM (n=280)", "63.2% Bỏ học", "103 SV An toàn vs. 177 SV Bỏ học\nNgưỡng chuyển tiếp nguy cấp cần can thiệp", C_AMBER),
        ("KIỆT SỨC CAO - HIGH (n=200)", "92.0% Bỏ học", "Chỉ 16 SV An toàn vs. 184 SV Bỏ học!\nBáo động đỏ: Nguy cơ bỏ học gần như tuyệt đối", C_CORAL),
    ]
    for idx, (b_hdr, b_big, b_sub, col) in enumerate(burn_kpis):
        x = 0.40 + idx * 4.22
        add_rect(s7, x, 1.35, 4.05, 1.55, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s7, x, 1.35, 4.05, 0.34, C_CARD_INNER, rounded=False)
        add_text(s7, x + 0.10, 1.38, 3.85, 0.28, b_hdr, font_size=10.5, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s7, x + 0.10, 1.72, 3.85, 0.48, b_big, font_size=20, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s7, x + 0.10, 2.22, 3.85, 0.62, b_sub, font_size=10, color=C_MUTED, align=PP_ALIGN.CENTER)

    add_figure_card(s7, os.path.join(FIG_VI, "02_target_and_burnout_distribution.png"), 0.40, 3.05, 12.53, 3.85, border_color=C_CYAN, pad=0.06)

    # ==========================================================================
    # SLIDE 08: EDA 2 - STUDY ENGAGEMENT VS. STRESS DISTRIBUTIONS
    # ==========================================================================
    s8 = create_master_slide(
        prs, 8, total_slides,
        "03 PHÂN TÍCH KHÁM PHÁ (EDA)",
        "Phát hiện EDA 2: Đối sánh Phân phối Tham gia Học tập & Căng thẳng Tâm lý",
        "Sinh viên bỏ học (Yes) bị hụt 7% chuyên cần (73.7% vs 80.7%), tự học ít hơn, động lực thấp hơn và căng thẳng/lo âu cao hơn rõ rệt"
    )
    add_figure_card(s8, os.path.join(FIG_VI, "03_stress_vs_engagement_distributions.png"), 0.40, 1.38, 9.15, 5.50, border_color=C_BLUE, pad=0.06)

    add_dark_card(
        s8, 9.72, 1.38, 3.21, 5.50,
        "CHÊNH LỆCH TRUNG BÌNH (μ)",
        [
            "★ Nhóm Gắn kết Học tập:",
            "• Attendance_Percent:",
            "  No: 80.68% | Yes: 73.70%",
            "• Study_Hours_Per_Day:",
            "  No: 3.43h  | Yes: 2.93h",
            "• Motivation_Score:",
            "  No: 5.90   | Yes: 5.11",
            "• Previous_GPA:",
            "  No: 7.02   | Yes: 6.72",
            "",
            "★ Nhóm Căng thẳng Tâm lý:",
            "• Stress_Level:",
            "  No: 5.17   | Yes: 5.81",
            "• Anxiety_Score:",
            "  No: 4.65   | Yes: 5.26",
            "• Financial_Stress:",
            "  No: 4.75   | Yes: 5.26",
            "• Family_Support:",
            "  No: 6.73   | Yes: 6.16",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=10.5
    )

    # ==========================================================================
    # SLIDE 09: EDA 3 - CORRELATION MATRIX & PROTECTIVE VS RISK FACTORS
    # ==========================================================================
    s9 = create_master_slide(
        prs, 9, total_slides,
        "03 PHÂN TÍCH KHÁM PHÁ (EDA)",
        "Phát hiện EDA 3: Ma trận Tương quan Pearson & Các Yếu tố Bảo vệ / Nguy cơ",
        "Xác định rõ ràng 5 yếu tố bảo vệ (tương quan âm) và 5 yếu tố rủi ro (tương quan dương) tác động lên cả Burnout và Dropout"
    )
    add_figure_card(s9, os.path.join(FIG_VI, "04_correlation_matrix_heatmap.png"), 0.40, 1.38, 7.35, 5.50, border_color=C_CYAN, pad=0.06)

    add_dark_card(
        s9, 7.95, 1.38, 4.98, 2.68,
        "5 YẾU TỐ BẢO VỆ (GIẢM NGUY CƠ BỎ HỌC)",
        [
            "• 1. Attendance_Percent: r = -0.244 (với Dropout) | -0.325 (với Burnout)",
            "• 2. Motivation_Score: r = -0.193 (với Dropout) | -0.249 (với Burnout)",
            "• 3. Study_Hours_Per_Day: r = -0.160 (với Dropout) | -0.257 (với Burnout)",
            "• 4. Family_Support_Score: r = -0.137 (với Dropout) | -0.205 (với Burnout)",
            "• 5. Previous_GPA & Exercise: r = -0.106 (Cải thiện sức bền học tập)",
        ],
        border_color=C_TEAL, header_color=C_TEAL, body_font=10.8
    )
    add_dark_card(
        s9, 7.95, 4.20, 4.98, 2.68,
        "5 YẾU TỐ NGUY CƠ (TĂNG NGUY CƠ BỎ HỌC)",
        [
            "• 1. Burnout_Numeric: r = +0.550 (Yếu tố trung gian mạnh nhất)",
            "• 2. Stress_Level: r = +0.164 (với Dropout) | +0.270 (với Burnout)",
            "• 3. Anxiety_Score: r = +0.145 (với Dropout) | +0.247 (với Burnout)",
            "• 4. Screen_Time_Hours: r = +0.127 (với Dropout) | +0.212 (với Burnout)",
            "• 5. Financial_Stress & Low Income: Nhóm thu nhập thấp bỏ học tới 62.1%",
        ],
        border_color=C_CORAL, header_color=C_CORAL, body_font=10.8
    )

    # ==========================================================================
    # SLIDE 10: MODEL BENCHMARKING VISUALIZATION
    # ==========================================================================
    s10 = create_master_slide(
        prs, 10, total_slides,
        "04 ĐỐI SÁNH 4 MÔ HÌNH ML",
        "So sánh Hiệu năng 4 Mô hình Học máy Tiêu biểu trên 2 Kịch bản Thực nghiệm",
        "Đánh giá bằng 5-Fold Stratified CV (2,000 SV): Random Forest dẫn đầu Kịch bản 1; XGBoost & Soft Voting dẫn đầu Kịch bản 2"
    )
    add_figure_card(s10, os.path.join(FIG_VI, "07_model_performance_comparison.png"), 0.40, 1.36, 12.53, 4.12, border_color=C_CYAN, pad=0.05)

    add_dark_card(
        s10, 0.40, 5.62, 6.15, 1.32,
        "KỊCH BẢN 1: ĐẦY ĐỦ ĐẶC TRƯNG (CÓ BURNOUT_LEVEL)",
        [
            "• Random Forest dẫn đầu toàn diện: CV ROC-AUC 89.01%, Test Acc 81.50%, Test F1 83.63%.",
            "• Soft Voting Ensemble (LR+RF+XGB) bám sát với CV ROC-AUC 88.89%, Test Acc 80.25%.",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=10.5
    )
    add_dark_card(
        s10, 6.78, 5.62, 6.15, 1.32,
        "KỊCH BẢN 2: CẢNH BÁO SỚM (ẨN BURNOUT_LEVEL)",
        [
            "• XGBoost dẫn đầu CV: Acc 79.95%, F1 82.40%, ROC-AUC 88.26%.",
            "• Soft Voting Ensemble đạt CV ROC-AUC 88.25% và Test ROC-AUC 87.39%.",
        ],
        border_color=C_TEAL, header_color=C_TEAL, body_font=10.5
    )

    # ==========================================================================
    # SLIDE 11: QUANTITATIVE COMPARISON MATRIX (TABLE)
    # ==========================================================================
    s11 = create_master_slide(
        prs, 11, total_slides,
        "04 BẢNG SỐ LIỆU THỰC NGHIỆM",
        "Bảng Tổng hợp Chỉ số Thực nghiệm 4 Mô hình Học máy Tiêu biểu",
        "Độ lệch giữa Kiểm định chéo 5-Fold (CV) và Tập kiểm thử độc lập (Test 20% = 400 SV) rất nhỏ, kiểm soát tốt Overfitting"
    )
    table_data = [
        [
            "Hạng & Mô hình Học máy",
            "KB1: CV Acc",
            "KB1: CV Recall",
            "KB1: CV F1",
            "KB1: CV AUC",
            "KB1: Test Acc",
            "KB1: Test F1",
            "KB1: Test AUC",
            "KB2: CV AUC (Sớm)",
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
                cell.fill.fore_color.rgb = RGBColor(20, 35, 58) # Deep navy header
            elif r_i == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(24, 45, 75) # Top 1 highlighted row
            elif r_i in [2, 3]:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(17, 30, 50)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(14, 24, 40)

            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.CENTER
                p.font.name = FONT_FAMILY
                p.font.size = Pt(10.5)
                p.font.bold = (r_i in [0, 1, 2, 3] or c_i == 0)
                if r_i == 0:
                    p.font.color.rgb = C_CYAN
                elif r_i == 1:
                    p.font.color.rgb = C_WHITE if c_i == 0 else C_TEAL
                else:
                    p.font.color.rgb = C_WHITE

    # ==========================================================================
    # SLIDE 12: ROC/PR CURVES & FEATURE IMPORTANCE
    # ==========================================================================
    s12 = create_master_slide(
        prs, 12, total_slides,
        "04 GIẢI THÍCH MÔ HÌNH",
        "Đường cong ROC / Precision-Recall & Xếp hạng Độ quan trọng Đặc trưng",
        "Các chỉ số tương tác Stress_Engagement_Ratio và Academic_Engagement_Index đứng đầu bảng xếp hạng ngay sau Burnout_Ordinal"
    )
    add_figure_card(s12, os.path.join(FIG_VI, "08_roc_and_precision_recall_curves.png"), 0.40, 1.36, 6.15, 4.10, border_color=C_CYAN, pad=0.05)
    add_figure_card(s12, os.path.join(FIG_VI, "10_feature_importance_analysis.png"), 6.78, 1.36, 6.15, 4.10, border_color=C_TEAL, pad=0.05)

    add_dark_card(
        s12, 0.40, 5.60, 12.53, 1.35,
        "Ý NGHĨA HỌC THUẬT & THỰC TIỄN TỪ PHÂN TÍCH ĐỘ QUAN TRỌNG ĐẶC TRƯNG",
        [
            "• Top 6 Đặc trưng Quyết định Nguy cơ Bỏ học: (1) Burnout_Ordinal, (2) Stress_Engagement_Ratio, (3) Academic_Engagement_Index, (4) Psychological_Stress_Index, (5) Attendance_Percent, (6) Net_Stress_Support_Balance.",
            "• Độ nhạy cao (Recall = 80.3% – 86.4%): Đảm bảo hệ thống phát hiện được hơn 4/5 số sinh viên thực sự có nguy cơ bỏ học để can thiệp kịp thời.",
            "• Khuyến cáo Học thuật: Feature Importance chỉ phản ánh mức đóng góp giảm độ vấy bẩn (impurity reduction/gain), không đồng nhất với quan hệ nhân quả.",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=10.5
    )

    # ==========================================================================
    # SLIDE 13: SYSTEM ARCHITECTURE & FASTAPI WEB DASHBOARD UI
    # ==========================================================================
    s13 = create_master_slide(
        prs, 13, total_slides,
        "05 ỨNG DỤNG WEB THỰC TẾ",
        "Kiến trúc Hệ thống 3 Tầng & Giao diện Web Dashboard Cảnh báo Sớm",
        "Ứng dụng Web (app.py) cho phép nhập 18 chỉ số hoặc nạp nhanh 3 mẫu hồ sơ, trả về suy luận kép và gợi ý can thiệp trong < 30ms"
    )
    add_figure_card(s13, os.path.join(FIG_VI, "12_system_architecture_diagram.png"), 0.40, 1.38, 7.85, 5.50, border_color=C_TEAL, pad=0.06)

    add_dark_card(
        s13, 8.45, 1.38, 4.48, 5.50,
        "TÍNH NĂNG GIAO DIỆN WEB (APP.PY)",
        [
            "★ 1. Ba Mẫu Hồ sơ Kiểm thử Nhanh (Presets):",
            "  • [An toàn] SV Tích cực: Nguy cơ Bỏ học 8.0%",
            "  • [Cảnh báo] SV Căng thẳng TB: Nguy cơ ~58%",
            "  • [Báo động] SV Kiệt sức Cao: Nguy cơ 95.8%!",
            "",
            "★ 2. Dự báo Kép Thời gian thực (<30ms):",
            "  • Dự báo tự động Burnout_Level (3 mức).",
            "  • Dự báo xác suất Dropout_Risk định chuẩn (%).",
            "",
            "★ 3. Giải thích AI & Khuyến nghị Can thiệp:",
            "  • Hiển thị cán cân Căng thẳng vs. Gắn kết.",
            "  • Đề xuất can thiệp cá nhân hóa: Lớp phụ đạo học tập, tham vấn tâm lý, hỗ trợ tài chính.",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=11.0
    )

    # ==========================================================================
    # SLIDE 14: CONCLUSION, FUTURE ROADMAP & Q&A
    # ==========================================================================
    s14 = create_master_slide(
        prs, 14, total_slides,
        "06 KẾT LUẬN & HỎI ĐÁP",
        "Tổng kết Đóng góp, Định hướng Phát triển Tương lai & Thảo luận (Q&A)",
        "Đồ án hoàn thành xuất sắc các mục tiêu nghiên cứu kỹ thuật — Sẵn sàng cho tích hợp hệ thống quản trị học tập đại học (LMS)"
    )
    add_dark_card(
        s14, 0.40, 1.40, 6.15, 3.85,
        "4 TỔNG KẾT ĐÓNG GÓP CỐT LÕI",
        [
            "• [1. Tiền xử lý 6 bước & Mở rộng Dữ liệu]: Từ 800 bản ghi gốc mở rộng lên 2,000 bản ghi sạch 100% (0 NaN); lượng hóa bước nhảy bỏ học theo Burnout (23.8% → 63.2% → 92.0%).",
            "• [2. Kỹ thuật Tạo Đặc trưng Chuyên ngành]: Thiết kế 6 chỉ số chuyên sâu giúp mô hình nắm bắt chính xác cán cân Căng thẳng – Gắn kết học tập.",
            "• [3. Hệ thống 4 Mô hình Học máy Tiêu biểu]: Random Forest, XGBoost & Soft Voting đạt CV ROC-AUC 89.01%, Test F1 83.63%, Test Acc 81.50%.",
            "• [4. Sản phẩm Thực tiễn Hoàn chỉnh]: Ứng dụng Web FastAPI hoàn chỉnh hỗ trợ cố vấn học đường sàng lọc sớm trong <30ms.",
        ],
        border_color=C_TEAL, header_color=C_TEAL, body_font=11.2
    )
    add_dark_card(
        s14, 6.78, 1.40, 6.15, 3.85,
        "4 ĐỊNH HƯỚNG PHÁT TRIỂN TƯƠNG LAI",
        [
            "• [Hướng 1 - Dữ liệu Chuỗi thời gian]: Thu thập chuyên cần và căng thẳng theo từng tuần học để huấn luyện mô hình tuần tự LSTM / Transformer.",
            "• [Hướng 2 - Giải thích Cục bộ SHAP]: Tích hợp biểu đồ SHAP Waterfall trực tiếp lên Web UI giải thích nguyên nhân cho từng sinh viên cụ thể.",
            "• [Hướng 3 - Kết nối LMS / Moodle API]: Tự động hóa quy trình sàng lọc định kỳ hàng tháng cho Phòng Công tác Sinh viên.",
            "• [Hướng 4 - Mở rộng Đa trường]: Thu thập và kiểm định mô hình trên 5,000–10,000 mẫu tại các trường đại học Việt Nam.",
        ],
        border_color=C_CYAN, header_color=C_CYAN, body_font=11.2
    )

    # Q&A Highlight Banner
    add_rect(s14, 0.40, 5.45, 12.53, 1.50, C_CARD_BG, border_rgb=C_AMBER, border_pt=2.0, rounded=True)
    add_rect(s14, 0.40, 5.45, 12.53, 0.36, C_CARD_INNER, rounded=False)
    add_text(
        s14, 0.60, 5.48, 12.13, 0.30,
        "CẢM ƠN QUÝ THẦY CÔ VÀ HỘI ĐỒNG CHẤM ĐỒ ÁN ĐÃ LẮNG NGHE! — THẢO LUẬN & HỎI ĐÁP",
        font_size=12, bold=True, color=C_AMBER, align=PP_ALIGN.CENTER
    )
    add_text(
        s14, 0.60, 5.92, 12.13, 0.45,
        "Nhóm 03: Đỗ Quang Minh (Trưởng nhóm) • Vũ Quang Huy (Thành viên)",
        font_size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER
    )
    add_text(
        s14, 0.60, 6.42, 12.13, 0.35,
        "Samsung Innovation Campus (SIC) AI Capstone Project  |  Web App: http://127.0.0.1:8000",
        font_size=12, italic=True, color=C_CYAN, align=PP_ALIGN.CENTER
    )

    prs.save(PPTX_OUT)
    print(f"Master Final Dark PPTX saved successfully -> {PPTX_OUT}")


if __name__ == "__main__":
    build_deck()

