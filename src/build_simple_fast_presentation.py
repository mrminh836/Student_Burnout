"""
Builds a streamlined, minimalist, fast-paced 10-Slide PowerPoint presentation
designed specifically for a 5 to 8 minute capstone defense.
- Fewer slides (10 core slides instead of 14 dense ones)
- Short, punchy bullet points (max 3-4 short lines per card)
- Larger typography (14-16pt body text, 22-28pt titles/KPIs)
- Preserves 100% of core scientific numbers and visual charts
"""

import os
import shutil
from PIL import Image as PilImage
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG_VI = os.path.join(BASE_DIR, "figures_vi")
FINAL_DIR = os.path.join(BASE_DIR, "FINAL_PRESENTATION")
VI_DIR = os.path.join(BASE_DIR, "SIC_Documents_VI")
DRV_DIR = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"

os.makedirs(FINAL_DIR, exist_ok=True)
os.makedirs(VI_DIR, exist_ok=True)

PPTX_FINAL = os.path.join(FINAL_DIR, "SIC_AI_Capstone_Final.pptx")
PPTX_VI = os.path.join(VI_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx")

# Palette: Premium Dark Navy AI Tech (Clean, High-Contrast, Airy)
C_DARK_BG = RGBColor(11, 18, 32)       # #0B1220 Canvas
C_CARD_BG = RGBColor(17, 28, 46)       # #111C2E Card surface
C_CARD_HEADER = RGBColor(23, 38, 62)   # #17263E Card header
C_BORDER = RGBColor(36, 52, 77)        # #24344D Subtle border
C_CYAN = RGBColor(0, 212, 255)         # #00D4FF Primary accent
C_TEAL = RGBColor(0, 229, 176)         # #00E5B0 Secondary accent
C_AMBER = RGBColor(255, 183, 3)        # #FFB703 Highlight gold
C_CORAL = RGBColor(255, 107, 107)      # #FF6B6B Alert red
C_BLUE = RGBColor(56, 189, 248)        # #38BDF8 Sky blue
C_WHITE = RGBColor(248, 250, 252)      # #F8FAFC Pure white text
C_MUTED = RGBColor(175, 190, 210)      # #AFBED2 Secondary text

FONT_NAME = "Segoe UI"


def add_rect(slide, left, top, width, height, fill_rgb, border_rgb=None, border_pt=1.2, rounded=False):
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


def add_text(slide, left, top, width, height, text, font_size=14, bold=False, italic=False, color=C_WHITE, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.06)
    tf.margin_right = Inches(0.06)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT_NAME
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def add_simple_card(slide, left, top, width, height, title, bullets, accent_color=C_CYAN, body_size=14.0):
    """Minimalist card with large, easy-to-scan bullet points for fast speaking."""
    add_rect(slide, left, top, width, height, C_CARD_BG, border_rgb=accent_color, border_pt=1.5, rounded=True)
    add_rect(slide, left, top, width, 0.46, C_CARD_HEADER, rounded=False)
    add_rect(slide, left, top + 0.44, width, 0.03, accent_color, rounded=False)
    add_text(slide, left + 0.14, top + 0.07, width - 0.28, 0.34, title, font_size=13.5, bold=True, color=accent_color)

    tb = slide.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.56), Inches(width - 0.30), Inches(height - 0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(7)
        r = p.add_run()
        r.text = line
        r.font.name = FONT_NAME
        r.font.size = Pt(body_size)
        if line.startswith("★") or "89.01%" in line or "88.26%" in line or "92.0%" in line:
            r.font.bold = True
            r.font.color.rgb = C_AMBER if "★" in line else C_CYAN
        elif line.startswith("→"):
            r.font.bold = True
            r.font.color.rgb = C_TEAL
        else:
            r.font.color.rgb = C_WHITE


def add_figure_card(slide, img_path, box_left, box_top, box_width, box_height, border_color=C_CYAN, pad=0.08):
    add_rect(slide, box_left, box_top, box_width, box_height, RGBColor(255, 255, 255), border_rgb=border_color, border_pt=1.5, rounded=True)
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


def create_header_footer(prs, slide_num, total_slides, badge, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, 13.333, 7.5, C_DARK_BG)
    # Top bar
    add_rect(slide, 0, 0, 13.333, 1.15, C_CARD_BG)
    add_rect(slide, 0, 1.13, 13.333, 0.03, C_CYAN)
    add_rect(slide, 0, 0, 0.12, 1.15, C_TEAL)

    # Pill badge
    add_rect(slide, 0.36, 0.12, 2.90, 0.26, C_CARD_HEADER, border_rgb=C_TEAL, border_pt=1.0, rounded=True)
    add_text(slide, 0.38, 0.13, 2.86, 0.24, badge, font_size=9.5, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    # Title & short takeaway
    add_text(slide, 0.36, 0.40, 12.6, 0.40, title, font_size=21, bold=True, color=C_WHITE)
    add_text(slide, 0.36, 0.79, 12.6, 0.28, subtitle, font_size=12, italic=True, color=C_CYAN)

    # Footer
    add_rect(slide, 0, 7.12, 13.333, 0.38, C_CARD_BG)
    add_text(slide, 0.38, 7.17, 10.0, 0.28, "SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE  |  NHÓM 03: ĐỖ QUANG MINH • VŨ QUANG HUY", font_size=9.5, bold=True, color=C_MUTED)
    add_text(slide, 11.2, 7.17, 1.8, 0.28, f"Slide {slide_num:02d} / {total_slides:02d}", font_size=10, bold=True, color=C_AMBER, align=PP_ALIGN.RIGHT)
    return slide


def build_streamlined_deck():
    total_slides = 10
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # ==========================================================================
    # SLIDE 01: TRANG BÌA (30 giây)
    # ==========================================================================
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s1, 0, 0, 13.333, 7.5, C_DARK_BG)
    add_rect(s1, 0, 0, 0.20, 7.5, C_CYAN)
    add_rect(s1, 0.20, 0, 0.10, 7.5, C_TEAL)

    add_rect(s1, 0.85, 0.65, 5.20, 0.40, C_CARD_BG, border_rgb=C_TEAL, border_pt=1.5, rounded=True)
    add_text(s1, 0.90, 0.69, 5.10, 0.32, "SAMSUNG INNOVATION CAMPUS — AI CAPSTONE", font_size=11.5, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    title_main = (
        "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY\n"
        "DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG & THAM GIA HỌC TẬP"
    )
    add_text(s1, 0.85, 1.25, 11.8, 1.45, title_main, font_size=26, bold=True, color=C_WHITE)
    add_text(s1, 0.85, 2.72, 11.5, 0.40, "Hệ thống Cảnh báo Sớm Học đường — Phân tích Kiệt sức (Burnout) & Nguy cơ Bỏ học", font_size=15, italic=True, color=C_CYAN)

    # Simple Team Box
    add_rect(s1, 0.85, 3.38, 11.6, 1.35, C_CARD_BG, border_rgb=C_CYAN, border_pt=1.5, rounded=True)
    add_text(s1, 1.10, 3.48, 11.1, 0.35, "NHÓM THỰC HIỆN (NHÓM 03):", font_size=13, bold=True, color=C_AMBER)
    add_text(
        s1, 1.10, 3.88, 11.1, 0.75,
        "• Đỗ Quang Minh (Trưởng nhóm) — Pipeline Học máy, Tạo Đặc trưng, Mô hình Cây & FastAPI Web\n"
        "• Vũ Quang Huy (Thành viên) — Phân tích Khám phá (EDA), Điền khuyết KNN, Mô hình Cơ sở & Giao diện UI",
        font_size=13.5, color=C_WHITE
    )

    kpis = [
        ("2,000", "Hồ sơ Sinh viên Sạch", C_TEAL),
        ("6 Chỉ số", "Đặc trưng Mới (25 → 34 cột)", C_CYAN),
        ("4 Mô hình", "LR • RF • XGB • Voting", C_AMBER),
        ("89.01%", "Đỉnh ROC-AUC (5-Fold CV)", C_CORAL),
    ]
    for idx, (val, lbl, col) in enumerate(kpis):
        x = 0.85 + idx * 2.95
        add_rect(s1, x, 5.10, 2.75, 1.55, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_text(s1, x, 5.25, 2.75, 0.65, val, font_size=27, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s1, x + 0.10, 6.00, 2.55, 0.50, lbl, font_size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 02: BỐI CẢNH, CƠ CHẾ & MỤC TIÊU (45 giây)
    # ==========================================================================
    s2 = create_header_footer(
        prs, 2, total_slides,
        "01 ĐẶT VẤN ĐỀ & MỤC TIÊU",
        "Tại Sao Sinh Viên Bỏ Học & Mục Tiêu Giải Pháp AI",
        "Bỏ học là hệ quả tích lũy: Gắn kết học tập giảm + Căng thẳng tăng → Kiệt sức (Burnout) → Bỏ học (Dropout)"
    )
    add_simple_card(
        s2, 0.40, 1.40, 4.05, 5.45,
        "1. NGUYÊN NHÂN GỐC RỄ",
        [
            "★ Suy giảm Học tập:",
            "• Đi học muộn / vắng học (< 70%).",
            "• Tự học ít, giảm động lực, nợ môn.",
            "",
            "★ Áp lực Tâm lý & Lối sống:",
            "• Căng thẳng, lo âu, áp lực tài chính.",
            "• Thiếu ngủ (< 6h), dùng màn hình nhiều.",
        ],
        accent_color=C_TEAL, body_size=14.0
    )
    add_simple_card(
        s2, 4.64, 1.40, 4.05, 5.45,
        "2. CẦU NỐI: KIỆT SỨC (BURNOUT)",
        [
            "★ 3 Cấp độ Kiệt sức Học đường:",
            "• Mức Thấp (Low): 40% sinh viên.",
            "• Mức Vừa (Medium): 35% sinh viên.",
            "• Mức Cao (High): 25% sinh viên.",
            "",
            "→ Tương quan mạnh nhất với quyết định bỏ học (r = +0.550).",
        ],
        accent_color=C_AMBER, body_size=14.0
    )
    add_simple_card(
        s2, 8.88, 1.40, 4.05, 5.45,
        "3. MỤC TIÊU ĐỒ ÁN (NHÓM 03)",
        [
            "★ Dự báo: Dropout_Risk (Yes 54.6% / No 45.4%)",
            "• Làm sạch 800 SV → Mở rộng 2,000 SV.",
            "• Tạo 6 chỉ số đo Căng thẳng vs Gắn kết.",
            "• So sánh 4 mô hình trên 2 kịch bản:",
            "  1. Đầy đủ (Có Burnout): AUC 89.01%",
            "  2. Cảnh báo sớm (Ẩn Burnout): 88.26%",
            "• Xây dựng Web App dự báo < 30ms.",
        ],
        accent_color=C_CYAN, body_size=13.5
    )

    # ==========================================================================
    # SLIDE 03: DỮ LIỆU & LÀM SẠCH 6 BƯỚC (45 giây)
    # ==========================================================================
    s3 = create_header_footer(
        prs, 3, total_slides,
        "02 DỮ LIỆU & TIỀN XỬ LÝ",
        "Bộ Dữ Liệu (800 → 2,000 SV) & Quy Trình Làm Sạch 6 Bước",
        "Xử lý 18 cột khuyết bằng KNNImputer (k=7) trong Pipeline — Đảm bảo 100% sạch (0 NaN) và không rò rỉ dữ liệu"
    )
    add_simple_card(
        s3, 0.40, 1.40, 5.10, 5.45,
        "QUY TRÌNH 6 BƯỚC CHUẨN HÓA DỮ LIỆU",
        [
            "★ Dữ liệu gốc: 800 SV × 25 cột (612 ô bị trống).",
            "• Bước 1: Kiểm tra & xóa trùng lặp Student_ID.",
            "• Bước 2: Điền khuyết số bằng KNNImputer (k=7).",
            "• Bước 3: Điền khuyết chữ bằng giá trị phổ biến (Mode).",
            "• Bước 4: Cắt ngoại lai (IQR) & chuẩn hóa kiểu số.",
            "• Bước 5: Tạo 6 chỉ số chuyên ngành + Mã hóa số.",
            "→ Bước 6: Mở rộng đa biến lên 2,000 SV × 34 cột (Sạch 100%, 0 lỗi NaN).",
        ],
        accent_color=C_CYAN, body_size=13.2
    )
    add_figure_card(s3, os.path.join(FIG_VI, "01_missing_values_analysis.png"), 5.70, 1.40, 7.23, 5.45, border_color=C_CYAN)

    # ==========================================================================
    # SLIDE 04: QUY TRÌNH ML & 6 ĐẶC TRƯNG MỚI (45 giây)
    # ==========================================================================
    s4 = create_header_footer(
        prs, 4, total_slides,
        "02 KỸ THUẬT ĐẶC TRƯNG",
        "Quy Trình Học Máy 5 Bước & 6 Chỉ Số Chuyên Ngành Mới",
        "Thiết kế bộ biến đổi DomainFeatureEngineer giúp mô hình nắm bắt cán cân Căng thẳng vs. Gắn kết học tập"
    )
    add_figure_card(s4, os.path.join(FIG_VI, "11_system_workflow_diagram.png"), 0.40, 1.35, 12.53, 2.65, border_color=C_TEAL, pad=0.05)

    fe_short = [
        ("1. Chỉ số Gắn kết Học tập", "Chuyên cần + Giờ học + Động lực + GPA", C_TEAL),
        ("2. Chỉ số Căng thẳng Tâm lý", "Stress + Lo âu + Tài chính + Bạn bè", C_CORAL),
        ("3. Tỷ số Căng thẳng / Gắn kết ★", "Căng thẳng / (Gắn kết + 1) → Top 2 Quan trọng", C_CYAN),
        ("4. Cán cân Căng thẳng Ròng", "Căng thẳng − (Gia đình + Xã hội + Thể dục)", C_BLUE),
        ("5. Tỷ số Mệt mỏi Số hóa", "Giờ dùng màn hình / (Giờ ngủ + 1)", C_AMBER),
        ("6. Gánh nặng Rủi ro Học vụ", "1.5 × Số môn nợ + Độ hụt chuyên cần", C_TEAL),
    ]
    for idx, (f_title, f_desc, col) in enumerate(fe_short):
        r_i, c_i = divmod(idx, 3)
        x = 0.40 + c_i * 4.22
        y = 4.18 + r_i * 1.40
        add_rect(s4, x, y, 4.05, 1.25, C_CARD_BG, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s4, x, y, 4.05, 0.36, C_CARD_HEADER, rounded=False)
        add_text(s4, x + 0.12, y + 0.04, 3.80, 0.28, f_title, font_size=12, bold=True, color=col)
        add_text(s4, x + 0.12, y + 0.48, 3.80, 0.65, f_desc, font_size=12.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 05: EDA 1 - TÁC ĐỘNG CỦA BURNOUT (40 giây)
    # ==========================================================================
    s5 = create_header_footer(
        prs, 5, total_slides,
        "03 PHÂN TÍCH DỮ LIỆU (EDA)",
        "Phát Hiện EDA 1: Kiệt Sức Tăng Làm Bỏ Học Tăng Gấp 4 Lần",
        "Khi sinh viên rơi vào trạng thái Kiệt sức Cao (High Burnout), tỷ lệ bỏ học vọt lên tới 92.0%"
    )
    burn_cards = [
        ("KIỆT SỨC THẤP (LOW)", "23.8% Bỏ học", "Trạng thái cân bằng tốt", C_TEAL),
        ("KIỆT SỨC VỪA (MEDIUM)", "63.2% Bỏ học", "Ngưỡng cảnh báo cần can thiệp", C_AMBER),
        ("KIỆT SỨC CAO (HIGH)", "92.0% Bỏ học!", "Báo động đỏ: Nguy cơ cực cao", C_CORAL),
    ]
    for idx, (hdr, big, sub, col) in enumerate(burn_cards):
        x = 0.40 + idx * 4.22
        add_rect(s5, x, 1.35, 4.05, 1.50, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s5, x, 1.35, 4.05, 0.36, C_CARD_HEADER, rounded=False)
        add_text(s5, x + 0.10, 1.38, 3.85, 0.30, hdr, font_size=11.5, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s5, x + 0.10, 1.75, 3.85, 0.50, big, font_size=22, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s5, x + 0.10, 2.32, 3.85, 0.40, sub, font_size=11.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    add_figure_card(s5, os.path.join(FIG_VI, "02_target_and_burnout_distribution.png"), 0.40, 3.00, 12.53, 3.90, border_color=C_CYAN, pad=0.06)

    # ==========================================================================
    # SLIDE 06: EDA 2 - TƯƠNG QUAN YẾU TỐ BẢO VỆ VS RỦI RO (40 giây)
    # ==========================================================================
    s6 = create_header_footer(
        prs, 6, total_slides,
        "03 PHÂN TÍCH DỮ LIỆU (EDA)",
        "Phát Hiện EDA 2: Nhóm Yếu Tố Bảo Vệ vs. Nhóm Yếu Tố Nguy Cơ",
        "Chuyên cần và Động lực giúp giữ chân sinh viên; Căng thẳng, Lo âu và Thu nhập thấp đẩy nhanh nguy cơ bỏ học"
    )
    add_figure_card(s6, os.path.join(FIG_VI, "04_correlation_matrix_heatmap.png"), 0.40, 1.38, 7.20, 5.48, border_color=C_CYAN, pad=0.06)

    add_simple_card(
        s6, 7.80, 1.38, 5.13, 2.65,
        "YẾU TỐ BẢO VỆ (GIẢM BỎ HỌC)",
        [
            "• 1. Chuyên cần cao (80.7% vs 73.7%): r = -0.244",
            "• 2. Động lực học tập cao: r = -0.193",
            "• 3. Giờ tự học đều đặn: r = -0.160",
            "• 4. Gia đình đồng hành hỗ trợ: r = -0.137",
        ],
        accent_color=C_TEAL, body_size=13.2
    )
    add_simple_card(
        s6, 7.80, 4.20, 5.13, 2.66,
        "YẾU TỐ NGUY CƠ (TĂNG BỎ HỌC)",
        [
            "• 1. Mức độ Kiệt sức (Burnout): r = +0.550",
            "• 2. Điểm Căng thẳng (Stress_Level): r = +0.164",
            "• 3. Điểm Lo âu & Áp lực tài chính: r = +0.145",
            "• 4. Nhóm thu nhập thấp bỏ học lên tới 62.1%",
        ],
        accent_color=C_CORAL, body_size=13.2
    )

    # ==========================================================================
    # SLIDE 07: KẾT QUẢ SO SÁNH 4 MÔ HÌNH HỌC MÁY (60 giây)
    # ==========================================================================
    s7 = create_header_footer(
        prs, 7, total_slides,
        "04 KẾT QUẢ MÔ HÌNH ML",
        "So Sánh 4 Mô Hình Học Máy Tiêu Biểu Trên 2 Kịch Bản",
        "Kiểm định chéo 5-Fold (2,000 SV): Random Forest dẫn đầu Kịch bản 1 (89.01%); XGBoost dẫn đầu Kịch bản 2 (88.26%)"
    )
    add_figure_card(s7, os.path.join(FIG_VI, "07_model_performance_comparison.png"), 0.40, 1.36, 7.50, 5.50, border_color=C_CYAN, pad=0.05)

    add_simple_card(
        s7, 8.10, 1.36, 4.83, 2.68,
        "KỊCH BẢN 1: ĐẦY ĐỦ (CÓ BURNOUT)",
        [
            "★ Hạng 1: Random Forest",
            "  • CV ROC-AUC: 89.01%",
            "  • Test Accuracy: 81.50% | F1: 83.63%",
            "• Hạng 2: Soft Voting (AUC 88.89%)",
            "• Hạng 3: XGBoost (AUC 88.64%)",
            "• Hạng 4: Logistic Regression (87.61%)",
        ],
        accent_color=C_CYAN, body_size=13.0
    )
    add_simple_card(
        s7, 8.10, 4.18, 4.83, 2.68,
        "KỊCH BẢN 2: CẢNH BÁO SỚM (ẨN BURNOUT)",
        [
            "★ Hạng 1: XGBoost",
            "  • CV ROC-AUC: 88.26% | F1: 82.40%",
            "• Hạng 2: Soft Voting (AUC 88.25%)",
            "→ Ý nghĩa: Dù chưa đo Burnout, mô hình vẫn cảnh báo sớm chính xác > 88% chỉ từ điểm chuyên cần & căng thẳng!",
        ],
        accent_color=C_TEAL, body_size=12.8
    )

    # ==========================================================================
    # SLIDE 08: GIẢI THÍCH MÔ HÌNH & TOP 6 ĐẶC TRƯNG (45 giây)
    # ==========================================================================
    s8 = create_header_footer(
        prs, 8, total_slides,
        "04 GIẢI THÍCH MÔ HÌNH",
        "Đường Cong ROC / PR & Top 6 Yếu Tố Quyết Định Nguy Cơ Bỏ Học",
        "Hai chỉ số mới do nhóm thiết kế (Tỷ số Căng thẳng/Gắn kết & Chỉ số Gắn kết) đứng ngay sau Burnout"
    )
    add_figure_card(s8, os.path.join(FIG_VI, "08_roc_and_precision_recall_curves.png"), 0.40, 1.36, 6.15, 4.15, border_color=C_CYAN, pad=0.05)
    add_figure_card(s8, os.path.join(FIG_VI, "10_feature_importance_analysis.png"), 6.78, 1.36, 6.15, 4.15, border_color=C_TEAL, pad=0.05)

    add_simple_card(
        s8, 0.40, 5.62, 12.53, 1.32,
        "TOP 6 ĐẶC TRƯNG QUAN TRỌNG NHẤT & Ý NGHĨA THỰC TIỄN",
        [
            "• Top 6: (1) Burnout_Ordinal → (2) Stress_Engagement_Ratio → (3) Academic_Engagement_Index → (4) Stress_Index → (5) Chuyên cần → (6) Cán cân Hỗ trợ.",
            "• Độ nhạy cao (Recall > 84%): Phát hiện đúng hơn 4/5 số sinh viên thực sự có nguy cơ bỏ học để nhà trường kịp thời hỗ trợ.",
        ],
        accent_color=C_AMBER, body_size=12.5
    )

    # ==========================================================================
    # SLIDE 09: ỨNG DỤNG WEB DASHBOARD CẢNH BÁO SỚM (45 giây)
    # ==========================================================================
    s9 = create_header_footer(
        prs, 9, total_slides,
        "05 SẢN PHẨM ỨNG DỤNG WEB",
        "Kiến Trúc Hệ Thống 3 Tầng & Web Dashboard Cảnh Báo Sớm (FastAPI)",
        "Nhập thông tin sinh viên hoặc chọn 3 mẫu thử nhanh — Trả kết quả dự báo kép & khuyến nghị trong < 30ms"
    )
    add_figure_card(s9, os.path.join(FIG_VI, "12_system_architecture_diagram.png"), 0.40, 1.38, 7.65, 5.48, border_color=C_TEAL, pad=0.06)

    add_simple_card(
        s9, 8.25, 1.38, 4.68, 5.48,
        "TÍNH NĂNG WEB DEMO (APP.PY)",
        [
            "★ 1. Ba Mẫu Sinh viên Thử Nhanh:",
            "• Mẫu An toàn: Nguy cơ 8.0% (Thấp)",
            "• Mẫu Cảnh báo: Nguy cơ ~58% (Vừa)",
            "• Mẫu Báo động: Nguy cơ 95.8% (Cao!)",
            "",
            "★ 2. Dự báo Kép & Gợi ý Tự động:",
            "• Chẩn đoán đồng thời Burnout + Dropout.",
            "• Đề xuất: Học phụ đạo, tư vấn tâm lý, hỗ trợ học bổng.",
        ],
        accent_color=C_CYAN, body_size=13.2
    )

    # ==========================================================================
    # SLIDE 10: KẾT LUẬN & HỎI ĐÁP - Q&A (30 giây)
    # ==========================================================================
    s10 = create_header_footer(
        prs, 10, total_slides,
        "06 KẾT LUẬN & HỎI ĐÁP",
        "Tổng Kết Đóng Góp & Định Hướng Phát Triển Tương Lai",
        "Hoàn thành 100% mục tiêu đồ án — Sẵn sàng tích hợp vào hệ thống Quản lý Học tập (LMS) của nhà trường"
    )
    add_simple_card(
        s10, 0.40, 1.40, 6.15, 3.80,
        "KẾT QUẢ ĐẠT ĐƯỢC CỐT LÕI",
        [
            "• 1. Chuẩn hóa 2,000 hồ sơ sinh viên sạch 100% (0 NaN).",
            "• 2. Thiết kế 6 chỉ số mới đo lường Căng thẳng vs Gắn kết.",
            "• 3. Huấn luyện 4 mô hình đạt đỉnh ROC-AUC 89.01%.",
            "• 4. Xây dựng hoàn chỉnh Web App cảnh báo sớm (< 30ms).",
        ],
        accent_color=C_TEAL, body_size=13.5
    )
    add_simple_card(
        s10, 6.78, 1.40, 6.15, 3.80,
        "HƯỚNG PHÁT TRIỂN TIẾP THEO",
        [
            "• 1. Tích hợp biểu đồ giải thích SHAP trực tiếp lên Web.",
            "• 2. Kết nối API tự động với hệ thống điểm danh LMS.",
            "• 3. Áp dụng mô hình chuỗi thời gian (LSTM) theo từng tuần.",
            "• 4. Mở rộng khảo sát thực tế tại các trường đại học VN.",
        ],
        accent_color=C_CYAN, body_size=13.5
    )

    add_rect(s10, 0.40, 5.40, 12.53, 1.50, C_CARD_BG, border_rgb=C_AMBER, border_pt=2.0, rounded=True)
    add_text(
        s10, 0.60, 5.52, 12.13, 0.45,
        "CẢM ƠN THẦY CÔ VÀ HỘI ĐỒNG ĐÃ LẮNG NGHE! — XIN MỜI CÂU HỎI PHẢN BIỆN (Q&A)",
        font_size=17, bold=True, color=C_AMBER, align=PP_ALIGN.CENTER
    )
    add_text(
        s10, 0.60, 6.08, 12.13, 0.45,
        "Nhóm 03: Đỗ Quang Minh (Trưởng nhóm) • Vũ Quang Huy (Thành viên)  |  Live Demo: http://127.0.0.1:8000",
        font_size=13.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER
    )

    # Save to all target locations
    prs.save(PPTX_FINAL)
    prs.save(PPTX_VI)
    if os.path.exists(DRV_DIR):
        try:
            prs.save(os.path.join(DRV_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx"))
            drv_vi = os.path.join(DRV_DIR, "Ban_Tieng_Viet_Full")
            if os.path.exists(drv_vi):
                prs.save(os.path.join(drv_vi, "SIC_AI_Capstone Project_Presentation Slide Template.pptx"))
        except Exception as e:
            print("Drive save note:", e)

    print(f"Saved streamlined 10-Slide presentation -> {PPTX_FINAL} and {PPTX_VI}")


if __name__ == "__main__":
    build_streamlined_deck()

