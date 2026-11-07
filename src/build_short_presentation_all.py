"""
Builds the complete 9-Slide Short Presentation package (target duration: 6m 30s, 5-8 min window)
and the 6-Slide Q&A Backup Deck in FINAL_SHORT_PRESENTATION/.

Outputs:
1. FINAL_SHORT_PRESENTATION/SIC_AI_Capstone_Short_Final.pptx (9 minimalist slides)
2. FINAL_SHORT_PRESENTATION/SIC_AI_Capstone_Short_Final.pdf
3. FINAL_SHORT_PRESENTATION/Q&A_Backup_Slides.pptx (6 technical backup slides)
4. FINAL_SHORT_PRESENTATION/Q&A_Backup_Slides.pdf
5. FINAL_SHORT_PRESENTATION/Slide_Images/ (Rendered PNGs for visual inspection)
"""

import os
import glob
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image as PilImage
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG_VI = os.path.join(BASE_DIR, "figures_vi")
OUT_DIR = os.path.join(BASE_DIR, "FINAL_SHORT_PRESENTATION")
IMG_DIR = os.path.join(OUT_DIR, "Slide_Images")
VI_DIR = os.path.join(BASE_DIR, "SIC_Documents_VI")
FINAL_DIR = os.path.join(BASE_DIR, "FINAL_PRESENTATION")

os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

PPTX_SHORT = os.path.join(OUT_DIR, "SIC_AI_Capstone_Short_Final.pptx")
PDF_SHORT = os.path.join(OUT_DIR, "SIC_AI_Capstone_Short_Final.pdf")
PPTX_BACKUP = os.path.join(OUT_DIR, "Q&A_Backup_Slides.pptx")
PDF_BACKUP = os.path.join(OUT_DIR, "Q&A_Backup_Slides.pdf")

# Design System Colors (Strictly following Section 3 specification)
C_DARK_BG = RGBColor(11, 18, 32)       # #0B1220 Dark Navy
C_CARD_BG = RGBColor(17, 28, 46)       # #111C2E Secondary Navy Card
C_CARD_HDR = RGBColor(23, 38, 62)      # #17263E Header surface
C_BORDER = RGBColor(38, 56, 84)        # #263854 Subtle border
C_CYAN = RGBColor(0, 212, 255)         # #00D4FF Cyan Accent
C_TEAL = RGBColor(0, 229, 176)         # #00E5B0 Teal Accent
C_WHITE = RGBColor(255, 255, 255)      # #FFFFFF Main White
C_GRAY = RGBColor(170, 184, 204)       # #AAB8CC Secondary Text Gray
C_AMBER = RGBColor(255, 183, 3)        # #FFB703 Highlight Warning/Gold
C_CORAL = RGBColor(255, 107, 107)      # #FF6B6B High Risk Alert

FONT_MAIN = "Segoe UI"


def generate_minimalist_slide_charts():
    """Generates two ultra-clean, dark-themed, large-font charts for Slide 5 & Slide 6."""
    plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    # 1. Slide 5 Chart: Dropout Rate by Burnout Level (Low 23.8%, Medium 63.2%, High 92.0%)
    fig, ax = plt.subplots(figsize=(7.8, 4.6), dpi=220)
    fig.patch.set_facecolor("#111C2E")
    ax.set_facecolor("#111C2E")

    tiers = ["Low Burnout\n(Kiệt sức Thấp)", "Medium Burnout\n(Kiệt sức Vừa)", "High Burnout\n(Kiệt sức Cao)"]
    dropout_rates = [23.8, 63.2, 92.0]
    safe_rates = [76.2, 36.8, 8.0]
    colors = ["#00E5B0", "#FFB703", "#FF6B6B"]

    x = np.arange(len(tiers))
    bars = ax.bar(x, dropout_rates, width=0.52, color=colors, edgecolor="#FFFFFF", linewidth=1.2, zorder=3)

    for bar, rate, safe in zip(bars, dropout_rates, safe_rates):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 3.0,
            f"{rate:.1f}%",
            ha="center",
            va="bottom",
            fontsize=20,
            fontweight="bold",
            color="#FFFFFF",
        )
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() / 2,
            f"An toàn:\n{safe:.1f}%",
            ha="center",
            va="center",
            fontsize=11,
            fontweight="bold",
            color="#0B1220",
        )

    ax.set_ylim(0, 110)
    ax.set_ylabel("Tỷ lệ Bỏ học Quan sát (%)", fontsize=13, fontweight="bold", color="#AAB8CC")
    ax.set_title("Tỷ lệ Sinh viên Bỏ học theo Mức độ Kiệt sức (Burnout_Level)", fontsize=15, fontweight="bold", color="#FFFFFF", pad=14)
    ax.set_xticks(x)
    ax.set_xticklabels(tiers, fontsize=13, fontweight="bold", color="#FFFFFF")
    ax.tick_params(axis="y", colors="#AAB8CC", labelsize=11)
    ax.grid(axis="y", linestyle="--", alpha=0.2, color="#AAB8CC", zorder=0)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    for spine in ["bottom", "left"]:
        ax.spines[spine].set_color("#263854")

    plt.tight_layout()
    s5_chart = os.path.join(FIG_VI, "short_s5_burnout_chart.png")
    plt.savefig(s5_chart, dpi=220, facecolor=fig.get_facecolor())
    plt.close(fig)

    # 2. Slide 6 Chart: 4 Models Comparison (5-Fold CV ROC-AUC %) Across 2 Scenarios
    fig, ax = plt.subplots(figsize=(7.8, 4.6), dpi=220)
    fig.patch.set_facecolor("#111C2E")
    ax.set_facecolor("#111C2E")

    models = ["Logistic\nRegression", "XGBoost", "Soft Voting\nEnsemble", "Random\nForest"]
    s1_auc = [87.61, 88.64, 88.89, 89.01]  # Scenario 1 (With Burnout_Level)
    s2_auc = [87.25, 88.26, 88.25, 87.97]  # Scenario 2 (Without Burnout_Level)

    x = np.arange(len(models))
    w = 0.34
    b1 = ax.bar(x - w / 2, s1_auc, width=w, label="Kịch bản 1: Có Burnout_Level (5-Fold CV)", color="#00D4FF", edgecolor="#FFFFFF", linewidth=0.8, zorder=3)
    b2 = ax.bar(x + w / 2, s2_auc, width=w, label="Kịch bản 2: Không có Burnout_Level (5-Fold CV)", color="#00E5B0", edgecolor="#FFFFFF", linewidth=0.8, zorder=3)

    for bar in b1:
        val = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            val + 0.18,
            f"{val:.2f}%",
            ha="center",
            va="bottom",
            fontsize=11.5,
            fontweight="bold",
            color="#00D4FF",
        )
    for bar in b2:
        val = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            val + 0.18,
            f"{val:.2f}%",
            ha="center",
            va="bottom",
            fontsize=11.5,
            fontweight="bold",
            color="#00E5B0",
        )

    ax.set_ylim(84.5, 90.8)
    ax.set_ylabel("5-Fold Stratified CV ROC-AUC (%)", fontsize=12.5, fontweight="bold", color="#AAB8CC")
    ax.set_title("So sánh ROC-AUC (5-Fold CV) của 4 Mô hình Học máy", fontsize=14.5, fontweight="bold", color="#FFFFFF", pad=12)
    ax.set_xticks(x)
    ax.set_xticklabels(models, fontsize=12.5, fontweight="bold", color="#FFFFFF")
    ax.tick_params(axis="y", colors="#AAB8CC", labelsize=11)
    ax.grid(axis="y", linestyle="--", alpha=0.2, color="#AAB8CC", zorder=0)
    leg = ax.legend(loc="upper left", frameon=True, facecolor="#0B1220", edgecolor="#263854", fontsize=10.5)
    for text in leg.get_texts():
        text.set_color("#FFFFFF")

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    for spine in ["bottom", "left"]:
        ax.spines[spine].set_color("#263854")

    plt.tight_layout()
    s6_chart = os.path.join(FIG_VI, "short_s6_model_chart.png")
    plt.savefig(s6_chart, dpi=220, facecolor=fig.get_facecolor())
    plt.close(fig)
    return s5_chart, s6_chart


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
    r.font.name = FONT_MAIN
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def add_card(slide, left, top, width, height, title, bullets, accent=C_CYAN, body_size=14.5, space_pt=8):
    add_rect(slide, left, top, width, height, C_CARD_BG, border_rgb=accent, border_pt=1.5, rounded=True)
    add_rect(slide, left, top, width, 0.48, C_CARD_HDR, rounded=False)
    add_rect(slide, left, top + 0.46, width, 0.03, accent, rounded=False)
    add_text(slide, left + 0.16, top + 0.08, width - 0.32, 0.34, title, font_size=14, bold=True, color=accent)

    tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.58), Inches(width - 0.36), Inches(height - 0.68))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_pt)
        r = p.add_run()
        r.text = line
        r.font.name = FONT_MAIN
        r.font.size = Pt(body_size)
        if line.startswith("★") or "89.01%" in line or "88.26%" in line or "92.0%" in line:
            r.font.bold = True
            r.font.color.rgb = C_CYAN if "89.01%" in line or "88.26%" in line else C_AMBER
        elif line.startswith("→") or line.startswith("✓"):
            r.font.bold = True
            r.font.color.rgb = C_TEAL
        elif line.startswith("  •") or line.startswith("  -"):
            r.font.color.rgb = C_GRAY
        else:
            r.font.color.rgb = C_WHITE


def add_image_box(slide, img_path, left, top, width, height, bg_rgb=C_CARD_BG, border_rgb=C_CYAN, pad=0.08):
    add_rect(slide, left, top, width, height, bg_rgb, border_rgb=border_rgb, border_pt=1.5, rounded=True)
    if not os.path.exists(img_path):
        return
    with PilImage.open(img_path) as im:
        w_px, h_px = im.size
    aspect = w_px / float(h_px)
    avail_w = width - 2 * pad
    avail_h = height - 2 * pad
    if avail_w / avail_h > aspect:
        final_h = avail_h
        final_w = avail_h * aspect
    else:
        final_w = avail_w
        final_h = avail_w / aspect
    x = left + (width - final_w) / 2.0
    y = top + (height - final_h) / 2.0
    slide.shapes.add_picture(img_path, Inches(x), Inches(y), width=Inches(final_w), height=Inches(final_h))


def create_base_slide(prs, slide_num, total_slides, badge_text, title_text, takeaway_text):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, 13.333, 7.5, C_DARK_BG)

    # Top Header Bar
    add_rect(slide, 0, 0, 13.333, 1.18, C_CARD_BG)
    add_rect(slide, 0, 1.16, 13.333, 0.03, C_CYAN)
    add_rect(slide, 0, 0, 0.12, 1.18, C_TEAL)

    # Badge
    add_rect(slide, 0.38, 0.12, 3.10, 0.27, C_CARD_HDR, border_rgb=C_TEAL, border_pt=1.0, rounded=True)
    add_text(slide, 0.40, 0.13, 3.06, 0.25, badge_text, font_size=9.5, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    # Title & Single Main Message
    add_text(slide, 0.38, 0.42, 12.5, 0.42, title_text, font_size=22, bold=True, color=C_WHITE)
    add_text(slide, 0.38, 0.82, 12.5, 0.28, takeaway_text, font_size=12.5, italic=True, color=C_CYAN)

    # Footer
    add_rect(slide, 0, 7.12, 13.333, 0.38, C_CARD_BG)
    add_text(slide, 0.38, 7.17, 10.0, 0.26, "SAMSUNG INNOVATION CAMPUS (SIC) — NHÓM 03: ĐỖ QUANG MINH • VŨ QUANG HUY", font_size=9.5, bold=True, color=C_GRAY)
    add_text(slide, 11.2, 7.17, 1.75, 0.26, f"Slide {slide_num:02d} / {total_slides:02d}", font_size=10, bold=True, color=C_CYAN, align=PP_ALIGN.RIGHT)
    return slide


def build_9_slide_short_deck(s5_chart, s6_chart):
    total_slides = 9
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # ==========================================================================
    # SLIDE 1 – TRANG BÌA (20 GIÂY)
    # ==========================================================================
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s1, 0, 0, 13.333, 7.5, C_DARK_BG)
    add_rect(s1, 0, 0, 0.22, 7.5, C_CYAN)
    add_rect(s1, 0.22, 0, 0.10, 7.5, C_TEAL)

    # Top SIC Pill
    add_rect(s1, 0.90, 0.80, 4.80, 0.44, C_CARD_BG, border_rgb=C_TEAL, border_pt=1.5, rounded=True)
    add_text(s1, 0.95, 0.85, 4.70, 0.34, "SAMSUNG INNOVATION CAMPUS (SIC)", font_size=12, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    # Project Title
    title_text = (
        "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY\n"
        "DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG VÀ THAM GIA HỌC TẬP"
    )
    add_text(s1, 0.90, 1.55, 11.6, 1.65, title_text, font_size=27, bold=True, color=C_WHITE)
    add_text(s1, 0.90, 3.25, 11.0, 0.45, "Đồ án Tốt nghiệp AI Capstone — Hệ thống Cảnh báo Sớm Nguy cơ Bỏ học", font_size=16, italic=True, color=C_CYAN)

    # Clean Group 03 Card
    add_rect(s1, 0.90, 4.10, 11.5, 2.20, C_CARD_BG, border_rgb=C_CYAN, border_pt=1.5, rounded=True)
    add_rect(s1, 0.90, 4.10, 11.5, 0.48, C_CARD_HDR, rounded=False)
    add_text(s1, 1.15, 4.18, 11.0, 0.34, "NHÓM THỰC HIỆN: NHÓM 03", font_size=14, bold=True, color=C_CYAN)

    add_text(s1, 1.15, 4.80, 5.4, 0.55, "• Đỗ Quang Minh", font_size=20, bold=True, color=C_WHITE)
    add_text(s1, 1.45, 5.32, 5.1, 0.45, "Trưởng nhóm (Team Leader)", font_size=14, color=C_TEAL)

    add_text(s1, 6.80, 4.80, 5.4, 0.55, "• Vũ Quang Huy", font_size=20, bold=True, color=C_WHITE)
    add_text(s1, 7.10, 5.32, 5.1, 0.45, "Thành viên (Team Member)", font_size=14, color=C_TEAL)

    add_text(s1, 11.2, 7.05, 1.75, 0.30, "Slide 01 / 09", font_size=10, bold=True, color=C_GRAY, align=PP_ALIGN.RIGHT)

    # ==========================================================================
    # SLIDE 2 – BÀI TOÁN CHÚNG TÔI GIẢI QUYẾT (40 GIÂY)
    # ==========================================================================
    s2 = create_base_slide(
        prs, 2, total_slides,
        "01 BÀI TOÁN NGHIÊN CỨU",
        "Bài Toán Chúng Tôi Giải Quyết",
        "Phát hiện sớm nguy cơ bỏ học ngay từ khi sinh viên mới xuất hiện dấu hiệu căng thẳng và giảm gắn kết học tập"
    )

    # Visual Flow Diagram at Top (Căng thẳng tăng + Gắn kết giảm -> Nguy cơ kiệt sức -> Nguy cơ bỏ học)
    flow_steps = [
        ("1. NGUYÊN NHÂN TÍCH LŨY", "Căng thẳng tăng cao\n+ Gắn kết học tập suy giảm", C_CYAN),
        ("2. TRẠNG THÁI TRUNG GIAN", "Nguy cơ Kiệt sức Học đường\n(Academic Burnout)", C_AMBER),
        ("3. HỆ QUẢ CUỐI CÙNG", "Nguy cơ Bỏ học Đại học\n(Student Dropout Risk)", C_CORAL),
    ]
    for idx, (f_hdr, f_body, col) in enumerate(flow_steps):
        x = 0.40 + idx * 4.35
        add_rect(s2, x, 1.42, 3.85, 1.85, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s2, x, 1.42, 3.85, 0.42, C_CARD_HDR, rounded=False)
        add_text(s2, x + 0.12, 1.48, 3.60, 0.30, f_hdr, font_size=12.5, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s2, x + 0.15, 2.00, 3.55, 1.10, f_body, font_size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        if idx < 2:
            add_text(s2, x + 3.88, 2.05, 0.44, 0.60, "➔", font_size=26, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    # 3 Question Cards at Bottom
    q_cards = [
        ("1. VÌ SAO SINH VIÊN BỎ HỌC?", [
            "• Không diễn ra đột ngột trong một ngày.",
            "• Là quá trình tích lũy áp lực kéo dài.",
            "• Sinh viên mất cân bằng giữa học tập và đời sống.",
        ], C_CYAN),
        ("2. TÁC ĐỘNG CỦA STRESS & HỌC TẬP", [
            "• Vắng học, giảm giờ tự học, nợ môn.",
            "• Căng thẳng, lo âu, áp lực tài chính tăng.",
            "→ Đẩy sinh viên vào trạng thái kiệt sức.",
        ], C_AMBER),
        ("3. TẠI SAO CẦN CẢNH BÁO SỚM?", [
            "• Khi sinh viên nộp đơn thôi học thì đã muộn.",
            "• AI giúp nhà trường nhận diện sớm rủi ro.",
            "→ Hỗ trợ học vụ và tâm lý đúng thời điểm.",
        ], C_TEAL),
    ]
    for idx, (q_t, q_b, col) in enumerate(q_cards):
        x = 0.40 + idx * 4.22
        add_card(s2, x, 3.55, 4.08, 3.30, q_t, q_b, accent=col, body_size=14.0, space_pt=10)

    # ==========================================================================
    # SLIDE 3 – DỮ LIỆU ĐẦU VÀO (40 GIÂY)
    # ==========================================================================
    s3 = create_base_slide(
        prs, 3, total_slides,
        "02 DỮ LIỆU ĐẦU VÀO",
        "Cấu Trúc Bộ Dữ Liệu & Quy Trình Mở Rộng Mẫu",
        "800 bản ghi gốc (25 thuộc tính) được làm sạch và mở rộng thống kê lên 2.000 hồ sơ (34 thuộc tính)"
    )

    # 4 Top KPI Cards
    s3_kpis = [
        ("800", "Bản ghi ban đầu\n(Dữ liệu gốc có 612 ô khuyết)", C_CYAN),
        ("25", "Thuộc tính gốc\n(18 cột có giá trị thiếu)", C_TEAL),
        ("2.000", "Hồ sơ sau xử lý & mở rộng\n(Sạch 100%, 0 giá trị NaN)", C_AMBER),
        ("34", "Thuộc tính hoàn chỉnh\n(Thêm 6 đặc trưng & mã hóa)", C_CYAN),
    ]
    for idx, (val, lbl, col) in enumerate(s3_kpis):
        x = 0.40 + idx * 3.18
        add_rect(s3, x, 1.40, 2.98, 1.65, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_text(s3, x, 1.50, 2.98, 0.65, val, font_size=28, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s3, x + 0.10, 2.22, 2.78, 0.72, lbl, font_size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Left Card: 4 Attribute Domains + Targets
    add_card(
        s3, 0.40, 3.28, 6.15, 3.60,
        "CÁC NHÓM DỮ LIỆU ĐẦU VÀO",
        [
            "• 1. Nhân khẩu học: Độ tuổi, Giới tính, Năm học, Khoa, Cư trú, Thu nhập.",
            "• 2. Học tập: Chuyên cần (%), Giờ tự học/ngày, GPA kỳ trước, Số môn nợ, Động lực.",
            "• 3. Lối sống: Giờ ngủ, Thời gian dùng màn hình, Thể dục, Hoạt động xã hội, Làm thêm.",
            "• 4. Căng thẳng & Hỗ trợ: Mức căng thẳng, Lo âu, Tài chính, Bạn bè, Gia đình.",
            "→ Biến mục tiêu: Burnout_Level (3 mức) & Dropout_Risk (Yes 54.6% / No 45.4%).",
        ],
        accent=C_TEAL, body_size=13.2, space_pt=7
    )

    # Right Card: Data Provenance Flow (Explicitly clarifying origin of 2,000 records)
    add_card(
        s3, 6.78, 3.28, 6.15, 3.60,
        "NGUỒN GỐC 2.000 HỒ SƠ (MINH BẠCH KHOA HỌC)",
        [
            "★ Bước 1: Làm sạch 800 bản ghi gốc",
            "  • Điền khuyết KNNImputer (k=7) + lọc ngoại lai IQR → 800 hồ sơ sạch.",
            "★ Bước 2: Mở rộng thêm 1.200 bản ghi tổng hợp",
            "  • Sinh mẫu từ phân phối chuẩn đa biến có phân tầng (Multivariate Gaussian) khớp theo từng nhóm (Burnout × Dropout) của 800 mẫu gốc.",
            "→ Lưu ý học thuật: 2.000 hồ sơ là bộ dữ liệu mở rộng thống kê phục vụ huấn luyện mô hình, không phải khảo sát thực địa đợt 2.",
        ],
        accent=C_AMBER, body_size=13.0, space_pt=6
    )

    # ==========================================================================
    # SLIDE 4 – QUY TRÌNH XÂY DỰNG AI (50 GIÂY)
    # ==========================================================================
    s4 = create_base_slide(
        prs, 4, total_slides,
        "03 QUY TRÌNH XÂY DỰNG AI",
        "Quy Trình Xây Dựng & Huấn Luyện Mô Hình Học Máy",
        "Đóng gói khép kín toàn bộ các bước trong Pipeline để ngăn ngừa rò rỉ dữ liệu (Data Leakage)"
    )

    # 5-Step Pipeline Flow Diagram
    steps_5 = [
        ("Bước 1", "Data Cleaning", "Làm sạch & Lọc ngoại lai IQR", C_CYAN),
        ("Bước 2", "Preprocessing", "KNNImputer (k=7) & Chuẩn hóa", C_TEAL),
        ("Bước 3", "Feature Engineering", "Tạo 6 Đặc trưng tương tác mới", C_AMBER),
        ("Bước 4", "Model Training", "Huấn luyện 4 Mô hình Học máy", C_CYAN),
        ("Bước 5", "Model Evaluation", "Stratified 5-Fold CV & Test 20%", C_TEAL),
    ]
    for idx, (st_num, st_en, st_vi, col) in enumerate(steps_5):
        x = 0.40 + idx * 2.54
        add_rect(s4, x, 1.45, 2.30, 1.85, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s4, x, 1.45, 2.30, 0.40, C_CARD_HDR, rounded=False)
        add_text(s4, x + 0.08, 1.50, 2.14, 0.30, st_num, font_size=12, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.08, 1.95, 2.14, 0.48, st_en, font_size=14.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.10, 2.50, 2.10, 0.65, st_vi, font_size=11.5, color=C_GRAY, align=PP_ALIGN.CENTER)
        if idx < 4:
            add_text(s4, x + 2.31, 2.10, 0.22, 0.45, "➔", font_size=16, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

    # 3 Highlight Cards at Bottom
    pillars = [
        ("1. ĐIỀN KHUYẾT KNNIMPUTER (k=7)", [
            "• Tìm 7 sinh viên có hồ sơ gần giống nhất để nội suy giá trị bị thiếu.",
            "• Đặt bên trong Pipeline để không làm rò rỉ dữ liệu giữa tập Train và Test.",
        ], C_CYAN),
        ("2. SÁU ĐẶC TRƯNG MỚI (6 FEATURES)", [
            "• Kết hợp các biến rời rạc thành chỉ số đo cán cân Căng thẳng vs. Học tập.",
            "• Nổi bật: Chỉ số Gắn kết, Chỉ số Căng thẳng và Tỷ số Căng thẳng / Gắn kết.",
        ], C_AMBER),
        ("3. STRATIFIED 5-FOLD CV", [
            "• Chia dữ liệu thành 5 phần xoay vòng kiểm định khách quan.",
            "• Giữ nguyên tỷ lệ nhãn bỏ học (54.6% Yes / 45.4% No) ở mọi lượt đánh giá.",
        ], C_TEAL),
    ]
    for idx, (p_t, p_b, col) in enumerate(pillars):
        x = 0.40 + idx * 4.22
        add_card(s4, x, 3.60, 4.08, 3.25, p_t, p_b, accent=col, body_size=14.0, space_pt=10)

    # ==========================================================================
    # SLIDE 5 – PHÂN TÍCH DỮ LIỆU (45 GIÂY)
    # ==========================================================================
    s5 = create_base_slide(
        prs, 5, total_slides,
        "04 PHÂN TÍCH DỮ LIỆU (EDA)",
        "Phát Hiện Quan Trọng Nhất: Tỷ Lệ Bỏ Học Theo Mức Độ Burnout",
        "Mức độ kiệt sức học đường càng cao thì tỷ lệ bỏ học quan sát được trong bộ dữ liệu càng lớn"
    )

    # Left: Big Clean Bar Chart
    add_image_box(s5, s5_chart, 0.40, 1.40, 7.65, 5.45, bg_rgb=C_CARD_BG, border_rgb=C_CYAN, pad=0.06)

    # Right: 3 Big Tier Callouts + Non-causal note
    tiers_info = [
        ("MỨC THẤP (LOW BURNOUT)", "23.8% Bỏ học", "76 / 320 SV bỏ học", C_TEAL),
        ("MỨC VỪA (MEDIUM BURNOUT)", "63.2% Bỏ học", "177 / 280 SV bỏ học", C_AMBER),
        ("MỨC CAO (HIGH BURNOUT)", "92.0% Bỏ học", "184 / 200 SV bỏ học", C_CORAL),
    ]
    for idx, (t_hdr, t_val, t_sub, col) in enumerate(tiers_info):
        y = 1.40 + idx * 1.42
        add_rect(s5, 8.30, y, 4.63, 1.28, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_text(s5, 8.48, y + 0.10, 4.30, 0.28, t_hdr, font_size=12, bold=True, color=col)
        add_text(s5, 8.48, y + 0.40, 2.30, 0.50, t_val, font_size=22, bold=True, color=C_WHITE)
        add_text(s5, 10.65, y + 0.48, 2.15, 0.45, t_sub, font_size=12, bold=True, color=C_GRAY, align=PP_ALIGN.RIGHT)

    # Scientific caution box at bottom right
    add_rect(s5, 8.30, 5.68, 4.63, 1.17, C_CARD_HDR, border_rgb=C_CYAN, border_pt=1.2, rounded=True)
    add_text(
        s5, 8.45, 5.76, 4.35, 1.00,
        "💡 Lưu ý khoa học:\nSố liệu phản ánh mối liên hệ thống kê quan sát được trong dữ liệu (r = +0.550), không tự kết luận là quan hệ nhân quả một chiều.",
        font_size=11.5, italic=True, color=C_WHITE
    )

    # ==========================================================================
    # SLIDE 6 – KẾT QUẢ HUẤN LUYỆN MÔ HÌNH (60 GIÂY)
    # ==========================================================================
    s6 = create_base_slide(
        prs, 6, total_slides,
        "05 KẾT QUẢ MÔ HÌNH HỌC MÁY",
        "So Sánh 4 Mô Hình Học Máy Trên 2 Kịch Bản (5-Fold CV)",
        "Random Forest dẫn đầu Kịch bản 1 (89.01% CV ROC-AUC); XGBoost dẫn đầu Kịch bản 2 Cảnh báo sớm (88.26% CV ROC-AUC)"
    )

    # Left: Big Comparison Bar Chart
    add_image_box(s6, s6_chart, 0.40, 1.40, 7.65, 5.45, bg_rgb=C_CARD_BG, border_rgb=C_CYAN, pad=0.06)

    # Right Top: Scenario 1 Card
    add_card(
        s6, 8.30, 1.40, 4.63, 2.60,
        "KỊCH BẢN 1: CÓ BURNOUT_LEVEL",
        [
            "★ Mô hình dẫn đầu: Random Forest",
            "• 5-Fold CV ROC-AUC: 89.01%",
            "• Hold-out Test (20%): ROC-AUC 88.42% | Acc 81.50% | F1 83.63%",
            "→ Phù hợp khi nhà trường đã có đánh giá mức độ kiệt sức.",
        ],
        accent=C_CYAN, body_size=13.0, space_pt=6
    )

    # Right Bottom: Scenario 2 Card
    add_card(
        s6, 8.30, 4.22, 4.63, 2.63,
        "KỊCH BẢN 2: KHÔNG CÓ BURNOUT_LEVEL",
        [
            "★ Mô hình dẫn đầu: XGBoost",
            "• 5-Fold CV ROC-AUC: 88.26%",
            "• Hold-out Test (20%): ROC-AUC 87.39% | Acc 78.50% | F1 80.28%",
            "→ Cảnh báo sớm hiệu quả ngay từ các chỉ số học tập & căng thẳng.",
        ],
        accent=C_TEAL, body_size=13.0, space_pt=6
    )

    # ==========================================================================
    # SLIDE 7 – SẢN PHẨM WEB DASHBOARD (60 GIÂY)
    # ==========================================================================
    s7 = create_base_slide(
        prs, 7, total_slides,
        "06 SẢN PHẨM ỨNG DỤNG THỰC TẾ",
        "Ứng Dụng Web Dashboard Cảnh Báo Sớm (FastAPI)",
        "Giao diện thực tế (app.py) hỗ trợ cố vấn học tập nhập thông tin, nhận dự báo kép và khuyến nghị tức thì"
    )

    # Top 3-Step Flow Banner: INPUT -> AI PREDICTION -> RESULT & RECOMMENDATION
    w_steps = [
        ("1. INPUT", "Nhập 18 chỉ số sinh viên hoặc chọn 3 mẫu thử nhanh", C_CYAN),
        ("2. AI PREDICTION", "Phân tích đặc trưng & Chạy mô hình dự báo kép", C_AMBER),
        ("3. RESULT & RECOMMENDATION", "Hiển thị mức Burnout, xác suất Dropout & Khuyến nghị", C_TEAL),
    ]
    for idx, (ws_t, ws_d, col) in enumerate(w_steps):
        x = 0.40 + idx * 4.30
        add_rect(s7, x, 1.35, 3.93, 0.92, C_CARD_BG, border_rgb=col, border_pt=1.8, rounded=True)
        add_text(s7, x + 0.12, 1.40, 3.68, 0.30, ws_t, font_size=12.5, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s7, x + 0.12, 1.72, 3.68, 0.48, ws_d, font_size=11.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        if idx < 2:
            add_text(s7, x + 3.95, 1.55, 0.32, 0.45, "➔", font_size=18, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

    # Left: Real Screenshot of app.py
    live_shot = os.path.join(FIG_VI, "13_live_web_dashboard_screenshot.png")
    add_image_box(s7, live_shot, 0.40, 2.42, 7.85, 4.45, bg_rgb=RGBColor(255, 255, 255), border_rgb=C_CYAN, pad=0.04)

    # Right: 5 Operational Steps on Dashboard
    add_card(
        s7, 8.45, 2.42, 4.48, 4.45,
        "5 BƯỚC HOẠT ĐỘNG TRÊN WEB APP",
        [
            "1. Nhập thông tin sinh viên: Chuyên cần, giờ học, GPA, căng thẳng, giấc ngủ...",
            "2. Phân tích dữ liệu: Tự động tính Chỉ số Gắn kết & Chỉ số Căng thẳng.",
            "3. Dự báo mức độ Burnout: Hiển thị xác suất Low / Medium / High.",
            "4. Dự báo nguy cơ Dropout: Cảnh báo phần trăm rủi ro bỏ học.",
            "5. Khuyến nghị hỗ trợ: Đề xuất phụ đạo học tập, tham vấn tâm lý hoặc hỗ trợ tài chính.",
        ],
        accent=C_TEAL, body_size=12.8, space_pt=6
    )

    # ==========================================================================
    # SLIDE 8 – KẾT LUẬN VÀ ĐÓNG GÓP (40 GIÂY)
    # ==========================================================================
    s8 = create_base_slide(
        prs, 8, total_slides,
        "07 KẾT LUẬN & HƯỚNG PHÁT TRIỂN",
        "Tổng Kết Đóng Góp & Định Hướng Phát Triển",
        "Hoàn thiện quy trình từ xử lý dữ liệu, huấn luyện mô hình học máy đến ứng dụng Web hỗ trợ cảnh báo sớm"
    )

    add_card(
        s8, 0.40, 1.45, 6.15, 5.35,
        "3 ĐÓNG GÓP CHÍNH CỦA ĐỒ ÁN",
        [
            "★ 1. Quy trình xử lý & phân tích dữ liệu bài bản",
            "  • Làm sạch 800 bản ghi gốc, điền khuyết KNNImputer (k=7) trong Pipeline và thiết kế 6 đặc trưng tương tác mới.",
            "★ 2. Huấn luyện & so sánh 4 mô hình Machine Learning",
            "  • Đánh giá trên 2 kịch bản (5-Fold CV): Random Forest đạt 89.01% ROC-AUC (Kịch bản 1) và XGBoost đạt 88.26% (Kịch bản 2).",
            "★ 3. Xây dựng ứng dụng Web cảnh báo sớm",
            "  • Triển khai Web Dashboard trực quan giúp dự báo đồng thời Burnout, Dropout và gợi ý hướng hỗ trợ sinh viên.",
        ],
        accent=C_TEAL, body_size=13.8, space_pt=8
    )

    add_card(
        s8, 6.78, 1.45, 6.15, 5.35,
        "3 HƯỚNG PHÁT TRIỂN TIẾP THEO",
        [
            "→ 1. Bổ sung dữ liệu khảo sát thực tế từ nhiều trường",
            "  • Thu thập và kiểm định ngoại kiểm (External Validation) trên dữ liệu sinh viên thực tế tại các trường đại học.",
            "→ 2. Tích hợp giải thích mô hình cục bộ với SHAP",
            "  • Hiển thị trực tiếp mức độ tác động của từng yếu tố lên kết quả dự báo của mỗi cá nhân sinh viên trên Web.",
            "→ 3. Kết nối hệ thống Quản lý Học tập (LMS)",
            "  • Tự động đồng bộ dữ liệu chuyên cần và điểm số định kỳ để sàng lọc sớm theo từng học kỳ.",
        ],
        accent=C_CYAN, body_size=13.8, space_pt=8
    )

    # ==========================================================================
    # SLIDE 9 – THANK YOU & Q&A (15 GIÂY)
    # ==========================================================================
    s9 = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s9, 0, 0, 13.333, 7.5, C_DARK_BG)
    add_rect(s9, 0, 0, 0.22, 7.5, C_TEAL)
    add_rect(s9, 0.22, 0, 0.10, 7.5, C_CYAN)

    add_rect(s9, 1.60, 1.35, 10.13, 4.80, C_CARD_BG, border_rgb=C_CYAN, border_pt=2.0, rounded=True)
    add_text(s9, 1.80, 1.95, 9.73, 0.90, "THANK YOU FOR LISTENING", font_size=36, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_text(s9, 1.80, 2.95, 9.73, 0.85, "Q & A", font_size=32, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
    add_text(
        s9, 1.80, 4.15, 9.73, 0.55,
        "Nhóm 03 – Samsung Innovation Campus (SIC)",
        font_size=18, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER
    )
    add_text(
        s9, 1.80, 4.75, 9.73, 0.50,
        "Đỗ Quang Minh (Trưởng nhóm)  •  Vũ Quang Huy (Thành viên)",
        font_size=15, color=C_GRAY, align=PP_ALIGN.CENTER
    )
    add_text(s9, 11.2, 7.05, 1.75, 0.30, "Slide 09 / 09", font_size=10, bold=True, color=C_GRAY, align=PP_ALIGN.RIGHT)

    prs.save(PPTX_SHORT)
    # Also update workspace main presentation copies so user has the exact 9-slide version everywhere
    try:
        prs.save( os.path.join(FINAL_DIR, "SIC_AI_Capstone_Final.pptx") )
        prs.save( os.path.join(VI_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx") )
    except Exception as e:
        print("Sync save note:", e)

    print(f"Saved 9-Slide Short Presentation -> {PPTX_SHORT}")


def build_backup_slides_deck():
    """Builds 6 technical backup slides for answering committee Q&A."""
    total_backup = 6
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # BACKUP 1: DATA LEAKAGE PREVENTION
    b1 = create_base_slide(
        prs, 1, total_backup,
        "PHỤ LỤC Q&A 01 — DATA LEAKAGE",
        "Kiểm Soát Rò Rỉ Dữ Liệu (Data Leakage) Trong Pipeline",
        "Toàn bộ bước điền khuyết, chuẩn hóa và tạo đặc trưng đều chỉ fit trên tập Train, tuyệt đối không dùng thông tin tập Test"
    )
    add_card(
        b1, 0.40, 1.40, 6.15, 5.40,
        "1. CƠ CHẾ ĐÓNG GÓI SCIKIT-LEARN PIPELINE",
        [
            "★ Tách biệt hoàn toàn Train (80%) và Test (20%):",
            "• Tập Test (400 hồ sơ) được tách riêng ngay từ đầu bằng Stratified Split.",
            "★ Trong mỗi lượt của 5-Fold Cross Validation:",
            "• Bước 1 (DomainFeatureEngineer): Chỉ tính trung vị (median) trên 4 folds huấn luyện.",
            "• Bước 2 (KNNImputer k=7): Chỉ tìm láng giềng trong 4 folds huấn luyện.",
            "• Bước 3 (StandardScaler & OneHotEncoder): Chỉ lấy μ, σ trên tập huấn luyện.",
        ],
        accent=C_CYAN, body_size=13.5, space_pt=7
    )
    add_card(
        b1, 6.78, 1.40, 6.15, 5.40,
        "2. KIỂM ĐỊNH KỊCH BẢN ẨN BURNOUT_LEVEL",
        [
            "★ Câu hỏi phản biện: Burnout_Level có gây rò rỉ nhãn mục tiêu không?",
            "• Burnout_Level có tương quan cao với Dropout_Risk (r = +0.550).",
            "• Để kiểm chứng mô hình không phụ thuộc thụ động vào Burnout_Level, nhóm thiết kế Kịch bản 2 loại bỏ hoàn toàn biến này.",
            "→ Kết quả Kịch bản 2: XGBoost vẫn đạt 88.26% CV ROC-AUC (chỉ giảm 0.75% so với Kịch bản 1), chứng minh các đặc trưng hành vi có năng lực dự báo độc lập.",
        ],
        accent=C_TEAL, body_size=13.5, space_pt=7
    )

    # BACKUP 2: KNNIMPUTER (k=7) & DATA CLEANING
    b2 = create_base_slide(
        prs, 2, total_backup,
        "PHỤ LỤC Q&A 02 — KNNIMPUTER",
        "Chi Tiết Xử Lý Dữ Liệu Khuyết Bằng KNNImputer (k=7)",
        "18/25 thuộc tính gốc có giá trị khuyết (1.6% – 6.4%) — Xóa dòng sẽ làm mất >40% dữ liệu"
    )
    add_image_box(b2, os.path.join(FIG_VI, "01_missing_values_analysis.png"), 0.40, 1.40, 7.25, 5.40, bg_rgb=RGBColor(255, 255, 255), border_rgb=C_CYAN, pad=0.06)
    add_card(
        b2, 7.85, 1.40, 5.08, 5.40,
        "TẠI SAO CHỌN KNNIMPUTER (k=7)?",
        [
            "• Tổng số ô khuyết trong 800 mẫu gốc: 612 ô.",
            "• Nếu điền bằng Trung bình (Mean) đơn biến sẽ làm biến dạng phân phối và mất tương quan đa chiều.",
            "• KNNImputer (k=7) đo khoảng cách Euclid giữa các sinh viên để ước lượng từ 7 hồ sơ tương đồng nhất.",
            "• Kết hợp lọc ngoại lai theo quy tắc IQR: [Q1 - 1.5×IQR, Q3 + 1.5×IQR].",
        ],
        accent=C_AMBER, body_size=13.5, space_pt=8
    )

    # BACKUP 3: 6 FEATURE ENGINEERING FORMULAS
    b3 = create_base_slide(
        prs, 3, total_backup,
        "PHỤ LỤC Q&A 03 — FEATURE ENGINEERING",
        "Công Thức Toán Học Của 6 Đặc Trưng Mới (DomainFeatureEngineer)",
        "Lượng hóa mối tương tác phi tuyến giữa Căng thẳng tâm lý, Hỗ trợ xã hội và Gắn kết học tập"
    )
    fe_list = [
        ("1. Academic_Engagement_Index", "0.35×(Attendance/10) + 0.25×Study + 0.25×Motivation + 0.15×GPA", C_TEAL),
        ("2. Psychological_Stress_Index", "0.35×Stress + 0.30×Anxiety + 0.20×Financial + 0.15×Peer_Pressure", C_CORAL),
        ("3. Stress_Engagement_Ratio", "Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)", C_CYAN),
        ("4. Net_Stress_Support_Balance", "Stress_Index − (0.50×Family + 0.30×Social + 0.20×Exercise)", C_AMBER),
        ("5. Digital_Fatigue_Ratio", "Screen_Time_Hours / (Sleep_Hours + 1.0)", C_CYAN),
        ("6. Academic_Risk_Burden", "1.50 × Backlogs + (100.0 − Attendance_Percent) / 20.0", C_TEAL),
    ]
    for idx, (f_name, f_eq, col) in enumerate(fe_list):
        r_i, c_i = divmod(idx, 2)
        x = 0.40 + c_i * 6.35
        y = 1.42 + r_i * 1.80
        add_rect(b3, x, y, 6.15, 1.60, C_CARD_BG, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(b3, x, y, 6.15, 0.42, C_CARD_HDR, rounded=False)
        add_text(b3, x + 0.15, y + 0.06, 5.85, 0.30, f_name, font_size=13, bold=True, color=col)
        add_text(b3, x + 0.15, y + 0.58, 5.85, 0.85, f_eq, font_size=13.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # BACKUP 4: MODEL EVALUATION MATRIX (CV vs TEST)
    b4 = create_base_slide(
        prs, 4, total_backup,
        "PHỤ LỤC Q&A 04 — MODEL EVALUATION",
        "Bảng So Sánh Chi Tiết Kết Quả 5-Fold CV Và Tập Kiểm Thử (Test Set 20%)",
        "Phân tách minh bạch giữa kết quả Cross-Validation (1.600 mẫu) và Hold-out Test (400 mẫu)"
    )
    table_data = [
        ["Mô hình Học máy", "KB1: CV Acc", "KB1: CV Recall", "KB1: CV F1", "KB1: CV AUC", "KB1: Test Acc", "KB1: Test F1", "KB1: Test AUC", "KB2: CV AUC"],
        ["1. Random Forest ★", "80.25%", "86.38%", "82.98%", "89.01%", "81.50%", "83.63%", "88.42%", "87.97%"],
        ["2. Soft Voting Ensemble", "79.70%", "84.74%", "82.27%", "88.89%", "80.25%", "82.17%", "88.32%", "88.25%"],
        ["3. XGBoost ★", "79.95%", "84.56%", "82.39%", "88.64%", "78.50%", "80.28%", "87.99%", "88.26%"],
        ["4. Logistic Regression", "78.45%", "81.64%", "80.62%", "87.61%", "78.25%", "80.18%", "87.60%", "87.25%"],
    ]
    t_shp = b4.shapes.add_table(len(table_data), len(table_data[0]), Inches(0.40), Inches(1.50), Inches(12.53), Inches(4.80))
    tbl = t_shp.table
    widths = [2.55, 1.20, 1.28, 1.20, 1.25, 1.25, 1.20, 1.25, 1.35]
    for c_i, w in enumerate(widths):
        tbl.columns[c_i].width = Inches(w)

    for r_i, row in enumerate(table_data):
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text = str(val)
            cell.fill.solid()
            if r_i == 0:
                cell.fill.fore_color.rgb = C_CARD_HDR
            elif r_i in [1, 3]:
                cell.fill.fore_color.rgb = RGBColor(22, 42, 68)
            else:
                cell.fill.fore_color.rgb = C_CARD_BG
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.CENTER
                p.font.name = FONT_MAIN
                p.font.size = Pt(11.5)
                p.font.bold = (r_i == 0 or c_i == 0 or c_i in [4, 7, 8])
                p.font.color.rgb = C_CYAN if r_i == 0 else (C_TEAL if c_i in [4, 8] else C_WHITE)

    # BACKUP 5: CROSS VALIDATION & GENERALIZATION
    b5 = create_base_slide(
        prs, 5, total_backup,
        "PHỤ LỤC Q&A 05 — CROSS VALIDATION",
        "Thiết Kế Stratified 5-Fold CV & Kiểm Soát Overfitting",
        "Độ chênh lệch ROC-AUC giữa 5-Fold CV và tập Test nhỏ hơn 0.7%, cho thấy mô hình ổn định trên dữ liệu nội bộ"
    )
    add_card(
        b5, 0.40, 1.40, 6.15, 5.40,
        "CHIẾN LƯỢC CHIA TẬP DỮ LIỆU",
        [
            "• Tổng quy mô: 2.000 hồ sơ (54.6% Yes / 45.4% No).",
            "• Tập Huấn luyện (Train Set - 80%): 1.600 hồ sơ.",
            "  - Chia thành 5 Folds phân tầng (mỗi Fold 320 hồ sơ).",
            "  - Huấn luyện trên 4 Folds (1.280 mẫu), kiểm định trên 1 Fold (320 mẫu).",
            "• Tập Kiểm thử Độc lập (Hold-out Test - 20%): 400 hồ sơ.",
            "  - Chỉ đánh giá một lần cuối cùng sau khi cố định siêu tham số.",
        ],
        accent=C_CYAN, body_size=13.5, space_pt=8
    )
    add_card(
        b5, 6.78, 1.40, 6.15, 5.40,
        "CẤU HÌNH CHỐNG HỌC VẸT (REGULARIZATION)",
        [
            "• Random Forest: Giới hạn độ sâu max_depth=8, min_samples_leaf=4, max_features='sqrt' trên 300 cây.",
            "• XGBoost: Giới hạn max_depth=3, learning_rate=0.035, subsample=0.85 và hệ số phạt L2 reg_lambda=2.0.",
            "→ Lưu ý học thuật: Kết quả CV và Test gần nhau phản ánh khả năng khái quát hóa nội bộ tốt; để triển khai thực tế diện rộng cần kiểm định thêm trên dữ liệu độc lập từ các trường khác.",
        ],
        accent=C_AMBER, body_size=13.5, space_pt=8
    )

    # BACKUP 6: FEATURE IMPORTANCE & ROC CURVES
    b6 = create_base_slide(
        prs, 6, total_backup,
        "PHỤ LỤC Q&A 06 — FEATURE IMPORTANCE",
        "Đường Cong ROC / Precision-Recall & Độ Quan Trọng Đặc Trưng",
        "Top 6 đặc trưng đóng góp lớn nhất vào khả năng phân loại của mô hình"
    )
    add_image_box(b6, os.path.join(FIG_VI, "08_roc_and_precision_recall_curves.png"), 0.40, 1.40, 6.15, 4.10, bg_rgb=RGBColor(255, 255, 255), border_rgb=C_CYAN, pad=0.05)
    add_image_box(b6, os.path.join(FIG_VI, "10_feature_importance_analysis.png"), 6.78, 1.40, 6.15, 4.10, bg_rgb=RGBColor(255, 255, 255), border_rgb=C_TEAL, pad=0.05)
    add_card(
        b6, 0.40, 5.62, 12.53, 1.28,
        "TOP 6 ĐẶC TRƯNG & LƯU Ý HỌC THUẬT",
        [
            "• Top 6: (1) Burnout_Ordinal, (2) Stress_Engagement_Ratio, (3) Academic_Engagement_Index, (4) Psychological_Stress_Index, (5) Attendance_Percent, (6) Net_Stress_Support_Balance.",
            "• Lưu ý: Feature Importance đo mức độ đóng góp phân tách nút trong cây quyết định, không đồng nhất với tác động nhân quả.",
        ],
        accent=C_CYAN, body_size=12.5, space_pt=4
    )

    prs.save(PPTX_BACKUP)
    print(f"Saved 6-Slide Q&A Backup Deck -> {PPTX_BACKUP}")


if __name__ == "__main__":
    s5_c, s6_c = generate_minimalist_slide_charts()
    build_9_slide_short_deck(s5_c, s6_c)
    build_backup_slides_deck()
