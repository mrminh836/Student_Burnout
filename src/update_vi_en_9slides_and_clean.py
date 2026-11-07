"""
Updates the 9-Slide Short Presentation (5-8 minutes) into both:
- SIC_Documents_VI/SIC_AI_Capstone Project_Presentation Slide Template.pptx (100% Vietnamese)
- SIC_Documents_EN/SIC_AI_Capstone Project_Presentation Slide Template.pptx (100% English)
And cleans up redundant files so that SIC_Documents_VI and SIC_Documents_EN contain
ONLY the 5 main official SIC report files:
1. SIC_AI_Capstone Project_Action Plan.docx
2. SIC_AI_Capstone Project_Evaluation Sheet.docx
3. SIC_AI_Capstone Project_Final Report.docx
4. SIC_AI_Capstone Project_Presentation Slide Template.pptx
5. SIC_AI_Capstone Project_Work Breakdown Structure.xlsx
"""

import os
import shutil
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
FIG_EN = os.path.join(BASE_DIR, "figures_en")
VI_DIR = os.path.join(BASE_DIR, "SIC_Documents_VI")
EN_DIR = os.path.join(BASE_DIR, "SIC_Documents_EN")
DRV_DIR = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"

os.makedirs(FIG_VI, exist_ok=True)
os.makedirs(FIG_EN, exist_ok=True)
os.makedirs(VI_DIR, exist_ok=True)
os.makedirs(EN_DIR, exist_ok=True)

# Design System Colors (Modern AI Technology - Minimalist Dark Navy)
C_DARK_BG = RGBColor(11, 18, 32)       # #0B1220 Dark Navy
C_CARD_BG = RGBColor(17, 28, 46)       # #111C2E Secondary Navy Card
C_CARD_HDR = RGBColor(23, 38, 62)      # #17263E Header surface
C_BORDER = RGBColor(38, 56, 84)        # #263854 Subtle border
C_CYAN = RGBColor(0, 212, 255)         # #00D4FF Cyan Accent
C_TEAL = RGBColor(0, 229, 176)         # #00E5B0 Teal Accent
C_WHITE = RGBColor(255, 255, 255)      # #FFFFFF Main White
C_GRAY = RGBColor(170, 184, 204)       # #AAB8CC Secondary Text Gray
C_AMBER = RGBColor(255, 183, 3)        # #FFB703 Highlight Gold
C_CORAL = RGBColor(255, 107, 107)      # #FF6B6B High Risk Alert

FONT_MAIN = "Segoe UI"


def generate_slide_charts(lang="vi"):
    is_vi = lang == "vi"
    out_dir = FIG_VI if is_vi else FIG_EN
    plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    # 1. Slide 5 Chart: Burnout vs Dropout
    fig, ax = plt.subplots(figsize=(7.8, 4.6), dpi=220)
    fig.patch.set_facecolor("#111C2E")
    ax.set_facecolor("#111C2E")

    tiers = (
        ["Low Burnout\n(Kiệt sức Thấp)", "Medium Burnout\n(Kiệt sức Vừa)", "High Burnout\n(Kiệt sức Cao)"]
        if is_vi
        else ["Low Burnout\n(n = 320)", "Medium Burnout\n(n = 280)", "High Burnout\n(n = 200)"]
    )
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
        safe_lbl = f"An toàn:\n{safe:.1f}%" if is_vi else f"Retained:\n{safe:.1f}%"
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() / 2,
            safe_lbl,
            ha="center",
            va="center",
            fontsize=11,
            fontweight="bold",
            color="#0B1220",
        )

    ax.set_ylim(0, 110)
    ax.set_ylabel("Tỷ lệ Bỏ học Quan sát (%)" if is_vi else "Observed Dropout Rate (%)", fontsize=13, fontweight="bold", color="#AAB8CC")
    ax.set_title(
        "Tỷ lệ Sinh viên Bỏ học theo Mức độ Kiệt sức (Burnout_Level)" if is_vi else "Observed Student Dropout Rate Across Burnout Tiers",
        fontsize=14.5,
        fontweight="bold",
        color="#FFFFFF",
        pad=14,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(tiers, fontsize=12.5, fontweight="bold", color="#FFFFFF")
    ax.tick_params(axis="y", colors="#AAB8CC", labelsize=11)
    ax.grid(axis="y", linestyle="--", alpha=0.2, color="#AAB8CC", zorder=0)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    for spine in ["bottom", "left"]:
        ax.spines[spine].set_color("#263854")

    plt.tight_layout()
    s5_chart = os.path.join(out_dir, "short_s5_burnout_chart.png")
    plt.savefig(s5_chart, dpi=220, facecolor=fig.get_facecolor())
    plt.close(fig)

    # 2. Slide 6 Chart: 4 Models Comparison
    fig, ax = plt.subplots(figsize=(7.8, 4.6), dpi=220)
    fig.patch.set_facecolor("#111C2E")
    ax.set_facecolor("#111C2E")

    models = ["Logistic\nRegression", "XGBoost", "Soft Voting\nEnsemble", "Random\nForest"]
    s1_auc = [87.61, 88.64, 88.89, 89.01]
    s2_auc = [87.25, 88.26, 88.25, 87.97]

    x = np.arange(len(models))
    w = 0.34
    lbl1 = "Kịch bản 1: Có Burnout_Level (5-Fold CV)" if is_vi else "Scenario 1: With Burnout_Level (5-Fold CV)"
    lbl2 = "Kịch bản 2: Không có Burnout_Level (5-Fold CV)" if is_vi else "Scenario 2: Without Burnout_Level (5-Fold CV)"
    b1 = ax.bar(x - w / 2, s1_auc, width=w, label=lbl1, color="#00D4FF", edgecolor="#FFFFFF", linewidth=0.8, zorder=3)
    b2 = ax.bar(x + w / 2, s2_auc, width=w, label=lbl2, color="#00E5B0", edgecolor="#FFFFFF", linewidth=0.8, zorder=3)

    for bar in b1:
        val = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, val + 0.18, f"{val:.2f}%", ha="center", va="bottom", fontsize=11.5, fontweight="bold", color="#00D4FF")
    for bar in b2:
        val = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, val + 0.18, f"{val:.2f}%", ha="center", va="bottom", fontsize=11.5, fontweight="bold", color="#00E5B0")

    ax.set_ylim(84.5, 90.8)
    ax.set_ylabel("5-Fold Stratified CV ROC-AUC (%)", fontsize=12.5, fontweight="bold", color="#AAB8CC")
    ax.set_title(
        "So sánh ROC-AUC (5-Fold CV) của 4 Mô hình Học máy" if is_vi else "5-Fold CV ROC-AUC Comparison Across 4 ML Models",
        fontsize=14.5,
        fontweight="bold",
        color="#FFFFFF",
        pad=12,
    )
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
    s6_chart = os.path.join(out_dir, "short_s6_model_chart.png")
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


def add_card(slide, left, top, width, height, title, bullets, accent=C_CYAN, body_size=14.0, space_pt=8):
    add_rect(slide, left, top, width, height, C_CARD_BG, border_rgb=accent, border_pt=1.5, rounded=True)
    add_rect(slide, left, top, width, 0.48, C_CARD_HDR, rounded=False)
    add_rect(slide, left, top + 0.46, width, 0.03, accent, rounded=False)
    add_text(slide, left + 0.16, top + 0.08, width - 0.32, 0.34, title, font_size=13.5, bold=True, color=accent)

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


def create_base_slide(prs, slide_num, total_slides, badge_text, title_text, takeaway_text, is_vi=True):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, 13.333, 7.5, C_DARK_BG)

    add_rect(slide, 0, 0, 13.333, 1.18, C_CARD_BG)
    add_rect(slide, 0, 1.16, 13.333, 0.03, C_CYAN)
    add_rect(slide, 0, 0, 0.12, 1.18, C_TEAL)

    add_rect(slide, 0.38, 0.12, 3.10, 0.27, C_CARD_HDR, border_rgb=C_TEAL, border_pt=1.0, rounded=True)
    add_text(slide, 0.40, 0.13, 3.06, 0.25, badge_text, font_size=9.5, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    add_text(slide, 0.38, 0.42, 12.5, 0.42, title_text, font_size=22, bold=True, color=C_WHITE)
    add_text(slide, 0.38, 0.82, 12.5, 0.28, takeaway_text, font_size=12.5, italic=True, color=C_CYAN)

    add_rect(slide, 0, 7.12, 13.333, 0.38, C_CARD_BG)
    ftr = (
        "SAMSUNG INNOVATION CAMPUS (SIC) — NHÓM 03: ĐỖ QUANG MINH • VŨ QUANG HUY"
        if is_vi
        else "SAMSUNG INNOVATION CAMPUS (SIC) — GROUP 03: DO QUANG MINH • VU QUANG HUY"
    )
    add_text(slide, 0.38, 7.17, 10.0, 0.26, ftr, font_size=9.5, bold=True, color=C_GRAY)
    add_text(slide, 11.2, 7.17, 1.75, 0.26, f"Slide {slide_num:02d} / {total_slides:02d}", font_size=10, bold=True, color=C_CYAN, align=PP_ALIGN.RIGHT)
    return slide


def build_deck_9_slides(lang="vi"):
    is_vi = lang == "vi"
    s5_chart, s6_chart = generate_slide_charts(lang)
    total_slides = 9
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # SLIDE 1 - COVER (20s)
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s1, 0, 0, 13.333, 7.5, C_DARK_BG)
    add_rect(s1, 0, 0, 0.22, 7.5, C_CYAN)
    add_rect(s1, 0.22, 0, 0.10, 7.5, C_TEAL)

    add_rect(s1, 0.90, 0.80, 4.80, 0.44, C_CARD_BG, border_rgb=C_TEAL, border_pt=1.5, rounded=True)
    add_text(s1, 0.95, 0.85, 4.70, 0.34, "SAMSUNG INNOVATION CAMPUS (SIC)", font_size=12, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    title_text = (
        "XÂY DỰNG, HUẤN LUYỆN VÀ SO SÁNH CÁC MÔ HÌNH HỌC MÁY\nDỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG VÀ THAM GIA HỌC TẬP"
        if is_vi
        else "BUILDING, TRAINING AND COMPARING MACHINE LEARNING MODELS\nFOR STUDENT DROPOUT PREDICTION FROM STRESS & ENGAGEMENT"
    )
    add_text(s1, 0.90, 1.55, 11.6, 1.65, title_text, font_size=26.5, bold=True, color=C_WHITE)
    sub_text = (
        "Đồ án Tốt nghiệp AI Capstone — Hệ thống Cảnh báo Sớm Nguy cơ Bỏ học"
        if is_vi
        else "AI Capstone Project — Proactive Student Burnout & Dropout Early Warning System"
    )
    add_text(s1, 0.90, 3.25, 11.0, 0.45, sub_text, font_size=16, italic=True, color=C_CYAN)

    add_rect(s1, 0.90, 4.10, 11.5, 2.20, C_CARD_BG, border_rgb=C_CYAN, border_pt=1.5, rounded=True)
    add_rect(s1, 0.90, 4.10, 11.5, 0.48, C_CARD_HDR, rounded=False)
    add_text(s1, 1.15, 4.18, 11.0, 0.34, "NHÓM THỰC HIỆN: NHÓM 03" if is_vi else "PRESENTED BY: GROUP 03", font_size=14, bold=True, color=C_CYAN)

    add_text(s1, 1.15, 4.80, 5.4, 0.55, "• Đỗ Quang Minh" if is_vi else "• Do Quang Minh", font_size=20, bold=True, color=C_WHITE)
    add_text(s1, 1.45, 5.32, 5.1, 0.45, "Trưởng nhóm (Team Leader)", font_size=14, color=C_TEAL)

    add_text(s1, 6.80, 4.80, 5.4, 0.55, "• Vũ Quang Huy" if is_vi else "• Vu Quang Huy", font_size=20, bold=True, color=C_WHITE)
    add_text(s1, 7.10, 5.32, 5.1, 0.45, "Thành viên (Team Member)", font_size=14, color=C_TEAL)
    add_text(s1, 11.2, 7.05, 1.75, 0.30, "Slide 01 / 09", font_size=10, bold=True, color=C_GRAY, align=PP_ALIGN.RIGHT)

    # SLIDE 2 - PROBLEM WE SOLVE (40s)
    s2 = create_base_slide(
        prs, 2, total_slides,
        "01 BÀI TOÁN NGHIÊN CỨU" if is_vi else "01 RESEARCH PROBLEM",
        "Bài Toán Chúng Tôi Giải Quyết" if is_vi else "The Problem We Solve",
        "Phát hiện sớm nguy cơ bỏ học ngay từ khi sinh viên mới xuất hiện dấu hiệu căng thẳng và giảm gắn kết học tập"
        if is_vi else "Detecting dropout risk early from declining academic engagement and rising psychological stress",
        is_vi=is_vi,
    )
    flow_steps = (
        [
            ("1. NGUYÊN NHÂN TÍCH LŨY", "Căng thẳng tăng cao\n+ Gắn kết học tập suy giảm", C_CYAN),
            ("2. TRẠNG THÁI TRUNG GIAN", "Nguy cơ Kiệt sức Học đường\n(Academic Burnout)", C_AMBER),
            ("3. HỆ QUẢ CUỐI CÙNG", "Nguy cơ Bỏ học Đại học\n(Student Dropout Risk)", C_CORAL),
        ]
        if is_vi
        else [
            ("1. CUMULATIVE DRIVERS", "Rising Psychological Stress\n+ Declining Study Engagement", C_CYAN),
            ("2. MEDIATING STATE", "Academic Burnout Risk\n(Low / Medium / High)", C_AMBER),
            ("3. FINAL OUTCOME", "University Dropout Risk\n(Student Attrition)", C_CORAL),
        ]
    )
    for idx, (f_hdr, f_body, col) in enumerate(flow_steps):
        x = 0.40 + idx * 4.35
        add_rect(s2, x, 1.42, 3.85, 1.85, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s2, x, 1.42, 3.85, 0.42, C_CARD_HDR, rounded=False)
        add_text(s2, x + 0.12, 1.48, 3.60, 0.30, f_hdr, font_size=12.5, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s2, x + 0.15, 2.00, 3.55, 1.10, f_body, font_size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        if idx < 2:
            add_text(s2, x + 3.88, 2.05, 0.44, 0.60, "➔", font_size=26, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)

    q_cards = (
        [
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
        if is_vi
        else [
            ("1. WHY DO STUDENTS DROP OUT?", [
                "• Not a sudden, isolated single-day event.",
                "• Cumulative result of prolonged pressure.",
                "• Loss of balance between study and life.",
            ], C_CYAN),
            ("2. IMPACT OF STRESS & ENGAGEMENT", [
                "• Low attendance, fewer study hours, backlogs.",
                "• Rising stress, anxiety, and financial pressure.",
                "→ Pushes students into academic burnout.",
            ], C_AMBER),
            ("3. WHY PROACTIVE EARLY WARNING?", [
                "• Waiting until withdrawal submission is too late.",
                "• AI identifies at-risk students proactively.",
                "→ Enables timely tutoring & counseling support.",
            ], C_TEAL),
        ]
    )
    for idx, (q_t, q_b, col) in enumerate(q_cards):
        x = 0.40 + idx * 4.22
        add_card(s2, x, 3.55, 4.08, 3.30, q_t, q_b, accent=col, body_size=14.0, space_pt=10)

    # SLIDE 3 - INPUT DATA (40s)
    s3 = create_base_slide(
        prs, 3, total_slides,
        "02 DỮ LIỆU ĐẦU VÀO" if is_vi else "02 INPUT DATASET",
        "Cấu Trúc Bộ Dữ Liệu & Quy Trình Mở Rộng Mẫu" if is_vi else "Dataset Architecture & Sample Expansion Provenance",
        "800 bản ghi gốc (25 thuộc tính) được làm sạch và mở rộng thống kê lên 2.000 hồ sơ (34 thuộc tính)"
        if is_vi else "800 raw records (25 attributes) cleaned and statistically expanded to 2,000 profiles (34 attributes)",
        is_vi=is_vi,
    )
    s3_kpis = (
        [
            ("800", "Bản ghi ban đầu\n(Dữ liệu gốc có 612 ô khuyết)", C_CYAN),
            ("25", "Thuộc tính gốc\n(18 cột có giá trị thiếu)", C_TEAL),
            ("2.000", "Hồ sơ sau xử lý & mở rộng\n(Sạch 100%, 0 giá trị NaN)", C_AMBER),
            ("34", "Thuộc tính hoàn chỉnh\n(Thêm 6 đặc trưng & mã hóa)", C_CYAN),
        ]
        if is_vi
        else [
            ("800", "Initial Raw Records\n(612 missing cells total)", C_CYAN),
            ("25", "Original Attributes\n(18 incomplete columns)", C_TEAL),
            ("2,000", "Processed & Expanded Profiles\n(100% clean, 0 NaN values)", C_AMBER),
            ("34", "Final Attributes\n(+6 engineered & encoded cols)", C_CYAN),
        ]
    )
    for idx, (val, lbl, col) in enumerate(s3_kpis):
        x = 0.40 + idx * 3.18
        add_rect(s3, x, 1.40, 2.98, 1.65, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_text(s3, x, 1.50, 2.98, 0.65, val, font_size=28, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s3, x + 0.10, 2.22, 2.78, 0.72, lbl, font_size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    add_card(
        s3, 0.40, 3.28, 6.15, 3.60,
        "CÁC NHÓM DỮ LIỆU ĐẦU VÀO" if is_vi else "INPUT ATTRIBUTE DOMAINS",
        [
            "• 1. Nhân khẩu học: Độ tuổi, Giới tính, Năm học, Khoa, Cư trú, Thu nhập.",
            "• 2. Học tập: Chuyên cần (%), Giờ tự học/ngày, GPA kỳ trước, Số môn nợ, Động lực.",
            "• 3. Lối sống: Giờ ngủ, Thời gian dùng màn hình, Thể dục, Hoạt động xã hội, Làm thêm.",
            "• 4. Căng thẳng & Hỗ trợ: Mức căng thẳng, Lo âu, Tài chính, Bạn bè, Gia đình.",
            "→ Biến mục tiêu: Burnout_Level (3 mức) & Dropout_Risk (Yes 54.6% / No 45.4%).",
        ] if is_vi else [
            "• 1. Demographics: Age, Gender, Year of Study, Department, Residence, Income.",
            "• 2. Academic: Attendance (%), Study Hours/Day, Previous GPA, Backlogs, Motivation.",
            "• 3. Lifestyle: Sleep Hours, Screen Time, Exercise, Social Activity, Part-Time Job.",
            "• 4. Stress & Support: Stress Level, Anxiety, Financial, Peer Pressure, Family Support.",
            "→ Targets: Burnout_Level (3 tiers) & Dropout_Risk (Yes 54.6% / No 45.4%).",
        ],
        accent=C_TEAL, body_size=13.2, space_pt=7
    )
    add_card(
        s3, 6.78, 3.28, 6.15, 3.60,
        "NGUỒN GỐC 2.000 HỒ SƠ (MINH BẠCH KHOA HỌC)" if is_vi else "PROVENANCE OF 2,000 PROFILES (SCIENTIFIC CLARITY)",
        [
            "★ Bước 1: Làm sạch 800 bản ghi gốc",
            "  • Điền khuyết KNNImputer (k=7) + lọc ngoại lai IQR → 800 hồ sơ sạch.",
            "★ Bước 2: Mở rộng thêm 1.200 bản ghi tổng hợp",
            "  • Sinh mẫu từ phân phối chuẩn đa biến có phân tầng (Multivariate Gaussian) khớp theo từng nhóm (Burnout × Dropout) của 800 mẫu gốc.",
            "→ Lưu ý học thuật: 2.000 hồ sơ là bộ dữ liệu mở rộng thống kê phục vụ huấn luyện mô hình, không phải khảo sát thực địa đợt 2.",
        ] if is_vi else [
            "★ Step 1: Clean 800 original raw records",
            "  • Impute via KNNImputer (k=7) + clip IQR outliers → 800 clean records.",
            "★ Step 2: Augment with 1,200 synthetic records",
            "  • Sampled from stratified Multivariate Gaussian distributions fitted per (Burnout × Dropout) stratum of the 800 cleaned records.",
            "→ Academic Note: The 2,000 dataset is statistically augmented for ML training, not a second field survey.",
        ],
        accent=C_AMBER, body_size=13.0, space_pt=6
    )

    # SLIDE 4 - AI PIPELINE WORKFLOW (50s)
    s4 = create_base_slide(
        prs, 4, total_slides,
        "03 QUY TRÌNH XÂY DỰNG AI" if is_vi else "03 AI DEVELOPMENT WORKFLOW",
        "Quy Trình Xây Dựng & Huấn Luyện Mô Hình Học Máy" if is_vi else "End-to-End Machine Learning Pipeline Workflow",
        "Đóng gói khép kín toàn bộ các bước trong Pipeline để ngăn ngừa rò rỉ dữ liệu (Data Leakage)"
        if is_vi else "Encapsulating all preprocessing and engineering steps inside Pipeline to prevent Data Leakage",
        is_vi=is_vi,
    )
    steps_5 = (
        [
            ("Bước 1", "Data Cleaning", "Làm sạch & Lọc ngoại lai IQR", C_CYAN),
            ("Bước 2", "Preprocessing", "KNNImputer (k=7) & Chuẩn hóa", C_TEAL),
            ("Bước 3", "Feature Engineering", "Tạo 6 Đặc trưng tương tác mới", C_AMBER),
            ("Bước 4", "Model Training", "Huấn luyện 4 Mô hình Học máy", C_CYAN),
            ("Bước 5", "Model Evaluation", "Stratified 5-Fold CV & Test 20%", C_TEAL),
        ]
        if is_vi
        else [
            ("Step 1", "Data Cleaning", "Deduplicate & IQR Outlier Clipping", C_CYAN),
            ("Step 2", "Preprocessing", "KNNImputer (k=7) & Scaling", C_TEAL),
            ("Step 3", "Feature Engineering", "6 Domain Interaction Indices", C_AMBER),
            ("Step 4", "Model Training", "Train 4 Flagship ML Models", C_CYAN),
            ("Step 5", "Model Evaluation", "Stratified 5-Fold CV & 20% Test", C_TEAL),
        ]
    )
    for idx, (st_num, st_en, st_vi, col) in enumerate(steps_5):
        x = 0.40 + idx * 2.54
        add_rect(s4, x, 1.45, 2.30, 1.85, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_rect(s4, x, 1.45, 2.30, 0.40, C_CARD_HDR, rounded=False)
        add_text(s4, x + 0.08, 1.50, 2.14, 0.30, st_num, font_size=12, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.08, 1.95, 2.14, 0.48, st_en, font_size=14.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.10, 2.50, 2.10, 0.65, st_vi, font_size=11.5, color=C_GRAY, align=PP_ALIGN.CENTER)
        if idx < 4:
            add_text(s4, x + 2.31, 2.10, 0.22, 0.45, "➔", font_size=16, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

    pillars = (
        [
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
        if is_vi
        else [
            ("1. KNNIMPUTER (k=7) IN PIPELINE", [
                "• Estimates missing values from the 7 most similar student profiles.",
                "• Fitted strictly inside the CV Pipeline to prevent Train-Test leakage.",
            ], C_CYAN),
            ("2. SIX ENGINEERED FEATURES", [
                "• Combines raw variables to quantify the Stress vs. Engagement balance.",
                "• Key indices: Engagement Index, Stress Index & Stress-to-Engagement Ratio.",
            ], C_AMBER),
            ("3. STRATIFIED 5-FOLD CV", [
                "• Evaluates models across 5 rotating validation folds objectively.",
                "• Preserves exact class balance (54.6% Yes / 45.4% No) in every fold.",
            ], C_TEAL),
        ]
    )
    for idx, (p_t, p_b, col) in enumerate(pillars):
        x = 0.40 + idx * 4.22
        add_card(s4, x, 3.60, 4.08, 3.25, p_t, p_b, accent=col, body_size=14.0, space_pt=10)

    # SLIDE 5 - DATA ANALYSIS / EDA (45s)
    s5 = create_base_slide(
        prs, 5, total_slides,
        "04 PHÂN TÍCH DỮ LIỆU (EDA)" if is_vi else "04 EXPLORATORY DATA ANALYSIS",
        "Phát Hiện Quan Trọng Nhất: Tỷ Lệ Bỏ Học Theo Mức Độ Burnout" if is_vi else "Key EDA Finding: Observed Dropout Rate by Burnout Level",
        "Mức độ kiệt sức học đường càng cao thì tỷ lệ bỏ học quan sát được trong bộ dữ liệu càng lớn"
        if is_vi else "Higher observed academic burnout levels correspond to substantially higher dropout rates in the dataset",
        is_vi=is_vi,
    )
    add_image_box(s5, s5_chart, 0.40, 1.40, 7.65, 5.45, bg_rgb=C_CARD_BG, border_rgb=C_CYAN, pad=0.06)

    tiers_info = (
        [
            ("MỨC THẤP (LOW BURNOUT)", "23.8% Bỏ học", "76 / 320 SV bỏ học", C_TEAL),
            ("MỨC VỪA (MEDIUM BURNOUT)", "63.2% Bỏ học", "177 / 280 SV bỏ học", C_AMBER),
            ("MỨC CAO (HIGH BURNOUT)", "92.0% Bỏ học", "184 / 200 SV bỏ học", C_CORAL),
        ]
        if is_vi
        else [
            ("LOW BURNOUT TIER", "23.8% Dropout", "76 / 320 dropped out", C_TEAL),
            ("MEDIUM BURNOUT TIER", "63.2% Dropout", "177 / 280 dropped out", C_AMBER),
            ("HIGH BURNOUT TIER", "92.0% Dropout", "184 / 200 dropped out", C_CORAL),
        ]
    )
    for idx, (t_hdr, t_val, t_sub, col) in enumerate(tiers_info):
        y = 1.40 + idx * 1.42
        add_rect(s5, 8.30, y, 4.63, 1.28, C_CARD_BG, border_rgb=col, border_pt=2.0, rounded=True)
        add_text(s5, 8.48, y + 0.10, 4.30, 0.28, t_hdr, font_size=12, bold=True, color=col)
        add_text(s5, 8.48, y + 0.40, 2.30, 0.50, t_val, font_size=22, bold=True, color=C_WHITE)
        add_text(s5, 10.65, y + 0.48, 2.15, 0.45, t_sub, font_size=12, bold=True, color=C_GRAY, align=PP_ALIGN.RIGHT)

    add_rect(s5, 8.30, 5.68, 4.63, 1.17, C_CARD_HDR, border_rgb=C_CYAN, border_pt=1.2, rounded=True)
    note_txt = (
        "💡 Lưu ý khoa học:\nSố liệu phản ánh mối liên hệ thống kê quan sát được trong dữ liệu (r = +0.550), không tự kết luận là quan hệ nhân quả một chiều."
        if is_vi
        else "💡 Scientific Note:\nReflects observed statistical association in the dataset (r = +0.550), not a standalone causal claim."
    )
    add_text(s5, 8.45, 5.76, 4.35, 1.00, note_txt, font_size=11.5, italic=True, color=C_WHITE)

    # SLIDE 6 - MODEL TRAINING RESULTS (60s)
    s6 = create_base_slide(
        prs, 6, total_slides,
        "05 KẾT QUẢ MÔ HÌNH HỌC MÁY" if is_vi else "05 MODEL BENCHMARKING RESULTS",
        "So Sánh 4 Mô Hình Học Máy Trên 2 Kịch Bản (5-Fold CV)" if is_vi else "Comparing 4 Machine Learning Models Across 2 Scenarios (5-Fold CV)",
        "Random Forest dẫn đầu Kịch bản 1 (89.01% CV ROC-AUC); XGBoost dẫn đầu Kịch bản 2 Cảnh báo sớm (88.26% CV ROC-AUC)"
        if is_vi else "Random Forest leads Scenario 1 (89.01% CV ROC-AUC); XGBoost leads Scenario 2 Early Warning (88.26% CV ROC-AUC)",
        is_vi=is_vi,
    )
    add_image_box(s6, s6_chart, 0.40, 1.40, 7.65, 5.45, bg_rgb=C_CARD_BG, border_rgb=C_CYAN, pad=0.06)

    add_card(
        s6, 8.30, 1.40, 4.63, 2.60,
        "KỊCH BẢN 1: CÓ BURNOUT_LEVEL" if is_vi else "SCENARIO 1: WITH BURNOUT_LEVEL",
        [
            "★ Mô hình dẫn đầu: Random Forest",
            "• 5-Fold CV ROC-AUC: 89.01%",
            "• Hold-out Test (20%): ROC-AUC 88.42% | Acc 81.50% | F1 83.63%",
            "→ Phù hợp khi nhà trường đã có đánh giá mức độ kiệt sức.",
        ] if is_vi else [
            "★ Leading Model: Random Forest",
            "• 5-Fold CV ROC-AUC: 89.01%",
            "• Hold-out Test (20%): ROC-AUC 88.42% | Acc 81.50% | F1 83.63%",
            "→ Optimal when burnout screening scores are available.",
        ],
        accent=C_CYAN, body_size=13.0, space_pt=6
    )
    add_card(
        s6, 8.30, 4.22, 4.63, 2.63,
        "KỊCH BẢN 2: KHÔNG CÓ BURNOUT_LEVEL" if is_vi else "SCENARIO 2: WITHOUT BURNOUT_LEVEL",
        [
            "★ Mô hình dẫn đầu: XGBoost",
            "• 5-Fold CV ROC-AUC: 88.26%",
            "• Hold-out Test (20%): ROC-AUC 87.39% | Acc 78.50% | F1 80.28%",
            "→ Cảnh báo sớm hiệu quả ngay từ các chỉ số học tập & căng thẳng.",
        ] if is_vi else [
            "★ Leading Model: XGBoost",
            "• 5-Fold CV ROC-AUC: 88.26%",
            "• Hold-out Test (20%): ROC-AUC 87.39% | Acc 78.50% | F1 80.28%",
            "→ Effective early warning directly from study & stress metrics.",
        ],
        accent=C_TEAL, body_size=13.0, space_pt=6
    )

    # SLIDE 7 - WEB DASHBOARD PRODUCT (60s)
    s7 = create_base_slide(
        prs, 7, total_slides,
        "06 SẢN PHẨM ỨNG DỤNG THỰC TẾ" if is_vi else "06 WEB DASHBOARD PRODUCT",
        "Ứng Dụng Web Dashboard Cảnh Báo Sớm (FastAPI)" if is_vi else "Interactive Early Warning Web Dashboard (FastAPI)",
        "Giao diện thực tế (app.py) hỗ trợ cố vấn học tập nhập thông tin, nhận dự báo kép và khuyến nghị tức thì"
        if is_vi else "Real FastAPI dashboard (app.py) providing dual burnout/dropout predictions and personalized advising actions",
        is_vi=is_vi,
    )
    w_steps = (
        [
            ("1. INPUT", "Nhập 18 chỉ số sinh viên hoặc chọn 3 mẫu thử nhanh", C_CYAN),
            ("2. AI PREDICTION", "Phân tích đặc trưng & Chạy mô hình dự báo kép", C_AMBER),
            ("3. RESULT & RECOMMENDATION", "Hiển thị mức Burnout, xác suất Dropout & Khuyến nghị", C_TEAL),
        ]
        if is_vi
        else [
            ("1. INPUT", "Enter 18 student indicators or select 3 quick presets", C_CYAN),
            ("2. AI PREDICTION", "Compute domain indices & run dual ML inference", C_AMBER),
            ("3. RESULT & RECOMMENDATION", "Display Burnout tier, Dropout risk % & Advising actions", C_TEAL),
        ]
    )
    for idx, (ws_t, ws_d, col) in enumerate(w_steps):
        x = 0.40 + idx * 4.30
        add_rect(s7, x, 1.35, 3.93, 0.92, C_CARD_BG, border_rgb=col, border_pt=1.8, rounded=True)
        add_text(s7, x + 0.12, 1.40, 3.68, 0.30, ws_t, font_size=12.5, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s7, x + 0.12, 1.72, 3.68, 0.48, ws_d, font_size=11.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        if idx < 2:
            add_text(s7, x + 3.95, 1.55, 0.32, 0.45, "➔", font_size=18, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

    live_shot = os.path.join(FIG_VI, "13_live_web_dashboard_screenshot.png")
    add_image_box(s7, live_shot, 0.40, 2.42, 7.85, 4.45, bg_rgb=RGBColor(255, 255, 255), border_rgb=C_CYAN, pad=0.04)

    add_card(
        s7, 8.45, 2.42, 4.48, 4.45,
        "5 BƯỚC HOẠT ĐỘNG TRÊN WEB APP" if is_vi else "5 WORKFLOW STEPS ON WEB APP",
        [
            "1. Nhập thông tin sinh viên: Chuyên cần, giờ học, GPA, căng thẳng, giấc ngủ...",
            "2. Phân tích dữ liệu: Tự động tính Chỉ số Gắn kết & Chỉ số Căng thẳng.",
            "3. Dự báo mức độ Burnout: Hiển thị xác suất Low / Medium / High.",
            "4. Dự báo nguy cơ Dropout: Cảnh báo phần trăm rủi ro bỏ học.",
            "5. Khuyến nghị hỗ trợ: Đề xuất phụ đạo học tập, tham vấn tâm lý hoặc hỗ trợ tài chính.",
        ] if is_vi else [
            "1. Input Student Data: Attendance, study hours, GPA, stress, sleep...",
            "2. Feature Analysis: Auto-computes Engagement & Stress indices.",
            "3. Burnout Prediction: Displays Low / Medium / High probabilities.",
            "4. Dropout Risk Prediction: Outputs calibrated attrition probability (%).",
            "5. Support Recommendations: Suggests tutoring, counseling, or financial aid.",
        ],
        accent=C_TEAL, body_size=12.8, space_pt=6
    )

    # SLIDE 8 - CONCLUSION & CONTRIBUTIONS (40s)
    s8 = create_base_slide(
        prs, 8, total_slides,
        "07 KẾT LUẬN & HƯỚNG PHÁT TRIỂN" if is_vi else "07 CONCLUSION & FUTURE WORK",
        "Tổng Kết Đóng Góp & Định Hướng Phát Triển" if is_vi else "Key Contributions & Future Development Roadmap",
        "Hoàn thiện quy trình từ xử lý dữ liệu, huấn luyện mô hình học máy đến ứng dụng Web hỗ trợ cảnh báo sớm"
        if is_vi else "End-to-end delivery from data preprocessing and ML benchmarking to an interactive early warning Web App",
        is_vi=is_vi,
    )
    add_card(
        s8, 0.40, 1.45, 6.15, 5.35,
        "3 ĐÓNG GÓP CHÍNH CỦA ĐỒ ÁN" if is_vi else "3 CORE PROJECT CONTRIBUTIONS",
        [
            "★ 1. Quy trình xử lý & phân tích dữ liệu bài bản",
            "  • Làm sạch 800 bản ghi gốc, điền khuyết KNNImputer (k=7) trong Pipeline và thiết kế 6 đặc trưng tương tác mới.",
            "★ 2. Huấn luyện & so sánh 4 mô hình Machine Learning",
            "  • Đánh giá trên 2 kịch bản (5-Fold CV): Random Forest đạt 89.01% ROC-AUC (Kịch bản 1) và XGBoost đạt 88.26% (Kịch bản 2).",
            "★ 3. Xây dựng ứng dụng Web cảnh báo sớm",
            "  • Triển khai Web Dashboard trực quan giúp dự báo đồng thời Burnout, Dropout và gợi ý hướng hỗ trợ sinh viên.",
        ] if is_vi else [
            "★ 1. Rigorous Data Processing & Analysis Pipeline",
            "  • Cleaned 800 raw records, embedded KNNImputer (k=7) inside Pipeline, and engineered 6 interaction features.",
            "★ 2. Benchmarking 4 Machine Learning Models",
            "  • Evaluated across 2 scenarios (5-Fold CV): Random Forest achieved 89.01% ROC-AUC (Scenario 1) & XGBoost achieved 88.26% (Scenario 2).",
            "★ 3. Interactive Early Warning Web Application",
            "  • Built a FastAPI Web Dashboard predicting both Burnout and Dropout risk with actionable recommendations.",
        ],
        accent=C_TEAL, body_size=13.5, space_pt=8
    )
    add_card(
        s8, 6.78, 1.45, 6.15, 5.35,
        "3 HƯỚNG PHÁT TRIỂN TIẾP THEO" if is_vi else "3 FUTURE DEVELOPMENT DIRECTIONS",
        [
            "→ 1. Bổ sung dữ liệu khảo sát thực tế từ nhiều trường",
            "  • Thu thập và kiểm định ngoại kiểm (External Validation) trên dữ liệu sinh viên thực tế tại các trường đại học.",
            "→ 2. Tích hợp giải thích mô hình cục bộ với SHAP",
            "  • Hiển thị trực tiếp mức độ tác động của từng yếu tố lên kết quả dự báo của mỗi cá nhân sinh viên trên Web.",
            "→ 3. Kết nối hệ thống Quản lý Học tập (LMS)",
            "  • Tự động đồng bộ dữ liệu chuyên cần và điểm số định kỳ để sàng lọc sớm theo từng học kỳ.",
        ] if is_vi else [
            "→ 1. Multi-Institution Real-World Data Validation",
            "  • Collect real survey datasets across multiple universities for external validation.",
            "→ 2. Local Explainability with SHAP Integration",
            "  • Render per-student SHAP waterfall explanations directly inside the Web Dashboard.",
            "→ 3. Learning Management System (LMS) Integration",
            "  • Connect with institutional LMS APIs for automated semester-by-semester risk screening.",
        ],
        accent=C_CYAN, body_size=13.5, space_pt=8
    )

    # SLIDE 9 - THANK YOU & Q&A (15s)
    s9 = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s9, 0, 0, 13.333, 7.5, C_DARK_BG)
    add_rect(s9, 0, 0, 0.22, 7.5, C_TEAL)
    add_rect(s9, 0.22, 0, 0.10, 7.5, C_CYAN)

    add_rect(s9, 1.60, 1.35, 10.13, 4.80, C_CARD_BG, border_rgb=C_CYAN, border_pt=2.0, rounded=True)
    add_text(s9, 1.80, 1.95, 9.73, 0.90, "THANK YOU FOR LISTENING", font_size=36, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_text(s9, 1.80, 2.95, 9.73, 0.85, "Q & A", font_size=32, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
    add_text(
        s9, 1.80, 4.15, 9.73, 0.55,
        "Nhóm 03 – Samsung Innovation Campus (SIC)" if is_vi else "Group 03 – Samsung Innovation Campus (SIC)",
        font_size=18, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER
    )
    add_text(
        s9, 1.80, 4.75, 9.73, 0.50,
        "Đỗ Quang Minh (Trưởng nhóm)  •  Vũ Quang Huy (Thành viên)" if is_vi else "Do Quang Minh (Team Leader)  •  Vu Quang Huy (Team Member)",
        font_size=15, color=C_GRAY, align=PP_ALIGN.CENTER
    )
    add_text(s9, 11.2, 7.05, 1.75, 0.30, "Slide 09 / 09", font_size=10, bold=True, color=C_GRAY, align=PP_ALIGN.RIGHT)

    out_folder = VI_DIR if is_vi else EN_DIR
    out_path = os.path.join(out_folder, "SIC_AI_Capstone Project_Presentation Slide Template.pptx")
    prs.save(out_path)
    print(f"Updated 9-Slide Presentation ({lang.upper()}) -> {out_path}")

    # Also update Drive folder if accessible
    if os.path.exists(DRV_DIR):
        sub_drv = os.path.join(DRV_DIR, "Ban_Tieng_Viet_Full" if is_vi else "Ban_Tieng_Anh_Full")
        if os.path.exists(sub_drv):
            try:
                prs.save(os.path.join(sub_drv, "SIC_AI_Capstone Project_Presentation Slide Template.pptx"))
            except Exception:
                pass


def clean_redundant_files():
    """Ensures SIC_Documents_VI and SIC_Documents_EN contain ONLY the 5 main SIC report files."""
    MAIN_5_FILES = {
        "SIC_AI_Capstone Project_Action Plan.docx",
        "SIC_AI_Capstone Project_Evaluation Sheet.docx",
        "SIC_AI_Capstone Project_Final Report.docx",
        "SIC_AI_Capstone Project_Presentation Slide Template.pptx",
        "SIC_AI_Capstone Project_Work Breakdown Structure.xlsx",
    }

    for folder in [VI_DIR, EN_DIR]:
        if not os.path.exists(folder):
            continue
        for fname in os.listdir(folder):
            if fname not in MAIN_5_FILES:
                fpath = os.path.join(folder, fname)
                try:
                    if os.path.isdir(fpath):
                        shutil.rmtree(fpath)
                    else:
                        os.remove(fpath)
                    print(f"Removed redundant file from {os.path.basename(folder)}: {fname}")
                except Exception as e:
                    print(f"Could not remove {fpath}: {e}")

    # Also clean extra root folders/files if present
    for extra_dir in ["FINAL_PRESENTATION", "FINAL_SHORT_PRESENTATION"]:
        p = os.path.join(BASE_DIR, extra_dir)
        if os.path.exists(p):
            try:
                shutil.rmtree(p)
                print(f"Removed extra directory: {extra_dir}")
            except Exception as e:
                print(f"Could not remove {p}: {e}")

    for extra_file in ["0.26.0", "Kich_Ban_Thuyet_Trinh_Tung_Slide_Nhom03.docx", "Tai_Lieu_On_Tap_Va_Bao_Ve_Do_An_Nhom03.docx"]:
        p = os.path.join(BASE_DIR, extra_file)
        if os.path.exists(p):
            try:
                os.remove(p)
                print(f"Removed extra root file: {extra_file}")
            except Exception as e:
                print(f"Could not remove {p}: {e}")


if __name__ == "__main__":
    build_deck_9_slides("vi")
    build_deck_9_slides("en")
    clean_redundant_files()
    print("DONE! Both SIC_Documents_VI and SIC_Documents_EN now have ONLY the 5 main report files!")

