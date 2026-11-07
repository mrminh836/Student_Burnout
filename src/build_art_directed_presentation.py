"""
ROLE: SENIOR PRESENTATION DESIGNER & POWERPOINT ART DIRECTOR
Complete Visual Redesign of the SIC AI Capstone Presentation.

Key Art Direction Principles:
1. 12 Distinct Layout Archetypes across 12 slides (Zero repetitive templates):
   - Slide 01: Hero Cover Layout (Deep Navy asymmetric split + geometric vector KPI grid)
   - Slide 02: Timeline / Connected Roadmap Agenda Layout (Horizontal 6-node vector track)
   - Slide 03: Split-Screen + Causal Process Flow Layout (Left dark impact panel + Right 3-stage chevron flow)
   - Slide 04: Asymmetric Two-Column + Gantt Chevron Timeline Layout (Objectives, 50/50 Roles, 4 Phases)
   - Slide 05: Big-Number Strip + 6-Step Vector Pipeline Layout (800 -> 2,000 Dataset & Missing Data Chart)
   - Slide 06: Bento Infographic Grid Layout (6 Engineered Features with formula boxes & role pills)
   - Slide 07: Data Visualization + Stacked Big-Number Staircase Layout (EDA 1: Burnout -> Dropout)
   - Slide 08: Comparison / Tug-of-War Layout (EDA 2: Protective Shield vs. Risk Drivers + Heatmap)
   - Slide 09: Benchmark Chart + Executive Matrix Table Layout (4 Models x 2 Scenarios, CV vs. Test)
   - Slide 10: Dual Visual Explainability Layout (ROC/PR Curves + Top 6 Feature Importance)
   - Slide 11: 3-Tier Architecture Ribbon + Browser Frame Showcase Layout (Live FastAPI Web App)
   - Slide 12: Dark Executive Conclusion & Q&A Finale Layout (3 Pillars + 3 Roadmap Horizons)
2. Strictly ZERO emojis: All badges, indicators, and icons are native PowerPoint vector shapes.
3. Cohesive Editorial Color System:
   - Deep Navy (#0B162C) for Hero Cover, Split Panels & Finale
   - Crisp Alabaster (#F8FAFC) & Pure White (#FFFFFF) for high-legibility content slides
   - Electric Blue (#0066FF) & Cyan (#00B4D8) as primary tech accents
   - Emerald Teal (#0D9488) for protective/positive metrics, Amber (#D97706) & Rose Coral (#E11D48) for risk alerts
"""

import os
import glob
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image as PilImage, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
FIG_VI = os.path.join(ROOT_DIR, "figures_vi")
FIG_EN = os.path.join(ROOT_DIR, "figures_en")
VI_DIR = os.path.join(ROOT_DIR, "SIC_Documents_VI")
EN_DIR = os.path.join(ROOT_DIR, "SIC_Documents_EN")
QA_IMG_DIR = os.path.join(ROOT_DIR, "reports", "qa_slide_renders")

os.makedirs(FIG_VI, exist_ok=True)
os.makedirs(FIG_EN, exist_ok=True)
os.makedirs(VI_DIR, exist_ok=True)
os.makedirs(EN_DIR, exist_ok=True)
os.makedirs(QA_IMG_DIR, exist_ok=True)

# ==============================================================================
# DESIGN SYSTEM PALETTE & TYPOGRAPHY
# ==============================================================================
C_NAVY_DARK = RGBColor(11, 22, 44)      # #0B162C Primary Deep Navy
C_NAVY_CARD = RGBColor(19, 34, 64)      # #132240 Secondary Dark Surface
C_SLATE_900 = RGBColor(15, 23, 42)      # #0F172A Ink Black / Dark Text
C_SLATE_700 = RGBColor(51, 65, 85)      # #334155 Body Text on Light
C_SLATE_500 = RGBColor(100, 116, 139)   # #64748B Muted Caption Text
C_SLATE_300 = RGBColor(203, 213, 225)   # #CBD5E1 Subtle Card Border
C_SLATE_100 = RGBColor(241, 245, 249)   # #F1F5F9 Soft Gray Surface
C_BG_LIGHT  = RGBColor(248, 250, 252)   # #F8FAFC Alabaster Slide Canvas
C_WHITE     = RGBColor(255, 255, 255)   # #FFFFFF Pure White

C_BLUE_ELEC = RGBColor(0, 102, 255)     # #0066FF Primary Accent: Electric Blue
C_BLUE_SOFT = RGBColor(224, 242, 254)   # #E0F2FE Soft Blue Tint
C_CYAN      = RGBColor(0, 180, 216)     # #00B4D8 Secondary Accent: Cyan
C_TEAL      = RGBColor(13, 148, 136)    # #0D9488 Positive / Protective Teal
C_TEAL_SOFT = RGBColor(204, 251, 241)   # #CCFBF1 Soft Teal Tint
C_AMBER     = RGBColor(217, 119, 6)     # #D97706 Warning / Medium Tier
C_AMBER_BRIGHT = RGBColor(255, 183, 3)  # #FFB703 Gold Accent on Dark
C_AMBER_SOFT = RGBColor(254, 243, 199)  # #FEF3C7 Soft Amber Tint
C_CORAL     = RGBColor(225, 29, 72)     # #E11D48 High Risk Alert Rose/Coral
C_CORAL_SOFT = RGBColor(255, 228, 230)  # #FFE4E6 Soft Coral Tint

FONT_HEADING = "Segoe UI"
FONT_BODY = "Segoe UI"


# ==============================================================================
# BESPOKE CHART GENERATOR (MATCHING DESIGN SYSTEM PALETTE)
# ==============================================================================
def generate_design_system_charts(lang="vi"):
    is_vi = lang == "vi"
    out_dir = FIG_VI if is_vi else FIG_EN
    plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    # 1. Slide 07 Chart: Burnout Level vs Dropout Rate (Clean Editorial Light Theme)
    fig, ax = plt.subplots(figsize=(7.2, 4.5), dpi=220)
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")

    tiers = (
        ["Low Burnout\n(Kiệt sức Thấp\nn = 320)", "Medium Burnout\n(Kiệt sức Vừa\nn = 280)", "High Burnout\n(Kiệt sức Cao\nn = 200)"]
        if is_vi
        else ["Low Burnout\n(n = 320)", "Medium Burnout\n(n = 280)", "High Burnout\n(n = 200)"]
    )
    rates = [23.8, 63.2, 92.0]
    safe = [76.2, 36.8, 8.0]
    bar_colors = ["#0D9488", "#D97706", "#E11D48"]

    x = np.arange(len(tiers))
    bars = ax.bar(x, rates, width=0.48, color=bar_colors, edgecolor="none", zorder=3)

    for bar, r_val, s_val in zip(bars, rates, safe):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 2.5,
            f"{r_val:.1f}%",
            ha="center",
            va="bottom",
            fontsize=18,
            fontweight="bold",
            color="#0F172A",
        )
        lbl_in = f"Duy trì:\n{s_val:.1f}%" if is_vi else f"Retained:\n{s_val:.1f}%"
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() / 2,
            lbl_in,
            ha="center",
            va="center",
            fontsize=10.5,
            fontweight="bold",
            color="#FFFFFF",
        )

    ax.set_ylim(0, 108)
    ax.set_ylabel("Tỷ lệ Bỏ học Quan sát (%)" if is_vi else "Observed Dropout Rate (%)", fontsize=12, fontweight="bold", color="#334155")
    ax.set_title(
        "Tỷ lệ Bỏ học Tăng vọt theo Cấp độ Kiệt sức Học đường" if is_vi else "Observed Dropout Rate Across Academic Burnout Tiers",
        fontsize=13.5,
        fontweight="bold",
        color="#0F172A",
        pad=12,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(tiers, fontsize=11.5, fontweight="bold", color="#0F172A")
    ax.tick_params(axis="y", colors="#64748B", labelsize=10.5)
    ax.grid(axis="y", linestyle="--", alpha=0.4, color="#CBD5E1", zorder=0)
    for sp in ["top", "right"]:
        ax.spines[sp].set_visible(False)
    for sp in ["bottom", "left"]:
        ax.spines[sp].set_color("#CBD5E1")

    plt.tight_layout()
    c_burn = os.path.join(out_dir, "art_burnout_progression.png")
    plt.savefig(c_burn, dpi=220, facecolor="#FFFFFF")
    plt.close(fig)

    # 2. Slide 09 Chart: 4 Models x 2 Scenarios (5-Fold CV ROC-AUC)
    fig, ax = plt.subplots(figsize=(7.5, 3.2), dpi=220)
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")

    models = ["Logistic Regression", "XGBoost", "Soft Voting Ensemble", "Random Forest"]
    s1 = [87.61, 88.64, 88.89, 89.01]
    s2 = [87.25, 88.26, 88.25, 87.97]

    x = np.arange(len(models))
    w = 0.34
    b1 = ax.bar(
        x - w / 2,
        s1,
        width=w,
        label="Kịch bản 1: Có Burnout_Level (5-Fold CV)" if is_vi else "Scenario 1: With Burnout_Level (5-Fold CV)",
        color="#0066FF",
        zorder=3,
    )
    b2 = ax.bar(
        x + w / 2,
        s2,
        width=w,
        label="Kịch bản 2: Ẩn Burnout_Level - Cảnh báo sớm (5-Fold CV)" if is_vi else "Scenario 2: Without Burnout_Level - Early Warning (5-Fold CV)",
        color="#0D9488",
        zorder=3,
    )

    for bar in b1:
        v = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, v + 0.12, f"{v:.2f}%", ha="center", va="bottom", fontsize=10.5, fontweight="bold", color="#0066FF")
    for bar in b2:
        v = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, v + 0.12, f"{v:.2f}%", ha="center", va="bottom", fontsize=10.5, fontweight="bold", color="#0D9488")

    ax.set_ylim(85.5, 90.3)
    ax.set_ylabel("5-Fold CV ROC-AUC (%)", fontsize=11, fontweight="bold", color="#334155")
    ax.set_xticks(x)
    ax.set_xticklabels(models, fontsize=11.5, fontweight="bold", color="#0F172A")
    ax.tick_params(axis="y", colors="#64748B", labelsize=10)
    ax.grid(axis="y", linestyle="--", alpha=0.4, color="#CBD5E1", zorder=0)
    ax.legend(loc="upper left", frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1", fontsize=9.5)
    for sp in ["top", "right"]:
        ax.spines[sp].set_visible(False)
    for sp in ["bottom", "left"]:
        ax.spines[sp].set_color("#CBD5E1")

    plt.tight_layout()
    c_mod = os.path.join(out_dir, "art_model_benchmark.png")
    plt.savefig(c_mod, dpi=220, facecolor="#FFFFFF")
    plt.close(fig)

    return c_burn, c_mod


# ==============================================================================
# VECTOR SHAPE & TYPOGRAPHY PRIMITIVES
# ==============================================================================
def add_shape(slide, shape_type, left, top, width, height, fill_rgb, border_rgb=None, border_pt=1.0, adj=None):
    shp = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_rgb
    if border_rgb:
        shp.line.color.rgb = border_rgb
        shp.line.width = Pt(border_pt)
    else:
        shp.line.fill.background()
    if adj is not None and hasattr(shp, "adjustments") and len(shp.adjustments) > 0:
        shp.adjustments[0] = adj
    return shp


def add_rect(slide, left, top, width, height, fill_rgb, border_rgb=None, border_pt=1.0, rounded=False):
    # Enforce sharp square/rectangle boxes (no rounded corners) per user request
    st = MSO_SHAPE.RECTANGLE
    return add_shape(slide, st, left, top, width, height, fill_rgb, border_rgb, border_pt, adj=None)


def add_oval_badge(slide, left, top, size, bg_rgb, text_str, text_rgb=C_WHITE, font_size=12):
    # Sharp square badge instead of oval
    shp = add_shape(slide, MSO_SHAPE.RECTANGLE, left, top, size, size, bg_rgb)
    tf = shp.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text_str
    r.font.name = FONT_HEADING
    r.font.size = Pt(font_size)
    r.font.bold = True
    r.font.color.rgb = text_rgb
    return shp


def add_pill_badge(slide, left, top, width, height, bg_rgb, text_str, text_rgb=C_WHITE, font_size=9.5, border_rgb=None):
    # Sharp rectangular badge instead of rounded pill
    shp = add_shape(slide, MSO_SHAPE.RECTANGLE, left, top, width, height, bg_rgb, border_rgb=border_rgb, border_pt=1.0, adj=None)
    tf = shp.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text_str
    r.font.name = FONT_HEADING
    r.font.size = Pt(font_size)
    r.font.bold = True
    r.font.color.rgb = text_rgb
    return shp


def add_text(slide, left, top, width, height, text, font_size=13, bold=False, italic=False, color=C_SLATE_900, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT_BODY
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def add_bullet_list(slide, left, top, width, height, items, font_size=12.5, space_pt=6, text_color=C_SLATE_700, highlight_color=C_SLATE_900):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.02)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_pt)
        if isinstance(item, tuple):
            prefix, body = item
            r1 = p.add_run()
            r1.text = prefix + " "
            r1.font.name = FONT_BODY
            r1.font.size = Pt(font_size)
            r1.font.bold = True
            r1.font.color.rgb = highlight_color

            r2 = p.add_run()
            r2.text = body
            r2.font.name = FONT_BODY
            r2.font.size = Pt(font_size)
            r2.font.bold = False
            r2.font.color.rgb = text_color
        else:
            r = p.add_run()
            r.text = item
            r.font.name = FONT_BODY
            r.font.size = Pt(font_size)
            r.font.color.rgb = text_color
    return tb


def add_fitted_image(slide, img_path, left, top, width, height, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.08, rounded=True):
    if bg_rgb is not None:
        add_rect(slide, left, top, width, height, bg_rgb, border_rgb=border_rgb, border_pt=1.0, rounded=rounded)
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


def apply_editorial_chrome(slide, slide_num, total_slides, section_label, title_text, subtitle_text, is_vi=True, dark_mode=False):
    """Applies consistent, clean editorial header and footer while leaving the canvas open for custom layouts."""
    bg = C_NAVY_DARK if dark_mode else C_BG_LIGHT
    add_rect(slide, 0, 0, 13.333, 7.5, bg)

    # Top thin brand bar
    add_rect(slide, 0, 0, 13.333, 0.06, C_BLUE_ELEC if not dark_mode else C_CYAN)

    # Category pill + Title + Subtitle
    pill_bg = C_NAVY_CARD if dark_mode else C_BLUE_SOFT
    pill_fg = C_CYAN if dark_mode else C_BLUE_ELEC
    add_pill_badge(slide, 0.50, 0.22, 2.75, 0.26, pill_bg, section_label, text_rgb=pill_fg, font_size=9.0, border_rgb=pill_fg if dark_mode else None)

    title_col = C_WHITE if dark_mode else C_SLATE_900
    sub_col = C_SLATE_300 if dark_mode else C_SLATE_500
    add_text(slide, 0.48, 0.52, 12.35, 0.42, title_text, font_size=22, bold=True, color=title_col)
    add_text(slide, 0.48, 0.92, 12.35, 0.28, subtitle_text, font_size=11.5, italic=False, color=sub_col)

    # Subtle divider line under header
    div_col = C_NAVY_CARD if dark_mode else C_SLATE_300
    add_rect(slide, 0.50, 1.24, 12.33, 0.015, div_col)

    # Clean Minimalist Footer
    ftr_line_col = C_NAVY_CARD if dark_mode else C_SLATE_300
    add_rect(slide, 0.50, 7.10, 12.33, 0.012, ftr_line_col)
    ftr_txt = (
        "SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE PROJECT  |  NHÓM 03: ĐỖ QUANG MINH • VŨ QUANG HUY"
        if is_vi
        else "SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE PROJECT  |  GROUP 03: DO QUANG MINH • VU QUANG HUY"
    )
    ftr_col = C_SLATE_500
    add_text(slide, 0.48, 7.15, 9.5, 0.25, ftr_txt, font_size=8.5, bold=True, color=ftr_col)
    add_text(slide, 11.3, 7.15, 1.55, 0.25, f"{slide_num:02d} / {total_slides:02d}", font_size=9.5, bold=True, color=C_CYAN if dark_mode else C_BLUE_ELEC, align=PP_ALIGN.RIGHT)


# ==============================================================================
# MASTER PRESENTATION BUILDER (12 BESPOKE SLIDE LAYOUTS)
# ==============================================================================
def build_art_directed_deck(lang="vi"):
    is_vi = lang == "vi"
    fig_dir = FIG_VI if is_vi else FIG_EN
    c_burn, c_mod = generate_design_system_charts(lang)

    total_slides = 12
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ==========================================================================
    # SLIDE 01: HERO COVER LAYOUT (Deep Navy Asymmetric Split + Geometric Vector Art)
    # ==========================================================================
    s1 = prs.slides.add_slide(blank)
    add_rect(s1, 0, 0, 13.333, 7.5, C_NAVY_DARK)

    # Right-side decorative architectural dark panel
    add_rect(s1, 8.35, 0, 4.983, 7.5, C_NAVY_CARD)
    add_rect(s1, 8.35, 0, 0.04, 7.5, C_BLUE_ELEC)

    # Top electric accent bar
    add_rect(s1, 0, 0, 8.35, 0.08, C_CYAN)

    # SIC Program Pill Badge
    add_pill_badge(
        s1, 0.70, 0.72, 4.45, 0.36, C_NAVY_CARD,
        "SAMSUNG INNOVATION CAMPUS — AI CAPSTONE" if is_vi else "SAMSUNG INNOVATION CAMPUS — AI CAPSTONE",
        text_rgb=C_CYAN, font_size=10.5, border_rgb=C_CYAN
    )

    # Hero Main Title (Left 60% column)
    hero_title = (
        "XÂY DỰNG, HUẤN LUYỆN & SO SÁNH\nCÁC MÔ HÌNH HỌC MÁY DỰ ĐOÁN\nTỶ LỆ BỎ HỌC DỰA TRÊN CĂNG THẲNG\nVÀ THAM GIA HỌC TẬP"
        if is_vi
        else "BUILDING, TRAINING & COMPARING\nMACHINE LEARNING MODELS FOR\nSTUDENT DROPOUT PREDICTION\nFROM STRESS & ENGAGEMENT"
    )
    add_text(s1, 0.68, 1.32, 7.45, 2.55, hero_title, font_size=26.5, bold=True, color=C_WHITE)

    # Subtitle with vertical cyan bar
    add_rect(s1, 0.70, 4.02, 0.06, 0.55, C_TEAL)
    hero_sub = (
        "Hệ thống Cảnh báo Sớm Học đường Thông minh — Phân tích Kiệt sức (Burnout)\n& Dự báo Nguy cơ Bỏ học Đại học (Student Attrition Early Warning System)"
        if is_vi
        else "Intelligent Academic Early Warning System — Analyzing Student Burnout\n& Predicting University Attrition Risk Across Dual Scenarios"
    )
    add_text(s1, 0.88, 3.98, 7.15, 0.65, hero_sub, font_size=12.5, color=C_SLATE_300)

    # Author & Role Cards (Bottom Left)
    authors = (
        [
            ("ĐỖ QUANG MINH", "Trưởng nhóm (Team Leader)", "Kiến trúc Pipeline • Tạo 6 Đặc trưng • Mô hình Cây • FastAPI", C_CYAN),
            ("VŨ QUANG HUY", "Thành viên (Team Member)", "Phân tích Thống kê EDA • Điền khuyết KNN • Mô hình Cơ sở • Web UI", C_TEAL),
        ]
        if is_vi
        else [
            ("DO QUANG MINH", "Team Leader (Group 03)", "Leakage-Free Pipeline • 6 Domain Indices • Tree Ensembles • FastAPI", C_CYAN),
            ("VU QUANG HUY", "Team Member (Group 03)", "Statistical EDA • KNN Imputation • Baseline Models • Web Dashboard", C_TEAL),
        ]
    )
    for idx, (a_name, a_role, a_desc, col) in enumerate(authors):
        x = 0.70 + idx * 3.72
        add_rect(s1, x, 4.92, 3.55, 1.75, C_NAVY_CARD, border_rgb=col, border_pt=1.2, rounded=True)
        add_rect(s1, x, 4.92, 3.55, 0.06, col)
        add_text(s1, x + 0.16, 5.08, 3.22, 0.32, a_name, font_size=14, bold=True, color=C_WHITE)
        add_text(s1, x + 0.16, 5.40, 3.22, 0.26, a_role, font_size=10.5, bold=True, color=col)
        add_text(s1, x + 0.16, 5.74, 3.22, 0.78, a_desc, font_size=10.0, color=C_SLATE_300)

    # Right Panel: 4 Geometric Hero KPI Tiles (2x2 Matrix + Vector Connector Nodes)
    add_text(
        s1, 8.75, 0.72, 4.20, 0.32,
        "CHỈ SỐ THỰC NGHIỆM CỐT LÕI" if is_vi else "KEY EXPERIMENTAL METRICS",
        font_size=11, bold=True, color=C_CYAN
    )
    hero_kpis = (
        [
            ("2,000", "Hồ sơ Sinh viên Sạch", "Mở rộng từ 800 bản ghi gốc (0 NaN)", C_CYAN),
            ("25 → 34", "Thuộc tính Dữ liệu", "Tích hợp 6 chỉ số chuyên ngành mới", C_TEAL),
            ("89.01%", "Đỉnh CV ROC-AUC", "Random Forest (Kịch bản 1 Đầy đủ)", C_AMBER_BRIGHT),
            ("88.26%", "Cảnh báo Sớm AUC", "XGBoost (Kịch bản 2 Ẩn Burnout)", C_BLUE_ELEC),
        ]
        if is_vi
        else [
            ("2,000", "Clean Student Profiles", "Expanded from 800 raw rows (0 NaN)", C_CYAN),
            ("25 → 34", "Feature Dimensions", "Includes 6 engineered domain indices", C_TEAL),
            ("89.01%", "Peak CV ROC-AUC", "Random Forest (Scenario 1 Full)", C_AMBER_BRIGHT),
            ("88.26%", "Early Warning AUC", "XGBoost (Scenario 2 No Burnout)", C_BLUE_ELEC),
        ]
    )
    for idx, (k_val, k_title, k_sub, col) in enumerate(hero_kpis):
        y = 1.22 + idx * 1.42
        add_rect(s1, 8.75, y, 4.15, 1.24, C_NAVY_DARK, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s1, 8.75, y, 0.10, 1.24, col)
        add_text(s1, 9.00, y + 0.12, 1.75, 0.55, k_val, font_size=24, bold=True, color=col)
        add_text(s1, 10.75, y + 0.16, 2.05, 0.35, k_title, font_size=12, bold=True, color=C_WHITE)
        add_text(s1, 10.75, y + 0.52, 2.05, 0.58, k_sub, font_size=9.5, color=C_SLATE_300)

    add_text(s1, 11.3, 7.12, 1.55, 0.25, "01 / 12", font_size=9.5, bold=True, color=C_CYAN, align=PP_ALIGN.RIGHT)

    # ==========================================================================
    # SLIDE 02: TIMELINE / CONNECTED ROADMAP AGENDA LAYOUT (Alternating Track)
    # ==========================================================================
    s2 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s2, 2, total_slides,
        "CẤU TRÚC BÁO CÁO" if is_vi else "EXECUTIVE AGENDA",
        "Lộ Trình Thuyết Trình Đồ Án (6 Phần Trọng Tâm)" if is_vi else "Presentation Roadmap Across 6 Core Modules",
        "Mạch nội dung khép kín từ Đặt vấn đề, Kiến trúc Dữ liệu, Tạo Đặc trưng, Phân tích EDA, Huấn luyện Mô hình đến Ứng dụng Web"
        if is_vi else "A structured progression from Problem Formulation and Data Engineering to 4-Model Benchmarking and Web Deployment",
        is_vi=is_vi, dark_mode=False
    )

    # Horizontal Center Timeline Track Bar at y = 4.15
    add_rect(s2, 0.90, 4.14, 11.53, 0.06, C_SLATE_300)

    agenda_nodes = (
        [
            ("01", "Đặt Vấn Đề &\nMục Tiêu", "Cơ chế tích lũy dẫn tới bỏ học & phân công Nhóm 03 (50/50).", C_BLUE_ELEC),
            ("02", "Dữ Liệu &\nLàm Sạch 6 Bước", "800 bản ghi gốc (612 ô khuyết) → KNN (k=7) → 2.000 hồ sơ sạch.", C_CYAN),
            ("03", "Kỹ Thuật Tạo\n6 Đặc Trưng", "Thiết kế 6 chỉ số đo cán cân Căng thẳng vs. Gắn kết học tập.", C_TEAL),
            ("04", "Phân Tích Khám\nPhá (EDA)", "Tỷ lệ bỏ học tăng gấp 4 lần theo Burnout (23.8% → 92.0%).", C_AMBER),
            ("05", "Đối Sánh 4 Mô\nHình Học Máy", "Thực nghiệm 2 kịch bản: Random Forest (89.01%) & XGBoost (88.26%).", C_CORAL),
            ("06", "Web Dashboard\n& Kết Luận", "Ứng dụng FastAPI dự báo kép <30ms & lộ trình tích hợp LMS.", C_NAVY_DARK),
        ]
        if is_vi
        else [
            ("01", "Problem &\nObjectives", "Cumulative attrition mechanism & Group 03 50/50 role split.", C_BLUE_ELEC),
            ("02", "Dataset & 6-Step\nPreprocessing", "800 raw rows (612 missing) → KNN (k=7) → 2,000 clean profiles.", C_CYAN),
            ("03", "6 Engineered\nDomain Indices", "Quantifying the interplay between Stress and Engagement.", C_TEAL),
            ("04", "Exploratory Data\nAnalysis (EDA)", "Dropout surges nearly 4x across Burnout tiers (23.8% → 92.0%).", C_AMBER),
            ("05", "4-Model ML\nBenchmarking", "Dual-scenario CV: Random Forest (89.01%) & XGBoost (88.26%).", C_CORAL),
            ("06", "Web Dashboard\n& Conclusion", "Live FastAPI app (<30ms inference) & future LMS roadmap.", C_NAVY_DARK),
        ]
    )

    for idx, (num_str, ag_title, ag_desc, col) in enumerate(agenda_nodes):
        cx = 1.35 + idx * 2.12
        is_top = (idx % 2 == 0)
        # Node circle on track
        add_oval_badge(s2, cx - 0.26, 4.17 - 0.26, 0.52, col, num_str, text_rgb=C_WHITE, font_size=13)

        card_y = 1.55 if is_top else 4.65
        stem_y = 3.65 if is_top else 4.43
        add_rect(s2, cx - 0.015, stem_y, 0.03, 0.25, col)

        # Card
        card_w = 1.98
        card_x = cx - card_w / 2
        add_rect(s2, card_x, card_y, card_w, 2.10, C_WHITE, border_rgb=C_SLATE_300, border_pt=1.0, rounded=True)
        add_rect(s2, card_x, card_y, card_w, 0.08, col)
        add_text(s2, card_x + 0.10, card_y + 0.16, card_w - 0.20, 0.62, ag_title, font_size=13, bold=True, color=C_SLATE_900, align=PP_ALIGN.CENTER)
        add_text(s2, card_x + 0.10, card_y + 0.82, card_w - 0.20, 1.15, ag_desc, font_size=10.5, color=C_SLATE_700, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 03: SPLIT-SCREEN + CAUSAL PROCESS FLOW LAYOUT (Problem Statement)
    # ==========================================================================
    s3 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s3, 3, total_slides,
        "01 ĐẶT VẤN ĐỀ NGHIÊN CỨU" if is_vi else "01 PROBLEM FORMULATION",
        "Bối Cảnh Bài Toán & Cơ Chế Hình Thành Nguy Cơ Bỏ Học" if is_vi else "Problem Context & Causal Mechanism of Student Dropout",
        "Bỏ học không diễn ra đột ngột mà là hệ quả tích lũy từ suy giảm Gắn kết Học tập và Căng thẳng Tâm lý kéo dài"
        if is_vi else "Student attrition is not a sudden event, but the cumulative outcome of declining engagement and chronic stress",
        is_vi=is_vi, dark_mode=False
    )

    # Left Split Panel (Deep Navy Hero Callout Panel)
    add_rect(s3, 0.50, 1.42, 3.85, 5.45, C_NAVY_DARK, rounded=True)
    add_rect(s3, 0.50, 1.42, 3.85, 0.08, C_CORAL)
    add_pill_badge(s3, 0.75, 1.72, 2.30, 0.28, C_NAVY_CARD, "BIẾN MỤC TIÊU" if is_vi else "TARGET OUTCOME", text_rgb=C_CORAL, font_size=9.5, border_rgb=C_CORAL)
    add_text(s3, 0.75, 2.12, 3.35, 0.90, "54.62%", font_size=44, bold=True, color=C_WHITE)
    add_text(
        s3, 0.75, 3.02, 3.35, 0.55,
        "Sinh viên trong mẫu thuộc nhóm\nCó Nguy cơ Bỏ học (Dropout_Risk = Yes)" if is_vi else "Students in cohort classified as\nAt-Risk of Dropout (Dropout_Risk = Yes)",
        font_size=12.5, bold=True, color=C_CYAN
    )
    add_rect(s3, 0.75, 3.72, 3.35, 0.02, C_NAVY_CARD)
    add_bullet_list(
        s3, 0.75, 3.90, 3.35, 2.75,
        [
            ("• Yes (Nguy cơ):", "437 / 800 SV (54.62%)" if is_vi else "437 / 800 students (54.62%)"),
            ("• No (An toàn):", "363 / 800 SV (45.38%)" if is_vi else "363 / 800 students (45.38%)"),
            ("• Thách thức:", "Nhà trường thường chỉ phát hiện khi sinh viên đã nộp đơn xin thôi học." if is_vi else "Institutions typically react only after formal withdrawal."),
            ("• Giải pháp AI:", "Dự báo sớm trên cả 2 kịch bản (Có và Không có biến Burnout)." if is_vi else "Dual-scenario AI early warning before crisis point."),
        ],
        font_size=11.5, space_pt=8, text_color=C_SLATE_300, highlight_color=C_WHITE
    )

    # Right Side: 3-Stage Horizontal Process Flow + Detailed Cards
    stages = (
        [
            ("GIAI ĐOẠN 1: NGUYÊN NHÂN GỐC", "Suy Giảm Học Tập &\nÁp Lực Tâm Lý", C_BLUE_ELEC, C_BLUE_SOFT, [
                ("Chuyên cần thấp:", "Vắng học thường xuyên (<70%)."),
                ("Tự học sụt giảm:", "Ít giờ tự học, giảm động lực."),
                ("Áp lực đa chiều:", "Stress học vụ, lo âu, gánh nặng tài chính, thiếu ngủ (<6h)."),
            ]),
            ("GIAI ĐOẠN 2: TRUNG GIAN", "Kiệt Sức Học Đường\n(Burnout_Level)", C_AMBER, C_AMBER_SOFT, [
                ("Cầu nối trực tiếp:", "Khi áp lực vượt quá sức chịu đựng."),
                ("Phân tầng 3 mức:", "Low (40%), Medium (35%), High (25%)."),
                ("Tương quan mạnh:", "Hệ số Pearson r = +0.550 với nguy cơ bỏ học."),
            ]),
            ("GIAI ĐOẠN 3: CAN THIỆP SỚM", "Hệ Thống Cảnh Báo\nSớm Bằng Học Máy", C_TEAL, C_TEAL_SOFT, [
                ("Kịch bản 1 (Đầy đủ):", "Chẩn đoán khi có Burnout (AUC 89.01%)."),
                ("Kịch bản 2 (Sớm):", "Dự báo chỉ từ hành vi thô (AUC 88.26%)."),
                ("Hỗ trợ đúng lúc:", "Phụ đạo, tư vấn tâm lý, học bổng."),
            ]),
        ]
        if is_vi
        else [
            ("STAGE 1: ROOT DRIVERS", "Declining Engagement\n& Chronic Stress", C_BLUE_ELEC, C_BLUE_SOFT, [
                ("Low Attendance:", "Frequent absences (<70%)."),
                ("Study Deficit:", "Reduced study hours & motivation."),
                ("Multi-Faceted Stress:", "Academic stress, anxiety, financial burden, sleep loss (<6h)."),
            ]),
            ("STAGE 2: MEDIATOR", "Academic Burnout\n(Burnout_Level)", C_AMBER, C_AMBER_SOFT, [
                ("Proximate Mediator:", "Triggered when stress exceeds support."),
                ("3 Severity Tiers:", "Low (40%), Medium (35%), High (25%)."),
                ("Strong Association:", "Pearson r = +0.550 with student dropout."),
            ]),
            ("STAGE 3: AI INTERVENTION", "Proactive Early\nWarning System", C_TEAL, C_TEAL_SOFT, [
                ("Scenario 1 (Full):", "Diagnostic with Burnout (89.01% AUC)."),
                ("Scenario 2 (Early):", "From raw behavior alone (88.26% AUC)."),
                ("Targeted Action:", "Tutoring, counseling, financial aid."),
            ]),
        ]
    )

    for idx, (st_badge, st_title, col, soft_col, bullets) in enumerate(stages):
        x = 4.62 + idx * 2.80
        # Top Chevron Header Box
        add_shape(s3, MSO_SHAPE.CHEVRON, x, 1.42, 2.65, 0.55, col)
        add_text(s3, x + 0.12, 1.52, 2.30, 0.35, st_badge, font_size=9.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

        # Lower Content Card
        add_rect(s3, x, 2.12, 2.65, 4.75, C_WHITE, border_rgb=col, border_pt=1.2, rounded=True)
        add_rect(s3, x + 0.15, 2.28, 2.35, 0.78, soft_col, rounded=True)
        add_text(s3, x + 0.20, 2.34, 2.25, 0.65, st_title, font_size=13, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_bullet_list(s3, x + 0.15, 3.22, 2.35, 3.45, bullets, font_size=11.5, space_pt=10, text_color=C_SLATE_700, highlight_color=C_SLATE_900)

    # ==========================================================================
    # SLIDE 04: TWO-COLUMN + GANTT TIMELINE LAYOUT (Objectives, Roles & Schedule)
    # ==========================================================================
    s4 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s4, 4, total_slides,
        "01 MỤC TIÊU & KẾ HOẠCH" if is_vi else "01 OBJECTIVES & WBS PLAN",
        "Mục Tiêu Nghiên Cứu, Phân Công Nhóm 03 & Tiến Độ WBS" if is_vi else "Research Objectives, Group 03 Role Split & WBS Timeline",
        "Phân chia công việc cân bằng 50% – 50% và hoàn thành 100% các gói công việc theo lộ trình 4 giai đoạn"
        if is_vi else "Balanced 50/50 technical workload distribution with 100% completion across all 4 WBS phases",
        is_vi=is_vi, dark_mode=False
    )

    # Top Left Column: 4 Numbered Objectives
    add_rect(s4, 0.50, 1.42, 6.00, 3.35, C_WHITE, border_rgb=C_SLATE_300, border_pt=1.0, rounded=True)
    add_rect(s4, 0.50, 1.42, 6.00, 0.08, C_BLUE_ELEC)
    add_text(s4, 0.72, 1.60, 5.50, 0.34, "4 MỤC TIÊU KỸ THUẬT CỐT LÕI" if is_vi else "4 CORE TECHNICAL OBJECTIVES", font_size=13.5, bold=True, color=C_BLUE_ELEC)

    objs = (
        [
            ("01", "Chuẩn hóa & Mở rộng Dữ liệu:", "Làm sạch 6 bước trên 800 bản ghi gốc và mở rộng thống kê lên 2.000 hồ sơ sạch (0 NaN)."),
            ("02", "Kỹ thuật Tạo Đặc trưng:", "Điền khuyết KNNImputer (k=7) trong Pipeline và thiết kế 6 chỉ số tương tác mới."),
            ("03", "Đối sánh 4 Mô hình ML:", "Huấn luyện LR, Random Forest, XGBoost, Soft Voting qua 5-Fold CV trên 2 kịch bản."),
            ("04", "Triển khai Web Dashboard:", "Xây dựng ứng dụng FastAPI dự báo kép Burnout & Dropout thời gian thực (<30ms)."),
        ]
        if is_vi
        else [
            ("01", "Data Standardization & Scaling:", "Clean 800 raw rows via 6 steps and expand statistically to 2,000 clean profiles (0 NaN)."),
            ("02", "Domain Feature Engineering:", "Embed KNNImputer (k=7) in Pipeline and formulate 6 Stress-Engagement indices."),
            ("03", "4-Model ML Benchmarking:", "Evaluate LR, Random Forest, XGBoost & Soft Voting via 5-Fold CV across 2 scenarios."),
            ("04", "Live Web Application:", "Deploy a FastAPI dashboard delivering dual Burnout & Dropout diagnostics (<30ms)."),
        ]
    )
    for idx, (o_num, o_bold, o_txt) in enumerate(objs):
        y = 2.05 + idx * 0.64
        add_oval_badge(s4, 0.72, y + 0.04, 0.36, C_BLUE_ELEC, o_num, font_size=10)
        add_bullet_list(s4, 1.18, y, 5.15, 0.58, [(o_bold, o_txt)], font_size=11.0, space_pt=2)

    # Top Right Column: Two Stacked Role Cards (50% / 50%)
    roles = (
        [
            ("ĐỖ QUANG MINH — TRƯỞNG NHÓM (50% KHỐI LƯỢNG)", C_NAVY_DARK, C_CYAN, [
                "• Thiết kế Pipeline chống rò rỉ dữ liệu & lớp DomainFeatureEngineer (6 chỉ số).",
                "• Huấn luyện nhóm mô hình Cây tập hợp & Boosting (Random Forest, XGBoost, Soft Voting).",
                "• Xây dựng Backend FastAPI (app.py) và đóng gói mô hình thương phẩm (.pkl).",
            ]),
            ("VŨ QUANG HUY — THÀNH VIÊN (50% KHỐI LƯỢNG)", C_TEAL, C_WHITE, [
                "• Kiểm toán dữ liệu khuyết (18/25 cột), phân tích thống kê EDA & trực quan hóa biểu đồ.",
                "• Thiết lập KNNImputer (k=7), lọc ngoại lai IQR & huấn luyện nhóm mô hình cơ sở (LR).",
                "• Thiết kế giao diện Web Dashboard Frontend & biên soạn bộ hồ sơ báo cáo SIC.",
            ]),
        ]
        if is_vi
        else [
            ("DO QUANG MINH — TEAM LEADER (50% WORKLOAD)", C_NAVY_DARK, C_CYAN, [
                "• Architected leakage-free ML Pipeline & DomainFeatureEngineer (6 composite indices).",
                "• Trained Tree Ensembles & Boosting models (Random Forest, XGBoost, Soft Voting).",
                "• Built FastAPI Backend service (app.py) and serialized production models (.pkl).",
            ]),
            ("VU QUANG HUY — TEAM MEMBER (50% WORKLOAD)", C_TEAL, C_WHITE, [
                "• Audited missing data (18/25 cols), performed statistical EDA & visual charting.",
                "• Configured KNNImputer (k=7), IQR outlier clipping & baseline classifiers (LR).",
                "• Designed Frontend Web UI and authored the official SIC Capstone documentation.",
            ]),
        ]
    )
    for idx, (r_title, hdr_bg, hdr_fg, r_lines) in enumerate(roles):
        y = 1.42 + idx * 1.74
        add_rect(s4, 6.75, y, 6.08, 1.60, C_WHITE, border_rgb=hdr_bg, border_pt=1.2, rounded=True)
        add_rect(s4, 6.75, y, 6.08, 0.38, hdr_bg)
        add_text(s4, 6.92, y + 0.05, 5.75, 0.28, r_title, font_size=11.0, bold=True, color=hdr_fg)
        add_bullet_list(s4, 6.90, y + 0.44, 5.80, 1.10, r_lines, font_size=10.8, space_pt=3)

    # Bottom Horizontal WBS Gantt / Chevron Timeline Bar
    add_rect(s4, 0.50, 4.95, 12.33, 1.92, C_NAVY_DARK, rounded=True)
    add_text(
        s4, 0.75, 5.06, 11.8, 0.30,
        "TIẾN ĐỘ THỰC HIỆN 4 GIAI ĐOẠN (WORK BREAKDOWN STRUCTURE — HOÀN THÀNH 100%)"
        if is_vi else "4-PHASE WORK BREAKDOWN STRUCTURE TIMELINE (100% MILESTONE COMPLETION)",
        font_size=11.0, bold=True, color=C_CYAN
    )
    wbs_phases = (
        [
            ("GIAI ĐOẠN 1 (19–20/10)", "Khởi tạo & Đặc tả Dữ liệu", "Action Plan, WBS, Kiểm toán 800 mẫu", C_BLUE_ELEC),
            ("GIAI ĐOẠN 2 (21–25/10)", "Tiền xử lý & Phân tích EDA", "KNNImputer (k=7), IQR, 6 Chỉ số mới", C_CYAN),
            ("GIAI ĐOẠN 3 (26–29/10)", "Huấn luyện & Kiểm định ML", "5-Fold CV trên 2 Kịch bản (2.000 mẫu)", C_TEAL),
            ("GIAI ĐOẠN 4 (30/10–05/11)", "Web App & Báo cáo Cuối", "FastAPI Dashboard, Báo cáo & Slide", C_AMBER_BRIGHT),
        ]
        if is_vi
        else [
            ("PHASE 1 (Oct 19–20)", "Initiation & Schema Audit", "Action Plan, WBS, 800-row audit", C_BLUE_ELEC),
            ("PHASE 2 (Oct 21–25)", "Preprocessing & EDA", "KNNImputer (k=7), IQR, 6 Indices", C_CYAN),
            ("PHASE 3 (Oct 26–29)", "Model Training & 5-Fold CV", "Dual-Scenario Benchmark (n=2,000)", C_TEAL),
            ("PHASE 4 (Oct 30–Nov 05)", "Web App & Final Report", "FastAPI UI, Final Report & Slides", C_AMBER_BRIGHT),
        ]
    )
    for idx, (p_date, p_name, p_sub, col) in enumerate(wbs_phases):
        x = 0.75 + idx * 2.98
        add_rect(s4, x, 5.44, 2.82, 1.25, C_NAVY_CARD, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s4, x, 5.44, 2.82, 0.32, col)
        add_text(s4, x + 0.08, 5.47, 2.66, 0.26, p_date, font_size=10.0, bold=True, color=C_NAVY_DARK if col == C_AMBER_BRIGHT or col == C_CYAN else C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.10, 5.82, 2.62, 0.34, p_name, font_size=11.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s4, x + 0.10, 6.18, 2.62, 0.42, p_sub, font_size=9.8, color=C_SLATE_300, align=PP_ALIGN.CENTER)

    # ==========================================================================
    # SLIDE 05: BIG-NUMBER STRIP + DATA PIPELINE LAYOUT (Dataset & Cleaning)
    # ==========================================================================
    s5 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s5, 5, total_slides,
        "02 DỮ LIỆU & TIỀN XỬ LÝ" if is_vi else "02 DATASET & PREPROCESSING",
        "Kiến Trúc Bộ Dữ Liệu (800 → 2.000 Hồ Sơ) & Xử Lý Dữ Liệu Khuyết" if is_vi else "Dataset Architecture (800 → 2,000 Profiles) & Missing Value Imputation",
        "Xử lý 612 ô khuyết trên 18 thuộc tính bằng KNNImputer (k=7) trong Pipeline và mở rộng đa biến có phân tầng"
        if is_vi else "Imputing 612 missing cells across 18 columns via in-Pipeline KNNImputer (k=7) and stratified multivariate expansion",
        is_vi=is_vi, dark_mode=False
    )

    # Top Big-Number Strip (5 Connected Metric Blocks)
    top_metrics = (
        [
            ("800", "Bản ghi Gốc ban đầu", "Khảo sát 25 thuộc tính", C_BLUE_ELEC),
            ("612", "Ô dữ liệu bị khuyết", "Nằm trên 18/25 cột (1.6%–6.4%)", C_CORAL),
            ("k = 7", "KNNImputer Đa biến", "Nội suy từ 7 SV tương đồng nhất", C_TEAL),
            ("2.000", "Hồ sơ Sau Mở rộng", "800 gốc sạch + 1.200 mẫu tổng hợp", C_AMBER),
            ("34", "Thuộc tính Hoàn chỉnh", "Sạch 100% (0 giá trị NaN)", C_NAVY_DARK),
        ]
        if is_vi
        else [
            ("800", "Initial Raw Records", "25 original attributes", C_BLUE_ELEC),
            ("612", "Missing Data Cells", "Across 18/25 cols (1.6%–6.4%)", C_CORAL),
            ("k = 7", "Multivariate KNN", "Imputed from 7 nearest peers", C_TEAL),
            ("2,000", "Expanded Profiles", "800 clean + 1,200 augmented", C_AMBER),
            ("34", "Final Attributes", "100% clean (0 NaN values)", C_NAVY_DARK),
        ]
    )
    for idx, (m_big, m_t, m_s, col) in enumerate(top_metrics):
        x = 0.50 + idx * 2.50
        add_rect(s5, x, 1.40, 2.35, 1.32, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s5, x, 1.40, 2.35, 0.07, col)
        add_text(s5, x + 0.08, 1.50, 2.18, 0.50, m_big, font_size=24, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s5, x + 0.08, 2.02, 2.18, 0.30, m_t, font_size=10.5, bold=True, color=C_SLATE_900, align=PP_ALIGN.CENTER)
        add_text(s5, x + 0.08, 2.32, 2.18, 0.32, m_s, font_size=9.2, color=C_SLATE_500, align=PP_ALIGN.CENTER)

    # Bottom Left: 5 Attribute Domains + Transparent Provenance Box
    add_rect(s5, 0.50, 2.90, 5.25, 3.98, C_WHITE, border_rgb=C_SLATE_300, border_pt=1.0, rounded=True)
    add_rect(s5, 0.50, 2.90, 5.25, 0.40, C_NAVY_DARK)
    add_text(
        s5, 0.68, 2.96, 4.90, 0.28,
        "5 NHÓM DỮ LIỆU & NGUỒN GỐC MỞ RỘNG MẪU" if is_vi else "5 DATA DOMAINS & SAMPLE EXPANSION PROVENANCE",
        font_size=11.0, bold=True, color=C_CYAN
    )
    add_bullet_list(
        s5, 0.68, 3.40, 4.90, 3.35,
        [
            ("1. Nhân khẩu học (7 biến):", "Tuổi (17–25), Giới tính, Năm học, Khoa, Cư trú, Thu nhập."),
            ("2. Học tập (5 biến):", "Chuyên cần (%), Giờ học/ngày, GPA kỳ trước, Môn nợ, Động lực."),
            ("3. Lối sống (5 biến):", "Giờ ngủ, Giờ màn hình, Thể dục, Hoạt động xã hội, Làm thêm."),
            ("4. Căng thẳng & Hỗ trợ (6 biến):", "Stress, Lo âu, Tài chính, Bạn bè, Gia đình, Tư vấn."),
            ("5. Biến mục tiêu (2 biến):", "Burnout_Level (Low/Med/High) & Dropout_Risk (Yes/No)."),
            ("Minh bạch nguồn gốc 2.000 hồ sơ:", "800 bản ghi gốc sau khi làm sạch 6 bước được mở rộng thêm 1.200 bản ghi từ phân phối chuẩn đa biến phân tầng (Multivariate Gaussian) phục vụ huấn luyện mô hình."),
        ] if is_vi else [
            ("1. Demographics (7 cols):", "Age (17–25), Gender, Year, Department, Residence, Income."),
            ("2. Study Engagement (5 cols):", "Attendance (%), Study Hours, Prior GPA, Backlogs, Motivation."),
            ("3. Lifestyle (5 cols):", "Sleep Hours, Screen Time, Exercise, Social Activity, Part-Time Job."),
            ("4. Stress & Support (6 cols):", "Stress, Anxiety, Financial, Peer Pressure, Family, Counseling."),
            ("5. Targets (2 cols):", "Burnout_Level (Low/Med/High) & Dropout_Risk (Yes/No)."),
            ("Provenance of 2,000 profiles:", "800 cleaned raw records augmented with 1,200 records sampled from stratified Multivariate Gaussian distributions for ML training."),
        ],
        font_size=10.5, space_pt=5
    )

    # Bottom Right: Missing Values Analysis Chart
    add_fitted_image(s5, os.path.join(fig_dir, "01_missing_values_analysis.png"), 5.95, 2.90, 6.88, 3.98, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.06)

    # ==========================================================================
    # SLIDE 06: BENTO INFOGRAPHIC GRID LAYOUT (Pipeline & 6 Engineered Features)
    # ==========================================================================
    s6 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s6, 6, total_slides,
        "03 KỸ THUẬT TẠO ĐẶC TRƯNG" if is_vi else "03 DOMAIN FEATURE ENGINEERING",
        "Quy Trình Học Máy 5 Bước & Thiết Kế 6 Đặc Trưng Chuyên Ngành" if is_vi else "5-Stage ML Pipeline & 6 Engineered Domain Interaction Indices",
        "Biến đổi các thuộc tính rời rạc thành 6 chỉ số tổng hợp đo lường trực tiếp cán cân Căng thẳng Tâm lý vs. Gắn kết Học tập"
        if is_vi else "Transforming raw attributes into 6 composite indices quantifying the balance between Psychological Stress and Engagement",
        is_vi=is_vi, dark_mode=False
    )

    # Top 5-Step Process Ribbon (Connected Vector Chevrons)
    pipe_steps = (
        [
            ("BƯỚC 1: LÀM SẠCH", "Xóa trùng ID & Cắt ngoại lai IQR", C_SLATE_700),
            ("BƯỚC 2: ĐIỀN KHUYẾT", "KNNImputer (k=7) & Mode trong Pipeline", C_BLUE_ELEC),
            ("BƯỚC 3: TẠO ĐẶC TRƯNG", "DomainFeatureEngineer (6 chỉ số mới)", C_TEAL),
            ("BƯỚC 4: CHUẨN HÓA DỮ LIỆU", "OneHotEncoder & StandardScaler", C_CYAN),
            ("BƯỚC 5: HUẤN LUYỆN & CV", "Stratified 5-Fold CV & Hold-out Test 20%", C_NAVY_DARK),
        ]
        if is_vi
        else [
            ("STEP 1: CLEANING", "Deduplicate & IQR Outlier Clipping", C_SLATE_700),
            ("STEP 2: IMPUTATION", "In-Pipeline KNNImputer (k=7) & Mode", C_BLUE_ELEC),
            ("STEP 3: ENGINEERING", "DomainFeatureEngineer (6 new indices)", C_TEAL),
            ("STEP 4: ENCODING & SCALING", "OneHotEncoder & StandardScaler", C_CYAN),
            ("STEP 5: TRAINING & CV", "Stratified 5-Fold CV & 20% Test Split", C_NAVY_DARK),
        ]
    )
    for idx, (ps_t, ps_s, col) in enumerate(pipe_steps):
        x = 0.50 + idx * 2.49
        add_rect(s6, x, 1.38, 2.36, 0.86, col, rounded=True)
        add_text(s6, x + 0.08, 1.44, 2.20, 0.28, ps_t, font_size=10.0, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text(s6, x + 0.08, 1.74, 2.20, 0.42, ps_s, font_size=9.5, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Bottom 2x3 Bento Grid of 6 Feature Cards
    fe_bento = (
        [
            ("F1", "Academic_Engagement_Index", "0.35×(Attendance/10) + 0.25×Study + 0.25×Motivation + 0.15×GPA", "Đo lường tổng hợp mức độ chuyên cần, nỗ lực tự học và kết quả học tập (Thang 0–10).", "Top 3 Đặc trưng Quan trọng nhất", C_TEAL),
            ("F2", "Psychological_Stress_Index", "0.35×Stress + 0.30×Anxiety + 0.20×Financial + 0.15×Peer_Pressure", "Đo lường tổng hợp áp lực tâm lý, lo âu, gánh nặng tài chính và áp lực đồng trang lứa.", "Top 4 Đặc trưng Quan trọng nhất", C_CORAL),
            ("F3", "Stress_Engagement_Ratio", "Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)", "Tỷ số mất cân bằng khi áp lực tâm lý vượt quá độ gắn kết học tập (cộng 1.0 tránh chia 0).", "Top 2 Đặc trưng Quan trọng nhất", C_BLUE_ELEC),
            ("F4", "Net_Stress_Support_Balance", "Stress_Index − (0.50×Family + 0.30×Social + 0.20×Exercise)", "Đo lường mức căng thẳng ròng còn lại sau khi đã trừ đi điểm tựa gia đình, xã hội & thể chất.", "Top 6 Đặc trưng Quan trọng nhất", C_CYAN),
            ("F5", "Digital_Fatigue_Ratio", "Screen_Time_Hours / (Sleep_Hours + 1.0)", "Đánh giá hội chứng mệt mỏi số hóa do thức khuya dùng thiết bị điện tử quá mức.", "Chỉ báo rủi ro lối sống", C_AMBER),
            ("F6", "Academic_Risk_Burden", "1.50 × Backlogs + (100.0 − Attendance_Percent) / 20.0", "Lượng hóa gánh nặng học vụ tích lũy từ số môn nợ và tỷ lệ vắng học trên lớp.", "Chỉ báo cảnh báo học vụ", C_NAVY_DARK),
        ]
        if is_vi
        else [
            ("F1", "Academic_Engagement_Index", "0.35×(Attendance/10) + 0.25×Study + 0.25×Motivation + 0.15×GPA", "Composite index of class attendance, daily study effort, motivation, and prior GPA.", "Rank #3 Feature Importance", C_TEAL),
            ("F2", "Psychological_Stress_Index", "0.35×Stress + 0.30×Anxiety + 0.20×Financial + 0.15×Peer_Pressure", "Composite index capturing academic stress, anxiety, financial burden, and peer pressure.", "Rank #4 Feature Importance", C_CORAL),
            ("F3", "Stress_Engagement_Ratio", "Psychological_Stress_Index / (Academic_Engagement_Index + 1.0)", "Non-linear ratio measuring when psychological stress overwhelms academic engagement.", "Rank #2 Feature Importance", C_BLUE_ELEC),
            ("F4", "Net_Stress_Support_Balance", "Stress_Index − (0.50×Family + 0.30×Social + 0.20×Exercise)", "Net remaining stress after subtracting family, social, and physical activity buffers.", "Rank #6 Feature Importance", C_CYAN),
            ("F5", "Digital_Fatigue_Ratio", "Screen_Time_Hours / (Sleep_Hours + 1.0)", "Quantifies digital exhaustion from excessive screen exposure relative to sleep duration.", "Lifestyle Risk Indicator", C_AMBER),
            ("F6", "Academic_Risk_Burden", "1.50 × Backlogs + (100.0 − Attendance_Percent) / 20.0", "Captures cumulative academic penalty from course backlogs and attendance deficit.", "Academic Probation Indicator", C_NAVY_DARK),
        ]
    )

    for idx, (f_code, f_name, f_math, f_desc, f_role, col) in enumerate(fe_bento):
        r_i, c_i = divmod(idx, 3)
        x = 0.50 + c_i * 4.18
        y = 2.40 + r_i * 2.28
        w, h = 3.96, 2.12

        add_rect(s6, x, y, w, h, C_WHITE, border_rgb=C_SLATE_300, border_pt=1.0, rounded=True)
        add_rect(s6, x, y, w, 0.06, col)
        add_oval_badge(s6, x + 0.14, y + 0.14, 0.36, col, f_code, font_size=10)
        add_text(s6, x + 0.56, y + 0.14, w - 0.68, 0.32, f_name, font_size=12, bold=True, color=C_SLATE_900)

        # Dark Formula Box
        add_rect(s6, x + 0.14, y + 0.56, w - 0.28, 0.48, C_NAVY_DARK, rounded=True)
        add_text(s6, x + 0.18, y + 0.62, w - 0.36, 0.36, f_math, font_size=9.5, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

        # Description & Role Pill
        add_text(s6, x + 0.14, y + 1.10, w - 0.28, 0.62, f_desc, font_size=9.8, color=C_SLATE_700)
        add_pill_badge(s6, x + 0.14, y + 1.74, 2.35, 0.26, C_SLATE_100, f_role, text_rgb=col, font_size=8.8, border_rgb=col)

    # ==========================================================================
    # SLIDE 07: DATA VISUALIZATION + BIG-NUMBER STAIRCASE LAYOUT (EDA 1)
    # ==========================================================================
    s7 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s7, 7, total_slides,
        "04 PHÂN TÍCH KHÁM PHÁ (EDA 1)" if is_vi else "04 EXPLORATORY DATA ANALYSIS (EDA 1)",
        "Phát Hiện Trọng Tâm: Tỷ Lệ Bỏ Học Theo Mức Độ Kiệt Sức (Burnout)" if is_vi else "Core EDA Finding: Observed Dropout Rate by Academic Burnout Tier",
        "Mức độ kiệt sức học đường càng cao thì tỷ lệ bỏ học quan sát được trong bộ dữ liệu càng lớn (tăng từ 23.8% lên 92.0%)"
        if is_vi else "Higher observed academic burnout levels correspond to a nearly 4-fold surge in student dropout rate (23.8% to 92.0%)",
        is_vi=is_vi, dark_mode=False
    )

    # Left: Bespoke High-Res Chart
    add_fitted_image(s7, c_burn, 0.50, 1.42, 7.35, 5.45, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.08)

    # Right: 3 Big-Number Callout Cards + Scientific Caution Box
    stair_cards = (
        [
            ("MỨC THẤP — LOW BURNOUT", "23.8%", "76 / 320 SV bỏ học  |  244 SV duy trì an toàn", C_TEAL, C_TEAL_SOFT),
            ("MỨC VỪA — MEDIUM BURNOUT", "63.2%", "177 / 280 SV bỏ học  |  Ngưỡng chuyển giao cần can thiệp", C_AMBER, C_AMBER_SOFT),
            ("MỨC CAO — HIGH BURNOUT", "92.0%", "184 / 200 SV bỏ học  |  Nguy cơ bỏ học gần như tuyệt đối", C_CORAL, C_CORAL_SOFT),
        ]
        if is_vi
        else [
            ("LOW BURNOUT TIER (40%)", "23.8%", "76 / 320 dropped out  |  244 retained safely", C_TEAL, C_TEAL_SOFT),
            ("MEDIUM BURNOUT TIER (35%)", "63.2%", "177 / 280 dropped out  |  Critical transition zone", C_AMBER, C_AMBER_SOFT),
            ("HIGH BURNOUT TIER (25%)", "92.0%", "184 / 200 dropped out  |  Extreme attrition risk", C_CORAL, C_CORAL_SOFT),
        ]
    )
    for idx, (st_lbl, st_big, st_sub, col, soft_bg) in enumerate(stair_cards):
        y = 1.42 + idx * 1.35
        add_rect(s7, 8.05, y, 4.78, 1.20, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s7, 8.05, y, 0.12, 1.20, col)
        add_rect(s7, 8.28, y + 0.14, 1.55, 0.92, soft_bg, rounded=True)
        add_text(s7, 8.32, y + 0.22, 1.48, 0.75, st_big, font_size=24, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_text(s7, 9.95, y + 0.16, 2.75, 0.32, st_lbl, font_size=11.0, bold=True, color=col)
        add_text(s7, 9.95, y + 0.50, 2.75, 0.58, st_sub, font_size=10.2, color=C_SLATE_700)

    # Bottom Right: Methodological Note (Non-Causal)
    add_rect(s7, 8.05, 5.52, 4.78, 1.35, C_NAVY_DARK, rounded=True)
    add_rect(s7, 8.05, 5.52, 4.78, 0.06, C_CYAN)
    add_text(
        s7, 8.22, 5.64, 4.45, 0.28,
        "LƯU Ý HỌC THUẬT (SCIENTIFIC RIGOR)" if is_vi else "METHODOLOGICAL NOTE (SCIENTIFIC RIGOR)",
        font_size=10.0, bold=True, color=C_CYAN
    )
    add_text(
        s7, 8.22, 5.94, 4.45, 0.85,
        "Hệ số tương quan Pearson r = +0.550 phản ánh mối liên hệ thống kê thuận chiều rất mạnh trong dữ liệu quan sát, không tự suy diễn là quan hệ nhân quả một chiều."
        if is_vi else "Pearson r = +0.550 reflects a strong positive statistical association in observational data, not a standalone one-way causal claim.",
        font_size=9.8, color=C_SLATE_300
    )

    # ==========================================================================
    # SLIDE 08: COMPARISON / TUG-OF-WAR LAYOUT (EDA 2: Protective vs Risk Factors)
    # ==========================================================================
    s8 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s8, 8, total_slides,
        "04 PHÂN TÍCH TƯƠNG QUAN (EDA 2)" if is_vi else "04 CORRELATION ANALYSIS (EDA 2)",
        "Ma Trận Tương Quan Pearson: Yếu Tố Bảo Vệ vs. Yếu Tố Nguy Cơ" if is_vi else "Pearson Correlation Matrix: Protective Factors vs. Risk Drivers",
        "Định lượng rõ nhóm biến bảo vệ giúp giữ chân sinh viên (tương quan âm) và nhóm biến đẩy nhanh nguy cơ bỏ học (tương quan dương)"
        if is_vi else "Quantifying protective buffers (negative correlation with dropout) versus attrition accelerators (positive correlation)",
        is_vi=is_vi, dark_mode=False
    )

    # Left: Correlation Matrix Heatmap
    add_fitted_image(s8, os.path.join(fig_dir, "04_correlation_matrix_heatmap.png"), 0.50, 1.42, 6.65, 5.45, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.06)

    # Right Side: Tug-of-War Comparison (Top Protective Column vs Bottom Risk Column)
    # Protective Card
    add_rect(s8, 7.35, 1.42, 5.48, 2.62, C_WHITE, border_rgb=C_TEAL, border_pt=1.5, rounded=True)
    add_rect(s8, 7.35, 1.42, 5.48, 0.40, C_TEAL)
    add_text(
        s8, 7.52, 1.48, 5.15, 0.28,
        "NHÓM YẾU TỐ BẢO VỆ (GIẢM NGUY CƠ BỎ HỌC — r < 0)" if is_vi else "PROTECTIVE FACTORS (REDUCE DROPOUT RISK — r < 0)",
        font_size=11.0, bold=True, color=C_WHITE
    )
    prot_items = (
        [
            ("1. Attendance_Percent (r = -0.244):", "Nhóm an toàn đi học 80.68% vs. Nhóm bỏ học 73.70% (-7%)."),
            ("2. Motivation_Score (r = -0.193):", "Động lực nội tại giữ vai trò điểm tựa duy trì việc học."),
            ("3. Study_Hours_Per_Day (r = -0.160):", "Nhóm an toàn tự học 3.43h/ngày vs. Nhóm bỏ học 2.93h/ngày."),
            ("4. Family_Support & GPA (r = -0.137):", "Sự đồng hành của gia đình giảm trực tiếp mức kiệt sức."),
        ]
        if is_vi
        else [
            ("1. Attendance_Percent (r = -0.244):", "Retained students average 80.68% vs. 73.70% for dropouts."),
            ("2. Motivation_Score (r = -0.193):", "Intrinsic motivation buffers against academic exhaustion."),
            ("3. Study_Hours_Per_Day (r = -0.160):", "Retained average 3.43h/day vs. 2.93h/day for dropouts."),
            ("4. Family_Support & GPA (r = -0.137):", "Strong family support mitigates burnout progression."),
        ]
    )
    add_bullet_list(s8, 7.52, 1.90, 5.15, 2.05, prot_items, font_size=10.5, space_pt=5, text_color=C_SLATE_700, highlight_color=C_TEAL)

    # Risk Card
    add_rect(s8, 7.35, 4.22, 5.48, 2.65, C_WHITE, border_rgb=C_CORAL, border_pt=1.5, rounded=True)
    add_rect(s8, 7.35, 4.22, 5.48, 0.40, C_CORAL)
    add_text(
        s8, 7.52, 4.28, 5.15, 0.28,
        "NHÓM YẾU TỐ NGUY CƠ (TĂNG NGUY CƠ BỎ HỌC — r > 0)" if is_vi else "RISK DRIVERS (INCREASE DROPOUT RISK — r > 0)",
        font_size=11.0, bold=True, color=C_WHITE
    )
    risk_items = (
        [
            ("1. Burnout_Numeric (r = +0.550):", "Biến trung gian có tương quan thuận lớn nhất với bỏ học."),
            ("2. Stress_Level (r = +0.164):", "Nhóm bỏ học có mức căng thẳng trung bình cao hơn (5.81 vs 5.17)."),
            ("3. Anxiety_Score (r = +0.145):", "Lo âu kéo dài kết hợp thiếu ngủ làm suy kiệt năng lượng học."),
            ("4. Financial_Stress & Thu nhập thấp:", "Sinh viên nhóm thu nhập thấp có tỷ lệ bỏ học tới 62.1%."),
        ]
        if is_vi
        else [
            ("1. Burnout_Numeric (r = +0.550):", "Strongest proximate correlate of student attrition."),
            ("2. Stress_Level (r = +0.164):", "Higher mean stress in dropout group (5.81 vs. 5.17)."),
            ("3. Anxiety_Score (r = +0.145):", "Chronic anxiety and sleep deficit accelerate burnout."),
            ("4. Financial_Stress & Low Income:", "Low-income bracket exhibits a 62.1% dropout rate."),
        ]
    )
    add_bullet_list(s8, 7.52, 4.70, 5.15, 2.05, risk_items, font_size=10.5, space_pt=5, text_color=C_SLATE_700, highlight_color=C_CORAL)

    # ==========================================================================
    # SLIDE 09: DATA VISUALIZATION + EXECUTIVE MATRIX TABLE LAYOUT (4 Models)
    # ==========================================================================
    s9 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s9, 9, total_slides,
        "05 KẾT QUẢ HUẤN LUYỆN MÔ HÌNH" if is_vi else "05 MODEL BENCHMARKING RESULTS",
        "Đối Sánh 4 Mô Hình Học Máy Trên 2 Kịch Bản (5-Fold CV & Test 20%)" if is_vi else "Benchmarking 4 Machine Learning Models Across 2 Scenarios (CV & Test)",
        "Tách bạch minh bạch kết quả Kiểm định chéo 5-Fold CV (n=1.600) và Tập kiểm thử độc lập 20% (n=400)"
        if is_vi else "Clearly separating 5-Fold Stratified CV metrics (n=1,600) from the 20% Hold-out Test Set (n=400)",
        is_vi=is_vi, dark_mode=False
    )

    # Top Left: Grouped Bar Chart
    add_fitted_image(s9, c_mod, 0.50, 1.38, 7.45, 3.05, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.05)

    # Top Right: Two Winner Highlight Cards (Scenario 1 vs Scenario 2)
    win_cards = (
        [
            ("KỊCH BẢN 1: CÓ BURNOUT_LEVEL", "Random Forest", "89.01%", "5-Fold CV ROC-AUC  |  Test AUC: 88.42%  |  Test Acc: 81.50%", C_BLUE_ELEC),
            ("KỊCH BẢN 2: ẨN BURNOUT (CẢNH BÁO SỚM)", "XGBoost", "88.26%", "5-Fold CV ROC-AUC  |  Test AUC: 87.39%  |  CV F1: 82.40%", C_TEAL),
        ]
        if is_vi
        else [
            ("SCENARIO 1: WITH BURNOUT_LEVEL", "Random Forest", "89.01%", "5-Fold CV ROC-AUC  |  Test AUC: 88.42%  |  Test Acc: 81.50%", C_BLUE_ELEC),
            ("SCENARIO 2: NO BURNOUT (EARLY WARNING)", "XGBoost", "88.26%", "5-Fold CV ROC-AUC  |  Test AUC: 87.39%  |  CV F1: 82.40%", C_TEAL),
        ]
    )
    for idx, (w_scen, w_mod, w_auc, w_sub, col) in enumerate(win_cards):
        y = 1.38 + idx * 1.56
        add_rect(s9, 8.12, y, 4.71, 1.48, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s9, 8.12, y, 4.71, 0.34, col)
        add_text(s9, 8.26, y + 0.04, 4.42, 0.26, w_scen, font_size=10.0, bold=True, color=C_WHITE)
        add_text(s9, 8.26, y + 0.42, 2.65, 0.42, w_mod, font_size=16, bold=True, color=C_SLATE_900)
        add_text(s9, 10.85, y + 0.38, 1.85, 0.50, w_auc, font_size=20, bold=True, color=col, align=PP_ALIGN.RIGHT)
        add_text(s9, 8.26, y + 0.94, 4.42, 0.42, w_sub, font_size=9.8, bold=True, color=C_SLATE_700)

    # Bottom: Full Quantitative Matrix Table (CV vs Test)
    tbl_rows = [
        [
            "Mô hình Học máy (ML Classifier)" if is_vi else "Machine Learning Model",
            "KB1: CV Acc",
            "KB1: CV Recall",
            "KB1: CV F1",
            "KB1: CV AUC",
            "KB1: Test Acc",
            "KB1: Test F1",
            "KB1: Test AUC",
            "KB2: CV AUC (Sớm)" if is_vi else "S2: Early CV AUC",
        ],
        ["1. Random Forest (Top 1 KB1)", "80.25%", "86.38%", "82.98%", "89.01%", "81.50%", "83.63%", "88.42%", "87.97%"],
        ["2. Soft Voting Ensemble (LR+RF+XGB)", "79.70%", "84.74%", "82.27%", "88.89%", "80.25%", "82.17%", "88.32%", "88.25%"],
        ["3. XGBoost (Top 1 KB2 Cảnh báo sớm)", "79.95%", "84.56%", "82.39%", "88.64%", "78.50%", "80.28%", "87.99%", "88.26%"],
        ["4. Logistic Regression (Baseline)", "78.45%", "81.64%", "80.62%", "87.61%", "78.25%", "80.18%", "87.60%", "87.25%"],
    ]
    t_shape = s9.shapes.add_table(len(tbl_rows), len(tbl_rows[0]), Inches(0.50), Inches(4.56), Inches(12.33), Inches(2.35))
    tbl = t_shape.table
    col_w = [3.05, 1.10, 1.18, 1.10, 1.15, 1.15, 1.10, 1.15, 1.35]
    for c_i, w in enumerate(col_w):
        tbl.columns[c_i].width = Inches(w)

    for r_i, row in enumerate(tbl_rows):
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text = str(val)
            cell.fill.solid()
            if r_i == 0:
                cell.fill.fore_color.rgb = C_NAVY_DARK
            elif r_i == 1:
                cell.fill.fore_color.rgb = C_BLUE_SOFT
            elif r_i == 3:
                cell.fill.fore_color.rgb = C_TEAL_SOFT
            else:
                cell.fill.fore_color.rgb = C_WHITE if r_i % 2 == 1 else C_SLATE_100

            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.CENTER
                p.font.name = FONT_BODY
                p.font.size = Pt(10.0)
                p.font.bold = (r_i in [0, 1, 3] or c_i in [0, 4, 7, 8])
                if r_i == 0:
                    p.font.color.rgb = C_WHITE
                elif r_i == 1 and c_i in [4, 5, 6, 7]:
                    p.font.color.rgb = C_BLUE_ELEC
                elif r_i == 3 and c_i == 8:
                    p.font.color.rgb = C_TEAL
                else:
                    p.font.color.rgb = C_SLATE_900

    # ==========================================================================
    # SLIDE 10: DUAL VISUAL EXPLAINABILITY LAYOUT (ROC/PR & Feature Importance)
    # ==========================================================================
    s10 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s10, 10, total_slides,
        "05 GIẢI THÍCH MÔ HÌNH HỌC MÁY" if is_vi else "05 MODEL EXPLAINABILITY",
        "Đường Cong ROC / Precision-Recall & Độ Quan Trọng Đặc Trưng" if is_vi else "ROC / Precision-Recall Curves & Feature Importance Analysis",
        "Hai đặc trưng mới do nhóm thiết kế (Stress_Engagement_Ratio & Academic_Engagement_Index) xếp ngay sau Burnout_Ordinal"
        if is_vi else "Engineered indices (Stress_Engagement_Ratio & Academic_Engagement_Index) rank immediately behind Burnout_Ordinal",
        is_vi=is_vi, dark_mode=False
    )

    add_fitted_image(s10, os.path.join(fig_dir, "08_roc_and_precision_recall_curves.png"), 0.50, 1.40, 6.02, 4.12, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.06)
    add_fitted_image(s10, os.path.join(fig_dir, "10_feature_importance_analysis.png"), 6.80, 1.40, 6.03, 4.12, bg_rgb=C_WHITE, border_rgb=C_SLATE_300, pad=0.06)

    # Bottom Executive Synthesis Strip (Top 6 Features + Recall + Caution)
    add_rect(s10, 0.50, 5.66, 12.33, 1.25, C_NAVY_DARK, rounded=True)
    add_rect(s10, 0.50, 5.66, 0.10, 1.25, C_CYAN)
    add_text(
        s10, 0.75, 5.74, 11.85, 0.28,
        "TỔNG HỢP PHÂN TÍCH ĐỘ QUAN TRỌNG ĐẶC TRƯNG & ĐỘ NHẠY MÔ HÌNH" if is_vi else "KEY TAKEAWAYS FROM FEATURE IMPORTANCE & MODEL SENSITIVITY",
        font_size=10.5, bold=True, color=C_CYAN
    )
    add_bullet_list(
        s10, 0.75, 6.04, 11.85, 0.80,
        [
            ("• Top 6 Đặc trưng Quyết định:", "(1) Burnout_Ordinal → (2) Stress_Engagement_Ratio → (3) Academic_Engagement_Index → (4) Psychological_Stress_Index → (5) Attendance_Percent → (6) Net_Stress_Support_Balance."),
            ("• Độ nhạy cao (Recall > 84%) & Lưu ý học thuật:", "Phát hiện đúng hơn 4/5 số sinh viên có nguy cơ bỏ học. Lưu ý: Feature Importance phản ánh mức đóng góp phân tách nút, không đồng nhất với quan hệ nhân quả."),
        ] if is_vi else [
            ("• Top 6 Predictive Features:", "(1) Burnout_Ordinal → (2) Stress_Engagement_Ratio → (3) Academic_Engagement_Index → (4) Psychological_Stress_Index → (5) Attendance_Percent → (6) Net_Stress_Support_Balance."),
            ("• High Sensitivity (Recall > 84%) & Methodological Note:", "Identifies >4/5 of at-risk students. Note: Feature importance measures node-split contribution, not causal effect."),
        ],
        font_size=10.0, space_pt=3, text_color=C_SLATE_300, highlight_color=C_WHITE
    )

    # ==========================================================================
    # SLIDE 11: ARCHITECTURE RIBBON + BROWSER SHOWCASE LAYOUT (Web Application)
    # ==========================================================================
    s11 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s11, 11, total_slides,
        "06 SẢN PHẨM ỨNG DỤNG WEB" if is_vi else "06 WEB APPLICATION PRODUCT",
        "Kiến Trúc 3 Tầng & Ứng Dụng Web Dashboard Cảnh Báo Sớm (FastAPI)" if is_vi else "3-Tier Architecture & Live Early Warning Web Dashboard (FastAPI)",
        "Triển khai thực tế trên app.py: Nhập 18 chỉ số sinh viên → Phân tích & Dự báo kép (<30ms) → Khuyến nghị hỗ trợ cá nhân hóa"
        if is_vi else "Implemented in app.py: Input 18 student indicators → Dual ML inference (<30ms) → Personalized advising recommendations",
        is_vi=is_vi, dark_mode=False
    )

    # Top 3-Step Vector Flow Ribbon: INPUT -> AI PREDICTION -> RESULT & RECOMMENDATION
    flow_3 = (
        [
            ("01", "INPUT LAYER (GIAO DIỆN NHẬP LIỆU)", "Nhập 18 chỉ số học tập, lối sống, tâm lý hoặc chọn 3 mẫu Preset thử nhanh", C_BLUE_ELEC),
            ("02", "AI PREDICTION ENGINE (FASTAPI + .PKL)", "Tự động tính 6 chỉ số mới & chạy mô hình dự báo kép trong < 30ms", C_NAVY_DARK),
            ("03", "RESULT & RECOMMENDATION (KẾT QUẢ)", "Hiển thị xác suất Burnout (3 mức), Dropout (%) & Khuyến nghị hỗ trợ", C_TEAL),
        ]
        if is_vi
        else [
            ("01", "INPUT LAYER (STUDENT INDICATORS)", "Enter 18 academic, lifestyle & stress metrics or select 3 one-click Presets", C_BLUE_ELEC),
            ("02", "AI PREDICTION ENGINE (FASTAPI + .PKL)", "Computes 6 engineered indices & runs dual ML inference in < 30ms", C_NAVY_DARK),
            ("03", "RESULT & RECOMMENDATION (ADVISING)", "Displays 3-tier Burnout probabilities, Dropout Risk % & Action plan", C_TEAL),
        ]
    )
    for idx, (f_num, f_t, f_s, col) in enumerate(flow_3):
        x = 0.50 + idx * 4.18
        add_rect(s11, x, 1.38, 3.96, 0.88, C_WHITE, border_rgb=col, border_pt=1.5, rounded=True)
        add_oval_badge(s11, x + 0.14, 1.54, 0.42, col, f_num, font_size=11)
        add_text(s11, x + 0.64, 1.44, 3.22, 0.28, f_t, font_size=10.0, bold=True, color=col)
        add_text(s11, x + 0.64, 1.72, 3.22, 0.46, f_s, font_size=9.5, color=C_SLATE_700)

    # Bottom Left: Browser Window Frame + Real Live Screenshot of app.py
    add_rect(s11, 0.50, 2.42, 7.75, 4.48, C_WHITE, border_rgb=C_SLATE_300, border_pt=1.2, rounded=True)
    # Browser Top Bar
    add_rect(s11, 0.50, 2.42, 7.75, 0.36, C_NAVY_DARK)
    add_oval_badge(s11, 0.68, 2.52, 0.14, C_CORAL, "", font_size=6)
    add_oval_badge(s11, 0.88, 2.52, 0.14, C_AMBER_BRIGHT, "", font_size=6)
    add_oval_badge(s11, 1.08, 2.52, 0.14, C_TEAL, "", font_size=6)
    add_text(s11, 1.35, 2.46, 6.50, 0.26, "http://127.0.0.1:8000 — FastAPI Student Burnout & Dropout Early Warning Dashboard (app.py)", font_size=9.0, bold=True, color=C_SLATE_300)

    live_shot = os.path.join(FIG_VI, "13_live_web_dashboard_screenshot.png")
    add_fitted_image(s11, live_shot, 0.56, 2.82, 7.63, 4.02, bg_rgb=None, border_rgb=None, pad=0.02, rounded=False)

    # Bottom Right: 3 Preset Profiles + 5 Operational Capabilities
    add_rect(s11, 8.45, 2.42, 4.38, 4.48, C_WHITE, border_rgb=C_SLATE_300, border_pt=1.0, rounded=True)
    add_rect(s11, 8.45, 2.42, 4.38, 0.40, C_NAVY_DARK)
    add_text(
        s11, 8.62, 2.48, 4.05, 0.28,
        "CHỨC NĂNG THỰC TẾ TRÊN WEB DASHBOARD" if is_vi else "LIVE WEB DASHBOARD CAPABILITIES",
        font_size=11.0, bold=True, color=C_CYAN
    )

    presets = (
        [
            ("Mẫu 1 — SV Tích cực:", "Nguy cơ Bỏ học 8.0% (An toàn)", C_TEAL, C_TEAL_SOFT),
            ("Mẫu 2 — SV Căng thẳng TB:", "Nguy cơ Bỏ học ~47–58% (Cảnh báo)", C_AMBER, C_AMBER_SOFT),
            ("Mẫu 3 — SV Kiệt sức Cao:", "Nguy cơ Bỏ học 95.8% (Báo động đỏ)", C_CORAL, C_CORAL_SOFT),
        ]
        if is_vi
        else [
            ("Preset 1 — Engaged Student:", "Dropout Risk 8.0% (Safe Tier)", C_TEAL, C_TEAL_SOFT),
            ("Preset 2 — Moderate Stress:", "Dropout Risk ~47–58% (Warning)", C_AMBER, C_AMBER_SOFT),
            ("Preset 3 — High Burnout:", "Dropout Risk 95.8% (Critical Alert)", C_CORAL, C_CORAL_SOFT),
        ]
    )
    for idx, (pr_t, pr_v, col, s_bg) in enumerate(presets):
        y = 2.94 + idx * 0.62
        add_rect(s11, 8.62, y, 4.04, 0.52, s_bg, border_rgb=col, border_pt=1.0, rounded=True)
        add_text(s11, 8.74, y + 0.04, 3.80, 0.22, pr_t, font_size=9.8, bold=True, color=col)
        add_text(s11, 8.74, y + 0.24, 3.80, 0.24, pr_v, font_size=10.2, bold=True, color=C_SLATE_900)

    add_bullet_list(
        s11, 8.62, 4.86, 4.04, 1.95,
        [
            ("• Dự báo kép tức thì:", "Trả kết quả xác suất Burnout (Low/Med/High) và Dropout_Risk (%)."),
            ("• Chỉ số giải thích:", "Hiển thị trực quan Chỉ số Gắn kết học tập & Chỉ số Căng thẳng."),
            ("• Khuyến nghị tự động:", "Đề xuất nhắc nhở chuyên cần, phụ đạo, tham vấn tâm lý hoặc học bổng."),
        ] if is_vi else [
            ("• Real-Time Dual Prediction:", "Outputs 3-tier Burnout probabilities and calibrated Dropout_Risk (%)."),
            ("• Explainable Indices:", "Displays computed Academic Engagement & Psychological Stress indices."),
            ("• Automated Advising:", "Generates targeted attendance, tutoring, counseling & financial referrals."),
        ],
        font_size=10.2, space_pt=5
    )

    # ==========================================================================
    # SLIDE 12: CONCLUSION & Q&A FINALE LAYOUT (Deep Navy Executive Finale)
    # ==========================================================================
    s12 = prs.slides.add_slide(blank)
    apply_editorial_chrome(
        s12, 12, total_slides,
        "07 KẾT LUẬN & HƯỚNG PHÁT TRIỂN" if is_vi else "07 CONCLUSION & FUTURE WORK",
        "Tổng Kết Đóng Góp Đồ Án & Định Hướng Phát Triển Tương Lai" if is_vi else "Summary of Contributions & Future Research Roadmap",
        "Hoàn thiện quy trình khép kín từ chuẩn hóa dữ liệu, tạo đặc trưng, huấn luyện 4 mô hình đến ứng dụng Web cảnh báo sớm"
        if is_vi else "Complete end-to-end delivery from data preprocessing and feature engineering to 4-model benchmarking and Web deployment",
        is_vi=is_vi, dark_mode=True
    )

    # Top 3 Accomplishment Pillars (3 Columns on Dark Navy)
    pillars_done = (
        [
            ("01", "DỮ LIỆU & 6 ĐẶC TRƯNG MỚI", "Chuẩn hóa 800 bản ghi gốc, điền khuyết KNNImputer (k=7) trong Pipeline, mở rộng lên 2.000 hồ sơ sạch và thiết kế 6 chỉ số tương tác chuyên ngành.", C_CYAN),
            ("02", "ĐỐI SÁNH 4 MÔ HÌNH HỌC MÁY", "Đánh giá khách quan qua 5-Fold CV trên 2 kịch bản: Random Forest dẫn đầu Kịch bản 1 (89.01% AUC) và XGBoost dẫn đầu Kịch bản 2 Cảnh báo sớm (88.26% AUC).", C_TEAL),
            ("03", "ỨNG DỤNG WEB CẢNH BÁO SỚM", "Triển khai Web Dashboard FastAPI (app.py) dự báo đồng thời mức độ Burnout và nguy cơ Dropout kèm khuyến nghị hỗ trợ cụ thể cho từng sinh viên.", C_AMBER_BRIGHT),
        ]
        if is_vi
        else [
            ("01", "DATA & 6 ENGINEERED INDICES", "Cleaned 800 raw records, embedded KNNImputer (k=7) inside Pipeline, expanded to 2,000 clean profiles, and designed 6 domain interaction indices.", C_CYAN),
            ("02", "4-MODEL ML BENCHMARKING", "Rigorous 5-Fold CV across 2 scenarios: Random Forest leads Scenario 1 (89.01% AUC) and XGBoost leads Scenario 2 Early Warning (88.26% AUC).", C_TEAL),
            ("03", "EARLY WARNING WEB DASHBOARD", "Built a live FastAPI Web App (app.py) delivering simultaneous Burnout and Dropout predictions with personalized student support recommendations.", C_AMBER_BRIGHT),
        ]
    )
    for idx, (p_num, p_t, p_d, col) in enumerate(pillars_done):
        x = 0.50 + idx * 4.18
        add_rect(s12, x, 1.42, 3.96, 2.45, C_NAVY_CARD, border_rgb=col, border_pt=1.5, rounded=True)
        add_rect(s12, x, 1.42, 3.96, 0.07, col)
        add_oval_badge(s12, x + 0.18, 1.62, 0.42, col, p_num, text_rgb=C_NAVY_DARK if col != C_BLUE_ELEC else C_WHITE, font_size=12)
        add_text(s12, x + 0.68, 1.64, 3.12, 0.34, p_t, font_size=12.0, bold=True, color=col)
        add_text(s12, x + 0.18, 2.15, 3.60, 1.55, p_d, font_size=10.8, color=C_SLATE_300)

    # Middle Strip: 3 Future Development Horizons
    add_text(
        s12, 0.50, 4.02, 12.33, 0.30,
        "ĐỊNH HƯỚNG PHÁT TRIỂN TIẾP THEO (FUTURE ROADMAP)" if is_vi else "FUTURE RESEARCH & DEPLOYMENT ROADMAP",
        font_size=11.0, bold=True, color=C_CYAN
    )
    horizons = (
        [
            ("HƯỚNG 1: DỮ LIỆU ĐA TRƯỜNG", "Thu thập và kiểm định ngoại kiểm trên dữ liệu khảo sát thực tế từ nhiều trường đại học.", C_BLUE_ELEC),
            ("HƯỚNG 2: GIẢI THÍCH CỤC BỘ SHAP", "Tích hợp biểu đồ SHAP trực tiếp lên giao diện Web cho từng hồ sơ sinh viên.", C_TEAL),
            ("HƯỚNG 3: KẾT NỐI HỆ THỐNG LMS", "Đồng bộ dữ liệu chuyên cần & điểm số định kỳ theo tuần qua API nhà trường.", C_CYAN),
        ]
        if is_vi
        else [
            ("HORIZON 1: MULTI-CAMPUS VALIDATION", "Collect real-world institutional survey data across multiple universities for external validation.", C_BLUE_ELEC),
            ("HORIZON 2: LOCAL SHAP EXPLAINABILITY", "Embed per-student SHAP waterfall explanations directly inside the Web Dashboard UI.", C_TEAL),
            ("HORIZON 3: INSTITUTIONAL LMS API", "Automate weekly attendance and grade ingestion via university LMS APIs.", C_CYAN),
        ]
    )
    for idx, (h_t, h_d, col) in enumerate(horizons):
        x = 0.50 + idx * 4.18
        add_rect(s12, x, 4.36, 3.96, 1.05, C_NAVY_CARD, border_rgb=C_SLATE_700, border_pt=1.0, rounded=True)
        add_rect(s12, x, 4.36, 0.08, 1.05, col)
        add_text(s12, x + 0.18, 4.44, 3.65, 0.28, h_t, font_size=10.5, bold=True, color=col)
        add_text(s12, x + 0.18, 4.74, 3.65, 0.58, h_d, font_size=9.8, color=C_WHITE)

    # Bottom Finale Banner: Thank You & Q&A
    add_rect(s12, 0.50, 5.58, 12.33, 1.35, C_NAVY_CARD, border_rgb=C_CYAN, border_pt=1.8, rounded=True)
    add_text(
        s12, 0.70, 5.68, 11.93, 0.45,
        "THANK YOU FOR LISTENING  —  Q & A",
        font_size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER
    )
    add_text(
        s12, 0.70, 6.18, 11.93, 0.34,
        "Nhóm 03 (Samsung Innovation Campus)  |  Đỗ Quang Minh (Trưởng nhóm)  •  Vũ Quang Huy (Thành viên)"
        if is_vi else "Group 03 (Samsung Innovation Campus)  |  Do Quang Minh (Team Leader)  •  Vu Quang Huy (Team Member)",
        font_size=12.5, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER
    )
    add_text(
        s12, 0.70, 6.52, 11.93, 0.30,
        "Trân trọng cảm ơn Quý Thầy Cô và Hội đồng Chấm Đồ án!  |  Web App Demo: http://127.0.0.1:8000"
        if is_vi else "Thank you to the Evaluation Committee for your time and feedback!  |  Web App Demo: http://127.0.0.1:8000",
        font_size=10.5, color=C_SLATE_300, align=PP_ALIGN.CENTER
    )

    # Save deliverables
    if is_vi:
        final_root_pptx = os.path.join(ROOT_DIR, "Final_Presentation.pptx")
        vi_pptx = os.path.join(VI_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx")
        prs.save(final_root_pptx)
        prs.save(vi_pptx)
        print(f"Saved VI Master Redesign -> {final_root_pptx} & {vi_pptx}")
        return final_root_pptx
    else:
        en_pptx = os.path.join(EN_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pptx")
        prs.save(en_pptx)
        print(f"Saved EN Master Redesign -> {en_pptx}")
        return en_pptx


def export_pdf_and_slide_overview(pptx_path):
    """Exports Preview.pdf, renders all 12 slides to PNG, and builds Slide_Overview.png contact sheet."""
    import win32com.client

    pdf_path = os.path.join(ROOT_DIR, "Preview.pdf")
    overview_path = os.path.join(ROOT_DIR, "Slide_Overview.png")

    for old in glob.glob(os.path.join(QA_IMG_DIR, "slide_*.png")):
        try:
            os.remove(old)
        except OSError:
            pass

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    slide_pngs = []
    try:
        pres = ppt_app.Presentations.Open(pptx_path, WithWindow=False)
        pres.SaveAs(pdf_path, 32)
        for idx, slide in enumerate(pres.Slides, start=1):
            sp = os.path.join(QA_IMG_DIR, f"slide_{idx:02d}.png")
            slide.Export(sp, "PNG", 1920, 1080)
            slide_pngs.append(sp)
            print(f"Rendered QA image: slide_{idx:02d}.png")
        pres.Close()
    finally:
        try:
            ppt_app.Quit()
        except Exception:
            pass

    # Build 4x3 High-Resolution Slide_Overview.png Contact Sheet
    cols, rows = 4, 3
    thumb_w, thumb_h = 640, 360
    pad = 28
    header_h = 85
    canvas_w = cols * thumb_w + (cols + 1) * pad
    canvas_h = header_h + rows * (thumb_h + 38) + (rows + 1) * pad

    overview = PilImage.new("RGB", (canvas_w, canvas_h), (11, 22, 44))
    draw = ImageDraw.Draw(overview)
    try:
        f_title = ImageFont.truetype("C:/Windows/Fonts/segoeuib.ttf", 28)
        f_lbl = ImageFont.truetype("C:/Windows/Fonts/segoeuib.ttf", 16)
    except Exception:
        f_title = ImageFont.load_default()
        f_lbl = ImageFont.load_default()

    draw.text(
        (pad, 24),
        "SIC AI CAPSTONE PROJECT — COMPLETE 12-SLIDE VISUAL REDESIGN OVERVIEW (GROUP 03)",
        fill=(0, 212, 255),
        font=f_title,
    )

    for idx, sp in enumerate(slide_pngs):
        r_i, c_i = divmod(idx, cols)
        x = pad + c_i * (thumb_w + pad)
        y = header_h + pad + r_i * (thumb_h + 38 + pad)
        with PilImage.open(sp) as im:
            im_resized = im.resize((thumb_w, thumb_h), PilImage.Resampling.LANCZOS)
        # Draw border & paste
        draw.rectangle([x - 2, y - 2, x + thumb_w + 2, y + thumb_h + 2], outline=(0, 180, 216), width=2)
        overview.paste(im_resized, (x, y))
        draw.text((x + 4, y + thumb_h + 8), f"Slide {idx + 1:02d} / {len(slide_pngs):02d}", fill=(248, 250, 252), font=f_lbl)

    overview.save(overview_path, "PNG")
    print(f"Saved composite overview -> {overview_path}")


if __name__ == "__main__":
    vi_deck = build_art_directed_deck("vi")
    build_art_directed_deck("en")
    export_pdf_and_slide_overview(vi_deck)
