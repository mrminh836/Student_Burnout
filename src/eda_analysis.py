"""
Exploratory Data Analysis (EDA) & Visualization Script
Project: Student Burnout & Dropout Risk Prediction (Nhóm 03: Đỗ Quang Minh, Vũ Quang Huy)
"""

import os
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

warnings.filterwarnings("ignore")

# Configure visual style
plt.rcParams["figure.dpi"] = 150
plt.rcParams["savefig.dpi"] = 200
plt.rcParams["font.sans-serif"] = ["Segoe UI", "DejaVu Sans", "Arial"]
plt.rcParams["axes.unicode_minus"] = False
sns.set_theme(style="whitegrid", palette="muted")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXPANDED_PATH = os.path.join(BASE_DIR, "data", "student_burnout_dropout_expanded.csv")
ORIG_PATH = os.path.join(BASE_DIR, "data", "student_burnout_dropout_dataset_2.csv")
DATA_PATH = EXPANDED_PATH if os.path.exists(EXPANDED_PATH) else ORIG_PATH
FIG_DIR = os.path.join(BASE_DIR, "figures")
os.makedirs(FIG_DIR, exist_ok=True)


def load_data():
    df = pd.read_csv(DATA_PATH)
    df["Dropout_Binary"] = (df["Dropout_Risk"] == "Yes").astype(int)
    df["Burnout_Numeric"] = df["Burnout_Level"].map({"Low": 0, "Medium": 1, "High": 2})
    return df


def plot_missing_values(df):
    missing = df.isnull().sum()
    if missing.sum() == 0 and os.path.exists(ORIG_PATH):
        df = pd.read_csv(ORIG_PATH)
        missing = df.isnull().sum()
    missing = missing[missing > 0].sort_values(ascending=True)
    missing_pct = (missing / len(df)) * 100

    fig, ax = plt.subplots(figsize=(10, 6.5))
    colors = sns.color_palette("YlOrRd", len(missing))
    bars = ax.barh(missing.index, missing_pct.values, color=colors, edgecolor="#333333", height=0.72)

    for bar, count, pct in zip(bars, missing.values, missing_pct.values):
        ax.text(
            bar.get_width() + 0.15,
            bar.get_y() + bar.get_height() / 2,
            f"{count} ({pct:.1f}%)",
            va="center",
            ha="left",
            fontsize=9.5,
            fontweight="bold",
            color="#222222",
        )

    ax.set_xlim(0, max(missing_pct.values) * 1.22)
    ax.set_xlabel("Tỷ lệ khuyết dữ liệu / Missing Percentage (%)", fontsize=11, fontweight="bold")
    ax.set_title(
        "Phân tích dữ liệu khuyết theo từng thuộc tính (Missing Data Analysis - 18/25 cột)",
        fontsize=13,
        fontweight="bold",
        pad=14,
    )
    ax.grid(axis="x", linestyle="--", alpha=0.5)
    ax.grid(axis="y", visible=False)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "01_missing_values_analysis.png"))
    plt.close()


def plot_target_and_burnout(df):
    fig, axes = plt.subplots(1, 3, figsize=(16, 5.2))

    # 1. Dropout Risk Distribution
    counts_drop = df["Dropout_Risk"].value_counts()
    palette_drop = {"No": "#2a9d8f", "Yes": "#e76f51"}
    bars1 = axes[0].bar(
        counts_drop.index,
        counts_drop.values,
        color=[palette_drop[k] for k in counts_drop.index],
        edgecolor="#222222",
        width=0.55,
    )
    for b in bars1:
        h = b.get_height()
        axes[0].text(
            b.get_x() + b.get_width() / 2,
            h + 8,
            f"{h} ({h/len(df)*100:.1f}%)",
            ha="center",
            fontweight="bold",
            fontsize=10.5,
        )
    axes[0].set_ylim(0, max(counts_drop.values) * 1.18)
    axes[0].set_title("Phân phối Nguy cơ Bỏ học\n(Dropout_Risk Distribution)", fontsize=12, fontweight="bold")
    axes[0].set_ylabel("Số lượng sinh viên (Count)", fontweight="bold")

    # 2. Burnout Level Distribution
    order_burn = ["Low", "Medium", "High"]
    counts_burn = df["Burnout_Level"].value_counts().reindex(order_burn)
    palette_burn = ["#457b9d", "#f4a261", "#e63946"]
    bars2 = axes[1].bar(
        order_burn,
        counts_burn.values,
        color=palette_burn,
        edgecolor="#222222",
        width=0.58,
    )
    for b in bars2:
        h = b.get_height()
        axes[1].text(
            b.get_x() + b.get_width() / 2,
            h + 6,
            f"{h} ({h/len(df)*100:.1f}%)",
            ha="center",
            fontweight="bold",
            fontsize=10.5,
        )
    axes[1].set_ylim(0, max(counts_burn.values) * 1.18)
    axes[1].set_title("Phân phối Mức độ Kiệt sức\n(Burnout_Level Distribution)", fontsize=12, fontweight="bold")
    axes[1].set_ylabel("Số lượng sinh viên (Count)", fontweight="bold")

    # 3. Crosstab Burnout vs Dropout (%)
    ct = pd.crosstab(df["Burnout_Level"], df["Dropout_Risk"], normalize="index").reindex(order_burn) * 100
    bottom = np.zeros(len(order_burn))
    for label, color in [("No", "#2a9d8f"), ("Yes", "#e76f51")]:
        vals = ct[label].values
        b = axes[2].bar(
            order_burn,
            vals,
            bottom=bottom,
            label=f"Dropout = {label}",
            color=color,
            edgecolor="#222222",
            width=0.58,
        )
        for i, v in enumerate(vals):
            if v > 7:
                axes[2].text(
                    i,
                    bottom[i] + v / 2,
                    f"{v:.1f}%",
                    ha="center",
                    va="center",
                    color="white",
                    fontweight="bold",
                    fontsize=10.5,
                )
        bottom += vals

    axes[2].set_ylim(0, 108)
    axes[2].set_title("Tỷ lệ Bỏ học theo Mức độ Kiệt sức\n(Dropout Rate by Burnout_Level)", fontsize=12, fontweight="bold")
    axes[2].set_ylabel("Tỷ lệ phần trăm (%)", fontweight="bold")
    axes[2].legend(loc="upper right", frameon=True)

    plt.suptitle("Tổng quan Biến mục tiêu: Mối liên hệ chặt chẽ giữa Burnout và Dropout Risk", fontsize=14, fontweight="bold", y=1.02)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "02_target_and_burnout_distribution.png"), bbox_inches="tight")
    plt.close()


def plot_stress_vs_engagement(df):
    features = [
        ("Attendance_Percent", "Tỷ lệ chuyên cần (Attendance %)"),
        ("Study_Hours_Per_Day", "Giờ tự học/ngày (Study Hours/Day)"),
        ("Motivation_Score", "Điểm động lực (Motivation Score)"),
        ("Previous_GPA", "Điểm GPA kỳ trước (Previous GPA)"),
        ("Stress_Level", "Mức độ căng thẳng (Stress Level)"),
        ("Anxiety_Score", "Điểm lo âu (Anxiety Score)"),
        ("Financial_Stress_Score", "Áp lực tài chính (Financial Stress)"),
        ("Family_Support_Score", "Hỗ trợ gia đình (Family Support)"),
    ]

    fig, axes = plt.subplots(2, 4, figsize=(18, 9))
    axes = axes.flatten()
    palette = {"No": "#2a9d8f", "Yes": "#e76f51"}

    for idx, (col, title) in enumerate(features):
        ax = axes[idx]
        sns.violinplot(
            data=df,
            x="Dropout_Risk",
            y=col,
            order=["No", "Yes"],
            palette=palette,
            inner=None,
            alpha=0.35,
            ax=ax,
        )
        sns.boxplot(
            data=df,
            x="Dropout_Risk",
            y=col,
            order=["No", "Yes"],
            width=0.28,
            palette=palette,
            boxprops=dict(edgecolor="#222"),
            showfliers=False,
            ax=ax,
        )
        means = df.groupby("Dropout_Risk")[col].mean()
        for i, k in enumerate(["No", "Yes"]):
            m_val = means[k]
            ax.scatter(i, m_val, color="yellow", edgecolor="black", s=60, zorder=5, label="Mean" if i == 0 else "")
            ax.text(i, m_val, f"  μ={m_val:.2f}", va="center", ha="left", fontsize=9, fontweight="bold", color="#111")

        ax.set_title(title, fontsize=11, fontweight="bold")
        ax.set_xlabel("Nguy cơ bỏ học (Dropout_Risk)", fontsize=9.5)
        ax.set_ylabel("")

    plt.suptitle(
        "So sánh Phân phối các Chỉ số Tham gia Học tập (Hàng trên) & Căng thẳng Tâm lý (Hàng dưới) theo Dropout Risk",
        fontsize=14,
        fontweight="bold",
        y=1.02,
    )
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "03_stress_vs_engagement_distributions.png"), bbox_inches="tight")
    plt.close()


def plot_correlation_heatmap(df):
    cols = [
        "Dropout_Binary",
        "Burnout_Numeric",
        "Attendance_Percent",
        "Study_Hours_Per_Day",
        "Motivation_Score",
        "Previous_GPA",
        "Family_Support_Score",
        "Exercise_Freq_Per_Week",
        "Sleep_Hours",
        "Social_Activity_Score",
        "Stress_Level",
        "Anxiety_Score",
        "Financial_Stress_Score",
        "Screen_Time_Hours",
        "Backlogs",
        "Peer_Pressure_Score",
        "Age",
        "Year_of_Study",
    ]
    corr = df[cols].corr()

    fig, ax = plt.subplots(figsize=(14, 11))
    mask = np.triu(np.ones_like(corr, dtype=bool), k=1)
    sns.heatmap(
        corr,
        mask=mask,
        annot=True,
        fmt=".2f",
        cmap="RdBu_r",
        vmin=-0.45,
        vmax=0.65,
        center=0,
        square=True,
        linewidths=0.6,
        cbar_kws={"shrink": 0.8, "label": "Hệ số tương quan Pearson (r)"},
        annot_kws={"size": 8.5, "weight": "bold"},
        ax=ax,
    )
    ax.set_title(
        "Ma trận Tương quan giữa các Chỉ số Căng thẳng, Tham gia Học tập, Burnout và Dropout Risk",
        fontsize=13.5,
        fontweight="bold",
        pad=15,
    )
    plt.xticks(rotation=35, ha="right", fontsize=9.5)
    plt.yticks(fontsize=9.5)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "04_correlation_matrix_heatmap.png"))
    plt.close()


def plot_categorical_factors(df):
    cat_cols = [
        ("Department", "Khoa / Ngành học (Department)"),
        ("Family_Income_Bracket", "Mức thu nhập gia đình (Income Bracket)"),
        ("Residence_Type", "Loại hình cư trú (Residence Type)"),
        ("Part_Time_Job", "Việc làm thêm (Part-Time Job)"),
        ("Counseling_Access", "Tiếp cận Tư vấn Tâm lý (Counseling)"),
        ("Year_of_Study", "Năm học (Year of Study)"),
    ]

    fig, axes = plt.subplots(2, 3, figsize=(17, 9.5))
    axes = axes.flatten()

    for idx, (col, title) in enumerate(cat_cols):
        ax = axes[idx]
        temp = df.dropna(subset=[col]).copy()
        rate = temp.groupby(col)["Dropout_Binary"].agg(["mean", "count"]).reset_index()
        rate["mean_pct"] = rate["mean"] * 100
        if col == "Family_Income_Bracket":
            order = ["Low", "Lower-Middle", "Middle", "Upper-Middle", "High"]
            rate = rate.set_index(col).reindex(order).reset_index()
        else:
            rate = rate.sort_values("mean_pct", ascending=False)

        bars = ax.bar(
            rate[col].astype(str),
            rate["mean_pct"],
            color=sns.color_palette("Blues_r", len(rate)),
            edgecolor="#222222",
            width=0.6,
        )
        ax.axhline(df["Dropout_Binary"].mean() * 100, color="#e63946", linestyle="--", linewidth=1.6, label="TB toàn trường (54.6%)")
        for b, cnt in zip(bars, rate["count"]):
            h = b.get_height()
            ax.text(
                b.get_x() + b.get_width() / 2,
                h + 1.2,
                f"{h:.1f}%\n(n={cnt})",
                ha="center",
                va="bottom",
                fontsize=8.5,
                fontweight="bold",
            )
        ax.set_ylim(0, 82)
        ax.set_title(title, fontsize=11, fontweight="bold")
        ax.set_ylabel("Tỷ lệ bỏ học (%)", fontsize=9.5)
        ax.tick_params(axis="x", rotation=20)
        if idx == 0:
            ax.legend(loc="upper right", fontsize=8.5)

    plt.suptitle("Phân tích Tỷ lệ Bỏ học theo các Yếu tố Nhân khẩu học & Môi trường Học tập", fontsize=14, fontweight="bold", y=1.02)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "05_categorical_factors_analysis.png"), bbox_inches="tight")
    plt.close()


def plot_multivariate_analysis(df):
    fig, axes = plt.subplots(1, 2, figsize=(15, 6))

    # Left: Attendance vs Stress colored by Burnout_Level
    palette_burn = {"Low": "#2a9d8f", "Medium": "#f4a261", "High": "#e63946"}
    sns.scatterplot(
        data=df,
        x="Attendance_Percent",
        y="Stress_Level",
        hue="Burnout_Level",
        hue_order=["Low", "Medium", "High"],
        style="Dropout_Risk",
        palette=palette_burn,
        s=75,
        alpha=0.82,
        ax=axes[0],
    )
    axes[0].set_title(
        "Tương tác giữa Tỷ lệ Chuyên cần & Căng thẳng theo Burnout và Dropout",
        fontsize=11.5,
        fontweight="bold",
    )
    axes[0].set_xlabel("Tỷ lệ chuyên cần - Attendance_Percent (%)", fontweight="bold")
    axes[0].set_ylabel("Mức độ căng thẳng - Stress_Level (0-10)", fontweight="bold")
    axes[0].legend(bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=8.5)

    # Right: Mean Stress & Engagement profile across (Burnout_Level, Dropout_Risk)
    profile_cols = [
        "Attendance_Percent",
        "Study_Hours_Per_Day",
        "Motivation_Score",
        "Stress_Level",
        "Anxiety_Score",
        "Financial_Stress_Score",
    ]
    df_norm = df.copy()
    for c in profile_cols:
        df_norm[c] = (df_norm[c] - df_norm[c].min()) / (df_norm[c].max() - df_norm[c].min())

    grouped = df_norm.groupby("Burnout_Level")[profile_cols].mean().reindex(["Low", "Medium", "High"]).T
    grouped.plot(kind="bar", ax=axes[1], color=["#2a9d8f", "#f4a261", "#e63946"], edgecolor="#222", width=0.75)
    axes[1].set_title(
        "Hồ sơ Chuẩn hóa (Min-Max) của Sinh viên theo 3 Mức độ Burnout",
        fontsize=11.5,
        fontweight="bold",
    )
    axes[1].set_ylabel("Giá trị trung bình chuẩn hóa (0 - 1)", fontweight="bold")
    axes[1].tick_params(axis="x", rotation=25)
    axes[1].legend(title="Burnout_Level")

    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "06_multivariate_burnout_engagement_scatter.png"), bbox_inches="tight")
    plt.close()


if __name__ == "__main__":
    print("Loading dataset and generating EDA visualizations...")
    df = load_data()
    plot_missing_values(df)
    plot_target_and_burnout(df)
    plot_stress_vs_engagement(df)
    plot_correlation_heatmap(df)
    plot_categorical_factors(df)
    plot_multivariate_analysis(df)
    print(f"Successfully generated 6 EDA figures in: {FIG_DIR}")
