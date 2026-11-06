"""
Generates two complete sets of 12 publication-quality figures (200 DPI):
1. figures_vi/ -> 100% Vietnamese labels, titles, and annotations
2. figures_en/ -> 100% English labels, titles, and annotations
"""

import os
import sys
import warnings
import joblib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.metrics import (
    accuracy_score,
    f1_score,
    roc_auc_score,
    roc_curve,
    precision_recall_curve,
    average_precision_score,
    confusion_matrix,
)

warnings.filterwarnings("ignore")

plt.rcParams["figure.dpi"] = 150
plt.rcParams["savefig.dpi"] = 200
plt.rcParams["font.sans-serif"] = ["Segoe UI", "DejaVu Sans", "Arial"]
plt.rcParams["axes.unicode_minus"] = False
sns.set_theme(style="whitegrid")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)
from src.train_and_compare_models import DomainFeatureEngineer, build_preprocessor, get_models  # noqa: E402
from sklearn.pipeline import Pipeline  # noqa: E402

DATA_PATH = os.path.join(BASE_DIR, "data", "student_burnout_dropout_dataset_2.csv")
METRICS_PATH = os.path.join(BASE_DIR, "reports", "model_comparison_metrics.csv")
FIG_VI = os.path.join(BASE_DIR, "figures_vi")
FIG_EN = os.path.join(BASE_DIR, "figures_en")
os.makedirs(FIG_VI, exist_ok=True)
os.makedirs(FIG_EN, exist_ok=True)


def generate_eda_figures(df, lang="vi"):
    out_dir = FIG_VI if lang == "vi" else FIG_EN
    is_vi = lang == "vi"

    # 01. Missing values
    missing = df.drop(columns=["Dropout_Binary", "Burnout_Numeric"]).isnull().sum()
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
        )
    ax.set_xlim(0, max(missing_pct.values) * 1.22)
    ax.set_xlabel("Tỷ lệ dữ liệu khuyết (%)" if is_vi else "Missing Data Percentage (%)", fontsize=11, fontweight="bold")
    ax.set_title(
        "Phân tích Tỷ lệ Dữ liệu Khuyết trên 18/25 Thuộc tính" if is_vi else "Missing Data Analysis Across 18/25 Attributes",
        fontsize=13,
        fontweight="bold",
        pad=14,
    )
    ax.grid(axis="x", linestyle="--", alpha=0.5)
    ax.grid(axis="y", visible=False)
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "01_missing_values_analysis.png"))
    plt.close()

    # 02. Target and Burnout distribution
    fig, axes = plt.subplots(1, 3, figsize=(16, 5.2))
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
        axes[0].text(b.get_x() + b.get_width() / 2, h + 8, f"{h} ({h/len(df)*100:.1f}%)", ha="center", fontweight="bold", fontsize=10.5)
    axes[0].set_ylim(0, 520)
    axes[0].set_title("Phân phối Nguy cơ Bỏ học\n(Dropout_Risk)" if is_vi else "Dropout Risk Distribution\n(Target Variable)", fontsize=12, fontweight="bold")
    axes[0].set_ylabel("Số lượng sinh viên" if is_vi else "Number of Students", fontweight="bold")

    order_burn = ["Low", "Medium", "High"]
    counts_burn = df["Burnout_Level"].value_counts().reindex(order_burn)
    palette_burn = ["#457b9d", "#f4a261", "#e63946"]
    bars2 = axes[1].bar(order_burn, counts_burn.values, color=palette_burn, edgecolor="#222222", width=0.58)
    for b in bars2:
        h = b.get_height()
        axes[1].text(b.get_x() + b.get_width() / 2, h + 6, f"{h} ({h/len(df)*100:.1f}%)", ha="center", fontweight="bold", fontsize=10.5)
    axes[1].set_ylim(0, 380)
    axes[1].set_title("Phân phối Mức độ Kiệt sức\n(Burnout_Level)" if is_vi else "Burnout Level Distribution\n(Low / Medium / High)", fontsize=12, fontweight="bold")
    axes[1].set_ylabel("Số lượng sinh viên" if is_vi else "Number of Students", fontweight="bold")

    ct = pd.crosstab(df["Burnout_Level"], df["Dropout_Risk"], normalize="index").reindex(order_burn) * 100
    bottom = np.zeros(len(order_burn))
    for label, color in [("No", "#2a9d8f"), ("Yes", "#e76f51")]:
        vals = ct[label].values
        lbl_str = f"Bỏ học = {label}" if is_vi else f"Dropout = {label}"
        axes[2].bar(order_burn, vals, bottom=bottom, label=lbl_str, color=color, edgecolor="#222222", width=0.58)
        for i, v in enumerate(vals):
            if v > 7:
                axes[2].text(i, bottom[i] + v / 2, f"{v:.1f}%", ha="center", va="center", color="white", fontweight="bold", fontsize=10.5)
        bottom += vals
    axes[2].set_ylim(0, 108)
    axes[2].set_title("Tỷ lệ Bỏ học theo Mức độ Kiệt sức" if is_vi else "Dropout Rate by Burnout Level (%)", fontsize=12, fontweight="bold")
    axes[2].set_ylabel("Tỷ lệ phần trăm (%)" if is_vi else "Percentage (%)", fontweight="bold")
    axes[2].legend(loc="upper right", frameon=True)

    plt.suptitle(
        "Tổng quan Biến mục tiêu: Mối liên hệ chặt chẽ giữa Burnout và Nguy cơ Bỏ học"
        if is_vi
        else "Target Variable Overview: Strong Association Between Burnout Level and Dropout Risk",
        fontsize=14,
        fontweight="bold",
        y=1.02,
    )
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "02_target_and_burnout_distribution.png"), bbox_inches="tight")
    plt.close()

    # 03. Stress vs Engagement distributions
    features = [
        ("Attendance_Percent", "Tỷ lệ chuyên cần (%)" if is_vi else "Attendance Percentage (%)"),
        ("Study_Hours_Per_Day", "Giờ tự học/ngày" if is_vi else "Study Hours Per Day"),
        ("Motivation_Score", "Điểm động lực (0-10)" if is_vi else "Motivation Score (0-10)"),
        ("Previous_GPA", "Điểm GPA kỳ trước" if is_vi else "Previous Semester GPA"),
        ("Stress_Level", "Mức độ căng thẳng (0-10)" if is_vi else "Stress Level (0-10)"),
        ("Anxiety_Score", "Điểm lo âu (0-10)" if is_vi else "Anxiety Score (0-10)"),
        ("Financial_Stress_Score", "Áp lực tài chính (0-10)" if is_vi else "Financial Stress Score (0-10)"),
        ("Family_Support_Score", "Hỗ trợ gia đình (0-10)" if is_vi else "Family Support Score (0-10)"),
    ]
    fig, axes = plt.subplots(2, 4, figsize=(18, 9))
    axes = axes.flatten()
    palette = {"No": "#2a9d8f", "Yes": "#e76f51"}
    for idx, (col, title) in enumerate(features):
        ax = axes[idx]
        sns.violinplot(data=df, x="Dropout_Risk", y=col, order=["No", "Yes"], palette=palette, inner=None, alpha=0.35, ax=ax)
        sns.boxplot(data=df, x="Dropout_Risk", y=col, order=["No", "Yes"], width=0.28, palette=palette, boxprops=dict(edgecolor="#222"), showfliers=False, ax=ax)
        means = df.groupby("Dropout_Risk")[col].mean()
        for i, k in enumerate(["No", "Yes"]):
            m_val = means[k]
            ax.scatter(i, m_val, color="yellow", edgecolor="black", s=60, zorder=5)
            ax.text(i, m_val, f"  μ={m_val:.2f}", va="center", ha="left", fontsize=9, fontweight="bold")
        ax.set_title(title, fontsize=11, fontweight="bold")
        ax.set_xlabel("Nguy cơ bỏ học (Dropout_Risk)" if is_vi else "Dropout Risk (No vs. Yes)", fontsize=9.5)
        ax.set_ylabel("")
    plt.suptitle(
        "So sánh Phân phối các Chỉ số Tham gia Học tập (Hàng trên) & Căng thẳng Tâm lý (Hàng dưới)"
        if is_vi
        else "Distribution Comparison of Study Engagement (Top) & Psychological Stress Indicators (Bottom)",
        fontsize=14,
        fontweight="bold",
        y=1.02,
    )
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "03_stress_vs_engagement_distributions.png"), bbox_inches="tight")
    plt.close()

    # 04. Correlation matrix
    cols = [
        "Dropout_Binary", "Burnout_Numeric", "Attendance_Percent", "Study_Hours_Per_Day",
        "Motivation_Score", "Previous_GPA", "Family_Support_Score", "Exercise_Freq_Per_Week",
        "Sleep_Hours", "Social_Activity_Score", "Stress_Level", "Anxiety_Score",
        "Financial_Stress_Score", "Screen_Time_Hours", "Backlogs", "Peer_Pressure_Score",
        "Age", "Year_of_Study",
    ]
    corr = df[cols].corr()
    fig, ax = plt.subplots(figsize=(14, 11))
    mask = np.triu(np.ones_like(corr, dtype=bool), k=1)
    sns.heatmap(
        corr, mask=mask, annot=True, fmt=".2f", cmap="RdBu_r", vmin=-0.45, vmax=0.65, center=0,
        square=True, linewidths=0.6,
        cbar_kws={"shrink": 0.8, "label": "Hệ số tương quan Pearson (r)" if is_vi else "Pearson Correlation Coefficient (r)"},
        annot_kws={"size": 8.5, "weight": "bold"}, ax=ax,
    )
    ax.set_title(
        "Ma trận Tương quan giữa các Chỉ số Căng thẳng, Tham gia Học tập, Burnout và Dropout Risk"
        if is_vi
        else "Pearson Correlation Matrix Across Stress, Study Engagement, Burnout, and Dropout Risk",
        fontsize=13.5, fontweight="bold", pad=15,
    )
    plt.xticks(rotation=35, ha="right", fontsize=9.5)
    plt.yticks(fontsize=9.5)
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "04_correlation_matrix_heatmap.png"))
    plt.close()

    # 05. Categorical factors
    cat_cols = [
        ("Department", "Khoa / Ngành học" if is_vi else "Academic Department"),
        ("Family_Income_Bracket", "Mức thu nhập gia đình" if is_vi else "Family Income Bracket"),
        ("Residence_Type", "Loại hình cư trú" if is_vi else "Residence Type"),
        ("Part_Time_Job", "Việc làm thêm" if is_vi else "Part-Time Job Status"),
        ("Counseling_Access", "Tiếp cận Tư vấn Tâm lý" if is_vi else "Counseling Access"),
        ("Year_of_Study", "Năm học (1 - 4)" if is_vi else "Year of Study (1 - 4)"),
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
        bars = ax.bar(rate[col].astype(str), rate["mean_pct"], color=sns.color_palette("Blues_r", len(rate)), edgecolor="#222222", width=0.6)
        ax.axhline(
            df["Dropout_Binary"].mean() * 100, color="#e63946", linestyle="--", linewidth=1.6,
            label="TB toàn trường (54.6%)" if is_vi else "Overall Avg (54.6%)",
        )
        for b, cnt in zip(bars, rate["count"]):
            h = b.get_height()
            ax.text(b.get_x() + b.get_width() / 2, h + 1.2, f"{h:.1f}%\n(n={cnt})", ha="center", va="bottom", fontsize=8.5, fontweight="bold")
        ax.set_ylim(0, 82)
        ax.set_title(title, fontsize=11, fontweight="bold")
        ax.set_ylabel("Tỷ lệ bỏ học (%)" if is_vi else "Dropout Rate (%)", fontsize=9.5)
        ax.tick_params(axis="x", rotation=20)
        if idx == 0:
            ax.legend(loc="upper right", fontsize=8.5)
    plt.suptitle(
        "Phân tích Tỷ lệ Bỏ học theo các Yếu tố Nhân khẩu học & Môi trường Học tập"
        if is_vi
        else "Dropout Rate Analysis Across Demographic & Environmental Factors",
        fontsize=14, fontweight="bold", y=1.02,
    )
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "05_categorical_factors_analysis.png"), bbox_inches="tight")
    plt.close()

    # 06. Multivariate analysis
    fig, axes = plt.subplots(1, 2, figsize=(15, 6))
    palette_burn = {"Low": "#2a9d8f", "Medium": "#f4a261", "High": "#e63946"}
    sns.scatterplot(
        data=df, x="Attendance_Percent", y="Stress_Level", hue="Burnout_Level",
        hue_order=["Low", "Medium", "High"], style="Dropout_Risk", palette=palette_burn, s=75, alpha=0.82, ax=axes[0],
    )
    axes[0].set_title(
        "Tương tác giữa Tỷ lệ Chuyên cần & Căng thẳng theo Burnout" if is_vi else "Attendance vs. Stress Level by Burnout & Dropout Status",
        fontsize=11.5, fontweight="bold",
    )
    axes[0].set_xlabel("Tỷ lệ chuyên cần - Attendance (%)" if is_vi else "Attendance Percentage (%)", fontweight="bold")
    axes[0].set_ylabel("Mức độ căng thẳng - Stress (0-10)" if is_vi else "Stress Level (0-10)", fontweight="bold")
    axes[0].legend(bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=8.5)

    profile_cols = ["Attendance_Percent", "Study_Hours_Per_Day", "Motivation_Score", "Stress_Level", "Anxiety_Score", "Financial_Stress_Score"]
    df_norm = df.copy()
    for c in profile_cols:
        df_norm[c] = (df_norm[c] - df_norm[c].min()) / (df_norm[c].max() - df_norm[c].min())
    grouped = df_norm.groupby("Burnout_Level")[profile_cols].mean().reindex(["Low", "Medium", "High"]).T
    grouped.plot(kind="bar", ax=axes[1], color=["#2a9d8f", "#f4a261", "#e63946"], edgecolor="#222", width=0.75)
    axes[1].set_title(
        "Hồ sơ Chuẩn hóa (Min-Max) theo 3 Mức độ Burnout" if is_vi else "Normalized Feature Profile Across Burnout Levels",
        fontsize=11.5, fontweight="bold",
    )
    axes[1].set_ylabel("Giá trị trung bình chuẩn hóa (0 - 1)" if is_vi else "Normalized Mean Value (0 - 1)", fontweight="bold")
    axes[1].tick_params(axis="x", rotation=25)
    axes[1].legend(title="Burnout_Level")
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "06_multivariate_burnout_engagement_scatter.png"), bbox_inches="tight")
    plt.close()


def generate_model_figures(df, metrics_df, lang="vi"):
    out_dir = FIG_VI if lang == "vi" else FIG_EN
    is_vi = lang == "vi"

    res_full = metrics_df[metrics_df["Scenario"].str.contains("Scenario 1")].copy()
    res_early = metrics_df[metrics_df["Scenario"].str.contains("Scenario 2")].copy()

    # 07. Model comparison
    fig, axes = plt.subplots(1, 2, figsize=(16, 6.5))
    df_plot = res_full.sort_values("CV_F1_Mean", ascending=True)
    y_pos = np.arange(len(df_plot))
    height = 0.26
    axes[0].barh(y_pos - height, df_plot["CV_Accuracy_Mean"] * 100, height, label="Accuracy (%)", color="#457b9d", edgecolor="#222")
    axes[0].barh(y_pos, df_plot["CV_F1_Mean"] * 100, height, label="F1-Score (%)", color="#2a9d8f", edgecolor="#222")
    axes[0].barh(y_pos + height, df_plot["CV_ROC_AUC_Mean"] * 100, height, label="ROC-AUC (%)", color="#e76f51", edgecolor="#222")
    for i, (_, row) in enumerate(df_plot.iterrows()):
        axes[0].text(row["CV_F1_Mean"] * 100 + 0.5, i, f"{row['CV_F1_Mean']*100:.1f}%", va="center", fontsize=8.5, fontweight="bold")
    axes[0].set_yticks(y_pos)
    axes[0].set_yticklabels(df_plot["Model"], fontweight="bold")
    axes[0].set_xlim(50, 90)
    axes[0].set_xlabel("Điểm đánh giá 5-Fold CV (%)" if is_vi else "5-Fold Cross-Validation Score (%)", fontweight="bold")
    axes[0].set_title(
        "So sánh 10 Mô hình Học máy (Kịch bản 1: Đầy đủ đặc trưng)" if is_vi else "10 ML Models Comparison (Scenario 1: Full Features)",
        fontsize=12, fontweight="bold",
    )
    axes[0].legend(loc="lower right")

    merged = pd.merge(
        res_full[["Model", "CV_ROC_AUC_Mean", "CV_F1_Mean"]],
        res_early[["Model", "CV_ROC_AUC_Mean", "CV_F1_Mean"]],
        on="Model", suffixes=("_Full", "_Early"),
    ).sort_values("CV_ROC_AUC_Mean_Full", ascending=True)
    y_pos2 = np.arange(len(merged))
    h2 = 0.35
    b1 = axes[1].barh(
        y_pos2 + h2 / 2, merged["CV_ROC_AUC_Mean_Full"] * 100, h2,
        label="Kịch bản 1: Có Burnout_Level (Full)" if is_vi else "Scenario 1: With Burnout_Level (Full)",
        color="#264653", edgecolor="#222",
    )
    b2 = axes[1].barh(
        y_pos2 - h2 / 2, merged["CV_ROC_AUC_Mean_Early"] * 100, h2,
        label="Kịch bản 2: Chỉ dùng Stress & Engagement (Sớm)" if is_vi else "Scenario 2: Without Burnout_Level (Early Warning)",
        color="#e9c46a", edgecolor="#222",
    )
    for bar in b1:
        w = bar.get_width()
        axes[1].text(w + 0.4, bar.get_y() + bar.get_height() / 2, f"{w:.1f}%", va="center", fontsize=8.5, fontweight="bold")
    for bar in b2:
        w = bar.get_width()
        axes[1].text(w + 0.4, bar.get_y() + bar.get_height() / 2, f"{w:.1f}%", va="center", fontsize=8, color="#444")
    axes[1].set_yticks(y_pos2)
    axes[1].set_yticklabels(merged["Model"], fontweight="bold")
    axes[1].set_xlim(55, 90)
    axes[1].set_xlabel("Điểm ROC-AUC trung bình 5-Fold CV (%)" if is_vi else "Mean 5-Fold CV ROC-AUC (%)", fontweight="bold")
    axes[1].set_title(
        "Đánh giá Vai trò của Biến trung gian Burnout_Level" if is_vi else "Impact of Intermediate Variable Burnout_Level (ROC-AUC)",
        fontsize=12, fontweight="bold",
    )
    axes[1].legend(loc="lower right", fontsize=9)
    plt.suptitle(
        "Kết quả Thực nghiệm & So sánh Hiệu năng 10 Mô hình Dự đoán Bỏ học"
        if is_vi
        else "Experimental Results & Performance Comparison Across 10 Machine Learning Models",
        fontsize=14, fontweight="bold", y=1.02,
    )
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "07_model_performance_comparison.png"), bbox_inches="tight")
    plt.close()

    # Train once on 80/20 split to get curves & confusion matrices
    X_raw = df.drop(columns=["Student_ID", "Dropout_Risk", "Dropout_Binary", "Burnout_Numeric"])
    y = df["Dropout_Binary"]
    X_train, X_test, y_train, y_test = train_test_split(X_raw, y, test_size=0.20, stratify=y, random_state=42)
    fe_sample = DomainFeatureEngineer(include_burnout=True).fit_transform(X_train)
    models = get_models(random_state=42)

    roc_data, pr_data, trained_pipes = {}, {}, {}
    for name, clf in models.items():
        pipe = Pipeline([("fe", DomainFeatureEngineer(include_burnout=True)), ("pre", build_preprocessor(fe_sample)), ("clf", clf)])
        pipe.fit(X_train, y_train)
        y_pred = pipe.predict(X_test)
        y_prob = pipe.predict_proba(X_test)[:, 1]
        fpr, tpr, _ = roc_curve(y_test, y_prob)
        prec_c, rec_c, _ = precision_recall_curve(y_test, y_prob)
        roc_data[name] = (fpr, tpr, roc_auc_score(y_test, y_prob))
        pr_data[name] = (rec_c, prec_c, average_precision_score(y_test, y_prob))
        trained_pipes[name] = (pipe, y_test, y_pred, y_prob)

    # 08. ROC and PR curves
    fig, axes = plt.subplots(1, 2, figsize=(15, 6.2))
    colors = sns.color_palette("tab10", len(roc_data))
    for (name, (fpr, tpr, auc_val)), color in zip(roc_data.items(), colors):
        lw = 2.6 if "Ensemble" in name or "Random Forest" in name or "Logistic" in name else 1.5
        axes[0].plot(fpr, tpr, label=f"{name} (AUC = {auc_val:.3f})", color=color, linewidth=lw)
    axes[0].plot([0, 1], [0, 1], "k--", alpha=0.5, label="Random Guess (0.500)")
    axes[0].set_xlabel("False Positive Rate (1 - Specificity)", fontweight="bold")
    axes[0].set_ylabel("True Positive Rate (Recall)", fontweight="bold")
    axes[0].set_title("Đường cong ROC - Tập Kiểm thử (Test Set)" if is_vi else "ROC Curves on Hold-Out Test Set (20%)", fontsize=12, fontweight="bold")
    axes[0].legend(loc="lower right", fontsize=8.5)

    for (name, (rec, prec, ap_val)), color in zip(pr_data.items(), colors):
        lw = 2.6 if "Ensemble" in name or "Random Forest" in name or "Logistic" in name else 1.5
        axes[1].plot(rec, prec, label=f"{name} (AP = {ap_val:.3f})", color=color, linewidth=lw)
    axes[1].set_xlabel("Recall (Độ nhạy)" if is_vi else "Recall (Sensitivity)", fontweight="bold")
    axes[1].set_ylabel("Precision (Độ chính xác dương)" if is_vi else "Precision (Positive Predictive Value)", fontweight="bold")
    axes[1].set_title("Đường cong Precision-Recall - Tập Kiểm thử" if is_vi else "Precision-Recall Curves on Hold-Out Test Set", fontsize=12, fontweight="bold")
    axes[1].legend(loc="lower left", fontsize=8.5)
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "08_roc_and_precision_recall_curves.png"), bbox_inches="tight")
    plt.close()

    # 09. Confusion Matrices
    top_models = ["Soft Voting Ensemble", "Random Forest", "XGBoost", "Logistic Regression"]
    fig, axes = plt.subplots(1, 4, figsize=(18, 4.3))
    for ax, name in zip(axes, top_models):
        _, y_t, y_p, _ = trained_pipes[name]
        cm = confusion_matrix(y_t, y_p)
        cm_pct = cm / cm.sum() * 100
        annot = np.array([[f"{v}\n({p:.1f}%)" for v, p in zip(r_v, r_p)] for r_v, r_p in zip(cm, cm_pct)])
        sns.heatmap(
            cm, annot=annot, fmt="", cmap="Blues", cbar=False,
            xticklabels=["Dự báo: No", "Dự báo: Yes"] if is_vi else ["Pred: No", "Pred: Yes"],
            yticklabels=["Thực tế: No", "Thực tế: Yes"] if is_vi else ["Actual: No", "Actual: Yes"],
            annot_kws={"size": 11, "weight": "bold"}, ax=ax,
        )
        acc = accuracy_score(y_t, y_p)
        f1 = f1_score(y_t, y_p)
        ax.set_title(f"{name}\nAcc: {acc*100:.1f}% | F1: {f1*100:.1f}%", fontsize=11, fontweight="bold")
    plt.suptitle(
        "Ma trận Nhầm lẫn (Confusion Matrix) của 4 Mô hình Tiêu biểu trên Tập Kiểm thử (n=160)"
        if is_vi
        else "Confusion Matrices of Top 4 Representative Models on Hold-Out Test Set (n=160)",
        fontsize=13.5, fontweight="bold", y=1.05,
    )
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "09_confusion_matrices_top_models.png"), bbox_inches="tight")
    plt.close()

    # 10. Feature Importance
    rf_pipe = trained_pipes["Random Forest"][0]
    xgb_pipe = trained_pipes["XGBoost"][0]
    pre = rf_pipe.named_steps["pre"]
    num_features = pre.transformers_[0][2]
    cat_ohe = pre.transformers_[1][1].named_steps["ohe"]
    cat_features = list(cat_ohe.get_feature_names_out(pre.transformers_[1][2]))
    all_features = num_features + cat_features

    rf_imp = pd.Series(rf_pipe.named_steps["clf"].feature_importances_, index=all_features).sort_values(ascending=True).tail(15)
    xgb_imp = pd.Series(xgb_pipe.named_steps["clf"].feature_importances_, index=all_features).sort_values(ascending=True).tail(15)
    fig, axes = plt.subplots(1, 2, figsize=(16, 6.5))
    rf_imp.plot(kind="barh", color=sns.color_palette("viridis", 15), edgecolor="#222", ax=axes[0])
    axes[0].set_title("Top 15 Đặc trưng Quan trọng nhất - Random Forest" if is_vi else "Top 15 Most Important Features - Random Forest", fontsize=12, fontweight="bold")
    axes[0].set_xlabel("Mức độ quan trọng (Gini Importance)" if is_vi else "Feature Importance (Gini Impurity)", fontweight="bold")

    xgb_imp.plot(kind="barh", color=sns.color_palette("magma", 15), edgecolor="#222", ax=axes[1])
    axes[1].set_title("Top 15 Đặc trưng Quan trọng nhất - XGBoost" if is_vi else "Top 15 Most Important Features - XGBoost", fontsize=12, fontweight="bold")
    axes[1].set_xlabel("Mức độ quan trọng (Gain Importance)" if is_vi else "Feature Importance (Relative Gain)", fontweight="bold")
    plt.suptitle(
        "Phân tích Mức độ Quan trọng của các Đặc trưng (Feature Importance Analysis)"
        if is_vi
        else "Feature Importance Analysis Across Tree Ensemble Models",
        fontsize=14, fontweight="bold", y=1.02,
    )
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "10_feature_importance_analysis.png"), bbox_inches="tight")
    plt.close()


def generate_architecture_diagrams(lang="vi"):
    out_dir = FIG_VI if lang == "vi" else FIG_EN
    is_vi = lang == "vi"

    # 11. Workflow Diagram
    fig, ax = plt.subplots(figsize=(14, 5.2), dpi=200)
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 5.2)
    ax.axis("off")

    if is_vi:
        steps = [
            ("1. Thu thập Dữ liệu", "800 Sinh viên x 25 thuộc tính\nNhân khẩu học, Học tập,\nLối sống & Tâm lý", "#264653"),
            ("2. Tiền xử lý & EDA", "Xử lý 18 cột khuyết (KNNImputer)\nMã hóa One-Hot / Ordinal\nChuẩn hóa StandardScaler", "#2a9d8f"),
            ("3. Tạo Đặc trưng Mới", "Academic Engagement Index\nPsychological Stress Index\nStress-Engagement Ratio", "#457b9d"),
            ("4. Huấn luyện Mô hình", "So sánh 10 thuật toán ML\nKiểm định chéo 5-Fold CV\nTrên 2 kịch bản thực nghiệm", "#e76f51"),
            ("5. Đánh giá & Web UI", "ROC-AUC ~80.5% | F1 ~78.9%\nHệ thống Cảnh báo Sớm\nFastAPI Web Dashboard", "#e63946"),
        ]
        wf_title = "QUY TRÌNH THỰC HIỆN DỰ ÁN HỌC MÁY (NHÓM 03: ĐỖ QUANG MINH, VŨ QUANG HUY)"
    else:
        steps = [
            ("1. Data Acquisition", "800 Students x 25 Attributes\nDemographics, Engagement,\nLifestyle & Stress Metrics", "#264653"),
            ("2. Preprocessing & EDA", "Impute 18 Missing Cols (KNN)\nOne-Hot & Ordinal Encoding\nStandardScaler Normalization", "#2a9d8f"),
            ("3. Feature Engineering", "Academic Engagement Index\nPsychological Stress Index\nStress-to-Engagement Ratio", "#457b9d"),
            ("4. Model Training", "Compare 10 ML Classifiers\n5-Fold Stratified CV\nAcross 2 Experimental Scenarios", "#e76f51"),
            ("5. Evaluation & Web UI", "ROC-AUC ~80.5% | F1 ~78.9%\nEarly Warning System\nFastAPI Web Dashboard", "#e63946"),
        ]
        wf_title = "MACHINE LEARNING PROJECT WORKFLOW (GROUP 03: DO QUANG MINH, VU QUANG HUY)"

    for i, (title, desc, color) in enumerate(steps):
        x = 0.5 + i * 2.68
        rect = patches.FancyBboxPatch((x, 1.0), 2.25, 3.1, boxstyle="round,pad=0.12,rounding_size=0.2", facecolor=color, edgecolor="#1d3557", linewidth=2)
        ax.add_patch(rect)
        ax.text(x + 1.125, 3.35, title, ha="center", va="center", color="white", fontsize=10.5, fontweight="bold")
        ax.plot([x + 0.2, x + 2.05], [2.8, 2.8], color="white", alpha=0.6, linewidth=1.2)
        ax.text(x + 1.125, 1.95, desc, ha="center", va="center", color="#f1faee", fontsize=8.8)
        if i < len(steps) - 1:
            ax.annotate("", xy=(x + 2.64, 2.55), xytext=(x + 2.28, 2.55), arrowprops=dict(arrowstyle="->", lw=2.8, color="#1d3557"))

    ax.text(7.0, 4.65, wf_title, ha="center", fontsize=13.5, fontweight="bold", color="#1d3557")
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "11_system_workflow_diagram.png"), bbox_inches="tight")
    plt.close()

    # 12. System Architecture Diagram
    fig, ax = plt.subplots(figsize=(14, 6.2), dpi=200)
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 6.2)
    ax.axis("off")

    bg = patches.FancyBboxPatch((0.3, 0.3), 13.4, 5.2, boxstyle="round,pad=0.1", facecolor="#f8f9fa", edgecolor="#adb5bd", linewidth=1.5)
    ax.add_patch(bg)

    if is_vi:
        sys_title = "KIẾN TRÚC HỆ THỐNG CẢNH BÁO SỚM NGUY CƠ BỎ HỌC & GIAO DIỆN WEB DASHBOARD"
        l1_title = "TẦNG GIAO DIỆN (WEB UI INPUT)"
        inputs = [
            "• Hồ sơ cá nhân: Tuổi, Giới tính, Khoa, Năm học",
            "• Tham gia học tập: Chuyên cần (%),",
            "  Giờ tự học/ngày, GPA, Động lực, Số môn nợ",
            "• Căng thẳng tâm lý: Mức Stress,",
            "  Lo âu, Áp lực tài chính, Áp lực bạn bè",
            "• Lối sống & Hỗ trợ: Giờ ngủ, Giờ màn hình,",
            "  Hỗ trợ gia đình, Tiếp cận tư vấn tâm lý",
        ]
        l2_title = "TẦNG XỬ LÝ & MÔ HÌNH AI (BACKEND)"
        engine = [
            "1. Bộ tạo đặc trưng DomainFeatureEngineer:",
            "   - Chỉ số Gắn kết (Engagement Index)",
            "   - Chỉ số Căng thẳng (Stress Index)",
            "   - Tỷ số Căng thẳng / Gắn kết",
            "2. Pipeline Tiền xử lý tự động:",
            "   - KNNImputer (k=7) + StandardScaler",
            "3. Bộ đôi Mô hình Học máy (Dual Engine):",
            "   - Mô hình 1: Dự báo Burnout_Level (3 mức)",
            "   - Mô hình 2: Dự báo Dropout_Risk (%)",
        ]
        l3_title = "KẾT QUẢ & KHUYẾN NGHỊ (OUTPUT)"
        outputs = [
            "• Xác suất Nguy cơ Bỏ học (%):",
            "  Phân tầng Thấp / Trung bình / Cao",
            "• Dự báo Mức độ Kiệt sức (Burnout):",
            "  Low / Medium / High",
            "• Biểu đồ Chỉ số Giải thích AI:",
            "  Điểm Gắn kết vs. Điểm Căng thẳng",
            "• Khuyến nghị Can thiệp Sớm:",
            "  Tư vấn tâm lý, phụ đạo học tập,",
            "  hỗ trợ tài chính cá nhân hóa",
        ]
    else:
        sys_title = "STUDENT DROPOUT EARLY WARNING SYSTEM ARCHITECTURE & WEB DASHBOARD"
        l1_title = "PRESENTATION LAYER (WEB UI INPUT)"
        inputs = [
            "• Student Profile: Age, Gender, Dept, Year",
            "• Study Engagement: Attendance (%),",
            "  Study Hours/Day, GPA, Motivation, Backlogs",
            "• Psychological Stress: Stress Level,",
            "  Anxiety, Financial Stress, Peer Pressure",
            "• Lifestyle & Support: Sleep, Screen Time,",
            "  Family Support, Counseling Access",
        ]
        l2_title = "AI PROCESSING ENGINE (FASTAPI BACKEND)"
        engine = [
            "1. DomainFeatureEngineer Module:",
            "   - Academic Engagement Index",
            "   - Psychological Stress Index",
            "   - Stress-to-Engagement Ratio",
            "2. Automated Preprocessing Pipeline:",
            "   - KNNImputer (k=7) + StandardScaler",
            "3. Dual ML Inference Engine:",
            "   - Model 1: 3-Class Burnout_Level Predictor",
            "   - Model 2: Dropout_Risk Ensemble Predictor",
        ]
        l3_title = "DIAGNOSTIC DASHBOARD (OUTPUT)"
        outputs = [
            "• Dropout Risk Probability (%):",
            "  Low / Moderate / High Risk Tiering",
            "• Predicted Burnout Level:",
            "  Low / Medium / High Probabilities",
            "• Explainable Domain Indices:",
            "  Engagement Score vs. Stress Score",
            "• Personalized Interventions:",
            "  Academic tutoring, counseling referral,",
            "  and financial aid recommendations",
        ]

    ax.text(7.0, 5.8, sys_title, ha="center", fontsize=13, fontweight="bold", color="#1d3557")

    b1 = patches.FancyBboxPatch((0.7, 0.8), 3.6, 4.2, boxstyle="round,pad=0.1", facecolor="#e9ecef", edgecolor="#495057", linewidth=1.8)
    ax.add_patch(b1)
    ax.text(2.5, 4.65, l1_title, ha="center", fontweight="bold", fontsize=10, color="#1d3557")
    for idx, line in enumerate(inputs):
        ax.text(0.9, 4.1 - idx * 0.45, line, fontsize=8.8, color="#212529")

    b2 = patches.FancyBboxPatch((5.1, 0.8), 3.8, 4.2, boxstyle="round,pad=0.1", facecolor="#d8f3dc", edgecolor="#2d6a4f", linewidth=1.8)
    ax.add_patch(b2)
    ax.text(7.0, 4.65, l2_title, ha="center", fontweight="bold", fontsize=10, color="#1b4332")
    for idx, line in enumerate(engine):
        ax.text(5.3, 4.15 - idx * 0.38, line, fontsize=8.8, color="#1b4332", fontweight="bold" if line[0].isdigit() else "normal")

    b3 = patches.FancyBboxPatch((9.7, 0.8), 3.6, 4.2, boxstyle="round,pad=0.1", facecolor="#ffe5d9", edgecolor="#9d0208", linewidth=1.8)
    ax.add_patch(b3)
    ax.text(11.5, 4.65, l3_title, ha="center", fontweight="bold", fontsize=10, color="#9d0208")
    for idx, line in enumerate(outputs):
        ax.text(9.9, 4.15 - idx * 0.38, line, fontsize=8.8, color="#6a040f")

    ax.annotate("", xy=(5.05, 2.9), xytext=(4.35, 2.9), arrowprops=dict(arrowstyle="->", lw=3, color="#1d3557"))
    ax.annotate("", xy=(9.65, 2.9), xytext=(8.95, 2.9), arrowprops=dict(arrowstyle="->", lw=3, color="#1d3557"))
    ax.text(4.7, 3.15, "JSON API", ha="center", fontsize=8.5, fontweight="bold")
    ax.text(9.3, 3.15, "Prediction", ha="center", fontsize=8.5, fontweight="bold")

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "12_system_architecture_diagram.png"), bbox_inches="tight")
    plt.close()


if __name__ == "__main__":
    df = pd.read_csv(DATA_PATH)
    df["Dropout_Binary"] = (df["Dropout_Risk"] == "Yes").astype(int)
    df["Burnout_Numeric"] = df["Burnout_Level"].map({"Low": 0, "Medium": 1, "High": 2})
    metrics_df = pd.read_csv(METRICS_PATH)

    for lang in ["vi", "en"]:
        print(f"Generating 12 figures for language: {lang.upper()}...")
        generate_eda_figures(df, lang=lang)
        generate_model_figures(df, metrics_df, lang=lang)
        generate_architecture_diagrams(lang=lang)

    print("Successfully generated both figures_vi/ (100% Vietnamese) and figures_en/ (100% English)!")
