"""
Trains the Burnout_Level classifier (3 classes: Low, Medium, High) for two-stage & UI prediction,
and generates architectural workflow & UI diagrams for the SIC Final Report and Presentation Slides.
"""

import os
import joblib
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

from sklearn.model_selection import StratifiedKFold, cross_validate
from sklearn.pipeline import Pipeline
from sklearn.ensemble import RandomForestClassifier, ExtraTreesClassifier, GradientBoostingClassifier, VotingClassifier
from sklearn.linear_model import LogisticRegression
from xgboost import XGBClassifier

import sys
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)
from src.train_and_compare_models import DomainFeatureEngineer, build_preprocessor

warnings.filterwarnings("ignore")

EXPANDED_PATH = os.path.join(BASE_DIR, "data", "student_burnout_dropout_expanded.csv")
ORIG_PATH = os.path.join(BASE_DIR, "data", "student_burnout_dropout_dataset_2.csv")
DATA_PATH = EXPANDED_PATH if os.path.exists(EXPANDED_PATH) else ORIG_PATH
MODEL_DIR = os.path.join(BASE_DIR, "models")
FIG_DIR = os.path.join(BASE_DIR, "figures")


def train_burnout_model():
    df = pd.read_csv(DATA_PATH)
    X_raw = df.drop(columns=["Student_ID", "Dropout_Risk", "Burnout_Level"])
    burn_map = {"Low": 0, "Medium": 1, "High": 2}
    y_burn = df["Burnout_Level"].map(burn_map)

    fe_sample = DomainFeatureEngineer(include_burnout=False).fit_transform(X_raw)

    clf = VotingClassifier(
        estimators=[
            ("lr", LogisticRegression(C=0.5, max_iter=1000, random_state=42)),
            ("rf", RandomForestClassifier(n_estimators=250, max_depth=8, min_samples_leaf=4, random_state=42)),
            ("xgb", XGBClassifier(n_estimators=150, learning_rate=0.04, max_depth=3, random_state=42, eval_metric="mlogloss")),
        ],
        voting="soft",
        weights=[1, 3, 2],
    )

    pipe = Pipeline(
        [
            ("fe", DomainFeatureEngineer(include_burnout=False)),
            ("pre", build_preprocessor(fe_sample)),
            ("clf", clf),
        ]
    )

    scores = cross_validate(pipe, X_raw, y_burn, cv=5, scoring=["accuracy", "f1_macro"])
    print(f"Burnout_Level 3-Class Model -> CV Acc: {scores['test_accuracy'].mean():.4f} | F1-Macro: {scores['test_f1_macro'].mean():.4f}")

    pipe.fit(X_raw, y_burn)
    joblib.dump(pipe, os.path.join(MODEL_DIR, "best_burnout_model.pkl"))
    print("Saved models/best_burnout_model.pkl")


def draw_workflow_and_system_diagrams():
    # 11. Workflow Diagram
    fig, ax = plt.subplots(figsize=(14, 5.2), dpi=200)
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 5.2)
    ax.axis("off")

    steps = [
        ("1. Thu thập & Mở rộng\n(Data Acquisition)", "2,000 Sinh viên x 25 thuộc tính\nNhân khẩu học, Học tập,\nLối sống & Tâm lý", "#264653"),
        ("2. Làm sạch 6 Bước\n(Data Cleaning)", "KNNImputer + Lọc ngoại lai IQR\nMã hóa One-Hot / Ordinal\nChuẩn hóa StandardScaler", "#2a9d8f"),
        ("3. Kỹ thuật Đặc trưng\n(Feature Engineering)", "Academic Engagement Index\nPsychological Stress Index\nStress-Engagement Ratio", "#457b9d"),
        ("4. Huấn luyện Mô hình\n(Model Training)", "So sánh 4 mô hình tiêu biểu:\nLR, RF, XGBoost, Soft Voting\n5-Fold Stratified CV", "#e76f51"),
        ("5. Đánh giá & Triển khai\n(Evaluation & Web UI)", "ROC-AUC ~89.1% | F1 ~83.0%\nHệ thống Cảnh báo Sớm\nFastAPI Web Dashboard", "#e63946"),
    ]

    for i, (title, desc, color) in enumerate(steps):
        x = 0.5 + i * 2.68
        rect = patches.FancyBboxPatch(
            (x, 1.0), 2.25, 3.1,
            boxstyle="round,pad=0.12,rounding_size=0.2",
            facecolor=color, edgecolor="#1d3557", linewidth=2
        )
        ax.add_patch(rect)
        ax.text(x + 1.125, 3.35, title, ha="center", va="center", color="white", fontsize=10.5, fontweight="bold")
        ax.plot([x + 0.2, x + 2.05], [2.8, 2.8], color="white", alpha=0.6, linewidth=1.2)
        ax.text(x + 1.125, 1.95, desc, ha="center", va="center", color="#f1faee", fontsize=8.8)

        if i < len(steps) - 1:
            ax.annotate(
                "", xy=(x + 2.64, 2.55), xytext=(x + 2.28, 2.55),
                arrowprops=dict(arrowstyle="->", lw=2.8, color="#1d3557")
            )

    ax.text(7.0, 4.65, "QUY TRÌNH THỰC HIỆN DỰ ÁN HỌC MÁY (ML PROJECT WORKFLOW - NHÓM 03)", ha="center", fontsize=13.5, fontweight="bold", color="#1d3557")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "11_system_workflow_diagram.png"), bbox_inches="tight")
    plt.close()

    # 12. System Architecture & UI Dashboard Diagram
    fig, ax = plt.subplots(figsize=(14, 6.2), dpi=200)
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 6.2)
    ax.axis("off")

    # Background container
    bg = patches.FancyBboxPatch((0.3, 0.3), 13.4, 5.2, boxstyle="round,pad=0.1", facecolor="#f8f9fa", edgecolor="#adb5bd", linewidth=1.5)
    ax.add_patch(bg)
    ax.text(7.0, 5.8, "KIẾN TRÚC HỆ THỐNG CẢNH BÁO SỚM NGUY CƠ BỎ HỌC (SYSTEM ARCHITECTURE & UI)", ha="center", fontsize=13, fontweight="bold", color="#1d3557")

    # Layer 1: Input Layer
    b1 = patches.FancyBboxPatch((0.7, 0.8), 3.6, 4.2, boxstyle="round,pad=0.1", facecolor="#e9ecef", edgecolor="#495057", linewidth=1.8)
    ax.add_patch(b1)
    ax.text(2.5, 4.65, "TẦNG GIAO DIỆN (WEB UI INPUT)", ha="center", fontweight="bold", fontsize=10.5, color="#1d3557")
    inputs = [
        "• Hồ sơ cá nhân: Age, Gender, Dept, Year",
        "• Tham gia học tập: Attendance (%),",
        "  Study Hours/Day, GPA, Motivation, Backlogs",
        "• Căng thẳng tâm lý: Stress Level,",
        "  Anxiety, Financial Stress, Peer Pressure",
        "• Lối sống & Hỗ trợ: Sleep, Screen Time,",
        "  Family Support, Counseling Access",
    ]
    for idx, line in enumerate(inputs):
        ax.text(0.9, 4.1 - idx * 0.45, line, fontsize=8.8, color="#212529")

    # Layer 2: AI Engine
    b2 = patches.FancyBboxPatch((5.1, 0.8), 3.8, 4.2, boxstyle="round,pad=0.1", facecolor="#d8f3dc", edgecolor="#2d6a4f", linewidth=1.8)
    ax.add_patch(b2)
    ax.text(7.0, 4.65, "TẦNG XỬ LÝ & MÔ HÌNH AI (BACKEND)", ha="center", fontweight="bold", fontsize=10.5, color="#1b4332")
    engine = [
        "1. DomainFeatureEngineer:",
        "   - Academic Engagement Index",
        "   - Psychological Stress Index",
        "   - Stress-to-Engagement Ratio",
        "2. Preprocessing Pipeline:",
        "   - KNNImputer (k=7) + StandardScaler",
        "3. Dual ML Inference Engine:",
        "   - Model 1: Dự báo Burnout_Level",
        "   - Model 2: Dự báo Dropout_Risk",
    ]
    for idx, line in enumerate(engine):
        ax.text(5.3, 4.15 - idx * 0.38, line, fontsize=8.8, color="#1b4332", fontweight="bold" if line[0].isdigit() else "normal")

    # Layer 3: Output Dashboard
    b3 = patches.FancyBboxPatch((9.7, 0.8), 3.6, 4.2, boxstyle="round,pad=0.1", facecolor="#ffe5d9", edgecolor="#9d0208", linewidth=1.8)
    ax.add_patch(b3)
    ax.text(11.5, 4.65, "KẾT QUẢ & KHUYẾN NGHỊ (OUTPUT)", ha="center", fontweight="bold", fontsize=10.5, color="#9d0208")
    outputs = [
        "• Xác suất Nguy cơ Bỏ học (%):",
        "  Phân tầng Thấp / Trung bình / Cao",
        "• Dự báo Mức độ Kiệt sức (Burnout):",
        "  Low / Medium / High",
        "• Biểu đồ Chỉ số Tổng hợp:",
        "  Điểm Gắn kết vs. Điểm Căng thẳng",
        "• Khuyến nghị Can thiệp Sớm:",
        "  Tư vấn tâm lý, phụ đạo học tập,",
        "  hỗ trợ tài chính cá nhân hóa",
    ]
    for idx, line in enumerate(outputs):
        ax.text(9.9, 4.15 - idx * 0.38, line, fontsize=8.8, color="#6a040f")

    ax.annotate("", xy=(5.05, 2.9), xytext=(4.35, 2.9), arrowprops=dict(arrowstyle="->", lw=3, color="#1d3557"))
    ax.annotate("", xy=(9.65, 2.9), xytext=(8.95, 2.9), arrowprops=dict(arrowstyle="->", lw=3, color="#1d3557"))
    ax.text(4.7, 3.15, "JSON API", ha="center", fontsize=8.5, fontweight="bold")
    ax.text(9.3, 3.15, "Prediction", ha="center", fontsize=8.5, fontweight="bold")

    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "12_system_architecture_diagram.png"), bbox_inches="tight")
    plt.close()
    print("Generated workflow and system architecture diagrams.")


if __name__ == "__main__":
    train_burnout_model()
    draw_workflow_and_system_diagrams()
