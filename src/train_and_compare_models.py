"""
Machine Learning Pipeline: Preprocessing, Feature Engineering, Model Training & Comparison
Project: Student Burnout & Dropout Risk Prediction (Nhóm 03: Đỗ Quang Minh, Vũ Quang Huy)
"""

import os
import json
import joblib
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.base import BaseEstimator, TransformerMixin
from sklearn.model_selection import StratifiedKFold, train_test_split, cross_validate, GridSearchCV
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer, KNNImputer
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.metrics import (
    accuracy_score,
    precision_score,
    recall_score,
    f1_score,
    roc_auc_score,
    roc_curve,
    precision_recall_curve,
    average_precision_score,
    confusion_matrix,
    classification_report,
)
from sklearn.calibration import CalibratedClassifierCV
from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import (
    RandomForestClassifier,
    ExtraTreesClassifier,
    GradientBoostingClassifier,
    VotingClassifier,
)
from sklearn.neural_network import MLPClassifier
from xgboost import XGBClassifier

warnings.filterwarnings("ignore")

plt.rcParams["figure.dpi"] = 150
plt.rcParams["savefig.dpi"] = 200
plt.rcParams["font.sans-serif"] = ["Segoe UI", "DejaVu Sans", "Arial"]
plt.rcParams["axes.unicode_minus"] = False
sns.set_theme(style="whitegrid")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(BASE_DIR, "data", "student_burnout_dropout_dataset_2.csv")
FIG_DIR = os.path.join(BASE_DIR, "figures")
MODEL_DIR = os.path.join(BASE_DIR, "models")
REPORT_DIR = os.path.join(BASE_DIR, "reports")

for d in [FIG_DIR, MODEL_DIR, REPORT_DIR]:
    os.makedirs(d, exist_ok=True)


class DomainFeatureEngineer(BaseEstimator, TransformerMixin):
    """
    Creates domain-specific features capturing Stress & Academic Engagement interactions.
    Operates cleanly inside scikit-learn Pipeline to prevent data leakage across folds.
    """

    def __init__(self, include_burnout=True):
        self.include_burnout = include_burnout

    def fit(self, X, y=None):
        self.num_medians_ = {}
        num_cols = [
            "Attendance_Percent",
            "Study_Hours_Per_Day",
            "Previous_GPA",
            "Backlogs",
            "Sleep_Hours",
            "Screen_Time_Hours",
            "Exercise_Freq_Per_Week",
            "Social_Activity_Score",
            "Financial_Stress_Score",
            "Family_Support_Score",
            "Stress_Level",
            "Anxiety_Score",
            "Motivation_Score",
            "Peer_Pressure_Score",
        ]
        for c in num_cols:
            if c in X.columns:
                self.num_medians_[c] = X[c].median()
        return self

    def transform(self, X):
        df = X.copy()

        # Temporary imputed series for composite index computation
        def get_col(name):
            return df[name].fillna(self.num_medians_.get(name, 0.0))

        att = get_col("Attendance_Percent")
        study = get_col("Study_Hours_Per_Day")
        mot = get_col("Motivation_Score")
        gpa = get_col("Previous_GPA")
        stress = get_col("Stress_Level")
        anx = get_col("Anxiety_Score")
        fin = get_col("Financial_Stress_Score")
        peer = get_col("Peer_Pressure_Score")
        fam = get_col("Family_Support_Score")
        soc = get_col("Social_Activity_Score")
        ex = get_col("Exercise_Freq_Per_Week")
        sleep = get_col("Sleep_Hours")
        screen = get_col("Screen_Time_Hours")
        backlogs = get_col("Backlogs")

        # 1. Academic Engagement Index (0 - 10 scale)
        df["Academic_Engagement_Index"] = 0.35 * (att / 10.0) + 0.25 * study + 0.25 * mot + 0.15 * gpa

        # 2. Psychological Stress Index (0 - 10 scale)
        df["Psychological_Stress_Index"] = 0.35 * stress + 0.30 * anx + 0.20 * fin + 0.15 * peer

        # 3. Stress-to-Engagement Ratio
        df["Stress_Engagement_Ratio"] = df["Psychological_Stress_Index"] / (df["Academic_Engagement_Index"] + 1.0)

        # 4. Net Stress vs Support Balance
        df["Net_Stress_Support_Balance"] = df["Psychological_Stress_Index"] - (0.5 * fam + 0.3 * soc + 0.2 * ex)

        # 5. Digital Fatigue Ratio (Screen Time vs Sleep)
        df["Digital_Fatigue_Ratio"] = screen / (sleep + 1.0)

        # 6. Academic Efficiency (GPA per Study Hour & Attendance)
        df["Study_Attendance_Product"] = (att / 100.0) * study
        df["Academic_Risk_Burden"] = backlogs * 1.5 + (100.0 - att) / 20.0

        # 7. Ordinal mapping for Family_Income_Bracket
        income_map = {"Low": 0, "Lower-Middle": 1, "Middle": 2, "Upper-Middle": 3, "High": 4}
        if "Family_Income_Bracket" in df.columns:
            df["Income_Ordinal"] = df["Family_Income_Bracket"].map(income_map).fillna(2)

        # 8. Ordinal mapping for Burnout_Level if included
        if self.include_burnout and "Burnout_Level" in df.columns:
            burn_map = {"Low": 0, "Medium": 1, "High": 2}
            df["Burnout_Ordinal"] = df["Burnout_Level"].map(burn_map).fillna(1)
        elif "Burnout_Level" in df.columns:
            df = df.drop(columns=["Burnout_Level"])

        return df


def build_preprocessor(X_sample):
    num_cols = X_sample.select_dtypes(include=[np.number]).columns.tolist()
    cat_cols = X_sample.select_dtypes(exclude=[np.number]).columns.tolist()

    num_pipeline = Pipeline(
        [
            ("imputer", KNNImputer(n_neighbors=7)),
            ("scaler", StandardScaler()),
        ]
    )
    cat_pipeline = Pipeline(
        [
            ("imputer", SimpleImputer(strategy="most_frequent")),
            ("ohe", OneHotEncoder(handle_unknown="ignore", sparse_output=False)),
        ]
    )

    return ColumnTransformer(
        [
            ("num", num_pipeline, num_cols),
            ("cat", cat_pipeline, cat_cols),
        ]
    )


def get_models(random_state=42):
    lr = LogisticRegression(C=0.5, max_iter=1000, random_state=random_state)
    svm = CalibratedClassifierCV(SVC(C=1.2, kernel="rbf", gamma="scale", random_state=random_state), cv=3)
    knn = KNeighborsClassifier(n_neighbors=17, weights="distance", p=2)
    dt = DecisionTreeClassifier(max_depth=5, min_samples_leaf=15, random_state=random_state)
    rf = RandomForestClassifier(
        n_estimators=300,
        max_depth=8,
        min_samples_split=8,
        min_samples_leaf=4,
        max_features="sqrt",
        random_state=random_state,
        n_jobs=-1,
    )
    et = ExtraTreesClassifier(
        n_estimators=300,
        max_depth=9,
        min_samples_split=8,
        min_samples_leaf=4,
        random_state=random_state,
        n_jobs=-1,
    )
    gb = GradientBoostingClassifier(
        n_estimators=150,
        learning_rate=0.04,
        max_depth=3,
        subsample=0.85,
        min_samples_leaf=10,
        random_state=random_state,
    )
    xgb = XGBClassifier(
        n_estimators=200,
        learning_rate=0.035,
        max_depth=3,
        subsample=0.85,
        colsample_bytree=0.8,
        reg_lambda=2.0,
        min_child_weight=5,
        random_state=random_state,
        eval_metric="logloss",
        n_jobs=-1,
    )
    mlp = MLPClassifier(
        hidden_layer_sizes=(64, 32),
        activation="relu",
        alpha=0.5,
        learning_rate_init=0.003,
        max_iter=600,
        early_stopping=True,
        random_state=random_state,
    )
    voting = VotingClassifier(
        estimators=[
            ("lr", lr),
            ("rf", rf),
            ("et", et),
            ("gb", gb),
            ("xgb", xgb),
            ("svm", svm),
        ],
        voting="soft",
        weights=[2, 3, 2, 2, 2, 2],
    )

    return {
        "Logistic Regression": lr,
        "SVM (RBF Kernel)": svm,
        "K-Nearest Neighbors": knn,
        "Decision Tree": dt,
        "Random Forest": rf,
        "Extra Trees": et,
        "Gradient Boosting": gb,
        "XGBoost": xgb,
        "MLP Neural Network": mlp,
        "Soft Voting Ensemble": voting,
    }


def evaluate_models_scenario(df, include_burnout=True, scenario_name="Full_Features"):
    X_raw = df.drop(columns=["Student_ID", "Dropout_Risk"])
    y = (df["Dropout_Risk"] == "Yes").astype(int)

    X_train, X_test, y_train, y_test = train_test_split(
        X_raw, y, test_size=0.20, stratify=y, random_state=42
    )

    fe_sample = DomainFeatureEngineer(include_burnout=include_burnout).fit_transform(X_train)
    models = get_models(random_state=42)

    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    results = []
    roc_data = {}
    pr_data = {}
    trained_pipelines = {}

    for name, clf in models.items():
        pipe = Pipeline(
            [
                ("fe", DomainFeatureEngineer(include_burnout=include_burnout)),
                ("pre", build_preprocessor(fe_sample)),
                ("clf", clf),
            ]
        )

        # 5-Fold Cross Validation on whole dataset for robust comparison
        cv_scores = cross_validate(
            pipe,
            X_raw,
            y,
            cv=cv,
            scoring=["accuracy", "precision", "recall", "f1", "roc_auc"],
            n_jobs=-1,
        )

        # Train on 80% split and evaluate on 20% hold-out test set
        pipe.fit(X_train, y_train)
        y_pred = pipe.predict(X_test)
        y_prob = pipe.predict_proba(X_test)[:, 1]

        test_acc = accuracy_score(y_test, y_pred)
        test_prec = precision_score(y_test, y_pred)
        test_rec = recall_score(y_test, y_pred)
        test_f1 = f1_score(y_test, y_pred)
        test_auc = roc_auc_score(y_test, y_prob)
        test_ap = average_precision_score(y_test, y_prob)

        fpr, tpr, _ = roc_curve(y_test, y_prob)
        prec_curve, rec_curve, _ = precision_recall_curve(y_test, y_prob)

        roc_data[name] = (fpr, tpr, test_auc)
        pr_data[name] = (rec_curve, prec_curve, test_ap)
        trained_pipelines[name] = (pipe, y_test, y_pred, y_prob)

        results.append(
            {
                "Scenario": scenario_name,
                "Model": name,
                "CV_Accuracy_Mean": cv_scores["test_accuracy"].mean(),
                "CV_Accuracy_Std": cv_scores["test_accuracy"].std(),
                "CV_Precision_Mean": cv_scores["test_precision"].mean(),
                "CV_Recall_Mean": cv_scores["test_recall"].mean(),
                "CV_F1_Mean": cv_scores["test_f1"].mean(),
                "CV_F1_Std": cv_scores["test_f1"].std(),
                "CV_ROC_AUC_Mean": cv_scores["test_roc_auc"].mean(),
                "CV_ROC_AUC_Std": cv_scores["test_roc_auc"].std(),
                "Test_Accuracy": test_acc,
                "Test_Precision": test_prec,
                "Test_Recall": test_rec,
                "Test_F1": test_f1,
                "Test_ROC_AUC": test_auc,
                "Test_PR_AUC": test_ap,
            }
        )

    res_df = pd.DataFrame(results).sort_values("CV_ROC_AUC_Mean", ascending=False).reset_index(drop=True)
    return res_df, roc_data, pr_data, trained_pipelines


def plot_model_comparison(res_full, res_early):
    fig, axes = plt.subplots(1, 2, figsize=(16, 6.5))

    # Plot 1: CV Metrics Bar Chart for Full Scenario
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
    axes[0].set_xlabel("Điểm đánh giá 5-Fold Cross-Validation (%)", fontweight="bold")
    axes[0].set_title("So sánh 10 Mô hình Học máy (Kịch bản 1: Đầy đủ đặc trưng)", fontsize=12, fontweight="bold")
    axes[0].legend(loc="lower right")

    # Plot 2: Comparison between Scenario 1 (With Burnout) vs Scenario 2 (Without Burnout)
    merged = pd.merge(
        res_full[["Model", "CV_ROC_AUC_Mean", "CV_F1_Mean"]],
        res_early[["Model", "CV_ROC_AUC_Mean", "CV_F1_Mean"]],
        on="Model",
        suffixes=("_Full", "_Early"),
    ).sort_values("CV_ROC_AUC_Mean_Full", ascending=True)

    y_pos2 = np.arange(len(merged))
    h2 = 0.35
    b1 = axes[1].barh(
        y_pos2 + h2 / 2,
        merged["CV_ROC_AUC_Mean_Full"] * 100,
        h2,
        label="Kịch bản 1: Có Burnout_Level (Full)",
        color="#264653",
        edgecolor="#222",
    )
    b2 = axes[1].barh(
        y_pos2 - h2 / 2,
        merged["CV_ROC_AUC_Mean_Early"] * 100,
        h2,
        label="Kịch bản 2: Chỉ dùng Stress & Engagement thô (Early Warning)",
        color="#e9c46a",
        edgecolor="#222",
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
    axes[1].set_xlabel("Điểm ROC-AUC trung bình 5-Fold CV (%)", fontweight="bold")
    axes[1].set_title("Đánh giá Vai trò của Biến trung gian Burnout_Level (ROC-AUC)", fontsize=12, fontweight="bold")
    axes[1].legend(loc="lower right", fontsize=9)

    plt.suptitle("Kết quả Thực nghiệm & So sánh Hiệu năng các Mô hình Dự đoán Bỏ học", fontsize=14, fontweight="bold", y=1.02)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "07_model_performance_comparison.png"), bbox_inches="tight")
    plt.close()


def plot_roc_and_pr_curves(roc_data, pr_data):
    fig, axes = plt.subplots(1, 2, figsize=(15, 6.2))
    colors = sns.color_palette("tab10", len(roc_data))

    for (name, (fpr, tpr, auc_val)), color in zip(roc_data.items(), colors):
        lw = 2.6 if "Ensemble" in name or "Random Forest" in name or "Logistic" in name else 1.5
        axes[0].plot(fpr, tpr, label=f"{name} (AUC = {auc_val:.3f})", color=color, linewidth=lw)

    axes[0].plot([0, 1], [0, 1], "k--", alpha=0.5, label="Random Guess (0.500)")
    axes[0].set_xlabel("False Positive Rate (1 - Specificity)", fontweight="bold")
    axes[0].set_ylabel("True Positive Rate (Sensitivity / Recall)", fontweight="bold")
    axes[0].set_title("Đường cong ROC (Receiver Operating Characteristic) - Tập Test", fontsize=12, fontweight="bold")
    axes[0].legend(loc="lower right", fontsize=8.5)

    for (name, (rec, prec, ap_val)), color in zip(pr_data.items(), colors):
        lw = 2.6 if "Ensemble" in name or "Random Forest" in name or "Logistic" in name else 1.5
        axes[1].plot(rec, prec, label=f"{name} (AP = {ap_val:.3f})", color=color, linewidth=lw)

    axes[1].set_xlabel("Recall (Độ nhạy)", fontweight="bold")
    axes[1].set_ylabel("Precision (Độ chính xác dương)", fontweight="bold")
    axes[1].set_title("Đường cong Precision-Recall - Tập Test", fontsize=12, fontweight="bold")
    axes[1].legend(loc="lower left", fontsize=8.5)

    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "08_roc_and_precision_recall_curves.png"), bbox_inches="tight")
    plt.close()


def plot_confusion_matrices(trained_pipelines):
    top_models = ["Soft Voting Ensemble", "Random Forest", "XGBoost", "Logistic Regression"]
    fig, axes = plt.subplots(1, 4, figsize=(18, 4.3))

    for ax, name in zip(axes, top_models):
        _, y_test, y_pred, _ = trained_pipelines[name]
        cm = confusion_matrix(y_test, y_pred)
        cm_pct = cm / cm.sum() * 100
        annot = np.array([[f"{v}\n({p:.1f}%)" for v, p in zip(r_v, r_p)] for r_v, r_p in zip(cm, cm_pct)])

        sns.heatmap(
            cm,
            annot=annot,
            fmt="",
            cmap="Blues",
            cbar=False,
            xticklabels=["Predict: No", "Predict: Yes"],
            yticklabels=["Actual: No", "Actual: Yes"],
            annot_kws={"size": 11, "weight": "bold"},
            ax=ax,
        )
        acc = accuracy_score(y_test, y_pred)
        f1 = f1_score(y_test, y_pred)
        ax.set_title(f"{name}\nAcc: {acc*100:.1f}% | F1: {f1*100:.1f}%", fontsize=11, fontweight="bold")

    plt.suptitle("Ma trận Nhầm lẫn (Confusion Matrix) của 4 Mô hình Tiêu biểu trên Tập Kiểm thử (n=160)", fontsize=13.5, fontweight="bold", y=1.05)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "09_confusion_matrices_top_models.png"), bbox_inches="tight")
    plt.close()


def plot_feature_importance(trained_pipelines):
    rf_pipe = trained_pipelines["Random Forest"][0]
    xgb_pipe = trained_pipelines["XGBoost"][0]

    pre = rf_pipe.named_steps["pre"]
    num_features = pre.transformers_[0][2]
    cat_ohe = pre.transformers_[1][1].named_steps["ohe"]
    cat_features = list(cat_ohe.get_feature_names_out(pre.transformers_[1][2]))
    all_features = num_features + cat_features

    rf_imp = pd.Series(rf_pipe.named_steps["clf"].feature_importances_, index=all_features).sort_values(ascending=True).tail(15)
    xgb_imp = pd.Series(xgb_pipe.named_steps["clf"].feature_importances_, index=all_features).sort_values(ascending=True).tail(15)

    fig, axes = plt.subplots(1, 2, figsize=(16, 6.5))

    rf_imp.plot(kind="barh", color=sns.color_palette("viridis", 15), edgecolor="#222", ax=axes[0])
    axes[0].set_title("Top 15 Đặc trưng Quan trọng nhất - Random Forest", fontsize=12, fontweight="bold")
    axes[0].set_xlabel("Mức độ quan trọng (Gini Importance)", fontweight="bold")

    xgb_imp.plot(kind="barh", color=sns.color_palette("magma", 15), edgecolor="#222", ax=axes[1])
    axes[1].set_title("Top 15 Đặc trưng Quan trọng nhất - XGBoost", fontsize=12, fontweight="bold")
    axes[1].set_xlabel("Mức độ quan trọng (Gain Importance)", fontweight="bold")

    plt.suptitle("Phân tích Mức độ Quan trọng của các Đặc trưng (Feature Importance Analysis)", fontsize=14, fontweight="bold", y=1.02)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "10_feature_importance_analysis.png"), bbox_inches="tight")
    plt.close()

    return rf_imp.sort_values(ascending=False), xgb_imp.sort_values(ascending=False)


if __name__ == "__main__":
    print("Loading dataset...")
    df = pd.read_csv(DATA_PATH)

    print("Evaluating Scenario 1: Full Feature Set (Including Burnout_Level & Engineered Indices)...")
    res_full, roc_full, pr_full, pipes_full = evaluate_models_scenario(
        df, include_burnout=True, scenario_name="Scenario 1: Full Features (With Burnout_Level)"
    )

    print("Evaluating Scenario 2: Early Warning Feature Set (Stress & Engagement Only, Without Burnout_Level)...")
    res_early, roc_early, pr_early, pipes_early = evaluate_models_scenario(
        df, include_burnout=False, scenario_name="Scenario 2: Early Warning (Without Burnout_Level)"
    )

    # Save CSV tables
    all_results = pd.concat([res_full, res_early], ignore_index=True)
    all_results.to_csv(os.path.join(REPORT_DIR, "model_comparison_metrics.csv"), index=False)

    print("\n--- SCENARIO 1 RESULTS (5-FOLD CV & TEST) ---")
    print(res_full[["Model", "CV_Accuracy_Mean", "CV_F1_Mean", "CV_ROC_AUC_Mean", "Test_Accuracy", "Test_F1", "Test_ROC_AUC"]].to_string(index=False))

    print("\n--- SCENARIO 2 RESULTS (5-FOLD CV & TEST) ---")
    print(res_early[["Model", "CV_Accuracy_Mean", "CV_F1_Mean", "CV_ROC_AUC_Mean", "Test_Accuracy", "Test_F1", "Test_ROC_AUC"]].to_string(index=False))

    # Generate comparison figures
    plot_model_comparison(res_full, res_early)
    plot_roc_and_pr_curves(roc_full, pr_full)
    plot_confusion_matrices(pipes_full)
    rf_top, xgb_top = plot_feature_importance(pipes_full)

    # Save best model pipeline
    best_model_name = res_full.iloc[0]["Model"]
    best_pipe = pipes_full[best_model_name][0]
    joblib.dump(best_pipe, os.path.join(MODEL_DIR, "best_dropout_model_full.pkl"))
    joblib.dump(pipes_early[res_early.iloc[0]["Model"]][0], os.path.join(MODEL_DIR, "best_dropout_model_early_warning.pkl"))

    print(f"\nSaved best models and all evaluation figures to {FIG_DIR} and {REPORT_DIR}")
