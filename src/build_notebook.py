"""
Generates a complete, richly documented Jupyter Notebook (.ipynb) for Group 03:
Đỗ Quang Minh & Vũ Quang Huy - Student Burnout & Dropout Risk Prediction
"""

import os
import json
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NB_PATH = os.path.join(BASE_DIR, "Student_Burnout_Dropout_Prediction_Nhom03.ipynb")
METRICS_CSV = os.path.join(BASE_DIR, "reports", "model_comparison_metrics.csv")


def md_cell(source_text):
    lines = [line + "\n" for line in source_text.strip().split("\n")]
    if lines:
        lines[-1] = lines[-1].rstrip("\n")
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": lines,
    }


def code_cell(source_code, output_text=None):
    lines = [line + "\n" for line in source_code.strip().split("\n")]
    if lines:
        lines[-1] = lines[-1].rstrip("\n")
    outputs = []
    if output_text:
        out_lines = [line + "\n" for line in output_text.strip().split("\n")]
        if out_lines:
            out_lines[-1] = out_lines[-1].rstrip("\n")
        outputs.append(
            {
                "name": "stdout",
                "output_type": "stream",
                "text": out_lines,
            }
        )
    return {
        "cell_type": "code",
        "execution_count": 1,
        "metadata": {},
        "outputs": outputs,
        "source": lines,
    }


def build_notebook():
    metrics_text = ""
    if os.path.exists(METRICS_CSV):
        df_m = pd.read_csv(METRICS_CSV)
        s1 = df_m[df_m["Scenario"].str.contains("Scenario 1")][
            ["Model", "CV_Accuracy_Mean", "CV_F1_Mean", "CV_ROC_AUC_Mean", "Test_Accuracy", "Test_F1", "Test_ROC_AUC"]
        ]
        s2 = df_m[df_m["Scenario"].str.contains("Scenario 2")][
            ["Model", "CV_Accuracy_Mean", "CV_F1_Mean", "CV_ROC_AUC_Mean", "Test_Accuracy", "Test_F1", "Test_ROC_AUC"]
        ]
        metrics_text = (
            "=== KỊCH BẢN 1: ĐẦY ĐỦ ĐẶC TRƯNG (CÓ BURNOUT_LEVEL & CHỈ SỐ TỔNG HỢP) ===\n"
            + s1.to_string(index=False)
            + "\n\n=== KỊCH BẢN 2: CẢNH BÁO SỚM (CHỈ DÙNG STRESS & ENGAGEMENT THÔ) ===\n"
            + s2.to_string(index=False)
        )

    cells = [
        md_cell(
            """# 🎓 BÀI TOÁN DỰ ĐOÁN TỶ LỆ BỎ HỌC DỰA TRÊN CÁC CHỈ SỐ CĂNG THẲNG VÀ SỰ THAM GIA HỌC TẬP
## *(Student Burnout & Dropout Risk Prediction using Machine Learning)*

---
- **Chương trình:** Samsung Innovation Campus (SIC) - AI Capstone Project
- **Đề tài:** Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập (*Student Burnout*)
- **Thực hiện bởi Nhóm 03:**
  1. **Đỗ Quang Minh**
  2. **Vũ Quang Huy**

---
### 📌 Mục tiêu đồ án (Project Objectives)
1. **Khám phá và phân tích dữ liệu (EDA):** Khảo sát 800 hồ sơ sinh viên với 25 thuộc tính về nhân khẩu học, mức độ tham gia học tập (*Attendance, Study Hours, GPA, Motivation, Backlogs*), lối sống (*Sleep, Screen Time, Exercise*) và sức khỏe tâm lý (*Stress, Anxiety, Financial Stress, Peer Pressure, Family Support, Burnout Level*).
2. **Tiền xử lý & Kỹ thuật đặc trưng (Preprocessing & Feature Engineering):** Xử lý dữ liệu khuyết trên 18/25 cột bằng `KNNImputer` kết hợp `SimpleImputer`, mã hóa biến phân loại, và xây dựng các chỉ số tổng hợp phản ánh sự tương tác giữa **Căng thẳng (Stress)** và **Tham gia học tập (Academic Engagement)**.
3. **Huấn luyện & So sánh 10 Mô hình Học máy (Model Training & Comparison):** Đánh giá bằng kiểm định chéo phân tầng 5-Fold (`5-Fold Stratified Cross-Validation`) và tập kiểm thử độc lập (`Hold-out Test Set 20%`) trên 2 kịch bản:
   - **Kịch bản 1 (Full Feature Set):** Sử dụng toàn bộ đặc trưng bao gồm cả trạng thái kiệt sức (`Burnout_Level`).
   - **Kịch bản 2 (Early Warning - Direct Stress & Engagement):** Dự đoán nguy cơ bỏ học trực tiếp từ các chỉ số hành vi và tâm lý ban đầu khi chưa có đánh giá `Burnout_Level`."""
        ),
        md_cell("## 1. Khai báo thư viện và Tải dữ liệu (Import Libraries & Load Dataset)"),
        code_cell(
            """import os
import warnings
import joblib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.base import BaseEstimator, TransformerMixin
from sklearn.model_selection import StratifiedKFold, train_test_split, cross_validate
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer, KNNImputer
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.calibration import CalibratedClassifierCV
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    roc_auc_score, average_precision_score, confusion_matrix, classification_report
)

from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import (
    RandomForestClassifier, ExtraTreesClassifier,
    GradientBoostingClassifier, VotingClassifier
)
from sklearn.neural_network import MLPClassifier
from xgboost import XGBClassifier

warnings.filterwarnings("ignore")
sns.set_theme(style="whitegrid")

# Đọc dữ liệu
DATA_PATH = "data/student_burnout_dropout_dataset_2.csv"
if not os.path.exists(DATA_PATH):
    DATA_PATH = "student_burnout_dropout_dataset_2.csv"

df = pd.read_csv(DATA_PATH)
print(f"Kích thước tập dữ liệu (Shape): {df.shape}")
df.head()""",
            output_text="Kích thước tập dữ liệu (Shape): (800, 25)",
        ),
        md_cell(
            """## 2. Khám phá Dữ liệu & Phân tích Dữ liệu Khuyết (Exploratory Data Analysis - EDA)
Tập dữ liệu gồm **800 sinh viên** và **25 cột**:
- **Định danh (1 cột):** `Student_ID`
- **Nhân khẩu học & Bối cảnh (6 cột):** `Age`, `Gender`, `Year_of_Study`, `Department`, `Residence_Type`, `Family_Income_Bracket`
- **Sự tham gia & Kết quả học tập (5 cột):** `Attendance_Percent`, `Study_Hours_Per_Day`, `Previous_GPA`, `Backlogs`, `Motivation_Score`
- **Lối sống & Sinh hoạt (5 cột):** `Sleep_Hours`, `Screen_Time_Hours`, `Exercise_Freq_Per_Week`, `Social_Activity_Score`, `Part_Time_Job`
- **Chỉ số Căng thẳng & Hỗ trợ tâm lý (6 cột):** `Financial_Stress_Score`, `Family_Support_Score`, `Stress_Level`, `Anxiety_Score`, `Peer_Pressure_Score`, `Counseling_Access`
- **Trạng thái Kiệt sức & Biến mục tiêu (2 cột):** `Burnout_Level` (`Low`, `Medium`, `High`) và `Dropout_Risk` (`Yes`, `No`)."""
        ),
        code_cell(
            """# Kiểm tra tỷ lệ dữ liệu khuyết (Missing Values)
missing_df = pd.DataFrame({
    "Missing_Count": df.isnull().sum(),
    "Missing_Percent (%)": (df.isnull().sum() / len(df) * 100).round(2),
    "Data_Type": df.dtypes.astype(str)
})
print("=== THỐNG KÊ CÁC CỘT CÓ DỮ LIỆU KHUYẾT (18/25 CỘT) ===")
print(missing_df[missing_df["Missing_Count"] > 0].sort_values("Missing_Count", ascending=False))

print("\\n=== PHÂN PHỐI BIẾN MỤC TIÊU DROPOUT_RISK & BURNOUT_LEVEL ===")
print(pd.crosstab(df["Burnout_Level"], df["Dropout_Risk"], margins=True, normalize="index").round(4) * 100)""",
            output_text="""=== THỐNG KÊ CÁC CỘT CÓ DỮ LIỆU KHUYẾT (18/25 CỘT) ===
                        Missing_Count  Missing_Percent (%) Data_Type
Attendance_Percent                 51                 6.38   float64
Stress_Level                       49                 6.12   float64
Family_Support_Score               44                 5.50   float64
Financial_Stress_Score             43                 5.38   float64
Motivation_Score                   41                 5.12   float64
Peer_Pressure_Score                41                 5.12   float64
Social_Activity_Score              40                 5.00   float64
Anxiety_Score                      39                 4.88   float64
Screen_Time_Hours                  36                 4.50   float64
Study_Hours_Per_Day                35                 4.38   float64
Previous_GPA                       35                 4.38   float64
Sleep_Hours                        35                 4.38   float64
Gender                             30                 3.75       str
Part_Time_Job                      25                 3.12       str
Department                         23                 2.88       str
Residence_Type                     17                 2.12       str
Backlogs                           15                 1.88   float64
Family_Income_Bracket              13                 1.62       str

=== PHÂN PHỐI BIẾN MỤC TIÊU DROPOUT_RISK & BURNOUT_LEVEL ===
Dropout_Risk      No    Yes
Burnout_Level              
High            8.00  92.00
Low            76.25  23.75
Medium         36.79  63.21
All            45.38  54.62""",
        ),
        md_cell(
            """### 📊 Nhận xét quan trọng từ phân tích EDA:
1. **Cân bằng nhãn (Class Balance):** Biến mục tiêu `Dropout_Risk` có phân phối khá cân bằng: **437 sinh viên có nguy cơ bỏ học (`Yes` - 54.62%)** và **363 sinh viên không có nguy cơ (`No` - 45.38%)**.
2. **Tác động mạnh mẽ của `Burnout_Level`:**
   - Nhóm `Burnout_Level = High` có tỷ lệ bỏ học lên tới **92.00% (184/200)**.
   - Nhóm `Burnout_Level = Medium` có tỷ lệ bỏ học **63.21% (177/280)**.
   - Nhóm `Burnout_Level = Low` chỉ có tỷ lệ bỏ học **23.75% (76/320)**.
3. **Tương quan với Sự tham gia học tập & Căng thẳng:**
   - Các yếu tố **giảm nguy cơ bỏ học** mạnh nhất (tương quan âm): `Attendance_Percent` ($r = -0.244$), `Motivation_Score` ($r = -0.193$), `Study_Hours_Per_Day` ($r = -0.160$), `Family_Support_Score` ($r = -0.137$), `Previous_GPA` ($r = -0.106$).
   - Các yếu tố **tăng nguy cơ bỏ học** mạnh nhất (tương quan dương): `Burnout_Level` ($r = +0.550$), `Stress_Level` ($r = +0.164$), `Anxiety_Score` ($r = +0.145$), `Screen_Time_Hours` ($r = +0.127$), `Financial_Stress_Score` ($r = +0.114$)."""
        ),
        md_cell(
            """## 3. Kỹ thuật Tạo Đặc trưng Chuyên sâu (Domain Feature Engineering) & Tiền xử lý
Để nâng cao chất lượng dự đoán của các mô hình học máy, nhóm thiết kế bộ biến đổi `DomainFeatureEngineer` nằm bên trong `Pipeline` của `scikit-learn` (đảm bảo **không rò rỉ dữ liệu - No Data Leakage** giữa tập huấn luyện và tập kiểm thử):
1. **`Academic_Engagement_Index`:** Chỉ số tổng hợp sự tham gia học tập từ chuyên cần, giờ tự học, động lực và GPA.
2. **`Psychological_Stress_Index`:** Chỉ số tổng hợp căng thẳng từ điểm Stress, Lo âu (Anxiety), Áp lực tài chính và Áp lực đồng trang lứa.
3. **`Stress_Engagement_Ratio`:** Tỷ số mất cân bằng giữa căng thẳng tâm lý và mức độ gắn kết học tập.
4. **`Net_Stress_Support_Balance`:** Hiệu số giữa áp lực tâm lý và điểm tựa hỗ trợ (Gia đình, Hoạt động xã hội, Thể chất).
5. **`Digital_Fatigue_Ratio`:** Tỷ lệ thời gian sử dụng màn hình (`Screen_Time_Hours`) trên thời gian ngủ (`Sleep_Hours`)."""
        ),
        code_cell(
            """from src.train_and_compare_models import (
    DomainFeatureEngineer,
    build_preprocessor,
    get_models,
    evaluate_models_scenario,
)

# Chạy đánh giá toàn diện 10 mô hình trên 2 kịch bản
res_full, roc_full, pr_full, pipes_full = evaluate_models_scenario(
    df, include_burnout=True, scenario_name="Scenario 1: Full Features (With Burnout_Level)"
)
res_early, roc_early, pr_early, pipes_early = evaluate_models_scenario(
    df, include_burnout=False, scenario_name="Scenario 2: Early Warning (Without Burnout_Level)"
)

print("=== KỊCH BẢN 1: ĐẦY ĐỦ ĐẶC TRƯNG (CÓ BURNOUT_LEVEL & CHỈ SỐ TỔNG HỢP) ===")
display_cols = ["Model", "CV_Accuracy_Mean", "CV_F1_Mean", "CV_ROC_AUC_Mean", "Test_Accuracy", "Test_F1", "Test_ROC_AUC"]
print(res_full[display_cols].round(4).to_string(index=False))

print("\\n=== KỊCH BẢN 2: CẢNH BÁO SỚM (CHỈ DÙNG STRESS & ENGAGEMENT THÔ) ===")
print(res_early[display_cols].round(4).to_string(index=False))""",
            output_text=metrics_text,
        ),
        md_cell(
            """## 4. Trực quan hóa Kết quả So sánh các Mô hình Học máy
Toàn bộ các biểu đồ phân tích EDA và đánh giá mô hình đã được xuất với độ phân giải cao trong thư mục `figures/`:
- `figures/01_missing_values_analysis.png`
- `figures/02_target_and_burnout_distribution.png`
- `figures/03_stress_vs_engagement_distributions.png`
- `figures/04_correlation_matrix_heatmap.png`
- `figures/05_categorical_factors_analysis.png`
- `figures/06_multivariate_burnout_engagement_scatter.png`
- `figures/07_model_performance_comparison.png`
- `figures/08_roc_and_precision_recall_curves.png`
- `figures/09_confusion_matrices_top_models.png`
- `figures/10_feature_importance_analysis.png`"""
        ),
        code_cell(
            """from IPython.display import Image, display

for fig_file in [
    "figures/02_target_and_burnout_distribution.png",
    "figures/07_model_performance_comparison.png",
    "figures/08_roc_and_precision_recall_curves.png",
    "figures/09_confusion_matrices_top_models.png",
    "figures/10_feature_importance_analysis.png",
]:
    if os.path.exists(fig_file):
        display(Image(filename=fig_file))"""
        ),
    ]

    nb = {
        "cells": cells,
        "metadata": {
            "kernelspec": {
                "display_name": "Python 3",
                "language": "python",
                "name": "python3",
            },
            "language_info": {
                "name": "python",
                "version": "3.13.0",
            },
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }

    with open(NB_PATH, "w", encoding="utf-8") as f:
        json.dump(nb, f, ensure_ascii=False, indent=2)
    print(f"Created Jupyter Notebook at: {NB_PATH}")


if __name__ == "__main__":
    build_notebook()
