# 🎓 Dự đoán Tỷ lệ Bỏ học dựa trên các Chỉ số Căng thẳng và Sự Tham gia Học tập
**(Student Burnout & Dropout Risk Prediction - Nhóm 03: Đỗ Quang Minh, Vũ Quang Huy)**

## 1. Cấu trúc Thư mục Dự án (Project Structure)
```text
Sudent_burnout/
├── data/
│   └── student_burnout_dropout_dataset_2.csv    # Bộ dữ liệu gốc (800 sinh viên x 25 thuộc tính)
├── figures/                                     # 10 biểu đồ phân tích EDA & Đánh giá mô hình (200 DPI)
│   ├── 01_missing_values_analysis.png
│   ├── 02_target_and_burnout_distribution.png
│   ├── 03_stress_vs_engagement_distributions.png
│   ├── 04_correlation_matrix_heatmap.png
│   ├── 05_categorical_factors_analysis.png
│   ├── 06_multivariate_burnout_engagement_scatter.png
│   ├── 07_model_performance_comparison.png
│   ├── 08_roc_and_precision_recall_curves.png
│   ├── 09_confusion_matrices_top_models.png
│   └── 10_feature_importance_analysis.png
├── models/                                      # Các Pipeline mô hình đã huấn luyện (.pkl)
│   ├── best_dropout_model_full.pkl
│   └── best_dropout_model_early_warning.pkl
├── reports/
│   └── model_comparison_metrics.csv             # Bảng kết quả thực nghiệm 10 mô hình x 2 kịch bản
├── src/
│   ├── eda_analysis.py                          # Script phân tích khám phá dữ liệu & xuất biểu đồ EDA
│   ├── train_and_compare_models.py              # Script Feature Engineering, huấn luyện & so sánh 10 mô hình
│   └── build_notebook.py                        # Script tự động đóng gói Jupyter Notebook
└── Student_Burnout_Dropout_Prediction_Nhom03.ipynb  # Jupyter Notebook trình bày toàn bộ quy trình
```

## 2. Hướng dẫn Chạy Chương trình (How to Run)
```bash
# 1. Chạy phân tích EDA và xuất 6 biểu đồ khám phá dữ liệu
python src/eda_analysis.py

# 2. Huấn luyện, kiểm định chéo 5-Fold và so sánh 10 mô hình học máy trên 2 kịch bản
python src/train_and_compare_models.py
```
