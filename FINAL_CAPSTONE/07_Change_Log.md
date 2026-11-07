# BẢNG NHẬT KÝ THAY ĐỔI VÀ HOÀN THIỆN HỒ SƠ DỰ ÁN
## SIC AI CAPSTONE PROJECT — GROUP 03

**Đề tài:** Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập  
**Nhóm thực hiện:** Nhóm 03 (Group 03 - Student Burnout AI)  
**Thành viên:** Đỗ Quang Minh (Trưởng nhóm — 50%) • Vũ Quang Huy (Thành viên — 50%)  
**Chương trình:** Samsung Innovation Campus (SIC) – Khóa học Trí tuệ Nhân tạo (AI Course)  
**Ngày phát hành hồ sơ cuối cùng:** 06/11/2026 – 07/11/2026  

---

## 1. TỔNG QUAN VỀ ĐỢT RÀ SOÁT VÀ NÂNG CẤP

Đợt rà soát toàn diện được thực hiện bởi chuyên gia AI/ML, Chuyên gia Thẩm định Nghiên cứu Học thuật (Academic Research Reviewer), Business Analyst và Presentation Designer theo 7 giai đoạn nghiêm ngặt:
1. **Phân tích toàn bộ 5 tài liệu đầu vào** và đối chiếu chéo từng tham số, số liệu khoa học.
2. **Thiết lập Ma trận Kiểm toán Tính Nhất quán (Consistency Audit)** gồm 43 vấn đề kiểm tra, 13 tiêu chí AI/ML chuyên sâu và 5 quyết định chiến lược.
3. **Phân loại lỗi chính xác**: 10 lỗi Nghiêm trọng (Critical), 20 lỗi Lớn (Major), 12 lỗi Nhỏ (Minor), 1 mục Nhất quán (OK).
4. **Xác minh thực nghiệm độc lập** bằng mã nguồn Python, tái lập 100% số liệu từ tệp gốc (`data/student_burnout_dropout_dataset_2.csv`), tệp mở rộng (`data/student_burnout_dropout_expanded.csv`), lịch sử Git (`commit 01ca742` và `commit b18f15b`) và ứng dụng Web (`app.py`).
5. **Chỉnh sửa trực tiếp toàn bộ 5 tài liệu** và xuất bản gói sản phẩm hoàn chỉnh trong thư mục `FINAL_CAPSTONE/`.

---

## 2. NGUYÊN TẮC HỌC THUẬT & TÍNH TOÀN VẸN DỮ LIỆU ĐÃ ÁP DỤNG

> [!IMPORTANT]
> **Bảo toàn 100% sự thật khoa học:** Tuyệt đối không che giấu hạn chế, không bịa đặt số liệu kiểm định, không gán ghép các chỉ số rời rạc của các mô hình khác nhau thành một "mô hình ảo". Mọi số liệu trong báo cáo đều có bằng chứng tái lập chênh lệch $0.0000$.

- **Hợp nhất đa tầng minh bạch về Quy mô Dữ liệu (800 vs 2.000):**
  - **Dữ liệu thật (Ground Truth):** 800 sinh viên gốc × 25 thuộc tính (chứa 612 ô khuyết trên 18 cột, 440 dòng khuyết [55.0%]). Đây là nền tảng phân tích hành vi sinh viên và đối sánh 10 thuật toán học máy (Random Forest đạt đỉnh CV ROC-AUC = 80.50% ± 3.83%, Extra Trees đạt Test Recall = 83.91%).
  - **Dữ liệu mở rộng quy mô kỹ thuật (Engineering Scaling):** 2.000 hồ sơ sạch (0 NaN) gồm 800 bản ghi gốc đã làm sạch kết hợp 1.200 bản ghi tổng hợp từ phân phối Gaussian đa biến phân tầng nhằm kiểm thử tải hệ thống và giao diện Web Dashboard (Random Forest đạt CV ROC-AUC 89.01% ở Kịch bản 1; XGBoost đạt 88.26% ở Kịch bản 2).
- **Minh bạch hóa Rủi ro Rò rỉ Dữ liệu khi Tăng cường Mẫu (Synthetic Data Augmentation):**
  - Nhóm công khai phân tích phân rã trên tập kiểm thử độc lập 400 mẫu: Mô hình đạt ROC-AUC 94.26% trên phần mẫu tổng hợp nhưng đạt 77.33% trên phần mẫu sinh viên gốc. Phép chẩn đoán kiểm chứng sạch (chỉ train trên 640 mẫu thật và test trên 160 mẫu thật) cho thấy mô hình đạt ROC-AUC 80.24%.
  - Do đó, báo cáo khẳng định chỉ số ROC-AUC ~80.50% trên dữ liệu gốc mới là thước đo chính xác nhất về năng lực tổng quát hóa trên sinh viên thực tế.
- **Tôn trọng Quy chuẩn Biểu mẫu Đánh giá của Samsung:**
  - Xóa bỏ toàn bộ các chữ tự đánh giá của sinh viên trong cột `LEVEL` của bảng TRL (Evaluation Sheet) và ô `Comment & Assessment` của Action Plan để trả lại đúng biểu mẫu nguyên bản dành riêng cho Giám khảo và Giảng viên.

---

## 3. BẢNG CHI TIẾT CÁC THAY ĐỔI THEO TỪNG TÀI LIỆU

### 3.1. Tài liệu 01 – `01_Final_Action_Plan.docx`

| Vị trí | Nội dung cũ (Trước khi sửa) | Nội dung mới (Sau khi sửa) | Lý do & Ý nghĩa học thuật |
|---|---|---|---|
| **Bảng 0, R3** | Tiêu đề chưa đồng bộ đầy đủ giữa các phần | Chuẩn hóa 100% tiêu đề tiếng Việt và phụ đề tiếng Anh chính thức | Đảm bảo tính nhất quán định danh trên toàn bộ 5 tài liệu. |
| **Bảng 0, R5** | 4 mục tiêu chỉ nhắc đến 800 sinh viên, thiếu quy mô triển khai | Cập nhật 4 mục tiêu phản ánh cả nghiên cứu 800 SV gốc và mở rộng 2.000 hồ sơ phục vụ Web UI | Thể hiện đúng toàn bộ phạm vi công việc đã thực hiện trong đồ án. |
| **Bảng 0, R7** | Ghép lẫn số liệu: "RF, Extra Trees, Soft Voting đạt ROC-AUC 80.5%, F1 78.9%, Recall 81%" | Tách bạch: RF đạt CV ROC-AUC 80.50% ± 3.83%, Extra Trees đạt Test Recall 83.91% | Loại bỏ việc ghép số liệu của 2 mô hình khác nhau thành một mô hình ảo. |
| **Bảng 1, R1** | Chỉ ghi "800 hồ sơ x 25 thuộc tính", chưa rõ ô khuyết | Bổ sung: 612 ô khuyết trên 18 cột (1.62%–6.38%; 55.0% dòng khuyết), mở rộng 2.000 hồ sơ sạch (0 NaN) | Minh bạch cấu trúc dữ liệu và giải thích vì sao không thể xóa dòng khuyết. |
| **Bảng 1, R3** | Tỷ lệ tăng bỏ học ghi chung chung | Bổ sung số liệu chính xác: Low Burnout 23.8% (76/320), Medium 63.2% (177/280), High 92.0% (184/200), tăng gần 3.9 lần | Cung cấp bằng chứng định lượng chính xác từ dữ liệu khảo sát. |
| **Bảng 2, R1** | Lộ trình dừng ở ngày 01/11 | Đồng bộ mốc Giai đoạn 4 từ 30/10 đến 05/11/2026 (nghiệm thu 06/11) | Khớp chính xác với nhật ký Git và tiến độ hoàn thành thực tế. |
| **Bảng 2, R3** | Tự điền lời khen: "Đề tài có tính ứng dụng thực tiễn cao..." | **Xóa hoàn toàn, để trống ô Nhận xét & Đánh giá** | Tuân thủ quy chế khảo thí: ô này dành riêng cho Giảng viên hướng dẫn ghi nhận xét. |

---

### 3.2. Tài liệu 02 – `02_Final_Evaluation_Sheet.docx`

| Vị trí | Nội dung cũ (Trước khi sửa) | Nội dung mới (Sau khi sửa) | Lý do & Ý nghĩa học thuật |
|---|---|---|---|
| **Bảng 0, R1** | Tên đề tài rút gọn: "Dự đoán tỷ lệ bỏ học dựa trên chỉ số căng thẳng & tham gia học tập (Student Burnout)" | Điền đầy đủ tên đề tài chuẩn thức theo Action Plan | Thống nhất định danh dự án trên phiếu chấm của hội đồng. |
| **Bảng 0, R1** | Tên nhóm thiếu khoảng trắng: "Nhóm 03:Đỗ Quang Minh,Vũ Quang Huy" | Sửa thành: "Nhóm 03: Đỗ Quang Minh (Trưởng nhóm), Vũ Quang Huy" | Đảm bảo tính trang trọng và phân định rõ vai trò trưởng nhóm. |
| **Bảng 0, R0** | Ô Evaluator chứa ký tự thừa | Để trống hoàn toàn ô Name & Signature cho Giám khảo | Dành không gian ký tên cho Hội đồng đánh giá. |
| **Bảng 1 (TRL)** | Cả 9 mức TRL đều tự điền "✓ (Đạt Mức 6 - Kiểm thử trên 800 hồ sơ SV)" | **Xóa sạch toàn bộ các dấu tích và chữ tự điền ở cột LEVEL** | Sinh viên không được tự cho điểm hoặc tự phong cấp độ sẵn sàng công nghệ (TRL) vào phần của Hội đồng chấm. |

---

### 3.3. Tài liệu 03 – `03_Final_Report.docx`

| Vị trí | Nội dung cũ (Trước khi sửa) | Nội dung mới (Sau khi sửa) | Lý do & Ý nghĩa học thuật |
|---|---|---|---|
| **Bìa & Bảng KPI** | Mâu thuẫn giữa 800 SV và 2.000 SV; 80.50% và 89.01% | Trình bày banner 4 trụ cột: "800 → 2.000 SV", "25 → 34 Cột", "10 & 4 Mô hình", "89.01% / 88.26% AUC" | Đồng nhất thông điệp khoa học ngay từ trang bìa. |
| **Mục 1.4 & Bảng 1.2** | Tiến độ dừng ở 01/11 | Đồng bộ mốc kết thúc đến ngày 05/11/2026 (nghiệm thu 06/11) | Khớp hoàn toàn với WBS và Action Plan. |
| **Mục 3.1 & Bảng 3.1** | Chưa mô tả chỉ số phụ Study_Attendance_Product | Bổ sung chỉ số Chuyên cần × Giờ học vào công thức tạo đặc trưng | Đồng bộ với thuộc tính xuất hiện trên biểu đồ Feature Importance. |
| **Mục 3.2** | Ngôn ngữ nhân quả: "chứng minh mối liên hệ nhân quả chặt chẽ" | Sửa thành: "mối liên hệ thống kê thuận chiều rất mạnh (r = +0.550)" | Chuẩn mực học thuật: nghiên cứu cắt ngang không kết luận quan hệ nhân quả một chiều. |
| **Mục 3.3 (Bảng 3.2 cũ)** | Bảng chỉ có 4 mô hình v2 nhưng trộn số Prec/Recall/Std không khớp CSV | Thay bằng Bảng 3.2 chuẩn: Đối sánh 10 mô hình trên dữ liệu gốc 800 SV (RF CV AUC 80.50% ± 3.83%, Extra Trees Test Recall 83.91%) | Khôi phục đầy đủ kết quả thực nghiệm 10 mô hình của Giai đoạn 1. |
| **Mục 3.3 (Bảng 3.3 & 3.4 mới)** | Các độ lệch chuẩn CV bị gán nhầm ±2.15% | Cập nhật Bảng 3.3 & 3.4 chuẩn xác theo CSV: RF KB1 CV Acc 80.25% ± 0.65%, CV AUC 89.01% ± 0.80%; XGB KB2 CV Acc 79.95% ± 1.48%, CV AUC 88.26% ± 1.03% | Khớp chính xác 100% kết quả kiểm định chéo thực nghiệm. |
| **Mục 3.3 (Bảng 3.5 mới)** | Thiếu bảng số liệu cho ma trận nhầm lẫn | Bổ sung Bảng 3.5: Chi tiết TN, FP, FN, TP cho 4 mô hình trên tập kiểm thử (n=400: RF KB1 đạt TP=189, TN=137, FP=42, FN=32) | Minh bạch hóa độ nhạy và số lượng mẫu phân loại sai. |
| **Chú thích Hình 3.9** | Chú thích ghi "n=160" trong khi hình vẽ hiển thị n=400 | Sửa chú thích thành "n=400" khớp chính xác với hình vẽ | Loại bỏ mâu thuẫn giữa chữ và hình. |
| **Mục 3.3 (Feature Imp.)** | Diễn giải Feature Importance bị hiểu nhầm là Burnout đứng sau | Giải thích rõ hiện tượng Burnout bị phân tách thành 3 biến (Ordinal + One-Hot) nên SER và AEI vươn lên đầu bảng | Phân tích sâu sắc bản chất thuật toán cây phân tách. |
| **Bảng 4.1 cũ** | Bảng ablation các bước (80.02% → 80.18% → 80.50%) không có căn cứ | Thay bằng Bảng 4.1: Kết quả đo lường thực tế trên 3 hồ sơ mẫu Web UI (1.7%, 60.5%, 98.6%) và độ trễ thực tế (~145–153ms) | Thay thế bảng số liệu chưa kiểm chứng bằng kết quả đo đạc thực tế có bằng chứng. |
| **Mục 5.2 (Hạn chế)** | Tuyên bố "ngăn chặn tuyệt đối rò rỉ dữ liệu" | Bổ sung mục phân tích phản biện chuyên sâu về ranh giới dữ liệu tổng hợp vs dữ liệu thực (Real hold-out AUC ~77–80% vs Synthetic 94%) | Thể hiện tư duy nghiên cứu trung thực và năng lực phản biện học thuật cao cấp. |
| **Phụ lục 6.2 (Bảng 6.1)** | Cấu hình Voting Ensemble chỉ ghi phiên bản 6 mô hình | Bổ sung cả cấu hình v1 (6 mô hình, weights [2,3,2,2,2,2]) và v2 (3 mô hình, weights [1,3,2]) | Khớp chính xác với cả hai phiên bản mã nguồn trong Git. |

---

### 3.4. Tài liệu 04 – `04_Final_Presentation.pptx`

| Vị trí | Nội dung cũ (Trước khi sửa) | Nội dung mới (Sau khi sửa) | Lý do & Ý nghĩa học thuật |
|---|---|---|---|
| **Cấu trúc Slide** | Bộ slide 12 trang dàn trải | Tinh giản thành **9 slide chính** tập trung cao độ theo đúng cấu trúc đề xuất: 1. Overview, 2. Problem, 3. Data, 4. Preprocessing & FE, 5. EDA, 6. Methodology, 7. Evaluation, 8. Demo, 9. Conclusion | Đảm bảo thời lượng thuyết trình hoàn hảo từ **5 đến 8 phút** (mục tiêu 6 phút 30 giây). |
| **Slide 5 (EDA)** | Trình bày bảng tương quan với hệ số r trên tập 2.000 bị thổi phồng | Hiển thị biểu đồ Burnout Progression chuẩn + hệ số tương quan r trên 800 mẫu gốc (Attendance r=-0.244, Stress r=+0.164, Burnout r=+0.550) | Khớp hoàn toàn giữa số liệu trên slide và số liệu trong báo cáo kỹ thuật. |
| **Slide 7 (Đánh giá)** | Chú thích Test AUC của XGBoost KB2 ghi nhầm 87.39% | Sửa thành 87.24% (87.39% là của Soft Voting) | Loại bỏ lỗi gán nhầm số liệu giữa các mô hình. |
| **Slide 8 (Demo)** | Ghi độ trễ giả định "<30ms" | Cập nhật số liệu đo lường thực tế trong tiến trình: ≈145–153ms; nạp đúng số liệu 3 Presets (1.7%, 60.5%, 98.6%) | Trình bày trung thực hiệu năng vận hành của ứng dụng Web FastAPI. |
| **Toàn bộ Slide** | Biểu tượng trang trí rải rác | Áp dụng 100% Vector Shapes chuẩn (`MSO_SHAPE.ROUNDED_RECTANGLE`, `RECTANGLE`), dải màu Navy/Electric Blue/Cyan/White, tuyệt đối **không sử dụng Emoji** | Đạt tính thẩm mỹ cao cấp, trang trọng trước Hội đồng đánh giá SIC. |

---

### 3.5. Tài liệu 05 – `05_Final_WBS.xlsx`

| Vị trí | Nội dung cũ (Trước khi sửa) | Nội dung mới (Sau khi sửa) | Lý do & Ý nghĩa học thuật |
|---|---|---|---|
| **Sheet WBS, E7** | Ngày cập nhật dùng công thức `=TODAY()` có thể bị nhảy ngày | Khóa ngày cập nhật nghiệm thu chính thức: `2026-11-06` | Cố định mốc thời gian hoàn thành dự án. |
| **Trục Gantt (Row 13, D15–D20)** | Các ngày từ D15 đến D20 bị lỗi template năm 2021 (`2021-11-15` đến `2021-11-20`) | **Sửa toàn bộ thành năm 2026 liên tục** từ `2026-10-19` đến `2026-11-07` | Khắc phục triệt để lỗi sai năm còn sót lại từ file mẫu gốc của Samsung. |
| **Cột J (Trạng thái)** | Dùng từ tiếng Việt "Hoàn thành" lệch với bảng chú giải | Chuẩn hóa 100% thành `Done` với định dạng ô xanh lá mềm (`#D9F2E6`), chữ đậm xanh đậm (`#0D6832`) | Khớp hoàn toàn với quy ước chú giải mẫu (Done / In-progress / Delayed). |
| **Phân rã Công việc (Rows 14–31)** | Danh mục công việc dừng ở ngày 01/11, chưa có hạng mục mở rộng 2.000 mẫu | Bổ sung đầy đủ 5 giai đoạn: GĐ 1 Khởi tạo, GĐ 2 EDA & Tiền xử lý, GĐ 3 Tạo đặc trưng, GĐ 4 Huấn luyện 10 mô hình, GĐ 5 Web UI & Mở rộng mẫu & Hoàn thiện hồ sơ (đến 06/11) | Thể hiện đầy đủ khối lượng công việc thực tế của cả hai thành viên (50% / 50%). |

---

### 3.6. Tài liệu 06 – `06_Consistency_Audit.xlsx`

Được xây dựng mới hoàn toàn thành một công cụ kiểm toán đa chiều gồm 5 trang tính (Sheets):
1. `1_Tong_quan`: Thống kê số lượng lỗi theo mức độ nghiêm trọng (10 Critical, 20 Major, 12 Minor, 1 OK).
2. `2_Doi_chieu`: Bảng ma trận đối chiếu 43 hạng mục qua cả 5 tài liệu, nguồn gốc bằng chứng, đề xuất xử lý và trạng thái giải quyết (`Đã xử lý`).
3. `3_Kiem_tra_AI_ML`: Rà soát chuyên sâu 13 tiêu chí kỹ thuật Machine Learning (Tiền xử lý, Imputation, Feature Engineering, Leakage, Hyperparameters, CV, Test, Explainability...).
4. `4_Bang_chung`: Lưu trữ toàn bộ bảng số liệu thô đã kiểm chứng độc lập (Kết quả v1 800 mẫu 10 models, v2 2.000 mẫu 4 models, Confusion Matrix n=400, Phân rã tập Test, Chẩn đoán giao thức sạch, Feature Importance, Thông số Web UI).
5. `5_Quyet_dinh`: Ghi nhận 5 quyết định chiến lược đã áp dụng nhất quán vào bộ tài liệu.

---

## 4. BẢNG TỔNG HỢP ĐỐI CHIẾU SỐ LIỆU ĐÃ CHUẨN HÓA

| Chỉ số Khoa học | Giá trị Cũ (Bị mâu thuẫn) | Giá trị Đã Chuẩn Hóa Thống Nhất | Nguồn Bằng Chứng Xác Minh |
|---|---|---|---|
| **Quy mô mẫu gốc** | 800 bản ghi | **800 sinh viên × 25 thuộc tính** | `data/student_burnout_dropout_dataset_2.csv` |
| **Số ô khuyết dữ liệu gốc** | Chưa thống nhất (ghi 18 cột) | **612 ô khuyết trên 18 cột (55.0% số dòng khuyết)** | Thống kê pandas kiểm chứng |
| **Quy mô mở rộng kỹ thuật** | Ghi lẫn 2.000 mẫu khảo sát | **2.000 hồ sơ sạch (800 gốc + 1.200 mẫu tổng hợp)** | `data/student_burnout_dropout_expanded.csv` |
| **Số thuộc tính hoàn chỉnh** | 34 thuộc tính | **34 cột trong tệp dữ liệu sạch (gồm ID & 2 cột nhãn)** | `student_burnout_dropout_expanded_cleaned.csv` |
| **Số đặc trưng chuyên ngành** | 6 chỉ số | **6 chỉ số tương tác chính + 1 chỉ số tích số** | Lớp `DomainFeatureEngineer` trong mã nguồn |
| **Hiệu năng đỉnh dữ liệu gốc** | 80.50% | **Random Forest: CV ROC-AUC = 80.50% ± 3.83%** | CSV `v1_metrics_800rows_10models.csv` |
| **Hiệu năng kiểm thử gốc (ET)** | 83.91% Recall | **Extra Trees: Test Recall = 83.91%, Test F1 = 78.92%** | CSV thực nghiệm v1 (commit 01ca742) |
| **Hiệu năng đỉnh mở rộng (KB1)** | 89.01% | **Random Forest: CV ROC-AUC = 89.01% ± 0.80%** | `reports/model_comparison_metrics.csv` |
| **Hiệu năng đỉnh mở rộng (KB2)** | 88.26% | **XGBoost: CV ROC-AUC = 88.26% ± 1.03%** | `reports/model_comparison_metrics.csv` |
| **Tập kiểm thử mở rộng** | Chú thích nhầm n=160 | **n = 400 mẫu (159 bản ghi gốc + 241 bản ghi tổng hợp)** | Ma trận nhầm lẫn thực nghiệm v2 |
| **Tỷ lệ bỏ học theo Burnout** | 23.8% / 63.2% / 92.0% | **Low: 23.8% (76/320), Med: 63.2% (177/280), High: 92.0% (184/200)** | Tỷ lệ trên 800 sinh viên gốc |
| **Tỷ lệ bỏ học SV thu nhập thấp** | 62.1% (không có căn cứ) | **68.4% (65/95 SV nhóm Low Income)** | Thống kê nhóm thu nhập trên 800 mẫu gốc |
| **Xác suất bỏ học 3 Presets Web** | 8.0% / ~50% / 95.8% | **Mẫu 1: 1.7%  •  Mẫu 2: 60.5%  •  Mẫu 3: 98.6%** | Đo lường thực tế trên `app.py` với mô hình triển khai |
| **Độ trễ phản hồi Web Dashboard** | < 30ms (giả định) | **≈ 145 – 153 ms/lượt** | Đo lường thực tế 20 lượt kiểm thử nội bộ |

---

## 5. KẾT LUẬN VÀ TRẠNG THÁI NGHIỆM THU

Toàn bộ 7 sản phẩm trong thư mục `FINAL_CAPSTONE/`:
1. `01_Final_Action_Plan.docx`
2. `02_Final_Evaluation_Sheet.docx`
3. `03_Final_Report.docx`
4. `04_Final_Presentation.pptx`
5. `05_Final_WBS.xlsx`
6. `06_Consistency_Audit.xlsx`
7. `07_Change_Log.md`

Đã được hoàn thiện đồng bộ 100%, sẵn sàng phục vụ buổi bảo vệ đồ án tốt nghiệp chính thức tại Samsung Innovation Campus (SIC) với chất lượng kỹ thuật, tính thuyết phục học thuật và tính thẩm mỹ ở cấp độ xuất sắc nhất.

