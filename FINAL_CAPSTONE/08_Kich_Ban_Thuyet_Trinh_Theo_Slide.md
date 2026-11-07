# KỊCH BẢN THUYẾT TRÌNH BẢO VỆ ĐỒ ÁN CHI TIẾT THEO TỪNG SLIDE
## SAMSUNG INNOVATION CAMPUS (SIC) — AI CAPSTONE PROJECT (NHÓM 03)

**Đề tài:** Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập  
**Thành viên:** Đỗ Quang Minh (Trưởng nhóm) • Vũ Quang Huy (Thành viên)  
**Thời lượng mục tiêu:** 6 phút 30 giây (khoảng an toàn 5–8 phút)  

---

## SLIDE 1: PROJECT OVERVIEW (TRANG BÌA) (Thời lượng: 25 giây)
- **Mục tiêu slide:** Mở đầu trang trọng, giới thiệu đề tài, nhóm thực hiện và neo ngay 4 chỉ số thực nghiệm nổi bật.

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Đứng thẳng, mắt nhìn bao quát Hội đồng, mỉm cười nhẹ và mở lời tự tin.]

"Kính thưa Quý Thầy Cô trong Hội đồng Đánh giá Đồ án Tốt nghiệp Samsung Innovation Campus!

Em là **Đỗ Quang Minh**, Trưởng nhóm 03, cùng bạn đồng hành là **Vũ Quang Huy**. Hôm nay, nhóm chúng em xin được báo cáo đồ án tốt nghiệp Capstone với đề tài: **'Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học dựa trên các chỉ số về căng thẳng và sự tham gia học tập'**.

[Hành động: Tay hướng về 4 thẻ KPI bên phải slide]
Đồ án của chúng em khảo sát từ **800 bản ghi sinh viên gốc**, làm sạch và mở rộng kỹ thuật lên **2.000 hồ sơ**, thiết kế **6 đặc trưng chuyên ngành mới**, và huấn luyện mô hình đạt hiệu năng đỉnh **ROC-AUC 89.01%** ở kịch bản đầy đủ và **88.26%** ở kịch bản cảnh báo sớm. Sau đây, em xin phép bắt đầu phần trình bày chi tiết!"

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Đề tài này có gì mới so với các bài toán dự đoán điểm số hay phân loại học lực thông thường?'
> -> Gợi ý trả lời: 'Dạ thưa Thầy Cô, điểm mới cốt lõi của đề tài là chúng em không chỉ nhìn vào điểm số GPA, mà lượng hóa tương tác phi tuyến giữa áp lực tâm lý (Stress, Lo âu, Tài chính) và mức độ gắn kết học tập (Chuyên cần, Giờ tự học) thông qua chỉ số Stress_Engagement_Ratio để cảnh báo sớm nguy cơ bỏ học ngay cả khi sinh viên chưa bị cảnh cáo học vụ.'

---

## SLIDE 2: PROBLEM STATEMENT (BÀI TOÁN CHÚNG TÔI GIẢI QUYẾT) (Thời lượng: 45 giây)
- **Mục tiêu slide:** Làm rõ cơ chế hình thành rủi ro: Bỏ học là kết quả tích lũy, không diễn ra đột ngột.

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 2. Chỉ tay vào con số 54.62% màu đỏ bên trái]

"Thưa Hội đồng, trong giáo dục đại học, tỷ lệ sinh viên có nguy cơ bỏ học trong tập dữ liệu khảo sát lên tới **54.62%**, tương đương 437 trên 800 sinh viên. Thực tế chỉ ra rằng, sinh viên **rất hiếm khi bỏ học đột ngột** chỉ vì một kỳ thi trượt. Thay vào đó, bỏ học là **hệ quả tích lũy kéo dài** qua 3 giai đoạn:

[Hành động: Chỉ tay lần lượt vào 3 khối tiến trình bên phải]
• **Giai đoạn 1 – Nguyên nhân gốc:** Bắt đầu từ sự suy giảm gắn kết học tập như chuyên cần sụt giảm dưới 70%, giảm giờ tự học, cộng hưởng cùng áp lực căng thẳng học vụ, lo âu và gánh nặng tài chính.
• **Giai đoạn 2 – Cầu nối trung gian:** Khi áp lực vượt quá ngưỡng chịu đựng, sinh viên rơi vào trạng thái **Kiệt sức Học đường (Academic Burnout)**. Biến này có tương quan thuận cực mạnh với nguy cơ bỏ học, hệ số Pearson đạt **r = +0.550**.
• **Giai đoạn 3 – Giải pháp của chúng em:** Thay vì đợi sinh viên nộp đơn xin thôi học mới xử lý bị động, nhóm xây dựng **Hệ thống Cảnh báo Sớm Học máy** để phát hiện rủi ro ngay từ đầu kỳ, giúp nhà trường kịp thời tư vấn tâm lý và phụ đạo học thuật."

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Tại sao lại chọn bài toán phân loại nhị phân Dropout_Risk mà không dự đoán trực tiếp điểm GPA?'
> -> Gợi ý trả lời: 'Dạ, mục tiêu ứng dụng thực tế của đồ án là giúp Phòng Công tác Sinh viên ra quyết định can thiệp (Actionable Decision). Việc phân loại rủi ro Dropout_Risk kèm mức độ xác suất giúp nhà trường phân bổ chính xác nguồn lực tư vấn và học bổng cho đúng nhóm đối tượng cấp thiết nhất.'

---

## SLIDE 3: DATASET & DATA UNDERSTANDING (DỮ LIỆU ĐẦU VÀO) (Thời lượng: 45 giây)
- **Mục tiêu slide:** Minh bạch cấu trúc 25 biến gốc, giải thích 612 ô khuyết và làm rõ nguồn gốc 2.000 hồ sơ.

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 3. Bạn Vũ Quang Huy bắt đầu trình bày]

"Tiếp theo, em xin phép trình bày về cấu trúc dữ liệu đầu vào.

[Hành động: Chỉ vào dải số liệu trên cùng: 800 -> 612 -> k=7 -> 2.000 -> 34]
Bộ dữ liệu nghiên cứu gốc gồm **800 bản ghi sinh viên × 25 thuộc tính**, bao quát 5 nhóm thông tin: Nhân khẩu học, Sự tham gia học tập, Lối sống sinh hoạt, Áp lực căng thẳng và Biến mục tiêu.

Qua kiểm toán chất lượng dữ liệu, nhóm phát hiện **612 ô khuyết nằm rải rác trên 18 cột**, chiếm tới 55% số dòng dữ liệu. Nếu áp dụng phương pháp xóa dòng truyền thống, chúng em sẽ mất hơn một nửa tập dữ liệu. Vì vậy, nhóm quyết định giữ lại toàn bộ mẫu và xử lý bằng thuật toán **KNNImputer với k=7** láng giềng gần nhất.

Đặc biệt, để kiểm thử khả năng mở rộng quy mô kỹ thuật và vận hành Web Dashboard tải lớn, từ 800 bản ghi gốc đã làm sạch, nhóm đã mở rộng thêm 1.200 bản ghi từ phân phối Gaussian đa biến phân tầng, tạo nên tập dữ liệu **2.000 hồ sơ sạch 100% không còn giá trị NaN**."

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: '2.000 hồ sơ này lấy từ đâu? Có phải khảo sát thực tế không?'
> -> Gợi ý trả lời: 'Dạ thưa Thầy Cô, nhóm xin được minh bạch hoàn toàn: Bộ dữ liệu khảo sát gốc là 800 sinh viên. 1.200 hồ sơ bổ sung là dữ liệu tổng hợp phân tầng (Multivariate Gaussian) được nhóm tạo ra nhằm kiểm thử khả năng mở rộng quy mô kỹ thuật và kiểm thử giao diện Web. Trong phần thực nghiệm, nhóm đã đánh giá riêng biệt cả trên 800 mẫu gốc và tập mở rộng để đảm bảo tính khách quan khoa học.'

---

## SLIDE 4: PREPROCESSING & FEATURE ENGINEERING (TIỀN XỬ LÝ & ĐẶC TRƯNG) (Thời lượng: 50 giây)
- **Mục tiêu slide:** Trình bày quy trình 5 bước và nêu bật 2 công thức đặc trưng quan trọng nhất (SER & AEI).

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 4. Bạn Đỗ Quang Minh tiếp tục trình bày]

"Về phương pháp xử lý, nhóm xây dựng quy trình khép kín gồm **5 bước chuẩn mực**:
Làm sạch lọc ngoại lai IQR -> Điền khuyết KNNImputer trong Pipeline -> Thiết kế đặc trưng -> Chuẩn hóa OneHot và StandardScaler -> Kiểm định chéo 5-Fold.

[Hành động: Hướng tay vào lưới 6 ô Bento bên dưới, đặc biệt nhấn mạnh vào F1 và F3]
Đóng góp kỹ thuật quan trọng nhất của nhóm là thiết kế **6 chỉ số tương tác chuyên ngành**:
• Tiêu biểu nhất là **F1 – Academic Engagement Index**: Lượng hóa mức độ gắn kết học tập từ chuyên cần, giờ tự học, động lực và GPA.
• Và **F3 – Stress Engagement Ratio**: Tỷ số phản ánh áp lực tâm lý chia cho mức độ gắn kết học tập. Khi căng thẳng vượt ngưỡng mà gắn kết học tập suy giảm, tỷ số này tăng vọt và trở thành **chỉ báo cảnh báo sớm nhạy bén nhất**.
• Bên cạnh đó, chúng em thiết kế chỉ số cân bằng áp lực và điểm tựa gia đình (F4), tỷ số mệt mỏi số (F5) và gánh nặng học vụ tích lũy từ nợ môn (F6)."

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Tại sao các hệ số trọng số 0.35, 0.25 trong công thức lại chọn như vậy? Có căn cứ toán học không?'
> -> Gợi ý trả lời: 'Dạ thưa Thầy Cô, các trọng số ban đầu được thiết lập dựa trên lý thuyết sư phạm của Tinto và Maslach kết hợp phân tích ma trận tương quan nội bộ (tương quan của từng biến với nhãn bỏ học). Sau đó, khi đưa vào mô hình Random Forest và XGBoost, các cây quyết định đã tự động tối ưu hóa các ngưỡng phân tách phi tuyến, khẳng định 2 chỉ số này đóng góp phân tách cao nhất trong toàn bộ mô hình.'

---

## SLIDE 5: EXPLORATORY DATA ANALYSIS (PHÂN TÍCH KHÁM PHÁ DỮ LIỆU) (Thời lượng: 50 giây)
- **Mục tiêu slide:** Nhấn mạnh phát hiện trọng tâm: tỷ lệ bỏ học tăng vọt theo Burnout (23.8% -> 63.2% -> 92.0%) và ma trận tương quan.

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 5. Bạn Vũ Quang Huy trình bày]

"Sau đây là những phát hiện thống kê quan trọng nhất từ quá trình khám phá dữ liệu (EDA) trên 800 sinh viên gốc:

[Hành động: Chỉ tay vào biểu đồ cột và 3 khối số lớn bên phải: 23.8% -> 63.2% -> 92.0%]
Phân tích chỉ ra mối liên hệ trực tiếp giữa mức độ kiệt sức và tỷ lệ bỏ học:
• Ở nhóm Kiệt sức Thấp, tỷ lệ bỏ học chỉ là **23.8%**.
• Khi chuyển sang Kiệt sức Vừa, tỷ lệ này tăng vọt lên **63.2%**.
• Và ở nhóm Kiệt sức Cao, tỷ lệ bỏ học lên tới **92.0%**, tức là 184 trên 200 sinh viên! Tỷ lệ này **tăng gấp gần 3.9 lần** so với nhóm kiệt sức thấp.

[Hành động: Chỉ vào khung phân tích tương quan phía dưới]
Đồng thời, ma trận tương quan Pearson khẳng định rõ hai cực đối lập:
• **Cực bảo vệ (giảm nguy cơ):** Chuyên cần đi học (r = -0.244, nhóm an toàn đi học 80.7% so với 73.7% của nhóm bỏ học) và Động lực nội tại (r = -0.193).
• **Cực nguy cơ (tăng nguy cơ):** Căng thẳng học vụ, lo âu kéo dài và áp lực tài chính. Đáng chú ý, sinh viên thuộc nhóm thu nhập thấp có tỷ lệ bỏ học lên tới **68.4%**."

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Hệ số tương quan r = +0.550 giữa Burnout và Bỏ học có đồng nghĩa với quan hệ nhân quả không?'
> -> Gợi ý trả lời: 'Dạ thưa Thầy Cô, nhóm hoàn toàn ý thức được đây là nghiên cứu quan sát cắt ngang (Cross-sectional Study). Hệ số r = +0.550 phản ánh mối liên hệ thống kê thuận chiều rất mạnh trong tập dữ liệu, chứ không khẳng định quan hệ nhân quả một chiều. Tuy nhiên, về mặt thực tiễn, đây là căn cứ sàng lọc định lượng vô cùng giá trị cho nhà trường.'

---

## SLIDE 6: AI/ML METHODOLOGY (PHƯƠNG PHÁP LUẬN HỌC MÁY) (Thời lượng: 45 giây)
- **Mục tiêu slide:** Làm rõ 2 kịch bản thực nghiệm đối chứng và cam kết chống rò rỉ dữ liệu (No Data Leakage).

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 6. Bạn Đỗ Quang Minh trình bày]

"Để giải quyết bài toán một cách khoa học, nhóm thiết kế **hai kịch bản thực nghiệm đối chứng**:

[Hành động: Chỉ vào 2 cột Kịch bản 1 và Kịch bản 2]
• **Kịch bản 1 – Chẩn đoán Toàn diện:** Sử dụng đầy đủ 32 biến đầu vào bao gồm cả biến kiệt sức Burnout_Level, giúp đánh giá độ chính xác chẩn đoán tối đa khi nhà trường có đợt khảo sát tâm lý chuyên sâu.
• **Kịch bản 2 – Cảnh báo Sớm Học đường:** Đây là trọng tâm ứng dụng thực tiễn. Chúng em **loại bỏ hoàn toàn biến Burnout_Level**, buộc mô hình phải dự báo nguy cơ bỏ học trực tiếp chỉ từ dữ liệu điểm danh, giờ học, GPA và áp lực thô. Kịch bản này cho phép phát hiện sớm ngay từ tuần đầu nhập học.

[Hành động: Nhấn mạnh vào khung cam kết khoa học bên dưới]
Đặc biệt, nhóm áp dụng nguyên tắc **chống rò rỉ dữ liệu tuyệt đối (No Data Leakage)**: Mọi khâu biến đổi dữ liệu từ điền khuyết KNN đến chuẩn hóa đều được đóng gói trong Pipeline của scikit-learn và chỉ khớp (fit) trên fold huấn luyện của từng vòng lặp kiểm định chéo phân tầng 5-Fold."

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Nếu Kịch bản 1 dùng Burnout_Level mà biến này lại tương quan mạnh với Bỏ học thì mô hình có bị thiên lệch (bias) không?'
> -> Gợi ý trả lời: 'Dạ chính xác thưa Thầy Cô! Đó chính là lý do vì sao nhóm bắt buộc phải thiết kế Kịch bản 2 độc lập. Kịch bản 2 chứng minh rằng ngay cả khi không có thông tin Burnout_Level, các mô hình học máy (tiêu biểu là XGBoost) vẫn đạt ROC-AUC tới 88.26%, khẳng định tính khả thi của hệ thống cảnh báo sớm trong thực tế.'

---

## SLIDE 7: MODEL PERFORMANCE & EVALUATION (KẾT QUẢ THỰC NGHIỆM) (Thời lượng: 50 giây)
- **Mục tiêu slide:** Trình bày bảng đối sánh 4 mô hình, số liệu ma trận nhầm lẫn chuẩn xác và độ quan trọng đặc trưng.

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 7. Bạn Đỗ Quang Minh trình bày]

"Kính thưa Hội đồng, đây là kết quả thực nghiệm đối sánh 4 mô hình học máy cốt lõi trên tập dữ liệu mở rộng 2.000 hồ sơ:

[Hành động: Chỉ vào Bảng đối sánh và biểu đồ Benchmark]
• Ở **Kịch bản 1 (Đầy đủ đặc trưng)**: **Random Forest** đạt hiệu năng dẫn đầu với **CV ROC-AUC = 89.01% ± 0.80%**, trên tập kiểm thử độc lập đạt **Test ROC-AUC = 88.42%** và **Test Accuracy = 81.50%**.
• Ở **Kịch bản 2 (Cảnh báo sớm)**: **XGBoost** thể hiện ưu thế vượt trội với **CV ROC-AUC = 88.26% ± 1.03%** và **Test ROC-AUC = 87.24%**.

[Hành động: Chỉ vào Ma trận nhầm lẫn n=400 và đường cong ROC]
Xét về Ma trận nhầm lẫn trên tập kiểm thử độc lập 400 mẫu, Random Forest phát hiện chính xác **189 trên 221 sinh viên có nguy cơ bỏ học**, đạt **Độ nhạy (Recall) lên tới 85.52%**, chỉ bỏ sót 32 trường hợp (FN).
Phân tích độ quan trọng đặc trưng khẳng định: Hai chỉ số tương tác do nhóm tự thiết kế là **Stress_Engagement_Ratio** và **Academic_Engagement_Index** xếp vị trí số 1 và số 2 trong các biến quyết định của Random Forest!"

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Tại sao trong bài toán này Độ nhạy (Recall) lại quan trọng hơn Độ chính xác (Precision)?'
> -> Gợi ý trả lời: 'Dạ thưa Thầy Cô, trong bài toán cảnh báo sớm bỏ học, sai lầm loại 2 (False Negative - bỏ sót sinh viên có nguy cơ) gây hậu quả nghiêm trọng hơn nhiều so với sai lầm loại 1 (False Positive - nhắc nhở nhầm sinh viên an toàn). Nếu bỏ sót, sinh viên có thể thôi học mà không được hỗ trợ; trong khi nếu nhắc nhở nhầm, sinh viên chỉ nhận thêm một buổi tham vấn tâm lý hoặc tư vấn học tập có lợi.'

---

## SLIDE 8: DEMO / SYSTEM ARCHITECTURE (KIẾN TRÚC & WEB DASHBOARD) (Thời lượng: 45 giây)
- **Mục tiêu slide:** Minh họa kiến trúc 3 tầng và demo trực tiếp kết quả đo lường thực tế trên Web FastAPI app.py.

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 8. Bạn Vũ Quang Huy trình bày]

"Để chuyển hóa mô hình nghiên cứu thành sản phẩm ứng dụng thực tế, nhóm đã phát triển hoàn chỉnh **Ứng dụng Web Dashboard Cảnh báo Sớm** bằng nền tảng **FastAPI** (tệp app.py) theo kiến trúc 3 tầng:

[Hành động: Chỉ vào sơ đồ 3 tầng bên trái và ảnh chụp giao diện bên phải]
• **Tầng 1 – Giao diện:** Cho phép nhập 18 chỉ số qua thanh trượt trực quan hoặc bấm chọn nhanh 3 hồ sơ mẫu.
• **Tầng 2 – Động cơ AI:** Vận hành song song mô hình dự báo Burnout 3 mức và mô hình dự báo nguy cơ bỏ học. Thời gian phản hồi đo lường nội bộ cực nhanh, chỉ xấp xỉ **145 đến 153 mili-giây** mỗi lượt dự báo.
• **Tầng 3 – Kết quả & Khuyến nghị:** Hiển thị trực quan xác suất rủi ro và tự động kích hoạt khuyến nghị sư phạm.

[Hành động: Chỉ vào 3 kết quả Preset bên dưới]
Kiểm thử thực tế trên 3 hồ sơ mẫu cho kết quả rất chính xác:
• Mẫu SV Tích cực: Nguy cơ bỏ học chỉ **1.7%** (An toàn).
• Mẫu SV Căng thẳng Trung bình: Nguy cơ bỏ học **60.5%** (Mức cảnh báo).
• Mẫu SV Kiệt sức Cao: Nguy cơ bỏ học lên tới **98.6%** (Báo động đỏ can thiệp khẩn cấp)!"

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Hệ thống này triển khai thực tế ở trường đại học như thế nào? Dữ liệu lấy từ đâu?'
> -> Gợi ý trả lời: 'Dạ thưa Thầy Cô, hệ thống được thiết kế dưới dạng RESTful API bằng FastAPI. Trong tương lai, hệ thống sẽ kết nối trực tiếp với cơ sở dữ liệu điểm danh và LMS (như Moodle/Canvas) của nhà trường qua API để tự động đồng bộ chuyên cần và điểm số mỗi tuần, cố vấn học tập chỉ cần đăng nhập là thấy ngay bảng xếp hạng rủi ro của sinh viên.'

---

## SLIDE 9: CONCLUSION & FUTURE WORK (TỔNG KẾT & ĐỊNH HƯỚNG) (Thời lượng: 35 giây)
- **Mục tiêu slide:** Tổng kết 3 đóng góp lớn, thẳng thắn nhìn nhận hạn chế khoa học và gửi lời cảm ơn Hội đồng.

### Lời thoại Thuyết trình Chi tiết (Verbatim Script):
[Hành động: Bấm chuyển Slide 9. Bạn Đỗ Quang Minh trình bày kết luận]

"Kính thưa Quý Thầy Cô trong Hội đồng, tóm lược lại đồ án của Nhóm 03 đã đạt được **3 đóng góp nổi bật**:

[Hành động: Nhìn thẳng vào Hội đồng, giọng dứt khoát và chân thành]
1. **Về Kỹ thuật:** Xây dựng hoàn chỉnh Pipeline học máy khép kín, thiết kế 6 đặc trưng tương tác chuyên ngành chứng minh được hiệu quả vượt trội.
2. **Về Học thuật:** Nhóm nhìn nhận thẳng thắn hạn chế của dữ liệu cắt ngang và dữ liệu tổng hợp. Trên dữ liệu sinh viên thực tế 800 mẫu, mô hình đạt ROC-AUC thực chất là **80.50%**.
3. **Về Ứng dụng:** Đóng gói thành công Web Dashboard FastAPI hoạt động ổn định, sẵn sàng tích hợp giải thích mô hình SHAP và kết nối hệ thống LMS của trường đại học trong tương lai.

[Hành động: Cúi đầu chào trang trọng]
Nhóm 03 chúng em xin trân trọng cảm ơn sự lắng nghe và hướng dẫn tận tình của Quý Thầy Cô chương trình Samsung Innovation Campus! Chúng em xin sẵn sàng lắng nghe các câu hỏi và nhận xét từ Hội đồng!"

> [!TIP]
> **Gợi ý Trả lời Phản biện Nhanh (Q&A):**  
> Thầy Cô có thể hỏi: 'Nếu có thêm thời gian 3 tháng nữa, nhóm sẽ ưu tiên làm gì tiếp theo?'
> -> Gợi ý trả lời: 'Dạ thưa Thầy Cô, nhóm sẽ ưu tiên 2 việc: Thứ nhất là tích hợp giải thích mô hình cục bộ SHAP Waterfall Plot lên giao diện Web để giải thích lý do cụ thể cho từng sinh viên; Thứ hai là thu thập dữ liệu chuỗi thời gian theo tuần (Longitudinal Data) để thử nghiệm mô hình LSTM dự báo xu hướng suy giảm phong độ học tập.'

---

## BỘ 5 CÂU HỎI PHẢN BIỆN CHUYÊN SÂU CỦA HỘI ĐỒNG SIC & CÁCH TRẢ LỜI

### Câu hỏi 1 (Về Rò rỉ Dữ liệu - Data Leakage): 'Nhóm làm thế nào để chứng minh mô hình không bị rò rỉ dữ liệu (Data Leakage) khi kiểm định chéo?'

> [!NOTE]
> Cách trả lời chuẩn xác:
> "Dạ thưa Thầy Cô, nhóm đã áp dụng 3 biện pháp kỹ thuật nghiêm ngặt:
> 1. Đóng gói khép kín trong scikit-learn Pipeline: Toàn bộ quá trình điền khuyết KNNImputer, chuẩn hóa StandardScaler và mã hóa OneHotEncoder nằm hoàn toàn trong Pipeline, chỉ thực hiện fit trên 4 fold huấn luyện và transform trên fold kiểm thử của từng vòng lặp.
> 2. Phân chia phân tầng (Stratified K-Fold): Tỷ lệ nhãn Dropout_Risk (54.6% Yes / 45.4% No) được bảo toàn đồng đều trên cả 5 folds.
> 3. Giữ tập kiểm thử độc lập (Hold-out Test 20%): Nhóm tách riêng 20% dữ liệu ngay từ đầu bằng random_state=42 và chỉ đưa vào đánh giá sau cùng, hoàn toàn không tham gia vào quá trình khớp tham số của mô hình."

### Câu hỏi 2 (Về Dữ liệu Tổng hợp 1.200 bản ghi): 'Tại sao nhóm lại mở rộng thêm 1.200 bản ghi tổng hợp? Liệu con số 89.01% AUC có bị thổi phồng không?'

> [!NOTE]
> Cách trả lời ghi điểm tuyệt đối (Thể hiện tính trung thực khoa học):
> "Dạ thưa Thầy Cô, nhóm xin được báo cáo hết sức trung thực:
> Mục tiêu ban đầu của việc sinh 1.200 hồ sơ bằng phân phối Gaussian đa biến phân tầng là nhằm kiểm thử khả năng mở rộng kỹ thuật và kiểm tra độ trễ của giao diện Web FastAPI.
> Tuy nhiên, qua kiểm toán chuyên sâu, nhóm phát hiện việc sinh mẫu có điều kiện theo nhãn đã vô tình làm tăng độ phân tách dữ liệu (trên phần tổng hợp test đạt tới 94.3% AUC, trong khi phần gốc đạt 77.3% AUC).
> Vì vậy, nhóm xin khẳng định trước Hội đồng: Con số phản ánh chính xác nhất năng lực tổng quát hóa trên sinh viên thật là ROC-AUC 80.50% (trên 800 sinh viên gốc). Nhóm đã ghi nhận điều này như một bài học hạn chế quan trọng trong báo cáo chính thức!"

### Câu hỏi 3 (Về Tính Phi Nhân Quả của Feature Importance): 'Mức độ quan trọng đặc trưng (Feature Importance) cao có đồng nghĩa với việc đặc trưng đó là nguyên nhân gây ra bỏ học không?'

> [!NOTE]
> Cách trả lời chuẩn mực học thuật:
> "Dạ hoàn toàn không thưa Thầy Cô. Feature Importance (như Gini Importance của Random Forest hay Gain của XGBoost) chỉ phản ánh mức độ đóng góp của biến trong việc giảm thiểu độ vẩn đục (impurity) khi phân tách các nút của cây quyết định. Đây là mối liên hệ thống kê trong cấu trúc mô hình, không đại diện cho mối quan hệ nhân quả (Causal Relationship). Để khẳng định quan hệ nhân quả, cần có nghiên cứu can thiệp đối chứng ngẫu nhiên (A/B Testing hoặc RCT) hoặc phân tích biến công cụ (Instrumental Variables)."

### Câu hỏi 4 (Về Lựa chọn Mô hình Triển khai): 'Tại sao Kịch bản 1 chọn Random Forest mà Kịch bản 2 lại chọn XGBoost để triển khai trên Web?'

> [!NOTE]
> Cách trả lời kỹ thuật vững vàng:
> "Dạ thưa Thầy Cô, ở Kịch bản 1 khi có mặt biến Burnout_Level, cấu trúc dữ liệu có tính phân tầng rõ rệt, cơ chế lấy mẫu ngẫu nhiên (Bagging) của Random Forest giúp kiểm soát phương sai rất tốt và đạt AUC 89.01%. Trong khi ở Kịch bản 2 (Cảnh báo sớm, không có biến Burnout), bài toán trở nên khó hơn với các tín hiệu yếu và phi tuyến từ hành vi học tập thô. Cơ chế tăng cường độ dốc (Boosting) của XGBoost với các siêu tham số phạt L2 (reg_lambda=2.0) và colsample_bytree=0.8 đã phát huy hiệu quả vượt trội trong việc khai phá các tương tác phức tạp này, đạt AUC 88.26%."

### Câu hỏi 5 (Về Giá trị Ứng dụng Thực tiễn): 'Nếu một trường đại học áp dụng hệ thống của nhóm, quy trình can thiệp sẽ diễn ra như thế nào?'

> [!NOTE]
> Cách trả lời giàu tính thực tiễn:
> "Dạ thưa Thầy Cô, hệ thống sẽ vận hành theo 3 cấp độ can thiệp tự động:
> • Cấp độ 1 (Rủi ro Thấp < 30%): Hệ thống ghi nhận hồ sơ học tập bình thường, chỉ gửi bản tin sinh hoạt và phương pháp học tập đại học định kỳ.
> • Cấp độ 2 (Rủi ro Trung bình 30%–70%): Hệ thống gửi thông báo cho Cố vấn học tập khoa để mời sinh viên trao đổi, tìm hiểu lý do vắng học và ghép vào nhóm học tập (Study Group) có gia sư phụ đạo.
> • Cấp độ 3 (Báo động đỏ > 70%): Kích hoạt quy trình phối hợp khẩn cấp giữa Phòng Đào tạo (giãn tiến độ học, bảo lưu tín chỉ), Phòng Công tác Sinh viên (xét học bổng hỗ trợ khó khăn) và Phòng Tham vấn Tâm lý (trị liệu kiệt sức học đường)."
