# HỆ THỐNG THIẾT KẾ THỊ GIÁC (VISUAL DESIGN SYSTEM)
## DỰ ÁN: DỰ BÁO KIỆT SỨC HỌC ĐƯỜNG & NGUY CƠ BỎ HỌC SINH VIÊN (SIC AI CAPSTONE)
**Đơn vị thực hiện:** Nhóm 03 (Đỗ Quang Minh – Trưởng nhóm, Vũ Quang Huy – Thành viên)  
**Vai trò thiết kế:** Senior Presentation Designer & PowerPoint Art Director  
**Tỷ lệ khung hình:** 16:9 Widescreen (`13.333" × 7.5"`)

---

## 1. PHÂN TÍCH BỘ SLIDE TRƯỚC KHI REDESIGN (AUDIT OF PREVIOUS DECK)

Qua quá trình đánh giá kiến trúc thông tin và thị giác của phiên bản trước, nhóm thiết kế xác định 4 vấn đề cốt lõi cần giải quyết triệt để:
1. **Sự lặp lại bố cục (Layout Monotony):** Các slide nội dung trước đó sử dụng lặp đi lặp lại cùng một kiểu khung thẻ chữ nhật (2 hoặc 3 cột) trên nền tối, khiến mạch trình bày thiếu điểm nhấn thị giác và gây mỏi mắt khi theo dõi liên tục.
2. **Độ tương phản biểu đồ chưa đồng bộ:** Một số biểu đồ phân tích có nền trắng đặt trên nền slide tối tạo cảm giác tách rời, thiếu sự hòa quyện của một hệ thống xuất bản học thuật (editorial design system).
3. **Sử dụng ký tự biểu tượng rời rạc:** Việc dùng ký tự đặc biệt/emoji thay cho hệ thống huy hiệu hình học vector (vector badges, numbered pills, chevrons) làm giảm tính trang trọng trước hội đồng khoa học.
4. **Cần phân bổ lại mật độ thông tin:** Các slide kỹ thuật (như 6 công thức Feature Engineering hoặc bảng so sánh CV vs. Test) cần bố cục Bento Grid và Matrix Table chuyên biệt để vừa bảo toàn 100% số liệu, vừa thoáng mắt.

---

## 2. BẢNG MÀU CHỦ ĐẠO (EDITORIAL HIGH-CONTRAST COLOR PALETTE)

Hệ thống thiết kế mới áp dụng triết lý **Editorial Contrast Rhythm**: kết hợp **Deep Navy (`#0B162C`)** ở trang bìa (Hero Cover), các khối nhấn (Impact Panels) và trang kết (Finale) với nền **Alabaster trắng sáng (`#F8FAFC` / `#FFFFFF`)** trên các trang phân tích dữ liệu, sơ đồ kỹ thuật và bảng biểu.

| Vai trò màu (Color Role) | Tên màu | Mã HEX | Mã RGB | Ứng dụng trong Slide |
| :--- | :--- | :---: | :---: | :--- |
| **Primary Brand / Dark Canvas** | **Deep Navy** | `#0B162C` | `(11, 22, 44)` | Nền Slide 01 (Cover), Slide 12 (Finale), cột Split-Screen Slide 03, thanh tiêu đề bảng và hộp công thức toán học. |
| **Secondary Dark Surface** | **Midnight Slate** | `#132240` | `(19, 34, 64)` | Bề mặt thẻ KPI trên nền tối (Slide 01, Slide 04 WBS, Slide 12). |
| **Editorial Light Canvas** | **Crisp Alabaster** | `#F8FAFC` | `(248, 250, 252)` | Nền chính cho các Slide nội dung, quy trình và phân tích dữ liệu (Slide 02 – Slide 11). |
| **Card Surface** | **Pure White** | `#FFFFFF` | `(255, 255, 255)` | Bề mặt thẻ thông tin, khung chứa biểu đồ và khung trình duyệt Web. |
| **Primary Accent** | **Electric Blue** | `#0066FF` | `(0, 102, 255)` | Màu nhấn chính cho Kịch bản 1, huy hiệu số thứ tự, đường kẻ thương hiệu và chỉ số dẫn đầu. |
| **Secondary Tech Accent** | **Vibrant Cyan** | `#00B4D8` | `(0, 180, 216)` | Màu nhấn công nghệ AI, văn bản công thức toán học và điểm nối sơ đồ. |
| **Positive / Protective Accent** | **Emerald Teal** | `#0D9488` | `(13, 148, 136)` | Nhóm yếu tố bảo vệ (Protective Factors), Kịch bản 2 Cảnh báo sớm và trạng thái An toàn (Low Burnout). |
| **Warning / Transition Accent** | **Warm Amber** | `#D97706` | `(217, 119, 6)` | Trạng thái trung gian (Medium Burnout 63.2%) và điểm nhấn cột mốc. |
| **High-Risk Alert Accent** | **Rose Coral** | `#E11D48` | `(225, 29, 72)` | Cảnh báo rủi ro cao (High Burnout 92.0%, nhóm yếu tố nguy cơ $r > 0$). |
| **Primary Ink Text** | **Slate 900** | `#0F172A` | `(15, 23, 42)` | Tiêu đề slide và số liệu chính trên nền sáng (tương phản chuẩn WCAG AAA). |
| **Body & Caption Text** | **Slate 700 / 500** | `#334155` | `(51, 65, 85)` | Nội dung giải thích, chú thích chân trang và nhãn trục đồ thị. |

---

## 3. HỆ THỐNG TYPOGRAPHY & PHÂN CẤP THÔNG TIN

* **Font Family duy nhất:** `Segoe UI` (tương thích hoàn hảo 100% tiếng Việt có dấu và ký hiệu toán học trên mọi phiên bản Microsoft PowerPoint Windows/macOS).
* **Thang phân cấp kích thước chữ (Type Scale):**
  * **Hero Cover Title & Big Numbers:** `26.5 pt – 44 pt` Bold (Tạo điểm hút thị giác ngay lập tức).
  * **Slide Header Title:** `22 pt` Bold (`#0F172A` trên nền sáng, `#FFFFFF` trên nền tối).
  * **Slide Takeaway Subtitle:** `11.5 pt` Regular (`#64748B` / `#CBD5E1`) — Tóm tắt luận điểm chính của từng slide trong đúng 1 dòng.
  * **Section / Card Heading:** `12 pt – 14 pt` Bold.
  * **Body Copy & Table Cells:** `10 pt – 11.5 pt` Regular/Bold (Khoảng cách đoạn `space_after = 5pt – 10pt`, tuyệt đối không tràn khung).
  * **Category Pill & Footer Metadata:** `8.5 pt – 9.5 pt` Bold.

---

## 4. HỆ THỐNG 12 BỐ CỤC ĐỘC BẢN (12 BESPOKE LAYOUT ARCHETYPES)

Mỗi slide trong bộ 12 slide được thiết kế với một cấu trúc hình học riêng biệt, phù hợp chính xác với bản chất dữ liệu của slide đó:

1. **Slide 01 – Hero Cover Layout:** Chia bất đối xứng 62% / 38% trên nền Deep Navy. Bên trái là cụm tiêu đề lớn và thẻ tác giả; bên phải là cột 4 thẻ KPI hình học có dải màu nhấn dọc.
2. **Slide 02 – Connected Roadmap Timeline Layout:** Thanh trục ngang xuyên suốt giữa màn hình với 6 nút tròn đánh số `01` $\rightarrow$ `06`, kết nối xen kẽ lên/xuống với 6 thẻ nội dung (thay thế hoàn toàn kiểu danh sách gạch đầu dòng nhàm chán).
3. **Slide 03 – Split-Screen + Causal Process Flow Layout:** Cột trái Deep Navy hiển thị con số lớn `54.62%` sinh viên có nguy cơ bỏ học; 2/3 bên phải là sơ đồ luồng nhân quả 3 giai đoạn dùng khối mũi tên `Chevron` kết nối với 3 cột phân tích.
4. **Slide 04 – Asymmetric Two-Column + Gantt Timeline Layout:** Nửa trên chia 2 cột (4 Mục tiêu đánh số tròn bên trái + 2 thẻ phân công 50/50 bên phải); nửa dưới là thanh tiến độ WBS 4 giai đoạn trên nền tối.
5. **Slide 05 – Big-Number Strip + Data Pipeline Layout:** Hàng ngang 5 thẻ số liệu lớn (`800` $\rightarrow$ `612` $\rightarrow$ `k=7` $\rightarrow$ `2.000` $\rightarrow$ `34`) ở phía trên; phía dưới chia đôi giữa khung giải thích 5 nhóm thuộc tính/nguồn gốc dữ liệu và biểu đồ phân tích dữ liệu khuyết.
6. **Slide 06 – Bento Infographic Grid Layout:** Dải quy trình 5 bước ở trên cùng kết hợp lưới Bento 2×3 gồm 6 mô-đun đặc trưng (`F1` $\rightarrow$ `F6`). Mỗi mô-đun có huy hiệu tròn, hộp công thức nền tối chữ Cyan và nhãn vai trò dự báo.
7. **Slide 07 – Data Visualization + Stacked Big-Number Staircase Layout:** Bên trái là biểu đồ cột lớn nền trắng tinh gọn; bên phải là 3 khối số lớn xếp tầng (`23.8%` $\rightarrow$ `63.2%` $\rightarrow$ `92.0%`) kèm hộp lưu ý học thuật phi nhân quả.
8. **Slide 08 – Comparison / Tug-of-War Layout:** Bên trái là Heatmap tương quan Pearson; bên phải đặt đối xứng trực diện khối **Yếu tố Bảo vệ (Teal, $r < 0$)** và khối **Yếu tố Nguy cơ (Coral, $r > 0$)**.
9. **Slide 09 – Benchmark Chart + Executive Matrix Table Layout:** Nửa trên kết hợp biểu đồ cột đôi ROC-AUC và 2 thẻ Quán quân từng kịch bản; nửa dưới là bảng ma trận 9 cột đối sánh đầy đủ 5-Fold CV và Hold-out Test 20%.
10. **Slide 10 – Dual Visual Explainability Layout:** Đặt song song 2 đồ thị giải thích mô hình (Đường cong ROC/PR và Top 15 Feature Importance) phía trên một thanh tổng kết ngang màu Deep Navy.
11. **Slide 11 – Architecture Ribbon + Browser Frame Showcase Layout:** Phía trên là dải băng 3 bước (`01 INPUT` $\rightarrow$ `02 AI ENGINE` $\rightarrow$ `03 RESULT`); phía dưới bên trái là khung trình duyệt Web (có thanh địa chỉ & 3 nút điều khiển) chứa ảnh chụp thực tế của `app.py`, bên phải là 3 thẻ Preset thử nghiệm nhanh.
12. **Slide 12 – Dark Executive Conclusion & Finale Layout:** Nền Deep Navy trang trọng với 3 trụ cột đóng góp lớn, 3 thẻ định hướng tương lai (Multi-campus, SHAP, LMS API) và banner `THANK YOU & Q&A`.

