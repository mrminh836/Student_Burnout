"""
Generates a super concise, easy-to-read Word speaking script (.docx)
for the 10-Slide presentation (5 to 8 minutes total speaking time).
"""

import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FINAL_DIR = os.path.join(BASE_DIR, "FINAL_PRESENTATION")
VI_DIR = os.path.join(BASE_DIR, "SIC_Documents_VI")
DRV_DIR = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"

SLIDE_SCRIPTS = [
    (
        "SLIDE 01: TRANG BÌA (Thời lượng: ~30 giây)",
        "Đỗ Quang Minh",
        "Em xin kính chào quý Thầy Cô trong Hội đồng chấm đồ án Samsung Innovation Campus! Em là Đỗ Quang Minh, cùng bạn Vũ Quang Huy thuộc Nhóm 03. Hôm nay nhóm em xin trình bày đề tài: Xây dựng, huấn luyện và so sánh các mô hình Học máy dự đoán tỷ lệ bỏ học dựa trên Căng thẳng và Tham gia học tập. Trên màn hình là 4 con số kết quả cốt lõi của nhóm: 2.000 hồ sơ sinh viên sạch, 6 chỉ số đặc trưng mới, 4 mô hình Học máy tiêu biểu và độ chính xác ROC-AUC đạt đỉnh 89,01%."
    ),
    (
        "SLIDE 02: ĐẶT VẤN ĐỀ & MỤC TIÊU (Thời lượng: ~45 giây)",
        "Đỗ Quang Minh",
        "Thưa Thầy Cô, sinh viên bỏ học không phải là chuyện xảy ra đột ngột, mà là hệ quả tích lũy từ hai phía: Một là sự tham gia học tập suy giảm như vắng học, ít tự học, nợ môn; Hai là áp lực tâm lý kéo dài như căng thẳng, lo âu và gánh nặng tài chính. Khi hai yếu tố này cộng hưởng sẽ đẩy sinh viên vào trạng thái trung gian là Kiệt sức học đường (Burnout) với 3 mức Thấp, Vừa và Cao, từ đó dẫn thẳng tới quyết định bỏ học. Vì vậy, mục tiêu của Nhóm 03 là xây dựng hệ thống AI dự báo nguy cơ bỏ học trên cả 2 kịch bản: Chẩn đoán đầy đủ và Cảnh báo sớm, đồng thời triển khai thành ứng dụng Web thực tế."
    ),
    (
        "SLIDE 03: DỮ LIỆU & LÀM SẠCH 6 BƯỚC (Thời lượng: ~45 giây)",
        "Vũ Quang Huy",
        "Em xin tiếp tục phần Dữ liệu ạ. Bộ dữ liệu gốc ban đầu gồm 800 sinh viên với 25 thuộc tính, trong đó có 18 cột bị khuyết tổng cộng 612 ô. Để không làm mất dữ liệu, nhóm em thực hiện quy trình làm sạch 6 bước: Xóa trùng lặp ID; Điền khuyết cột số bằng thuật toán KNN 7 láng giềng gần nhất ngay trong Pipeline; Điền khuyết cột chữ bằng giá trị phổ biến nhất; Lọc ngoại lai bằng khoảng tứ phân vị IQR; Tạo đặc trưng mới và mở rộng đa biến có phân tầng lên 2.000 hồ sơ sinh viên sạch hoàn toàn 100%, không còn bất kỳ giá trị lỗi nào."
    ),
    (
        "SLIDE 04: QUY TRÌNH HỌC MÁY & 6 ĐẶC TRƯNG MỚI (Thời lượng: ~45 giây)",
        "Vũ Quang Huy",
        "Trên Slide 4 là quy trình Học máy 5 bước khép kín và đóng góp kỹ thuật quan trọng nhất của nhóm: Thiết kế 6 chỉ số chuyên ngành mới để đo lường cán cân giữa Căng thẳng và Gắn kết học tập. Nổi bật nhất là Chỉ số Gắn kết học tập, Chỉ số Căng thẳng tâm lý, và đặc biệt là Tỷ số Căng thẳng trên Gắn kết — đây là biến quan trọng đứng thứ 2 toàn hệ thống, giúp mô hình nhận diện ngay khi áp lực vượt quá sức chịu đựng của sinh viên."
    ),
    (
        "SLIDE 05: PHÁT HIỆN EDA 1 — TÁC ĐỘNG CỦA KIỆT SỨC (Thời lượng: ~40 giây)",
        "Vũ Quang Huy",
        "Qua phân tích khám phá dữ liệu, phát hiện lớn nhất của nhóm nằm ở Slide 5: Mức độ kiệt sức quyết định trực tiếp đến tỷ lệ bỏ học. Ở nhóm Kiệt sức Thấp, tỷ lệ bỏ học chỉ là 23,8%. Sang mức Kiệt sức Vừa, tỷ lệ này tăng lên 63,2%. Và khi sinh viên rơi vào trạng thái Kiệt sức Cao, tỷ lệ bỏ học vọt lên tới 92,0% — tức là cứ 10 bạn kiệt sức cao thì hơn 9 bạn có nguy cơ bỏ học, gấp gần 4 lần nhóm thấp!"
    ),
    (
        "SLIDE 06: PHÁT HIỆN EDA 2 — YẾU TỐ BẢO VỆ VS. NGUY CƠ (Thời lượng: ~40 giây)",
        "Vũ Quang Huy",
        "Ở Slide 6, ma trận tương quan chỉ rõ hai nhóm yếu tố đối lập: Nhóm yếu tố bảo vệ giúp giảm bỏ học đứng đầu là Chuyên cần cao, Động lực học tập, Giờ tự học và Hỗ trợ gia đình. Ngược lại, nhóm yếu tố nguy cơ làm tăng bỏ học gồm Kiệt sức, Điểm căng thẳng, Lo âu và Áp lực tài chính — đặc biệt nhóm sinh viên thu nhập thấp có tỷ lệ bỏ học lên tới 62,1%. Em xin nhường lời lại cho bạn Minh trình bày phần Mô hình Học máy ạ."
    ),
    (
        "SLIDE 07: KẾT QUẢ SO SÁNH 4 MÔ HÌNH HỌC MÁY (Thời lượng: ~55 giây)",
        "Đỗ Quang Minh",
        "Cảm ơn Huy. Để đảm bảo tính khách quan, nhóm em huấn luyện và đánh giá chéo 5-Fold trên 4 mô hình đại diện cho 4 trường phái: Logistic Regression, Random Forest, XGBoost và Soft Voting Ensemble trong 2 kịch bản. Ở Kịch bản 1 (Đầy đủ đặc trưng bao gồm cả Burnout), Random Forest dẫn đầu với ROC-AUC đạt 89,01%, độ chính xác trên tập Test đạt 81,50% và F1-Score đạt 83,63%. Ở Kịch bản 2 (Ẩn hoàn toàn biến Burnout để mô phỏng cảnh báo sớm từ đầu kỳ), mô hình XGBoost vẫn đạt ROC-AUC lên tới 88,26%. Điều này chứng minh hệ thống cảnh báo sớm rất hiệu quả ngay cả khi chưa khảo sát tâm lý."
    ),
    (
        "SLIDE 08: GIẢI THÍCH MÔ HÌNH & TOP 6 YẾU TỐ QUYẾT ĐỊNH (Thời lượng: ~40 giây)",
        "Đỗ Quang Minh",
        "Tại Slide 8, đường cong ROC và biểu đồ Độ quan trọng đặc trưng cho thấy 2 chỉ số mới do nhóm tự thiết kế là Tỷ số Căng thẳng trên Gắn kết và Chỉ số Gắn kết học tập đứng ngay vị trí Top 2 và Top 3 sau biến Kiệt sức. Đặc biệt, mô hình đạt độ nhạy Recall trên 84%, giúp nhà trường phát hiện đúng hơn 4/5 số sinh viên thực sự có nguy cơ bỏ học để can thiệp kịp thời."
    ),
    (
        "SLIDE 09: ỨNG DỤNG WEB DASHBOARD CẢNH BÁO SỚM (Thời lượng: ~45 giây)",
        "Đỗ Quang Minh",
        "Không dừng lại ở lý thuyết, nhóm em đã đóng gói mô hình thành ứng dụng Web FastAPI hoàn chỉnh theo kiến trúc 3 tầng trên Slide 9. Giao diện Web có sẵn 3 mẫu hồ sơ sinh viên để thử nhanh: Mẫu An toàn (nguy cơ 8%), Mẫu Cảnh báo (58%) và Mẫu Báo động đỏ (95,8%). Chỉ trong chưa đầy 30 mili-giây, hệ thống trả về dự báo kép cả mức độ Kiệt sức lẫn xác suất Bỏ học, kèm theo khuyến nghị cá nhân hóa như xếp lớp phụ đạo, tư vấn tâm lý hoặc hỗ trợ học bổng."
    ),
    (
        "SLIDE 10: KẾT LUẬN & HỎI ĐÁP (Thời lượng: ~30 giây)",
        "Đỗ Quang Minh",
        "Tóm lại ở Slide 10, Nhóm 03 đã hoàn thành trọn vẹn từ chuẩn hóa 2.000 hồ sơ sạch, thiết kế 6 đặc trưng mới, huấn luyện 4 mô hình đạt đỉnh 89,01% ROC-AUC đến xây dựng Web App thực tế. Trong tương lai, nhóm định hướng tích hợp giải thích SHAP trực tiếp lên Web và kết nối API với hệ thống LMS của nhà trường. Bài trình bày gói gọn trong 6 phút của nhóm em xin kết thúc tại đây. Chúng em rất mong nhận được câu hỏi và góp ý của quý Thầy Cô ạ!"
    ),
]


def build_doc():
    doc = Document()
    title = doc.add_heading("KỊCH BẢN THUYẾT TRÌNH NGẮN GỌN (5 – 8 PHÚT) — 10 SLIDES", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("Đề tài: Dự báo Kiệt sức Học đường & Nguy cơ Bỏ học Sinh viên (Nhóm 03: Đỗ Quang Minh & Vũ Quang Huy)")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].bold = True

    tip = doc.add_paragraph(
        "💡 MẸO NÓI NHANH & TỰ NHIÊN (TỔNG THỜI GIAN ~6.5 PHÚT):\n"
        "• Bộ slide mới đã được rút gọn từ 14 slide xuống đúng 10 slide cực kỳ thoáng, chữ to (14pt), ít chữ.\n"
        "• Mỗi slide chỉ cần nói trong khoảng 35 – 50 giây theo đúng lời thoại ngắn gọn dưới đây."
    )
    tip.runs[0].font.color.rgb = RGBColor(0, 102, 153)

    for s_title, speaker, text in SLIDE_SCRIPTS:
        h = doc.add_heading(f"{s_title} — Người nói: {speaker}", level=2)
        h.runs[0].font.color.rgb = RGBColor(29, 53, 87)
        p = doc.add_paragraph(f"🎤 Lời nói: \"{text}\"")
        p.runs[0].font.size = Pt(11.5)

    out_paths = [
        os.path.join(FINAL_DIR, "Kich_Ban_Thuyet_Trinh_Ngan_Gon_5_8_Phut.docx"),
        os.path.join(VI_DIR, "Kich_Ban_Thuyet_Trinh_Ngan_Gon_5_8_Phut.docx"),
        os.path.join(VI_DIR, "Kich_Ban_Thuyet_Trinh_Tung_Slide_Chi_Tiet.docx"),
    ]
    if os.path.exists(DRV_DIR):
        out_paths.append(os.path.join(DRV_DIR, "Kich_Ban_Thuyet_Trinh_Ngan_Gon_5_8_Phut.docx"))

    for path in out_paths:
        try:
            doc.save(path)
            print("Saved short script ->", path)
        except Exception as e:
            print("Skip save:", path, e)


if __name__ == "__main__":
    build_doc()

