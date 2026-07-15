# Đề tài 50: Quy trình kiểm thử bảo mật cho một sản phẩm IoT

## 1. Mô tả đề tài
Tiểu luận này tập trung nghiên cứu và xây dựng một **quy trình kiểm thử bảo mật tổng quát và chuẩn hóa**, bao phủ đầy đủ các lớp cấu trúc kiến trúc của một hệ sinh thái IoT (bao gồm: phần cứng, firmware, mạng, API/cloud, ứng dụng di động, và cơ chế cập nhật). 

Để minh chứng tính khả thi của quy trình đề xuất, nghiên cứu áp dụng mô hình hóa hệ thống và giả lập dữ liệu thực nghiệm trên một sản phẩm giả định cụ thể là **Hệ thống Camera IP gia đình**. Hệ thống giả định này được phân tích cấu trúc dựa trên các phương pháp luận bảo mật quốc tế từ OWASP:
*   **OWASP ISTG:** Khung sườn cốt lõi để định hình quy trình kiểm thử từng bước.
*   **OWASP ISVS:** Danh mục các yêu cầu an toàn dùng làm tiêu chí đánh giá đạt/không đạt cho từng điểm kiểm thử.
*   **OWASP FSTM:** Phương pháp luận định hình các giai đoạn phân tích bảo mật firmware.

---

## 2. Thành viên thực hiện
*   **Họ và tên học viên:** Chu Nguyễn Thái Tuấn
*   **Mã học viên:** 231A390004
*   **Ngành đào tạo:** An toàn thông tin
*   **Học phần:** Bảo mật trong IoT
*   **Đơn vị đào tạo:** Trường Đại học Văn Hiến
*   **Giảng viên giảng dạy:** Hồ Nhựt Minh

---

## 3. Nguồn tài liệu đã dùng (References)
Dự án được hoàn thiện dựa trên việc trích dẫn và tham chiếu các nguồn học thuật và tiêu chuẩn kỹ thuật sau:
1.  Brian Russell, Drew Van Duren, *Practical Internet of Things Security*, Packt Publishing, 2016.
2.  OWASP Foundation, *OWASP IoT Security Testing Guide (ISTG)*. Đường dẫn: [https://github.com/OWASP/owasp-istg](https://github.com/OWASP/owasp-istg).
3.  OWASP Foundation, *OWASP IoT Security Verification Standard (ISVS)*. Đường dẫn: [https://github.com/OWASP/IoT-Security-Verification-Standard-ISVS](https://github.com/OWASP/IoT-Security-Verification-Standard-ISVS).
4.  OWASP Foundation, *OWASP Firmware Security Testing Methodology (FSTM)*. Đường dẫn: [https://github.com/scriptingxss/owasp-fstm](https://github.com/scriptingxss/owasp-fstm).
5.  OWASP Foundation, *OWASP IoTGoat Project*. Đường dẫn: [https://github.com/OWASP/IoTGoat](https://github.com/OWASP/IoTGoat).
6.  OWASP Foundation, *OWASP Internet of Things Project*. Đường dẫn: [https://github.com/OWASP/www-project-internet-of-things](https://github.com/OWASP/www-project-internet-of-things).
7.  Trường Đại học Văn Hiến, Khoa Công nghệ Thông tin, *Hướng dẫn chi tiết tiểu luận Bảo mật IoT — 50 đề tài có GitHub, cách làm, cách trình bày và thang điểm chi tiết cho sinh viên*, TP. Hồ Chí Minh, 2026.

---

## 4. Giới hạn an toàn & Phạm vi pháp lý (Safety Boundaries)
Để bảo đảm tính hợp pháp, an toàn thông tin mạng và đạo đức nghề nghiệp, dự án tuân thủ nghiêm ngặt các giới hạn cam kết sau:
*   **Tính chất dữ liệu:** Toàn bộ quá trình xây dựng, chạy thử quy trình và các lỗ hổng/lỗi an ninh phát hiện đều được thực hiện dưới dạng dữ liệu giả lập và mô hình hóa cục bộ trên môi trường Lab học tập lý thuyết
*   **Hạn chế can thiệp:** Dự án tuyệt đối **không thực hiện hành vi tấn công, quét lỗ hổng (scanning) hoặc khai thác (exploiting)** bất kỳ thiết bị vật lý thật hay hạ tầng Cloud Server thương mại nào đang hoạt động ngoài thực tế của bên thứ ba
*   **An toàn vật lý:** Quá trình đánh giá lớp phần cứng vật lý (cổng debug UART/JTAG, giao diện vật lý) chỉ dừng lại ở mức độ lý thuyết và quan sát trực quan, không can thiệp tháo rã thiết bị thật gây nguy cơ chập cháy điện
