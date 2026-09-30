# DATA GOVERNANCE LOG

Tài liệu này ghi nhận mọi thao tác quản trị, làm sạch và biến đổi dữ liệu nhằm đảm bảo tính minh bạch và khả năng lặp lại (reproducibility) của nghiên cứu.

## Nguyên tắc cốt lõi:
1. **Bất khả xâm phạm dữ liệu thô:** Dữ liệu trong thư mục `raw_survey_data/` và `raw_interview_data/` là dữ liệu gốc, **tuyệt đối không được ghi đè hay chỉnh sửa trực tiếp**.
2. **Ghi chép minh bạch:** Mọi bước xử lý từ file raw sang file cleaned phải được ghi log tại đây.
3. **Ẩn danh:** Không ghi thông tin định danh chuyên gia vào bản thảo nếu chưa được phép, sử dụng mã CG01, CG02...
4. **Trích dẫn:** Mỗi bảng/hình tạo từ dữ liệu phải trích dẫn rõ file nguồn và biến được dùng.

---

## Lịch sử xử lý dữ liệu

| Ngày (YYYY-MM-DD) | Thao tác / Sự kiện | Người thực hiện | File Input | File Output | Ghi chú |
| :--- | :--- | :--- | :--- | :--- | :--- |
| 2026-07-11 | Khởi tạo Hệ thống Quản trị Dữ liệu Nghiên cứu | AI Assistant | - | Toàn bộ thư mục 05_Data_Governance | Đã thiết lập thành công. Chờ học viên nạp dữ liệu. |
