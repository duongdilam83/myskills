# MySkills Repository

Kho lưu trữ các kỹ năng (Agent Skills) tự động hóa phục vụ công việc và học tập.

---

## Danh mục kỹ năng

### 1. [Google Sheets Teacher Feedback Generator](./google-sheets-teacher-feedback/)
- **Mã kỹ năng:** `google-sheets-teacher-feedback`
- **Mô tả:** Tự động kết nối Google Sheets bằng Service Account (PowerShell/.NET), đọc điểm số BTVN của học sinh, tự động sinh nhận xét sư phạm ngẫu nhiên 3–6 gạch đầu dòng theo tiêu chuẩn giáo viên, ghi trực tiếp vào Google Sheets, định dạng Wrap Text, Auto Row Height và tự động gửi email thông báo kết quả học tập cho phụ huynh (nếu có cột Email phụ huynh).
- **Tài liệu chi tiết:** [Xem SKILL.md](./google-sheets-teacher-feedback/SKILL.md)

---

## Hướng dẫn cài đặt & Sử dụng

### 1. Cấu hình Google Service Account
1. Chuẩn bị file JSON Service Account từ Google Cloud Console.
2. Lưu file vào thư mục **`Documents`** (ví dụ: `mathplus-508610-78878d97806f.json`) hoặc lưu trực tiếp vào thư mục kỹ năng/dự án với tên `service_account.json`. Script sẽ **tự động phát hiện** file Service Account trong thư mục `Documents` nếu không truyền đường dẫn.
3. Chia sẻ bảng tính Google Sheets cho email của Service Account (ví dụ: `mathplus@mathplus-508610.iam.gserviceaccount.com`) với quyền **Editor**.

### 2. Cấu hình Gửi Email cho phụ huynh (Tùy chọn)
1. Để gửi email tự động qua Gmail hoặc Outlook, tạo file `email_config.json` và lưu vào thư mục **`Documents`** hoặc thư mục `scripts/` (xem file mẫu tại `scripts/email_config.example.json`).
2. Script sẽ **tự động phát hiện** file cấu hình email trong thư mục `Documents` và tự động gửi email cho phụ huynh có địa chỉ email trong bảng tính. Nếu học sinh không có email, script sẽ tự động bỏ qua.

### 3. Tích hợp toàn cục vào Antigravity
Kỹ năng trong kho lưu trữ này đã được liên kết toàn cục vào hệ thống Antigravity (`~/.gemini/config/plugins/myskills` và `~/.gemini/config/skills.json`). Bạn có thể gọi kỹ năng này trong **bất kỳ dự án hay phiên làm việc nào của Antigravity** mà không cần cấu hình lại.

### 4. Chạy kỹ năng
Xem hướng dẫn chi tiết trong từng thư mục kỹ năng tương ứng.
