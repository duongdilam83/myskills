# MySkills Repository

Kho lưu trữ các kỹ năng (Agent Skills) tự động hóa phục vụ công việc và học tập.

---

## Danh mục kỹ năng

### 1. [Google Sheets Teacher Feedback Generator](./google-sheets-teacher-feedback/)
- **Mã kỹ năng:** `google-sheets-teacher-feedback`
- **Mô tả:** Tự động kết nối Google Sheets bằng Service Account (PowerShell/.NET), đọc điểm số BTVN của học sinh, tự động sinh nhận xét sư phạm ngẫu nhiên 3–6 gạch đầu dòng theo tiêu chuẩn giáo viên, ghi trực tiếp vào Google Sheets và định dạng Wrap Text, Auto Row Height.
- **Tài liệu chi tiết:** [Xem SKILL.md](./google-sheets-teacher-feedback/SKILL.md)

---

## Hướng dẫn cài đặt & Sử dụng

### Cấu hình Google Service Account
1. Chuẩn bị file JSON Service Account từ Google Cloud Console.
2. Đổi tên thành `service_account.json` và lưu vào thư mục dự án (file này được tự động bảo vệ qua `.gitignore`, không bao giờ bị push lên GitHub).
3. Chia sẻ bảng tính Google Sheets cho email của Service Account với quyền **Editor**.

### Chạy kỹ năng
Xem hướng dẫn chi tiết trong từng thư mục kỹ năng tương ứng.
