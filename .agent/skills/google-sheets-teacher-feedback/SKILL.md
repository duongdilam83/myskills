---
name: google-sheets-teacher-feedback
description: >-
  Tự động kết nối Google Sheets bằng Service Account (PowerShell/.NET), đọc điểm số BTVN của học sinh,
  sinh nhận xét sư phạm ngẫu nhiên từ 3 đến 6 gạch đầu dòng mang tính khích lệ,
  cập nhật trực tiếp vào Google Sheets API kèm định dạng tự động xuống dòng (Wrap Text), co giãn chiều cao (Auto Height),
  và tự động gửi email báo cáo tình hình học tập kèm nhận xét cho phụ huynh (nếu có email).
---

# Google Sheets Teacher Feedback Generator

## Overview
Kỹ năng tự động hóa quy trình quản lý điểm số, tạo nhận xét học sinh và gửi thông báo cho phụ huynh:
- Tự động xác thực Google Sheets API bằng khóa bí mật **Google Cloud Service Account** (sinh JWT Bearer token trực tiếp bằng PowerShell/.NET, không yêu cầu cài đặt Python hay thư viện ngoài).
- Đọc dữ liệu điểm bài tập về nhà (BTVN) và danh sách học sinh theo từng lớp.
- Sinh nhận xét chuẩn mực sư phạm với **3 đến 6 gạch đầu dòng ngẫu nhiên** (không cố định số dòng), phân loại chính xác theo thang điểm:
  - **Dưới 5 điểm:** "Cần cố gắng" kèm nhắc nhở ôn lại kiến thức cơ bản.
  - **5 - 6 điểm:** "Con cần học tập chăm chỉ hơn", rèn luyện thêm bài tập.
  - **6 - 7 điểm:** "Con có ý thức làm BTVN", cẩn thận hơn trong bài làm.
  - **7 - 8 điểm:** "Con đã biết cách làm bài, cần chú ý đọc kỹ đề và cẩn thận hơn".
  - **8 - 9 điểm:** "Con hiểu bài nhưng cần cẩn thận hơn trong tính toán".
  - **10 điểm (hoặc 9 - 10 điểm):** "Bài làm tốt, con tiếp tục phát huy nhé!".
- Văn phong xúc tích, khích lệ người học: điểm cao không khen quá mức, điểm thấp không chê bai gây nản lòng, mỗi học sinh đều có nhận xét biến thể tự nhiên, tránh trùng lặp máy móc.
- Tự động cấu hình định dạng ô: kích hoạt chế độ **Wrap Text (tự động xuống dòng)** và **Auto Row Height (tự động căn chỉnh chiều cao dòng)** để bảng tính luôn gọn gàng, trực quan.
- **Tự động gửi email cho phụ huynh:** Quét cột `Email phụ huynh` (tự động dò cột hoặc chỉ định index). Nếu có email, hệ thống soạn thư sư phạm trang trọng (gồm cả định dạng HTML và văn bản thuần túy) gửi qua SMTP (.NET). Nếu ô email trống, tự động bỏ qua.

## Quick Start

### 1. Chuẩn bị thông tin xác thực
1. **Google Service Account:** Đặt file khóa JSON vào thư mục **`Documents`** (ví dụ: `C:\Users\Admin\Documents\mathplus-508610-78878d97806f.json`) hoặc để trong thư mục dự án với tên `service_account.json`. Script sẽ **tự động quét và nhận diện**.
2. **Cấu hình Email gửi đi (Tùy chọn - Hỗ trợ 2 cách):**
   - **Cách 1: Google OAuth 2.0 (Khuyên dùng - Không cần mật khẩu):**
     1. Tải file `credentials.json` (OAuth 2.0 Client ID loại Desktop App từ Google Cloud Console) vào thư mục `Documents` (`C:\Users\Admin\Documents\credentials.json`).
     2. Chạy script xác thực trình duyệt một lần duy nhất:
        ```powershell
        powershell -ExecutionPolicy Bypass -File scripts/authorize_gmail.ps1
        ```
     3. Trình duyệt tự mở, bạn bấm chọn tài khoản Gmail và nhấn **Cho phép (Allow)**. File `token.json` sẽ tự động lưu vào `Documents`. Từ các lần sau, script tự động gửi mail qua Gmail REST API mà **không cần nhập mật khẩu bao giờ nữa**.
   - **Cách 2: SMTP Mật khẩu ứng dụng (App Password):**
     Đặt file `email_config.json` vào thư mục **`Documents`** hoặc thư mục `scripts/` (xem mẫu `scripts/email_config.example.json`):
     ```json
     {
       "smtp_server": "smtp.gmail.com",
       "smtp_port": 587,
       "sender_email": "thaytoan.mathplus@gmail.com",
       "sender_password": "xxxx xxxx xxxx xxxx",
       "sender_display_name": "Thầy giáo Toán - MathPlus"
     }
     ```
3. Chia sẻ bảng tính Google Sheets cho email của Service Account (`client_email`) với quyền **Editor (Người chỉnh sửa)**.
4. Kích hoạt **Google Sheets API**, **Google Drive API** và **Gmail API** trên Google Cloud Console.

### 2. Chạy kịch bản tự động cập nhật
Chạy lệnh PowerShell (script tự động nhận Service Account và cấu hình Email từ `Documents`):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/update_feedback.ps1 `
    -SpreadsheetId "1FJixJmSa8vTH1ub8yiue3cJrUh1igf3NpCs5Q4M1w78" `
    -SheetName "Tháng 09/2026" `
    -SheetGid 1490594752 `
    -StartRowIndex 15 `
    -EndRowIndex 21 `
    -CommentColumnIndex 8
```

*Ghi chú:* 
- Script tự động quét tìm cột có tiêu đề `Email` hoặc `Email phụ huynh` trên dòng tiêu đề (mặc định dòng 15).
- Bạn cũng có thể chỉ định trực tiếp vị trí cột qua `-EmailColumnIndex <số_thứ_tự_0_indexed>`.
- Nếu chưa có file `email_config.json`, script sẽ cập nhật nhận xét lên Sheet và thông báo bỏ qua bước gửi mail.

## Workflow

### 1. Xác thực Google Cloud Service Account
- Sử dụng thuật toán `RS256` mã hóa RSA-SHA256 để ký chữ ký số trên JWT Claim Set.
- Gửi yêu cầu POST tới endpoint `https://oauth2.googleapis.com/token` để nhận `access_token` hợp lệ trong 1 giờ.

### 2. Thu thập điểm BTVN & Quét cột Email
- Quét dòng tiêu đề tìm cột `Email phụ huynh` (nếu chưa chỉ định).
- Gọi Google Sheets API v4 qua endpoint `GET /v4/spreadsheets/{spreadsheetId}/values/{range}` để đọc đồng thời Tên học sinh, Điểm BTVN, và Email phụ huynh.

### 3. Phân tích và sinh nhận xét sư phạm
- Mỗi học sinh được chọn ngẫu nhiên số lượng dòng từ **3 đến 6 dòng** (hàm `Get-Random -Minimum 3 -Maximum 7`).
- Tổng hợp từ ngân hàng nhận xét chuẩn mực sư phạm theo thang điểm.

### 4. Ghi dữ liệu và định dạng ô
- Gửi dữ liệu bằng phương thức `PUT /v4/spreadsheets/{spreadsheetId}/values/{range}?valueInputOption=USER_ENTERED`.
- Gọi endpoint `POST /v4/spreadsheets/{spreadsheetId}:batchUpdate` cấu hình Wrap Text và tự động căn chỉnh chiều cao dòng (Auto Row Height).

### 5. Gửi Email thông báo cho phụ huynh
- Kiểm tra email từng học sinh:
  - **Không có email:** Tự động bỏ qua và ghi log rõ ràng.
  - **Có email:** Soạn nội dung email trang trọng chuẩn mực sư phạm (gồm HTML và Plain Text) theo định dạng:
    > **Kính gửi phụ huynh bạn [Tên],**  
    > Thầy xin gửi tới quý phụ huynh thông tin kết quả bài tập về nhà và nhận xét đánh giá của con trong buổi học vừa qua:  
    > - **Điểm bài tập về nhà:** [Điểm] / 10  
    > - **Chi tiết nhận xét từ thầy:**  
    > [Nội dung nhận xét gạch đầu dòng]  
    >   
    > Thầy rất mong quý phụ huynh cùng đồng hành, nhắc nhở và động viên con để con ngày càng tiến bộ hơn nữa...  
    > Trân trọng,  
    > **[Tên giáo viên / Trung tâm]**
- Gửi thư trực tiếp qua giao thức bảo mật SMTP SSL (.NET `System.Net.Mail`).

## Common Mistakes & Best Practices
1. **Thiếu quyền chia sẻ:** Luôn đảm bảo đã chia sẻ file Google Sheets cho email của Service Account (dạng `...@...iam.gserviceaccount.com`).
2. **Ký tự tiếng Việt trên PowerShell 5.1:** Khi gửi chuỗi UTF-8 qua WebClient hoặc Invoke-RestMethod, cần chỉ định rõ `Content-Type: application/json; charset=utf-8` và mã hóa UTF-8 bytes.
3. **Mã hóa URL Range:** Tên sheet có khoảng trắng hoặc dấu gạch chéo (ví dụ `'Tháng 09/2026'!A1:Z50`) bắt buộc phải được URL-encode bằng `[System.Uri]::EscapeDataString()`.
