---
name: google-sheets-teacher-feedback
description: >-
  Tự động kết nối Google Sheets bằng Service Account (PowerShell/.NET), đọc điểm số BTVN của học sinh,
  sinh nhận xét sư phạm ngẫu nhiên từ 3 đến 6 gạch đầu dòng mang tính khích lệ,
  cập nhật trực tiếp vào Google Sheets API kèm định dạng tự động xuống dòng (Wrap Text) và co giãn chiều cao (Auto Height).
---

# Google Sheets Teacher Feedback Generator

## Overview
Kỹ năng tự động hóa quy trình quản lý điểm số và tạo nhận xét học sinh cho giáo viên/trung tâm dạy học trên **Google Sheets**:
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

## Quick Start

### 1. Chuẩn bị thông tin xác thực
1. Đặt file khóa JSON của Service Account vào thư mục dự án (ví dụ `service_account.json`).
2. Chia sẻ bảng tính Google Sheets cho email của Service Account với quyền **Editor (Người chỉnh sửa)**.
3. Kích hoạt **Google Sheets API** và **Google Drive API** trên Google Cloud Console.

### 2. Chạy kịch bản tự động cập nhật
Chạy lệnh PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/update_feedback.ps1 `
    -ServiceAccountPath "service_account.json" `
    -SpreadsheetId "1FJixJmSa8vTH1ub8yiue3cJrUh1igf3NpCs5Q4M1w78" `
    -SheetName "Tháng 09/2026" `
    -SheetGid 1490594752 `
    -StartRowIndex 15 `
    -EndRowIndex 21 `
    -CommentColumnIndex 8
```

## Workflow

### 1. Xác thực Google Cloud Service Account
- Sử dụng thuật toán `RS256` mã hóa RSA-SHA256 để ký chữ ký số trên JWT Claim Set.
- Gửi yêu cầu POST tới endpoint `https://oauth2.googleapis.com/token` để nhận `access_token` hợp lệ trong 1 giờ.

### 2. Thu thập điểm BTVN
- Gọi Google Sheets API v4 qua endpoint `GET /v4/spreadsheets/{spreadsheetId}/values/{range}`.
- Trích xuất tên học sinh và điểm số ở cột `Điểm BTVN` (xử lý cả dấu chấm `.` và dấu phẩy `,`).

### 3. Phân tích và sinh nhận xét sư phạm
- Mỗi học sinh được chọn ngẫu nhiên số lượng dòng từ **3 đến 6 dòng** (hàm `Get-Random -Minimum 3 -Maximum 7`).
- Tổng hợp từ ngân hàng nhận xét bao gồm:
  - Ý thức làm bài / tinh thần tự giác.
  - Mức độ tiếp thu kiến thức và tư duy toán học.
  - Lưu ý tính toán, soát bài, cẩn thận.
  - Lời khích lệ, động viên của thầy giáo.

### 4. Ghi dữ liệu và định dạng ô
- Gửi dữ liệu bằng phương thức `PUT /v4/spreadsheets/{spreadsheetId}/values/{range}?valueInputOption=USER_ENTERED`.
- Gọi endpoint `POST /v4/spreadsheets/{spreadsheetId}:batchUpdate`:
  - `repeatCell`: Cấu hình `userEnteredFormat.wrapStrategy = "WRAP"`.
  - `autoResizeDimensions`: Tự động co giãn chiều cao các hàng (`dimension = "ROWS"`).

## Common Mistakes & Best Practices
1. **Thiếu quyền chia sẻ:** Luôn đảm bảo đã chia sẻ file Google Sheets cho email của Service Account (dạng `...@...iam.gserviceaccount.com`).
2. **Ký tự tiếng Việt trên PowerShell 5.1:** Khi gửi chuỗi UTF-8 qua WebClient hoặc Invoke-RestMethod, cần chỉ định rõ `Content-Type: application/json; charset=utf-8` và mã hóa UTF-8 bytes.
3. **Mã hóa URL Range:** Tên sheet có khoảng trắng hoặc dấu gạch chéo (ví dụ `'Tháng 09/2026'!A1:Z50`) bắt buộc phải được URL-encode bằng `[System.Uri]::EscapeDataString()`.
