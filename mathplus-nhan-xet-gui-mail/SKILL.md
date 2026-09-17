---
name: mathplus-nhan-xet-gui-mail
description: "Tự động đọc bảng điểm học sinh MathPlus trên Google Sheets, dựa theo cột 'Điểm BTVN' để sinh nhận xét sư phạm vào cột 'Điểm trên lớp + Nhận xét' (tự động giữ nguyên nếu đã có sẵn nhận xét của thầy cô, không tạo lại), cập nhật Google Sheets và soạn/gửi email thông báo kết quả học tập cho phụ huynh (BẮT BUỘC Human can thiệp phê duyệt trước khi gửi email)."
---

# MathPlus: Nhận Xét BTVN & Gửi Email Phụ Huynh

## Tổng quan
Skill này tự động hóa quy trình quản lý chất lượng học tập tại **Trung tâm Giáo dục MathPlus Academy**:
1. **Đọc bảng điểm Google Sheets**: Hỗ trợ link Google Sheets có chứa `gid` hoặc chỉ định sheet cụ thể (ví dụ: `Tháng 09/2026`). Tự động nhận diện phân khối các lớp (`Lớp 2`, `Lớp 3`, `Lớp 4`, `Lớp 5`) và các buổi học theo ngày (`05/09`, `12/09`, `19/09`, `26/09`...).
2. **Sinh nhận xét sư phạm theo Điểm BTVN**:
   - **Bảo toàn nhận xét của thầy cô**: Nếu ô `"Điểm trên lớp + Nhận xét"` đã có sẵn nhận xét của thầy/cô, hệ thống **tự động giữ nguyên và không tạo nhận xét cho dòng đó nữa** (bỏ qua bước sinh mới, trừ khi truyền cờ `--overwrite`).
   - **Tạo mới với các ô còn trống**: Dựa vào thang điểm BTVN (Xuất sắc >=9.5, Giỏi 8.0-9.0, Khá 7.0-7.5, Trung bình 5.0-6.5, hoặc Chưa nộp/0/thiếu), tự động sinh nhận xét gồm 4 gạch đầu dòng chuẩn mực sư phạm (ý thức chuẩn bị bài, năng lực tiếp thu, điểm cần cải thiện và lời động viên).
3. **Cập nhật Google Sheet**: Chỉ ghi nhận xét mới vào đúng hàng và cột `"Điểm trên lớp + Nhận xét"` của các học sinh chưa có nhận xét (không đè lên nhận xét sẵn có của giáo viên).
4. **Soạn thảo email thông báo phụ huynh**: Sử dụng template HTML chuyên nghiệp mang nhận diện thương hiệu MathPlus Academy (xanh lá `#2e5311`, logo, bảng điểm, nhận xét chi tiết gồm cả nhận xét viết tay của thầy cô hoặc nhận xét sinh tự động, lời nhắn gửi từ trung tâm).
5. ⚠️ **CHỐT CHẶN BẮT BUỘC (Human-in-the-loop)**:
   - **Tuyệt đối KHÔNG ĐƯỢC PHÉP tự động gửi email** mà không có xác nhận từ người dùng.
   - Script và Agent **phải dừng lại**, hiển thị bảng preview chi tiết (Học sinh, Lớp, Điểm, Nhận xét, Email nhận) và **yêu cầu người dùng duyệt/xác nhận**.
   - Chỉ khi người dùng phản hồi đồng ý ("Gửi", "OK", "Xác nhận"), Agent mới thực thi lệnh gửi với cờ `--confirm`.


---

## Hướng dẫn Vận hành cho AI Agent

Khi nhận được yêu cầu từ người dùng (ví dụ: *"Nhận xét và gửi mail cho link sheet https://docs.google.com/spreadsheets/d/.../edit?gid=1490594752"*):

### Bước 1: Xem trước & Sinh nhận xét (Preview)
Chạy script trợ lý ở chế độ `preview`:
```bash
python mathplus_sheet_assistant.py \
  --sheet-url "<GOOGLE_SHEET_URL>" \
  --action preview \
  --export-preview mathplus_preview.html
```
*(Có thể truyền thêm `--date "12/09"` nếu muốn xử lý một ngày học cụ thể; nếu không truyền, script tự động chọn ngày gần nhất có điểm BTVN).*

### Bước 2: Cập nhật nhận xét vào Google Sheet
Sau khi kiểm tra danh sách nhận xét được tạo, chạy lệnh cập nhật Google Sheet:
```bash
python mathplus_sheet_assistant.py \
  --sheet-url "<GOOGLE_SHEET_URL>" \
  --action update-sheet
```
*(Nếu muốn ghi đè các ô nhận xét đã có sẵn, thêm cờ `--overwrite`).*

### Bước 3: Trình bày Báo cáo Preview & DỪNG LẠI YÊU CẦU HUMAN DUYỆT
Agent **BẮT BUỘC** hiển thị bảng tóm tắt kết quả cho người dùng ngay trong hội thoại:
- Bảng danh sách học sinh:
  | STT | Học sinh | Lớp | Email phụ huynh | Điểm BTVN | Trích đoạn nhận xét | Trạng thái gửi |
  | --- | --- | --- | --- | --- | --- | --- |
- Trình bày mẫu email xem trước (hoặc đường dẫn `mathplus_preview.html`).
- **Đặt câu hỏi xác nhận rõ ràng**:
  > *"Tôi đã cập nhật đầy đủ nhận xét vào Google Sheet và chuẩn bị danh sách email thông báo kết quả buổi học cho Quý Phụ huynh như bảng trên. Xin bạn vui lòng kiểm tra và xác nhận: Bạn có đồng ý gửi email thông báo tới các phụ huynh trên không?"*

### Bước 4: Thực hiện gửi Email (Chỉ khi Human đồng ý)
- **Nếu người dùng xác nhận "Đồng ý" / "Gửi"**:
  Chạy lệnh gửi email kèm cờ an toàn `--confirm`:
  ```bash
  python mathplus_sheet_assistant.py \
    --sheet-url "<GOOGLE_SHEET_URL>" \
    --action send-email \
    --confirm
  ```
- **Nếu người dùng yêu cầu chỉnh sửa**:
  Điều chỉnh nhận xét theo ý người dùng rồi chạy lại cập nhật Google Sheet trước khi hỏi lại.
- **Nếu người dùng từ chối / chưa muốn gửi**:
  Dừng lại, lưu trữ danh sách nhận xét trên Google Sheet và kết thúc an toàn.

---

## Các Cơ Chế Gửi Email (Hỗ trợ KHÔNG CẦN NHẬP MẬT KHẨU)

Hệ thống hỗ trợ 4 cơ chế gửi qua tham số `--mailer`:

### 1. Cơ chế Windows Outlook (`--mailer outlook`) - Zero-Password trên Windows
- Tự động tận dụng tài khoản Microsoft Outlook đã đăng nhập sẵn trên máy tính Windows.
- Không cần khai báo mật khẩu hay app password.
- Thêm cờ `--open-drafts` nếu muốn mở cửa sổ soạn thảo Outlook để người dùng tự duyệt và bấm gửi.

### 2. Cơ chế macOS Mail (`--mailer macos-mail`) - Zero-Password trên macOS
- Tự động tận dụng tài khoản Google hoặc iCloud đã đăng nhập sẵn trong ứng dụng **Mail** của macOS.
- Thêm cờ `--open-drafts` nếu muốn hiển thị nháp trực tiếp trên màn hình macOS Mail.

### 3. Cơ chế Google Service Account (`--mailer service-account`)
- Dành cho tổ chức sử dụng **Google Workspace** (tên miền `@mathplus.vn`).
- Quản trị viên bật tính năng **Domain-Wide Delegation (DWD)** trong Google Admin Console cho Service Account với scope `https://www.googleapis.com/auth/gmail.send`.
- Khi đó Service Account có thể tự động gửi mail đại diện cho email phòng ban (ví dụ `cskh@mathplus.vn`):
  ```bash
  python mathplus_sheet_assistant.py \
    --sheet-url "<GOOGLE_SHEET_URL>" \
    --action send-email \
    --mailer service-account \
    --impersonate-user "cskh@mathplus.vn" \
    --confirm
  ```

### 4. Cơ chế SMTP Truyền Thống (`--mailer smtp`)
- Sử dụng file cấu hình `config.json` (hoặc `~/Documents/email_config.json`, biến môi trường `SMTP_USER` / `SMTP_PASS`) với Google App Password (16 ký tự).
- Xem mẫu tại `config.example.json`.

---

## Cấu hình Google Sheets API
Script tự động tìm file `credentials.json` của Google Service Account tại các vị trí:
- Cùng thư mục với script: `./credentials.json`
- Thư mục cấu hình toàn cục: `~/.gemini/config/skills/mathplus-nhan-xet-gui-mail/credentials.json`
- Thư mục người dùng: `~/Documents/credentials.json` hoặc `~/Documents/service_account.json`
Tài khoản Service Account cần được share quyền Editor vào Google Sheet cần xử lý.


