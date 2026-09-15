<#
.SYNOPSIS
    Xác thực Google OAuth 2.0 một lần duy nhất qua trình duyệt để cấp quyền gửi email Gmail.

.DESCRIPTION
    Script mở trình duyệt đăng nhập tài khoản Google của bạn, nhận mã ủy quyền (Authorization Code)
    và lưu Refresh Token vào C:\Users\Admin\Documents\token.json.
    Không yêu cầu lưu mật khẩu, tuân thủ tiêu chuẩn bảo mật Google OAuth 2.0.
#>

[CmdletBinding()]
param (
    [string]$CredentialsPath = "credentials.json",
    [string]$TokenOutputPath = "token.json",
    [int]$Port = 8080
)

$ErrorActionPreference = "Stop"

# 1. Tìm file credentials.json (hoặc client_secret_*.json)
$credFile = $null
if (Test-Path $CredentialsPath) {
    $credFile = $CredentialsPath
} else {
    $docPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) $CredentialsPath
    if (Test-Path $docPath) {
        $credFile = $docPath
    } else {
        $found = Get-ChildItem -Path ([Environment]::GetFolderPath('MyDocuments')) -Filter "*client_secret*.json" -File | Select-Object -First 1
        if (-not $found) {
            $found = Get-ChildItem -Path ([Environment]::GetFolderPath('MyDocuments')) -Filter "*credential*.json" -File | Select-Object -First 1
        }
        if ($found) {
            $credFile = $found.FullName
        }
    }
}

if (-not $credFile) {
    Write-Host @"
[LỖI] Không tìm thấy file OAuth Client Credentials (credentials.json hoặc client_secret_*.json)!

HƯỚNG DẪN TẠO FILE OAUTH 2.0:
1. Truy cập Google Cloud Console: https://console.cloud.google.com/apis/credentials?project=mathplus-508610
2. Nhấn '+ CREATE CREDENTIALS' (Tạo thông tin xác thực) -> Chọn 'OAuth client ID'.
3. Chọn Application type (Loại ứng dụng): 'Desktop app' (Ứng dụng cho máy tính).
4. Nhấn 'CREATE' (Tạo) -> Tải file JSON về máy.
5. Lưu file vào thư mục Documents: C:\Users\Admin\Documents\credentials.json
6. Chạy lại script này: powershell -File scripts/authorize_gmail.ps1
"@ -ForegroundColor Yellow
    exit 1
}

Write-Host "[INFO] Đang nạp thông tin client từ: $credFile"
$credJson = Get-Content $credFile -Raw -Encoding UTF8 | ConvertFrom-Json
$clientData = if ($credJson.installed) { $credJson.installed } else { $credJson.web }
if (-not $clientData) {
    throw "Cấu trúc file credentials.json không hợp lệ (cần trường 'installed' hoặc 'web')."
}

$clientId = $clientData.client_id
$clientSecret = $clientData.client_secret
$redirectUri = "http://127.0.0.1:$Port/callback/"

# 2. Khởi tạo Local HTTP Listener để đón mã xác thực
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$Port/callback/")
try {
    $listener.Start()
} catch {
    throw "Không thể lắng nghe trên cổng $Port. Hãy đóng ứng dụng khác đang chiếm cổng hoặc chỉ định -Port khác: $_"
}

# 3. Tạo link đăng nhập Google OAuth 2.0
$scope = [System.Uri]::EscapeDataString("https://www.googleapis.com/auth/gmail.send https://www.googleapis.com/auth/userinfo.email")
$authUrl = "https://accounts.google.com/o/oauth2/v2/auth?client_id=$clientId&redirect_uri=$([System.Uri]::EscapeDataString($redirectUri))&response_type=code&scope=$scope&access_type=offline&prompt=consent"

Write-Host "`n[BƯỚC 1] Đang mở trình duyệt để đăng nhập tài khoản Google..." -ForegroundColor Cyan
Write-Host "Nếu trình duyệt không tự mở, hãy truy cập link sau:`n$authUrl`n"
Start-Process $authUrl

Write-Host "[BƯỚC 2] Đang chờ bạn bấm 'Cho phép' trên trình duyệt..." -ForegroundColor Yellow
$context = $listener.GetContext()
$request = $context.Request
$response = $context.Response

$authCode = $request.QueryString["code"]
$authError = $request.QueryString["error"]

# Trả về trang HTML thông báo trên trình duyệt
$htmlMessage = if ($authCode) {
    @"
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Xác thực thành công</title></head>
<body style="font-family: Arial, sans-serif; text-align: center; padding-top: 50px;">
    <h2 style="color: #2e7d32;">🎉 Xác thực Google OAuth thành công!</h2>
    <p>MathPlus đã nhận được mã ủy quyền từ Google. Bạn có thể đóng tab trình duyệt này lại.</p>
</body>
</html>
"@
} else {
    @"
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Xác thực thất bại</title></head>
<body style="font-family: Arial, sans-serif; text-align: center; padding-top: 50px;">
    <h2 style="color: #d32f2f;">❌ Xác thực thất bại</h2>
    <p>Lỗi: $authError</p>
</body>
</html>
"@
}

$buffer = [System.Text.Encoding]::UTF8.GetBytes($htmlMessage)
$response.ContentType = "text/html; charset=utf-8"
$response.ContentLength64 = $buffer.Length
$response.OutputStream.Write($buffer, 0, $buffer.Length)
$response.OutputStream.Close()
$listener.Stop()

if (-not $authCode) {
    Write-Error "Không nhận được mã xác thực. Lỗi: $authError"
    exit 1
}

Write-Host "[BƯỚC 3] Đã nhận Authorization Code. Đang xin cấp Token từ Google..." -ForegroundColor Cyan

# 4. Trao đổi Authorization Code lấy Access Token & Refresh Token
$tokenBody = @{
    code = $authCode
    client_id = $clientId
    client_secret = $clientSecret
    redirect_uri = $redirectUri
    grant_type = "authorization_code"
}

$tokenRes = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $tokenBody -ContentType "application/x-www-form-urlencoded"

# Lấy thông tin email của người dùng
$userEmail = ""
try {
    $userInfo = Invoke-RestMethod -Uri "https://www.googleapis.com/oauth2/v2/userinfo" -Headers @{ Authorization = "Bearer $($tokenRes.access_token)" }
    $userEmail = $userInfo.email
} catch { }

# 5. Lưu token vào Documents
$finalTokenPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) $TokenOutputPath
$tokenData = @{
    client_id = $clientId
    client_secret = $clientSecret
    email = $userEmail
    access_token = $tokenRes.access_token
    refresh_token = $tokenRes.refresh_token
    token_type = $tokenRes.token_type
    expires_at = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() + $tokenRes.expires_in
} | ConvertTo-Json -Depth 5

[System.IO.File]::WriteAllText($finalTokenPath, $tokenData, [System.Text.Encoding]::UTF8)

Write-Host "`n[THÀNH CÔNG RỰC RỠ] 🎉🎉🎉" -ForegroundColor Green
Write-Host "Đã lưu thông tin xác thực OAuth 2.0 tại: $finalTokenPath" -ForegroundColor Green
if ($userEmail) {
    Write-Host "Tài khoản Gmail được ủy quyền: $userEmail" -ForegroundColor Green
}
Write-Host "Từ bây giờ, hệ thống sẽ tự động gửi email cho phụ huynh qua Gmail API mà KHÔNG BAO GIỜ cần nhập mật khẩu!`n" -ForegroundColor Green