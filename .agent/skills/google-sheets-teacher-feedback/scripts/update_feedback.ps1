<#
.SYNOPSIS
    Tự động đọc điểm BTVN từ Google Sheets, sinh nhận xét học sinh (3-6 gạch đầu dòng ngẫu nhiên),
    ghi vào cột nhận xét, định dạng text wrap & auto height, và gửi email nhận xét cho phụ huynh.

.DESCRIPTION
    Script sử dụng Google Service Account để xác thực qua JWT Bearer token và tương tác với Google Sheets API v4.
    Sau khi ghi nhận xét, script tự động quét cột Email phụ huynh và gửi email báo cáo tình hình học tập qua SMTP (.NET).
    Không yêu cầu cài đặt Python hay thư viện ngoài, chạy hoàn toàn bằng PowerShell và .NET Framework có sẵn trên Windows.

.PARAMETER ServiceAccountPath
    Đường dẫn đến file JSON của Service Account. Mặc định tự động tìm trong thư mục hiện tại hoặc Documents.

.PARAMETER SpreadsheetId
    ID của Google Spreadsheet cần cập nhật.

.PARAMETER SheetName
    Tên của Sheet (tab) cần thao tác (ví dụ: "Tháng 09/2026").

.PARAMETER TargetRange
    Dải ô nhận xét cần cập nhật (ví dụ: "I16:I21").

.PARAMETER ScoreRange
    Dải ô điểm BTVN tương ứng (ví dụ: "H16:H21").

.PARAMETER EmailColumnIndex
    Chỉ số cột (0-indexed) chứa Email phụ huynh. Mặc định là -1 (tự động quét dòng tiêu đề).

.PARAMETER EmailConfigPath
    Đường dẫn đến file cấu hình SMTP JSON. Mặc định tự động tìm trong thư mục hiện tại hoặc Documents.
#>

[CmdletBinding()]
param (
    [string]$ServiceAccountPath = "service_account.json",
    [string]$SpreadsheetId = "1FJixJmSa8vTH1ub8yiue3cJrUh1igf3NpCs5Q4M1w78",
    [string]$SheetName = "Tháng 09/2026",
    [int]$SheetGid = 1490594752,
    [int]$StartRowIndex = 15,     # 0-indexed (Row 16)
    [int]$EndRowIndex = 21,       # 0-indexed exclusive (Row 21)
    [int]$CommentColumnIndex = 8, # 0-indexed (Column I: Điểm trên lớp + Nhận xét)
    [int]$ScoreColumnIndex = 7,   # 0-indexed (Column H: Điểm BTVN)
    [int]$NameColumnIndex = 1,    # 0-indexed (Column B: Tên)
    [int]$HeaderRowIndex = 14,    # 0-indexed (Row 15: Dòng tiêu đề cột)
    [int]$EmailColumnIndex = -1,  # 0-indexed (-1 để tự động dò tìm cột "Email")
    [string]$EmailColumnHeader = "Email",
    # Cấu hình gửi Mail qua SMTP
    [string]$EmailConfigPath = "email_config.json",
    [string]$SmtpServer = "smtp.gmail.com",
    [int]$SmtpPort = 587,
    [string]$SenderEmail = "",
    [string]$SenderPassword = "",
    [string]$SenderDisplayName = "Thầy giáo Toán - MathPlus"
)

$ErrorActionPreference = "Stop"

function Base64UrlEncode([byte[]]$bytes) {
    return [System.Convert]::ToBase64String($bytes).Replace('+', '-').Replace('/', '_').TrimEnd('=')
}

function Base64UrlEncodeString([string]$str) {
    return Base64UrlEncode([System.Text.Encoding]::UTF8.GetBytes($str))
}

function Get-ColumnLetter([int]$colIndex) {
    $dividend = $colIndex + 1
    $colName = ""
    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $colName = [char](65 + $modulo) + $colName
        $dividend = [int](($dividend - $modulo) / 26)
    }
    return $colName
}

# 1. Kiểm tra file Service Account
if (-not (Test-Path $ServiceAccountPath)) {
    $docPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) $ServiceAccountPath
    if (Test-Path $docPath) {
        $ServiceAccountPath = $docPath
    } else {
        $saInDocs = Get-ChildItem -Path ([Environment]::GetFolderPath('MyDocuments')) -Filter "*.json" -File | Where-Object {
            try {
                $content = Get-Content $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                return ($content.type -eq "service_account")
            } catch { return $false }
        } | Select-Object -First 1

        if ($saInDocs) {
            $ServiceAccountPath = $saInDocs.FullName
            Write-Host "[INFO] Tự động nhận diện Service Account tại Documents: $ServiceAccountPath"
        } else {
            throw "Không tìm thấy file Service Account tại '$ServiceAccountPath' hoặc trong thư mục Documents!"
        }
    }
}

$sa = Get-Content $ServiceAccountPath -Raw -Encoding UTF8 | ConvertFrom-Json
$clientEmail = $sa.client_email
Write-Host "[INFO] Đang xác thực với Service Account: $clientEmail"

# 2. Tạo JWT và xin cấp Access Token từ Google OAuth2
$pk = $sa.private_key
$pkClean = $pk -replace '-----BEGIN PRIVATE KEY-----', '' -replace '-----END PRIVATE KEY-----', '' -replace '\s+', ''
$keyBytes = [System.Convert]::FromBase64String($pkClean)

$cngKey = [System.Security.Cryptography.CngKey]::Import($keyBytes, [System.Security.Cryptography.CngKeyBlobFormat]::Pkcs8PrivateBlob)
$rsa = New-Object System.Security.Cryptography.RSACng($cngKey)

$header = @{ alg = "RS256"; typ = "JWT" } | ConvertTo-Json -Compress
$now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$payload = @{
    iss = $clientEmail
    scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive"
    aud = "https://oauth2.googleapis.com/token"
    exp = $now + 3600
    iat = $now
} | ConvertTo-Json -Compress

$headerB64 = Base64UrlEncodeString $header
$payloadB64 = Base64UrlEncodeString $payload
$dataToSign = [System.Text.Encoding]::UTF8.GetBytes("$headerB64.$payloadB64")
$sig = $rsa.SignData($dataToSign, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
$sigB64 = Base64UrlEncode $sig
$jwt = "$headerB64.$payloadB64.$sigB64"

$tokenRes = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body @{
    grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
    assertion = $jwt
} -ContentType "application/x-www-form-urlencoded"

$token = $tokenRes.access_token
Write-Host "[SUCCESS] Lấy token xác thực Google Sheets thành công!"

# 3. Nạp cấu hình gửi Email (OAuth 2.0 Gmail Token hoặc SMTP)
$gmailOAuthReady = $false
$gmailAccessToken = $null
$gmailSenderEmail = $null

$tokenFile = $null
$docTokenPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "token.json"
if (Test-Path "token.json") {
    $tokenFile = "token.json"
} elseif (Test-Path $docTokenPath) {
    $tokenFile = $docTokenPath
}

if ($tokenFile) {
    try {
        $tok = Get-Content $tokenFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $nowEpoch = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        if ($tok.refresh_token -and ($tok.expires_at -lt ($nowEpoch + 60) -or -not $tok.access_token)) {
            Write-Host "[INFO] Đang làm mới Google OAuth 2.0 Access Token..."
            $refreshRes = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body @{
                client_id = $tok.client_id
                client_secret = $tok.client_secret
                refresh_token = $tok.refresh_token
                grant_type = "refresh_token"
            } -ContentType "application/x-www-form-urlencoded"

            $tok.access_token = $refreshRes.access_token
            $tok.expires_at = $nowEpoch + $refreshRes.expires_in
            [System.IO.File]::WriteAllText($tokenFile, ($tok | ConvertTo-Json -Depth 5), [System.Text.Encoding]::UTF8)
        }

        $gmailAccessToken = $tok.access_token
        $gmailSenderEmail = $tok.email
        $gmailOAuthReady = $true
        Write-Host "[INFO] Đã xác thực Google OAuth 2.0 cho tài khoản: $gmailSenderEmail (Không cần mật khẩu)" -ForegroundColor Green
    } catch {
        Write-Warning "Không làm mới được token OAuth 2.0: $_"
    }
}

$emailConfigFile = $null
if (-not $gmailOAuthReady) {
if (Test-Path $EmailConfigPath) {
    $emailConfigFile = $EmailConfigPath
} else {
    $docEmailPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) $EmailConfigPath
    if (Test-Path $docEmailPath) {
        $emailConfigFile = $docEmailPath
    } else {
        $foundConfig = Get-ChildItem -Path ([Environment]::GetFolderPath('MyDocuments')) -Filter "*email*.json" -File | Where-Object {
            try {
                $c = Get-Content $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                return ($c.sender_email -ne $null -or $c.smtp_server -ne $null)
            } catch { return $false }
        } | Select-Object -First 1
        if ($foundConfig) {
            $emailConfigFile = $foundConfig.FullName
        }
    }
}

if ($emailConfigFile) {
    try {
        $cfg = Get-Content $emailConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $SenderEmail -and $cfg.sender_email) { $SenderEmail = $cfg.sender_email }
        if (-not $SenderPassword -and $cfg.sender_password) { $SenderPassword = $cfg.sender_password }
        if ($cfg.smtp_server) { $SmtpServer = $cfg.smtp_server }
        if ($cfg.smtp_port) { $SmtpPort = [int]$cfg.smtp_port }
        if ($cfg.sender_display_name) { $SenderDisplayName = $cfg.sender_display_name }
        Write-Host "[INFO] Đã nạp cấu hình gửi mail từ: $emailConfigFile ($SenderEmail)"
    } catch {
        Write-Warning "Không đọc được file cấu hình email: $_"
    }
}
}

# 4. Tự động nhận diện cột Email nếu $EmailColumnIndex = -1
if ($EmailColumnIndex -lt 0) {
    Write-Host "[INFO] Đang dò tìm cột Email phụ huynh..."
    # 4.1 Thử quét dòng tiêu đề lớp ($HeaderRowIndex) và dòng tiêu đề chung (Dòng 2)
    $rowsToScan = @($HeaderRowIndex + 1, 2)
    foreach ($rNum in $rowsToScan) {
        $headerRangeA1 = [System.Uri]::EscapeDataString("'$SheetName'!$rNum`:$rNum")
        $headerUrl = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/$headerRangeA1"
        try {
            $headerData = Invoke-RestMethod -Uri $headerUrl -Method Get -Headers @{ Authorization = "Bearer $token" }
            if ($headerData.values -and $headerData.values.Count -gt 0) {
                $headers = $headerData.values[0]
                for ($colIdx = 0; $colIdx -lt $headers.Count; $colIdx++) {
                    $hText = "$($headers[$colIdx])".Trim()
                    if ($hText -match "(?i)email|thư điện tử") {
                        $EmailColumnIndex = $colIdx
                        Write-Host "[INFO] Tự động phát hiện cột '$hText' tại cột $(Get-ColumnLetter $colIdx) (Dòng $rNum, index $colIdx)"
                        break
                    }
                }
            }
        } catch { }
        if ($EmailColumnIndex -ge 0) { break }
    }

    # 4.2 Nếu tiêu đề không ghi rõ, quét các ô dữ liệu học sinh xem cột nào chứa ký tự @email
    if ($EmailColumnIndex -lt 0) {
        try {
            $sampleRangeA1 = [System.Uri]::EscapeDataString("'$SheetName'!A$($StartRowIndex + 1):Z$EndRowIndex")
            $sampleUrl = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/$sampleRangeA1"
            $sampleData = Invoke-RestMethod -Uri $sampleUrl -Method Get -Headers @{ Authorization = "Bearer $token" }
            if ($sampleData.values) {
                foreach ($sRow in $sampleData.values) {
                    for ($c = 0; $c -lt $sRow.Count; $c++) {
                        if ("$($sRow[$c])" -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                            $EmailColumnIndex = $c
                            Write-Host "[INFO] Tự động nhận diện cột Email qua dữ liệu tại cột $(Get-ColumnLetter $c) (index $c)"
                            break
                        }
                    }
                    if ($EmailColumnIndex -ge 0) { break }
                }
            }
        } catch { }
    }

    if ($EmailColumnIndex -lt 0) {
        Write-Host "[INFO] Bảng tính hiện tại chưa có cột '$EmailColumnHeader' (script sẽ chỉ cập nhật nhận xét và ghi nhận không có email)."
    }
}

# 5. Hàm tạo nhận xét ngẫu nhiên 3-6 gạch đầu dòng
function Generate-TeacherFeedback([double]$score, [string]$name) {
    $lineCount = Get-Random -Minimum 3 -Maximum 7

    $bankIntro = @(
        "Con có ý thức tự giác làm bài tập về nhà rất tốt.",
        "Con nộp bài đúng hạn và hoàn thành đầy đủ bài tập được giao.",
        "Ý thức chuẩn bị bài và làm BTVN ở nhà đáng khen ngợi.",
        "Con chăm chỉ làm bài tập, bài vở trình bày tương đối sạch đẹp."
    )

    if ($score -lt 5.0) {
        $bankIntro = @(
            "Con cần chú ý nộp bài đầy đủ và đúng hạn hơn.",
            "Ý thức tự giác làm bài tập về nhà của con cần được cải thiện.",
            "Con cần tập trung hơn khi làm bài tập ở nhà."
        )
        $bankKnowledge = @(
            "Con cần dành nhiều thời gian hơn để ôn lại các kiến thức cơ bản trên lớp.",
            "Chưa nắm chắc các dạng bài cơ bản, còn nhầm lẫn nhiều công thức.",
            "Kỹ năng giải toán còn yếu, cần rèn luyện thêm từ các bài toán đơn giản.",
            "Con cần xem kỹ lại lý thuyết trước khi bắt tay vào làm bài."
        )
        $bankCaution = @(
            "Các phép tính còn nhiều sai sót, con cần tính toán cẩn thận từng bước.",
            "Cần chú ý đọc kỹ yêu cầu của đề bài để tránh lạc đề.",
            "Khi làm bài con cần nháp cẩn thận và kiểm tra lại kết quả trước khi ghi vào vở."
        )
        $bankEncouragement = @(
            "Con cần cố gắng nhiều hơn nữa nhé, thầy tin con sẽ tiến bộ nếu chăm chỉ!",
            "Đừng nản lòng con nhé, hãy chủ động hỏi thầy những chỗ chưa hiểu.",
            "Chăm chỉ rèn luyện từng ngày sẽ giúp con vượt qua khó khăn trong môn học này."
        )
    }
    elseif ($score -ge 5.0 -and $score -lt 6.0) {
        $bankKnowledge = @(
            "Con đã nắm được một phần kiến thức nhưng chưa thực sự vững chắc.",
            "Con cần học tập chăm chỉ hơn và làm thêm nhiều bài tập tương tự để nhớ dạng bài.",
            "Cần ôn tập lại các phương pháp giải cơ bản để thao tác nhanh hơn.",
            "Hiểu được bài nhưng khi vận dụng vào giải toán còn lúng túng."
        )
        $bankCaution = @(
            "Cần chú ý hơn đến các bước biến đổi và trình bày lời giải.",
            "Tính toán còn để xảy ra sai sót ở các bài toán có nhiều phép tính.",
            "Hãy dành thời gian soát lại toàn bộ bài làm sau khi hoàn thành."
        )
        $bankEncouragement = @(
            "Con có tiến bộ nhưng cần nỗ lực thêm nhiều hơn nữa để đạt kết quả tốt hơn.",
            "Thầy mong con chăm chỉ làm bài hơn ở các buổi học sau nhé!",
            "Cố gắng lên con nhé, chỉ cần chú ý hơn một chút là điểm số sẽ cải thiện rõ rệt."
        )
    }
    elseif ($score -ge 6.0 -and $score -lt 7.0) {
        $bankKnowledge = @(
            "Con nắm được kiến thức nền tảng và biết vận dụng vào các bài toán cơ bản.",
            "Con có ý thức làm BTVN, nắm được hướng giải nhưng đôi chỗ còn chưa hoàn thiện.",
            "Cần trau dồi thêm để giải quyết các câu hỏi nâng cao một cách tự tin hơn.",
            "Tư duy làm bài khá tốt, cần luyện tập thêm để thuần thục kỹ năng làm bài."
        )
        $bankCaution = @(
            "Cần cẩn thận hơn trong bài làm, tránh để mất điểm ở các chi tiết nhỏ.",
            "Khâu tính toán cần được kiểm tra kỹ lưỡng hơn, tránh chủ quan.",
            "Trình bày lời giải cần mạch lạc và chặt chẽ hơn một chút."
        )
        $bankEncouragement = @(
            "Con làm khá tốt, tiếp tục cố gắng để vươn lên điểm khá giỏi nhé!",
            "Thầy nhận thấy con đang có chiều hướng tiến bộ, phát huy tinh thần này con nhé!",
            "Rèn luyện thêm tính cẩn thận con sẽ đạt được điểm số cao hơn."
        )
    }
    elseif ($score -ge 7.0 -and $score -lt 8.0) {
        $bankKnowledge = @(
            "Con đã biết cách làm bài, nắm vững các phương pháp giải trọng tâm.",
            "Tư duy bài toán nhanh nhẹn và áp dụng công thức chính xác.",
            "Con hiểu rõ bài giảng trên lớp và hoàn thành tốt phần lớn các bài tập.",
            "Khả năng tiếp thu tốt, có tinh thần học hỏi tích cực."
        )
        $bankCaution = @(
            "Con cần chú ý đọc kỹ đề bài hơn để không bỏ sót các dữ kiện quan trọng.",
            "Cần cẩn thận hơn trong khâu tính toán để bài làm đạt độ chính xác tối đa.",
            "Chú ý cách trình bày các bước lập luận sao cho ngắn gọn và thuyết phục nhất."
        )
        $bankEncouragement = @(
            "Kết quả bài làm rất đáng khích lệ, cố gắng trau chuốt thêm để đạt điểm 9, 10 nhé!",
            "Thầy rất khen ngợi tinh thần học của con, tiếp tục giữ vững phong độ nhé!",
            "Con hoàn toàn có thể bứt phá lên mức xuất sắc nếu rèn thêm tính tỉ mỉ."
        )
    }
    elseif ($score -ge 8.0 -and $score -lt 9.0) {
        $bankKnowledge = @(
            "Con hiểu bài rất nhanh và nắm chắc các dạng bài từ cơ bản đến nâng cao.",
            "Tư duy toán học sắc sảo, tìm ra hướng giải bài toán một cách nhanh chóng.",
            "Bài tập về nhà được hoàn thành rất tốt, chất lượng bài làm cao.",
            "Con chủ động tìm tòi và áp dụng các cách giải bài linh hoạt."
        )
        $bankCaution = @(
            "Con hiểu bài nhưng cần cẩn thận hơn một chút trong tính toán để tránh nhầm số.",
            "Tránh tâm lý chủ quan ở các câu hỏi dễ để không mất điểm đáng tiếc.",
            "Nên kiểm tra lại từng dòng biến đổi trước khi kết luận đáp số."
        )
        $bankEncouragement = @(
            "Bài làm rất tốt, thầy tin con sẽ sớm đạt điểm tối đa ở các bài tiếp theo!",
            "Tiếp tục phát huy phong độ và sự tự tin này con nhé!",
            "Chúc mừng con với kết quả học tập rất ấn tượng ở buổi học này."
        )
    }
    else {
        # 9.0 đến 10.0
        $bankIntro = @(
            "Con có ý thức tự giác học tập tuyệt vời, bài tập về nhà hoàn thành xuất sắc.",
            "Bài làm chuẩn mực, thể hiện tinh thần tự học và chuẩn bị bài rất chu đáo.",
            "Vở sạch chữ đẹp, trình bày lời giải rất khoa học và rõ ràng."
        )
        $bankKnowledge = @(
            "Bài làm tốt, con tiếp tục phát huy nhé!",
            "Nắm rất vững toàn bộ kiến thức, giải quyết chính xác cả những bài toán khó.",
            "Khả năng tư duy logic và suy luận toán học rất xuất sắc.",
            "Lập luận chặt chẽ, cách làm bài thông minh và sáng tạo."
        )
        $bankCaution = @(
            "Chú ý một vài chi tiết nhỏ trong khâu trình bày để hoàn hảo hơn.",
            "Tiếp tục giữ gìn sự cẩn thận, không nên chủ quan ở các phép tính đơn giản.",
            "Rèn luyện thêm các bài toán mở rộng để nâng cao tư duy."
        )
        $bankEncouragement = @(
            "Bài làm rất tốt, con tiếp tục duy trì phong độ nhé!",
            "Thầy rất hài lòng với bài làm của con, tiếp tục phát huy nhé!",
            "Xuất sắc! Con hãy luôn giữ niềm đam mê với môn Toán nhé!"
        )
    }

    $pool = @()
    $pool += ($bankIntro | Get-Random)
    $pool += ($bankKnowledge | Get-Random)
    $pool += ($bankCaution | Get-Random)
    $pool += ($bankEncouragement | Get-Random)

    $extraPool = ($bankKnowledge + $bankCaution + $bankEncouragement) | Where-Object { $pool -notcontains $_ }
    while ($pool.Count -lt $lineCount -and $extraPool.Count -gt 0) {
        $item = $extraPool | Get-Random
        $pool += $item
        $extraPool = $extraPool | Where-Object { $_ -ne $item }
    }

    $formattedLines = $pool[0..($lineCount - 1)] | ForEach-Object { "- $_" }
    return ($formattedLines -join "`n")
}

# 6. Đọc dữ liệu bảng tính
Write-Host "[INFO] Đang đọc dữ liệu bảng tính..."
$maxColIdx = [Math]::Max($CommentColumnIndex, [Math]::Max($ScoreColumnIndex, $NameColumnIndex))
if ($EmailColumnIndex -gt $maxColIdx) {
    $maxColIdx = $EmailColumnIndex
}
$endColLetter = Get-ColumnLetter $maxColIdx

$readRange = [System.Uri]::EscapeDataString("'$SheetName'!A$($StartRowIndex + 1):$endColLetter$EndRowIndex")
$readUrl = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/$readRange"
$sheetData = Invoke-RestMethod -Uri $readUrl -Method Get -Headers @{ Authorization = "Bearer $token" }

$updates = @()
$actualStartRow = $StartRowIndex + 1
$studentResults = @()

for ($i = 0; $i -lt $sheetData.values.Count; $i++) {
    $row = $sheetData.values[$i]
    $stt = if ($row.Count -gt 0) { $row[0] } else { "" }
    $name = if ($row.Count -gt $NameColumnIndex) { $row[$NameColumnIndex] } else { "" }
    $scoreStr = if ($row.Count -gt $ScoreColumnIndex) { $row[$ScoreColumnIndex] } else { "" }
    $parentEmail = ""
    if ($EmailColumnIndex -ge 0 -and $row.Count -gt $EmailColumnIndex) {
        $parentEmail = "$($row[$EmailColumnIndex])".Trim()
    }
    $currentRowNum = $actualStartRow + $i

    if ([string]::IsNullOrWhiteSpace($scoreStr)) {
        $updates += ,@("")
        continue
    }

    $cleanScore = $scoreStr.Replace(',', '.')
    [double]$val = 0
    if ([double]::TryParse($cleanScore, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val)) {
        $feedback = Generate-TeacherFeedback -score $val -name $name
        $lineNum = ($feedback -split "`n").Count
        Write-Host "Row $($currentRowNum): $name ($val điểm) -> Sinh $lineNum gạch đầu dòng"
        $updates += ,@($feedback)

        $studentResults += [PSCustomObject]@{
            RowNum = $currentRowNum
            Name = $name
            Score = $val
            Feedback = $feedback
            Email = $parentEmail
        }
    } else {
        $updates += ,@("")
    }
}

# 7. Ghi nhận xét vào Google Sheets
$colLetter = Get-ColumnLetter $CommentColumnIndex
$updateRangeA1 = "'$SheetName'!$colLetter$actualStartRow`:$colLetter$EndRowIndex"
$updatePayload = @{
    range = $updateRangeA1
    majorDimension = "ROWS"
    values = $updates
} | ConvertTo-Json -Depth 5

$escapedUpdateRange = [System.Uri]::EscapeDataString($updateRangeA1)
$updateUrl = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/${escapedUpdateRange}?valueInputOption=USER_ENTERED"

$payloadBytes = [System.Text.Encoding]::UTF8.GetBytes($updatePayload)
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("Authorization", "Bearer $token")
$webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
$resBytes = $webClient.UploadData($updateUrl, "PUT", $payloadBytes)
Write-Host "[SUCCESS] Đã cập nhật nhận xét vào Google Sheets!"

# 8. Cấu hình Wrap Text và Auto Row Height
$batchBody = @{
    requests = @(
        @{
            repeatCell = @{
                range = @{
                    sheetId = $SheetGid
                    startRowIndex = $StartRowIndex
                    endRowIndex = $EndRowIndex
                    startColumnIndex = $CommentColumnIndex
                    endColumnIndex = ($CommentColumnIndex + 1)
                }
                cell = @{
                    userEnteredFormat = @{
                        wrapStrategy = "WRAP"
                    }
                }
                fields = "userEnteredFormat.wrapStrategy"
            }
        },
        @{
            autoResizeDimensions = @{
                dimensions = @{
                    sheetId = $SheetGid
                    dimension = "ROWS"
                    startIndex = $StartRowIndex
                    endIndex = $EndRowIndex
                }
            }
        }
    )
} | ConvertTo-Json -Depth 10

$batchUrl = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId`:batchUpdate"
$batchBytes = [System.Text.Encoding]::UTF8.GetBytes($batchBody)
$webClient2 = New-Object System.Net.WebClient
$webClient2.Headers.Add("Authorization", "Bearer $token")
$webClient2.Headers.Add("Content-Type", "application/json; charset=utf-8")
$webClient2.UploadData($batchUrl, "POST", $batchBytes) | Out-Null
Write-Host "[SUCCESS] Đã cấu hình định dạng Wrap Text & Auto Row Height thành công!"

# 9. Gửi Email thông báo cho phụ huynh
Write-Host "`n[INFO] --- BẮT ĐẦU XỬ LÝ GỬI EMAIL CHO PHỤ HUYNH ---"

$canSendEmail = $false
$sendMethod = "NONE"

if ($gmailOAuthReady) {
    $canSendEmail = $true
    $sendMethod = "GMAIL_API"
} elseif (-not [string]::IsNullOrWhiteSpace($SenderEmail) -and -not [string]::IsNullOrWhiteSpace($SenderPassword)) {
    $canSendEmail = $true
    $sendMethod = "SMTP"
} else {
    Write-Host "[THÔNG BÁO] Chưa cấu hình xác thực gửi mail." -ForegroundColor Yellow
    Write-Host "  -> Cách 1 (Khuyên dùng - Google OAuth 2.0 không cần mật khẩu):" -ForegroundColor Cyan
    Write-Host "     Tải credentials.json vào Documents và chạy: powershell -File scripts/authorize_gmail.ps1" -ForegroundColor Cyan
    Write-Host "  -> Cách 2 (SMTP Mật khẩu ứng dụng):" -ForegroundColor Yellow
    Write-Host "     Tạo file email_config.json trong Documents (xem mẫu scripts/email_config.example.json)`n" -ForegroundColor Yellow
}

$smtpClient = $null
if ($sendMethod -eq "SMTP") {
    try {
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtpClient.EnableSsl = $true
        $smtpClient.Credentials = New-Object System.Net.NetworkCredential($SenderEmail, $SenderPassword)
        $smtpClient.Timeout = 15000
    } catch {
        Write-Error "Lỗi khởi tạo SMTP Client: $_"
        $canSendEmail = $false
    }
}

foreach ($item in $studentResults) {
    $stName = $item.Name
    $stScore = $item.Score
    $stFeedback = $item.Feedback
    $stEmail = $item.Email
    $stRow = $item.RowNum

    if ([string]::IsNullOrWhiteSpace($stEmail)) {
        Write-Host "Row $($stRow): Học sinh '$stName' - Không có email phụ huynh, bỏ qua bước gửi mail." -ForegroundColor DarkGray
        continue
    }

    if ($stEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        Write-Host "Row $($stRow): Học sinh '$stName' - Địa chỉ email '$stEmail' không hợp lệ, bỏ qua." -ForegroundColor Red
        continue
    }

    Write-Host "Row $($stRow): Học sinh '$stName' - Tìm thấy email: $stEmail" -ForegroundColor Cyan

    if (-not $canSendEmail) {
        Write-Host "  -> [BỎ QUA] Chưa có thông tin xác thực gửi mail (thiếu email_config.json)." -ForegroundColor Yellow
        continue
    }

    $subject = "[MathPlus] Báo cáo tình hình học tập và nhận xét BTVN của con $stName"
    
    $bodyText = @"
Kính gửi phụ huynh bạn $stName,

Thầy xin gửi tới quý phụ huynh thông tin kết quả bài tập về nhà và nhận xét đánh giá của con trong buổi học vừa qua:

- Điểm bài tập về nhà: $stScore / 10
- Chi tiết nhận xét từ thầy:
$stFeedback

Thầy rất mong quý phụ huynh cùng đồng hành, nhắc nhở và động viên con để con ngày càng tiến bộ hơn nữa trong các buổi học tiếp theo.
Nếu quý phụ huynh cần trao đổi thêm bất kỳ thông tin nào về quá trình học của con, xin vui lòng liên hệ trực tiếp với thầy nhé!

Trân trọng,
$SenderDisplayName
"@

    $bodyHtml = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333; margin: 0; padding: 15px; background-color: #f4f6f8;">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center">
                <table width="600" border="0" cellspacing="0" cellpadding="25" style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.08); border-top: 5px solid #1a73e8;">
                    <tr>
                        <td>
                            <h2 style="color: #1a73e8; margin-top: 0; font-size: 20px;">BÁO CÁO KẾT QUẢ BÀI TẬP VỀ NHÀ</h2>
                            <p style="font-size: 15px;">Kính gửi phụ huynh bạn <strong>$stName</strong>,</p>
                            <p style="font-size: 14px; color: #555555;">
                                Thầy xin gửi tới quý phụ huynh thông tin kết quả bài tập về nhà và nhận xét đánh giá của con trong buổi học vừa qua:
                            </p>
                            
                            <div style="background-color: #f8f9fa; border-left: 4px solid #1a73e8; padding: 15px; margin: 18px 0; border-radius: 4px;">
                                <p style="margin: 0 0 10px 0; font-size: 15px;">
                                    <strong>Điểm bài tập về nhà:</strong> 
                                    <span style="font-size: 18px; color: #d93025; font-weight: bold; margin-left: 5px;">$stScore / 10</span>
                                </p>
                                <p style="margin: 0 0 8px 0; font-size: 14px; font-weight: bold; color: #333333;">
                                    Chi tiết nhận xét từ thầy:
                                </p>
                                <div style="white-space: pre-line; line-height: 1.6; font-size: 14px; color: #444444; padding-left: 5px;">
$stFeedback
                                </div>
                            </div>
                            
                            <p style="font-size: 14px; color: #555555;">
                                Thầy rất mong quý phụ huynh cùng đồng hành, nhắc nhở và động viên con để con ngày càng tiến bộ hơn nữa trong các buổi học tiếp theo.
                            </p>
                            <p style="font-size: 14px; color: #555555;">
                                Nếu quý phụ huynh cần trao đổi thêm bất kỳ thông tin nào về quá trình học của con, xin vui lòng liên hệ trực tiếp với thầy nhé!
                            </p>
                            
                            <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;">
                            
                            <p style="margin: 0; font-size: 13px; color: #777777; line-height: 1.4;">
                                Trân trọng,<br>
                                <strong style="color: #1a73e8; font-size: 14px;">$SenderDisplayName</strong>
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"@

    # Gửi qua Gmail API (OAuth 2.0)
    if ($sendMethod -eq "GMAIL_API") {
        try {
            $senderHeader = if ($gmailSenderEmail) { "From: $SenderDisplayName <$gmailSenderEmail>`r`n" } else { "" }
            $subjectB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($subject))
            $rawMime = "${senderHeader}To: $stEmail`r`nSubject: =?utf-8?B?$subjectB64?=`r`nMIME-Version: 1.0`r`nContent-Type: text/html; charset=utf-8`r`n`r`n$bodyHtml"
            
            $rawBytes = [System.Text.Encoding]::UTF8.GetBytes($rawMime)
            $rawUrlEncoded = Base64UrlEncode $rawBytes

            $gmailPayload = @{ raw = $rawUrlEncoded } | ConvertTo-Json
            $gmailSendUrl = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"
            $gmailHeaders = @{
                Authorization = "Bearer $gmailAccessToken"
                "Content-Type" = "application/json; charset=utf-8"
            }

            Invoke-RestMethod -Uri $gmailSendUrl -Method Post -Headers $gmailHeaders -Body $gmailPayload | Out-Null
            Write-Host "  [THÀNH CÔNG] Đã gửi email qua Gmail API (OAuth 2.0) tới: $stEmail" -ForegroundColor Green
        } catch {
            Write-Error "  [LỖI] Gửi qua Gmail API thất bại tới $stEmail : $_"
        }
    }
    # Gửi qua SMTP
    elseif ($sendMethod -eq "SMTP") {
        $mail = $null
        try {
            $mail = New-Object System.Net.Mail.MailMessage
            $mail.From = New-Object System.Net.Mail.MailAddress($SenderEmail, $SenderDisplayName, [System.Text.Encoding]::UTF8)
            $mail.To.Add($stEmail)
            $mail.Subject = $subject
            $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
            $mail.Body = $bodyHtml
            $mail.BodyEncoding = [System.Text.Encoding]::UTF8
            $mail.IsBodyHtml = $true

            $plainView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($bodyText, [System.Text.Encoding]::UTF8, "text/plain")
            $mail.AlternateViews.Add($plainView)

            $smtpClient.Send($mail)
            Write-Host "  [THÀNH CÔNG] Đã gửi email qua SMTP tới: $stEmail" -ForegroundColor Green
        } catch {
            Write-Error "  [LỖI] Gửi qua SMTP thất bại tới $stEmail : $_"
        } finally {
            if ($mail) { $mail.Dispose() }
        }
    }
}

if ($smtpClient) {
    $smtpClient.Dispose()
}
Write-Host "[INFO] Hoàn thành quy trình kiểm tra và gửi email cho phụ huynh!`n"