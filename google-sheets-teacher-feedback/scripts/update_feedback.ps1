<#
.SYNOPSIS
    Tự động đọc điểm BTVN từ Google Sheets, sinh nhận xét học sinh (3-6 gạch đầu dòng ngẫu nhiên),
    ghi vào cột nhận xét và định dạng text wrap, auto height.

.DESCRIPTION
    Script sử dụng Google Service Account để xác thực qua JWT Bearer token và tương tác với Google Sheets API v4.
    Không yêu cầu cài đặt Python hay thư viện ngoài, chạy hoàn toàn bằng PowerShell và .NET Framework có sẵn trên Windows.

.PARAMETER ServiceAccountPath
    Đường dẫn đến file JSON của Service Account. Mặc định: "service_account.json".

.PARAMETER SpreadsheetId
    ID của Google Spreadsheet cần cập nhật.

.PARAMETER SheetName
    Tên của Sheet (tab) cần thao tác (ví dụ: "Tháng 09/2026").

.PARAMETER TargetRange
    Dải ô nhận xét cần cập nhật (ví dụ: "I16:I21").

.PARAMETER ScoreRange
    Dải ô điểm BTVN tương ứng (ví dụ: "H16:H21").
#>

[CmdletBinding()]
param (
    [string]$ServiceAccountPath = "service_account.json",
    [string]$SpreadsheetId = "1FJixJmSa8vTH1ub8yiue3cJrUh1igf3NpCs5Q4M1w78",
    [string]$SheetName = "Tháng 09/2026",
    [int]$SheetGid = 1490594752,
    [int]$StartRowIndex = 15, # 0-indexed (Row 16)
    [int]$EndRowIndex = 21,   # 0-indexed exclusive (Row 21)
    [int]$CommentColumnIndex = 8 # 0-indexed (Column I)
)

$ErrorActionPreference = "Stop"

function Base64UrlEncode([byte[]]$bytes) {
    return [System.Convert]::ToBase64String($bytes).Replace('+', '-').Replace('/', '_').TrimEnd('=')
}

function Base64UrlEncodeString([string]$str) {
    return Base64UrlEncode([System.Text.Encoding]::UTF8.GetBytes($str))
}

# 1. Kiểm tra file Service Account
if (-not (Test-Path $ServiceAccountPath)) {
    throw "Không tìm thấy file Service Account tại: $ServiceAccountPath"
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
Write-Host "[SUCCESS] Lấy token xác thực thành công!"

# 3. Hàm tạo nhận xét ngẫu nhiên 3-6 gạch đầu dòng
function Generate-TeacherFeedback([double]$score, [string]$name) {
    # Ngẫu nhiên số dòng từ 3 đến 6 dòng
    $lineCount = Get-Random -Minimum 3 -Maximum 7

    # Nhóm câu theo từng mức điểm và khía cạnh nhận xét
    $bankIntro = @(
        "Con có ý thức tự giác làm bài tập về nhà rất tốt.",
        "Con nộp bài đúng hạn và hoàn thành đầy đủ bài tập được giao.",
        "Ý thức chuẩn bị bài và làm BTVN ở nhà đáng khen ngợi.",
        "Con chăm chỉ làm bài tập, bài vở trình bày tương đối sạch đẹp."
    )

    $bankKnowledge = @()
    $bankCaution = @()
    $bankEncouragement = @()

    if ($score -lt 5.0) {
        $bankKnowledge = @(
            "Con còn hổng một số kiến thức cơ bản của bài học.",
            "Chưa nắm chắc phương pháp giải các dạng toán trọng tâm.",
            "Cần dành thêm nhiều thời gian ôn lại lý thuyết trước khi làm bài."
        )
        $bankCaution = @(
            "Cần chú ý đọc kỹ đề bài để hiểu đúng yêu cầu.",
            "Tránh làm vội vàng, cần làm nháp cẩn thận từng bước.",
            "Chưa nắm được cách trình bày mạch lạc, con cần rèn thêm."
        )
        $bankEncouragement = @(
            "Cần cố gắng nhiều hơn ở các bài học tiếp theo.",
            "Con cần chủ động hỏi thầy cô khi chưa hiểu bài nhé.",
            "Thầy tin nếu con tập trung và kiên trì hơn sẽ tiến bộ rõ rệt!"
        )
    }
    elseif ($score -ge 5.0 -and $score -lt 6.0) {
        $bankKnowledge = @(
            "Con đã nắm được một phần kiến thức nhưng chưa thực sự chắc chắn.",
            "Đã biết áp dụng công thức nhưng đôi chỗ còn lúng túng.",
            "Cần ôn tập kỹ lại các dạng toán cơ bản đã học."
        )
        $bankCaution = @(
            "Con cần học tập chăm chỉ hơn và dành thêm thời gian làm bài tập.",
            "Chú ý rèn luyện thêm kỹ năng tính toán cơ bản.",
            "Cần cẩn thận hơn, kiểm tra lại bài trước khi nộp."
        )
        $bankEncouragement = @(
            "Cố gắng chăm chỉ hơn con nhé, thầy tin con sẽ làm tốt hơn!",
            "Chỉ cần chăm chỉ rèn luyện thêm con sẽ nâng cao được điểm số.",
            "Hãy nỗ lực hơn nữa ở những bài sau nhé con!"
        )
    }
    elseif ($score -ge 6.0 -and $score -lt 7.0) {
        $bankKnowledge = @(
            "Con có ý thức làm BTVN và nắm được sườn bài cơ bản.",
            "Hiểu được dạng toán chính nhưng vận dụng chưa thật linh hoạt.",
            "Đã biết cách giải các bài toán vừa sức."
        )
        $bankCaution = @(
            "Cần trau dồi thêm để giải quyết các câu hỏi nâng cao.",
            "Chú ý các bước tính toán trung gian để tránh mất điểm.",
            "Nên rèn thêm tính cẩn thận và lập luận chặt chẽ hơn."
        )
        $bankEncouragement = @(
            "Ý thức học tập tốt, con cố gắng phát huy để tiến bộ hơn nhé!",
            "Cố gắng trau dồi thêm con sẽ bứt phá ở bài kiểm tra tới.",
            "Thầy khen tinh thần làm bài, tiếp tục nỗ lực nhé!"
        )
    }
    elseif ($score -ge 7.0 -and $score -lt 8.0) {
        $bankKnowledge = @(
            "Con đã biết cách làm bài và nắm được phương pháp giải.",
            "Tư duy toán học tốt, nắm chắc các kiến thức cốt lõi.",
            "Có khả năng tự giải quyết các bài toán độc lập."
        )
        $bankCaution = @(
            "Cần chú ý đọc kỹ đề và cẩn thận hơn trong từng phép tính.",
            "Đôi chỗ còn chủ quan ở các câu hỏi dễ, cần soát bài kỹ.",
            "Cần chú ý cách trình bày cho ngắn gọn và mạch lạc hơn."
        )
        $bankEncouragement = @(
            "Con có tiềm năng tốt, chú ý thêm một chút sẽ đạt điểm cao hơn nhé!",
            "Kết quả bài làm tốt, tiếp tục rèn luyện để hoàn thiện hơn.",
            "Thầy tin con sẽ sớm đạt điểm tối đa nếu cẩn thận hơn!"
        )
    }
    elseif ($score -ge 8.0 -and $score -lt 9.0) {
        $bankKnowledge = @(
            "Con hiểu bài tốt, nắm chắc phương pháp và làm bài đúng hướng.",
            "Tiếp thu bài nhanh nhẹn, tư duy logic khá sắc sảo.",
            "Làm chủ tốt các dạng toán đã được học trên lớp."
        )
        $bankCaution = @(
            "Con hiểu bài nhưng cần cẩn thận hơn trong tính toán để tránh lỗi nhỏ.",
            "Kỹ năng tính nhẩm tốt nhưng cần soát lại để không nhầm dấu/số.",
            "Trình bày bài nên nắn nót, rõ ràng hơn một chút nhé."
        )
        $bankEncouragement = @(
            "Bài làm đạt kết quả tốt, con rèn thêm tính tỉ mỉ nhé!",
            "Thầy rất khen ngợi con, cố gắng phát huy ở các bài sau nhé!",
            "Giữ vững tinh thần học tập tích cực này con nhé!"
        )
    }
    else { # >= 9.0
        $bankKnowledge = @(
            "Bài làm rất tốt, con hiểu sâu sắc kiến thức bài học.",
            "Tư duy logic nhạy bén, giải toán nhanh và chính xác.",
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

    # Chọn lọc ngẫu nhiên các ý để tạo thành số dòng yêu cầu
    $pool = @()
    $pool += ($bankIntro | Get-Random)
    $pool += ($bankKnowledge | Get-Random)
    $pool += ($bankCaution | Get-Random)
    $pool += ($bankEncouragement | Get-Random)

    # Thêm các ý phụ nếu cần 5 hoặc 6 dòng
    $extraPool = ($bankKnowledge + $bankCaution + $bankEncouragement) | Where-Object { $pool -notcontains $_ }
    while ($pool.Count -lt $lineCount -and $extraPool.Count -gt 0) {
        $item = $extraPool | Get-Random
        $pool += $item
        $extraPool = $extraPool | Where-Object { $_ -ne $item }
    }

    $formattedLines = $pool[0..($lineCount - 1)] | ForEach-Object { "- $_" }
    return ($formattedLines -join "`n")
}

Write-Host "[INFO] Đang đọc dữ liệu bảng tính..."
$readRange = [System.Uri]::EscapeDataString("'$SheetName'!A$($StartRowIndex + 1):I$EndRowIndex")
$readUrl = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/$readRange"
$sheetData = Invoke-RestMethod -Uri $readUrl -Method Get -Headers @{ Authorization = "Bearer $token" }

$updates = @()
$actualStartRow = $StartRowIndex + 1

for ($i = 0; $i -lt $sheetData.values.Count; $i++) {
    $row = $sheetData.values[$i]
    $stt = if ($row.Count -gt 0) { $row[0] } else { "" }
    $name = if ($row.Count -gt 1) { $row[1] } else { "" }
    $scoreStr = if ($row.Count -gt 7) { $row[7] } else { "" }
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
        Write-Host "Row $currentRowNum: $name ($val điểm) -> Sinh $lineNum gạch đầu dòng"
        $updates += ,@($feedback)
    } else {
        $updates += ,@("")
    }
}

# 4. Ghi nhận xét vào Google Sheets
$colLetter = [char](65 + $CommentColumnIndex)
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

# 5. Cấu hình Wrap Text và Auto Row Height
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
