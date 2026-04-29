# メール送信スクリプト
#
#   PowerShell -File send_mail.ps1 -MailSubject "件名" -MailBody "本文" -MailAttach "添付ファイルパス"

param(
    [string]$MailSubject,
    [string]$MailBody,
    [string]$MailAttach
)

# 設定項目（SMTPサーバー）
$SmtpServerName = "192.168.100.100"
$SmtpServerPort = 587

# 設定項目（SMTP認証）
$SmtpAuthenticate = $false
$SendUserName    = ""
$SendPassword    = ""

# 設定項目（メールアドレス）　※送信先は半角カンマ区切りで複数設定可
$FromMailAddress  = "foo@sample.com"
$ToMailAddresses  = "foo@ken.com,bar@kz.com"
$BccMailAddresses = "zoo@zzz.com,yoot@yyy.com"

# 設定項目（TLS暗号化）
$SmtpUseSsl = $false


# -------- メイン処理 --------

Add-Type -AssemblyName System.Net.Mail

$encoding = [System.Text.Encoding]::GetEncoding("iso-2022-jp")

# 本文の CrLf 文字列を実際の改行に置換し、日時を追加
$dateTime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
$body = $MailBody -ireplace "CrLf", "`r`n"
$body = $body + "`r`n日時：$dateTime`r`n"

# メールオブジェクト作成
$mail = New-Object System.Net.Mail.MailMessage
$mail.From            = New-Object System.Net.Mail.MailAddress($FromMailAddress)
$mail.Subject         = $MailSubject
$mail.SubjectEncoding = $encoding
$mail.Body            = $body
$mail.BodyEncoding    = $encoding
$mail.IsBodyHtml      = $false

# 宛先設定
foreach ($addr in $ToMailAddresses.Split(",")) {
    $trimmed = $addr.Trim()
    if ($trimmed) { $mail.To.Add($trimmed) }
}

# BCC設定
foreach ($addr in $BccMailAddresses.Split(",")) {
    $trimmed = $addr.Trim()
    if ($trimmed) { $mail.Bcc.Add($trimmed) }
}

# 添付ファイル
if ($MailAttach) {
    if (Test-Path $MailAttach) {
        $mail.Attachments.Add((New-Object System.Net.Mail.Attachment($MailAttach)))
    } else {
        Write-Error "添付ファイルが見つかりません: $MailAttach"
        exit 1
    }
}

# SMTPクライアント設定
$smtp = New-Object System.Net.Mail.SmtpClient($SmtpServerName, $SmtpServerPort)
$smtp.EnableSsl              = $SmtpUseSsl
$smtp.UseDefaultCredentials  = $false

if ($SmtpAuthenticate) {
    $smtp.Credentials = New-Object System.Net.NetworkCredential($SendUserName, $SendPassword)
}

# 送信
try {
    $smtp.Send($mail)
    Write-Host "メール送信完了"
} catch {
    Write-Error "メール送信失敗: $_"
    exit 1
} finally {
    $mail.Dispose()
    $smtp.Dispose()
}
