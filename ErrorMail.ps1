param(
    [Parameter(Mandatory = $true)][string]$Subject,
    [Parameter(Mandatory = $true)][string]$Body,
    [string]$Attachment = ""
)

# ── SMTP設定 ──────────────────────────────────────────────────────────────────
$SmtpServer      = "192.168.100.100"
$SmtpPort        = 587
# 暗号化モード: "None" | "STARTTLS" | "SSL"
#   None     : 平文（ポート25 / 587など）
#   STARTTLS : 接続後にSTARTTLSで暗号化（ポート587が一般的）
#   SSL      : 最初からTLS（SMTPS、ポート465が一般的）
$EncryptionMode  = "None"
$Authenticate    = $false
$UserName        = ""
$Password        = ""

# ── メールアドレス設定（送信先は配列で複数指定可） ────────────────────────────
$From = "foo@sample.com"
$To   = @("foo@ken.com", "bar@kz.com")
$Bcc  = @("zoo@zzz.com", "yoot@yyy.com")

# ── MailKit DLL読み込み ───────────────────────────────────────────────────────
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Add-Type -Path (Join-Path $ScriptDir "lib\MimeKit.dll")
Add-Type -Path (Join-Path $ScriptDir "lib\MailKit.dll")

# ── 本文整形（CrLf置換 + 日時付加） ──────────────────────────────────────────
$Body  = $Body -replace "CrLf", "`r`n"
$Body += "`r`n日時：" + (Get-Date -Format "yyyy/MM/dd HH:mm:ss")

# ── メッセージ構築 ─────────────────────────────────────────────────────────────
$message = New-Object MimeKit.MimeMessage

$message.From.Add([MimeKit.MailboxAddress]::Parse($From))
foreach ($addr in $To)  { $message.To.Add([MimeKit.MailboxAddress]::Parse($addr)) }
foreach ($addr in $Bcc) { $message.Bcc.Add([MimeKit.MailboxAddress]::Parse($addr)) }
$message.Subject = $Subject

$encoding = [System.Text.Encoding]::GetEncoding("iso-2022-jp")
$textPart = New-Object MimeKit.TextPart("plain")
$textPart.SetText($encoding, $Body)

$streams = [System.Collections.Generic.List[System.IO.Stream]]::new()
try {
    if ($Attachment -and (Test-Path -LiteralPath $Attachment)) {
        $stream = [System.IO.File]::OpenRead($Attachment)
        $streams.Add($stream)

        $attPart = New-Object MimeKit.MimePart
        $attPart.Content                  = New-Object MimeKit.MimeContent($stream)
        $attPart.ContentDisposition       = New-Object MimeKit.ContentDisposition([MimeKit.ContentDisposition]::Attachment)
        $attPart.ContentTransferEncoding  = [MimeKit.ContentEncoding]::Base64
        $attPart.FileName                 = [System.IO.Path]::GetFileName($Attachment)

        $multipart = New-Object MimeKit.Multipart("mixed")
        $multipart.Add($textPart)
        $multipart.Add($attPart)
        $message.Body = $multipart
    } else {
        $message.Body = $textPart
    }

    # ── SMTP送信 ────────────────────────────────────────────────────────────────
    $secureOpt = switch ($EncryptionMode) {
        "SSL"      { [MailKit.Security.SecureSocketOptions]::SslOnConnect }
        "STARTTLS" { [MailKit.Security.SecureSocketOptions]::StartTls }
        default    { [MailKit.Security.SecureSocketOptions]::None }
    }

    $client = New-Object MailKit.Net.Smtp.SmtpClient
    try {
        $client.Connect($SmtpServer, $SmtpPort, $secureOpt)
        if ($Authenticate) { $client.Authenticate($UserName, $Password) }
        $client.Send($message)
        Write-Host "送信完了"
    } finally {
        $client.Disconnect($true)
        $client.Dispose()
    }

} catch {
    Write-Error "メール送信に失敗しました: $_"
    exit 1
} finally {
    foreach ($s in $streams) { $s.Dispose() }
}
