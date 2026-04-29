# send_mail.ps1 を作成しました。元の JScript との対応は以下のとおりです。


  | 元の JScript                         | PowerShell 版                                      |
  |:-------------------------------------|:---------------------------------------------------|
  | WScript.Arguments.Unnamed(n)         | param() でパラメータ受け取り                       |
  | CDO.Message                          | System.Net.Mail.MailMessage                        |
  | WScript.Shell + @echo %DATE% %TIME%  | Get-Date                                           |
  | /CrLf/gi の正規表換                  | -ireplace "CrLf"                                   |
  | TextBodyPart.Charset = "ISO-2022-JP" | [System.Text.Encoding]::GetEncoding("iso-2022-jp") |
  | mail.AddAttachment()                 | System.Net.Mail.Attachment                         |

## 実行方法:

### 通常実行
    ```
    PowerShell -File send_mail.ps1 -MailSubject "件名" -MailBody "本文" -MailAttach "C:\path\to\file.txt"
    ```

### 添付なし
    ```
    PowerShell -File send_mail.ps1 -MailSubject "件名" -MailBody "本文CrLf改行あり"
    ```

### ExecutionPolicy が制限されている場合
    ```
    PowerShell -ExecutionPolicy Bypass -File send_mail.ps1 ...
    ```
