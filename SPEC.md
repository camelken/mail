# エラーメール送信スクリプト 仕様書

## 概要

バッチ処理でエラーが発生した際に、SMTPサーバー経由でエラー通知メールを送信するスクリプト群。  
本文エンコードに **ISO-2022-JP** を使用し、当日のログファイルを添付して送信する。

---

## ファイル構成

```
mail/
├── ErrorMail.bat       # 起動バッチ（ログ取得 → PS1呼び出し）
├── ErrorMail.ps1       # メール送信スクリプト本体（MailKit使用）
├── SPEC.md             # 本仕様書
└── lib/                # MailKit DLLフォルダ（要セットアップ）
    ├── MailKit.dll
    ├── MimeKit.dll
    └── BouncyCastle.Cryptography.dll
```

---

## 前提条件・セットアップ

### 必要環境

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / Windows Server 2016 以降 |
| PowerShell | 5.1 以上 |
| .NET Framework | 4.8 以上（または .NET 6+） |

### MailKit DLLの配置

NuGet から MailKit パッケージを取得し、`lib\` フォルダに配置する。

```powershell
# PowerShellから取得する場合
Install-Package MailKit -Destination .\lib -SkipDependencies
```

または NuGet CLI を使う場合：

```cmd
nuget install MailKit -OutputDirectory lib
```

取得後、以下の DLL を `lib\` 直下にコピーする：

- `MailKit.dll`
- `MimeKit.dll`
- `BouncyCastle.Cryptography.dll`

---

## 設定項目

`ErrorMail.ps1` 冒頭のパラメータを環境に合わせて変更する。

### SMTPサーバー設定

| 変数名 | 既定値 | 説明 |
|--------|--------|------|
| `$SmtpServer` | `192.168.100.100` | SMTPサーバーのホスト名またはIPアドレス |
| `$SmtpPort` | `587` | SMTPポート番号 |
| `$EncryptionMode` | `"None"` | 暗号化モード（下表参照） |

**`$EncryptionMode` の設定値：**

| 値 | 説明 | 一般的なポート |
|----|------|---------------|
| `"None"` | 平文（暗号化なし） | 25 / 587 |
| `"STARTTLS"` | 接続後に STARTTLS で暗号化 | 587 |
| `"SSL"` | 最初から TLS（SMTPS） | 465 |

### SMTP認証設定

| 変数名 | 既定値 | 説明 |
|--------|--------|------|
| `$Authenticate` | `$false` | SMTP認証を使用する場合は `$true` |
| `$UserName` | `""` | SMTP認証ユーザー名 |
| `$Password` | `""` | SMTP認証パスワード |

### メールアドレス設定

| 変数名 | 既定値 | 説明 |
|--------|--------|------|
| `$From` | `foo@sample.com` | 送信元メールアドレス |
| `$To` | `@("foo@ken.com", "bar@kz.com")` | 送信先（配列で複数指定可） |
| `$Bcc` | `@("zoo@zzz.com", "yoot@yyy.com")` | BCC（配列で複数指定可） |

---

## スクリプト仕様

### ErrorMail.bat

バッチ処理のエラーハンドラから呼び出すエントリーポイント。

**処理フロー：**

1. カレントディレクトリを保存（`%mypath%`）
2. 件名・本文・ログフォルダパスを変数にセット
3. ログフォルダ（`E:\LOG`）に移動し、**当日作成されたファイル**のパスを取得  
   （`forfiles /d 0` で当日付のファイルを検索）
4. `ErrorMail.ps1` を PowerShell で実行

**呼び出しコマンド：**

```cmd
powershell -ExecutionPolicy Bypass -NoProfile -File "%mypath%\ErrorMail.ps1" ^
    -Subject %mail_subject% -Body %mail_body% -Attachment %mail_attach%
```

---

### ErrorMail.ps1

MailKit を使用してメールを送信する PowerShell スクリプト。

#### パラメータ

| パラメータ名 | 必須 | 型 | 説明 |
|--------------|------|----|------|
| `-Subject` | 必須 | String | メール件名 |
| `-Body` | 必須 | String | メール本文（`CrLf` は改行に置換） |
| `-Attachment` | 任意 | String | 添付ファイルの絶対パス |

#### 本文の改行記法

本文中の `CrLf`（大文字・小文字問わず）は送信時に `\r\n`（CRLF）に自動変換される。

**例：**
```
"1行目CrLf2行目CrLfCrLf4行目"
↓
1行目
2行目

4行目
```

#### 本文末尾への日時付加

送信時に本文末尾へ以下の形式で日時が自動追記される：

```
日時：2026/04/30 12:34:56
```

#### メールエンコード

| 項目 | 設定値 |
|------|--------|
| 文字コード | ISO-2022-JP（JIS） |
| Content-Transfer-Encoding | 7bit |
| 添付ファイル | Base64 |

#### メッセージ構造

| 条件 | MIMEタイプ |
|------|-----------|
| 添付なし | `text/plain` |
| 添付あり | `multipart/mixed`（本文 + 添付） |

#### 処理フロー

```
1. パラメータ受取
2. MailKit / MimeKit DLL ロード
3. 本文の CrLf 置換 + 日時付加
4. MimeMessage 構築
   ├── From / To / Bcc / Subject セット
   ├── TextPart（ISO-2022-JP）生成
   └── 添付ファイルがあれば multipart/mixed に組み立て
5. SmtpClient.Connect()
6. 認証（$Authenticate = $true の場合のみ）
7. Send()
8. Disconnect()
9. ストリーム・クライアント解放（finally）
```

#### エラー処理

- 送信失敗時は `Write-Error` でエラー内容を標準エラーへ出力し、`exit 1` で終了
- ストリームとSMTPクライアントは例外発生時も `finally` ブロックで確実に解放

---

## 手動実行例

```powershell
# 添付なし
.\ErrorMail.ps1 -Subject "テスト" -Body "本文1行目CrLf本文2行目"

# 添付あり
.\ErrorMail.ps1 -Subject "エラー発生" -Body "異常終了しました。" -Attachment "C:\LOG\app_20260430.log"
```

---

## 注意事項

- `lib\` フォルダの DLL が存在しない場合、スクリプトは即時エラー終了する
- `E:\LOG` に当日付のファイルが複数存在する場合、`forfiles` が最後に返したファイルが添付される
- 添付ファイルパスが空または存在しないパスの場合、添付なしでメールを送信する
- SMTP認証を使用しない環境では `$Authenticate = $false` のままにする（既定値）
