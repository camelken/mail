メール送信用のPowerShellスクリプト。

# MailKit DLL インストール手順

## 配置先

```
mail/
└── lib/
    ├── MailKit.dll
    ├── MimeKit.dll
    └── BouncyCastle.Cryptography.dll
```

---

## 方法 1：nuget.exe を使う（推奨）

### 1. nuget.exe を入手

[https://www.nuget.org/downloads](https://www.nuget.org/downloads) から `nuget.exe` をダウンロードし、PATH の通った場所に置く。

### 2. パッケージ取得

```cmd
cd mail
nuget install MailKit -OutputDirectory packages
```

MailKit・MimeKit・BouncyCastle.Cryptography がまとめてダウンロードされる。

### 3. DLL をコピー

PowerShell のバージョンに合わせてフォルダを選ぶ：

| PowerShell | .NET バージョン | コピー元フォルダ |
|---|---|---|
| 5.1 | .NET Framework 4.x | `net462` または `net48` |
| 7.x | .NET 6 以降 | `net6.0` または `net8.0` |

```cmd
:: PowerShell 5.1 の場合（net462）
md lib
copy packages\MailKit.*\lib\net462\MailKit.dll                                      lib\
copy packages\MimeKit.*\lib\net462\MimeKit.dll                                      lib\
copy packages\BouncyCastle.Cryptography.*\lib\net462\BouncyCastle.Cryptography.dll  lib\
```

```cmd
:: PowerShell 7.x の場合（net6.0）
md lib
copy packages\MailKit.*\lib\net6.0\MailKit.dll                                      lib\
copy packages\MimeKit.*\lib\net6.0\MimeKit.dll                                      lib\
copy packages\BouncyCastle.Cryptography.*\lib\net6.0\BouncyCastle.Cryptography.dll  lib\
```

---

## 方法 2：dotnet CLI を使う（.NET 6+ 環境）

```cmd
cd mail
dotnet new classlib -o _tmp --no-restore
cd _tmp
dotnet add package MailKit
dotnet restore
cd ..
```

復元後、グローバルキャッシュに展開される：

```
%USERPROFILE%\.nuget\packages\mailkit\<version>\lib\net6.0\MailKit.dll
%USERPROFILE%\.nuget\packages\mimekit\<version>\lib\net6.0\MimeKit.dll
%USERPROFILE%\.nuget\packages\bouncycastle.cryptography\<version>\lib\net6.0\BouncyCastle.Cryptography.dll
```

上記を `lib\` にコピーし、`_tmp\` フォルダは削除してよい。

---

## 方法 3：PowerShell PackageManagement

```powershell
# NuGet プロバイダーが未登録の場合
Install-PackageProvider -Name NuGet -Force

# MailKit を packages\ に展開
Save-Package -Name MailKit -Path .\packages -ProviderName NuGet
```

展開後は方法 1 と同様に、対象フレームワークのフォルダから DLL を `lib\` にコピーする。

---

## 動作確認

```powershell
Add-Type -Path ".\lib\MimeKit.dll"
Add-Type -Path ".\lib\MailKit.dll"
[MailKit.Net.Smtp.SmtpClient]  # エラーなければ OK
```
