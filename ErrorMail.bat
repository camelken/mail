@echo off

rem 変数宣言
setlocal
set mypath=%CD%
set mail_subject="エラーが発生しました"
set mail_body="異常終了しました。CrLfCrLf詳細は添付のログファイルを確認してください。CrLfCrLfこのメールはシステムから送信しています。"
set log_folder="E:\LOG"

rem ログフォルダに移動して本日作成されたログファイルパスを取得
cd %log_folder%
set mail_attach=""
for /f %%a in ('forfiles /c "cmd /c echo @path" /d 0') do @set mail_attach=%%a

cd %mypath%
powershell -ExecutionPolicy Bypass -NoProfile -File "%mypath%\ErrorMail.ps1" -Subject %mail_subject% -Body %mail_body% -Attachment %mail_attach%

endlocal

exit
