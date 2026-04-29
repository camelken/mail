// メール送信モジュール
//
//   cscript //nologo send_mail.js "件名" "本文"


// 引数を取得
var mail_subject = WScript.Arguments.Unnamed(0);
var mail_body    = WScript.Arguments.Unnamed(1);
var mail_attach  = WScript.Arguments.Unnamed(2);


// 設定項目（SMTPサーバー）
var smtp_server_name  = "192.168.100.100";
var smtp_server_port  = 587;

// 設定項目（SMTP認証）
var smtp_authenticate = false;
var send_user_name = "";
var send_password = "";

// 設定項目（メールアドレス）　※送信先は半角カンマ区切りで複数設定可
var from_mail_address = "foo@sample.com";
var to_mail_addresses =  "foo@ken.com,bar@kz.com";
var bcc_mail_addresses = "zoo@zzz.com,yoot@yyy.com";

// 設定項目（TLS暗号化）
var smtp_use_ssl = false;

// -------- ローカルマシンでのコマンドの実行結果を取得する関数 --------


var ws = WScript.CreateObject("WScript.Shell");

// コマンド実行結果を行ごとに配列として取得
function cmd_output_arr( str_cmd )
{
	// コマンド実行
	var proc = ws.Exec( "cmd /c " + str_cmd );
	
	// 終了まで待つ
	while( proc.Status == 0 )
	{
		WScript.Sleep(100);
	}
	
	// 出力を取得
	var str_out = proc.StdOut.ReadAll();
	
	// 末尾の空行を削除
	var arr = str_out.split("\r\n");
	arr.pop();
	
	return arr;
}

// コマンド実行結果を文字列として取得
function cmd_output( str_cmd )
{
	return cmd_output_arr( str_cmd ).join("\r\n");
}



// -------- メイン処理 --------


var mail = WScript.CreateObject("CDO.Message");
var schemas = "http://schemas.microsoft.com/cdo/configuration/";
var reg = /CrLf/gi;

// メール内容に関する設定
mail.From     = from_mail_address;
mail.To       = to_mail_addresses;
mail.Bcc      = bcc_mail_addresses;
mail.Subject  = mail_subject;
mail.TextBody = mail_body.replace(reg, "\r\n") + "\r\n"
	+ "日時："     + cmd_output( "@echo %DATE% %TIME%" )  + "\r\n"
;
mail.TextBodyPart.Charset = "ISO-2022-JP";

//添付ファイル
mail.AddAttachment(mail_attach);

// メール送信に関する設定(XP Proなら不要)
mail.Configuration.Fields.Item( schemas + "sendusing" ) = 2;
mail.Configuration.Fields.Item( schemas + "smtpconnectiontimeout" ) = 30;

mail.Configuration.Fields.Item( schemas + "smtpserver" ) = smtp_server_name;
mail.Configuration.Fields.Item( schemas + "smtpserverport" ) = smtp_server_port;

// mail.Configuration.Fields.Item( schemas + "smtpauthenticate" ) = smtp_authenticate;
// mail.Configuration.Fields.Item( schemas + "sendusername" ) = send_user_name;
// mail.Configuration.Fields.Item( schemas + "sendpassword" ) = send_password;


// mail.Configuration.Fields.Item( schemas + "smtpusessl" ) = smtp_use_ssl;

mail.Configuration.Fields.Update();

// 送信
mail.Send();