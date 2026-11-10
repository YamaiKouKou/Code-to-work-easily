#importども。
import datetime
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#この下のimportはメールに添付する場合のみ用いる。(今回の場合は用いない。)
from email.mime.application import MIMEApplication

#プログラムを実行した日の翌々週の日曜日を取得する関数定義
def GetShiftDay():
    #今日の日付取得
    dt_shift = datetime.datetime.now()
    #日付から曜日取得
    dt_week = dt_shift.weekday()
    #今日の前週の日曜 -(1+dt_week) までの値を求め、そこから2週間後(+14)の日を取得するための加算値を求める
    week_num = 13 - dt_week
    #上で求めた加算値を今日の日付に足す
    dt_shift_sunday = dt_shift + datetime.timedelta(days=week_num)
    #空の配列StrShift_Titleを定義
    StrShift_Title = {}
    #配列の1番目にメールの題名用の日付を格納
    StrShift_Title[0] = dt_shift_sunday.strftime("%Y%m%d")
    #配列の2~8番目にメールの本文用の日付を順番に格納
    for i in range(1, 9, 1):
        StrShift_Title[i] = dt_shift_sunday + datetime.timedelta(days=i-1)
        #日付の文字列から、0の自動埋め込みを削除
        StrShift_Title[i] = StrShift_Title[i].strftime("%#m/%#d")
    #メールの題名の文字列に合わせたものを戻り値として返す
    return StrShift_Title

#テキストファイルからoutlookのメールアドレスを取得する関数定義
def GetEmailAddress_and_Password():
    #アドレスが記載されているテキストファイルがあるディレクトリに移動
    os.chdir("C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連")
    #テキストファイルを読み取り専用で開き、1行ずつ抽出
    with open("Address_and_Password.txt", "r") as f:
        lines = f.readlines()
        #配列linesをまとめて変数address_and_passwordに代入(なにが格納されているかを明示的にするため)
        address_and_password = lines
        #アドレスとパスワードの文字列の配列を戻り値として返す([0]にはアドレス、[1]にはパスワード、[2]には相手のアドレスがそれぞれ格納されている)
        return address_and_password

#メール作成する関数定義
def CreateEmail():
    #メールの本文内で用いる日付文字列の配列を変数で受け取る
    getshiftdays = GetShiftDay()
    address_and_password = GetEmailAddress_and_Password()
    #以下、メールの原料ども。
    message = MIMEMultipart()
    message_subject = getshiftdays[0] + " 週の作業希望日程日時 井上航輔"
    message_address = address_and_password[0].strip()
    others_message_address = address_and_password[2].strip()
    message_body ="""
お疲れ様です、井上航輔です。

{}を送付します。

{}: 13:00~18:00
{}: 作業なし
{}: 作業なし
{}: 作業なし
{}: 作業なし
{}: 作業なし
{}: 13:00~18:00
    """.format(getshiftdays[0] + " 週の作業希望日程日時", getshiftdays[1],
               getshiftdays[2], getshiftdays[3], getshiftdays[4],
               getshiftdays[5], getshiftdays[6], getshiftdays[7],
               getshiftdays[8])
    #メールの題名
    message['Subject'] = message_subject
    #メールの送り先
    message['To'] = others_message_address
    #メールの送り主
    message['From'] = message_address
    #メールに本文添付(MIME形式)
    message.attach(MIMEText(message_body, 'plain'))
    #メールの情報を返す
    return message

#メールを送る関数定義
def SendEmail():
    #sendを入力した場合のみ以下のコードは実行される
    ans = input("メールを送る場合は'send'を入力してね。")
    if ans == "send":
        #関数GetEmailAddress_and_Passwordで取得したアドレスとパスワードを代入
        address_and_password = GetEmailAddress_and_Password()
        #変数msgに関数CreateEmailで作成したメールの情報を格納
        msg = CreateEmail()
        #SMTPサーバにアクセス
        server = smtplib.SMTP('smtp.office365.com', 587)
        #STARTTLSの暗号化通信に用いてメールを送る準備
        server.starttls()
        #addressにaddress_and_password[0]を入れる(strip()によりテキストの行の空白文字を消去している)
        address = address_and_password[0].strip()
        #passwordにaddress_and_password[1]を入れる(strip()によりテキストの行の空白文字を消去している)
        password = address_and_password[1].strip()
        #SMTPサーバ(office365のメール)にログイン
        server.login(address, password)
        #メールを相手に送信
        server.send_message(msg)
        #メールサーバを閉じる
        server.close()

SendEmail()       
