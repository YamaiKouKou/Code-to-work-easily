import smtplib
import datetime
import os.path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

os.chdir("C://Users//sukes//OneDrive//デスクトップ//バイト関連//1.作業リスト")

dt_now = datetime.datetime.now().date().strftime("%Y%m%d")


email_address = '23ec015@ms.dendai.ac.jp'
outlook_password = '@Kousuke20043711'

body = """質問解決D.B.運営様

お疲れ様です、本日の作業内容を添付しますので確認をよろしくお願いいたします。

井上航輔
"""

subject = dt_now + " 作業報告 井上航輔"

mail_to = 'db_info@risu-kobetsu.co.jp'

msg = MIMEMultipart()
msg['Subject'] = subject
msg['To'] = mail_to
msg['From'] = email_address

msgText = MIMEText(body, 'plain')
msg.attach(msgText)

xlsx = MIMEApplication(open("20240513作業リスト福田次郎改定版.xlsx", 'rb').read())
xlsx.add_header('Content-Disposition', 'attachment', filename = "20240513作業リスト福田次郎改定版.xlsx")
msg.attach(xlsx)

server = smtplib.SMTP('smtp.office365.com', 587)
server.ehlo()
server.starttls()
server.ehlo()

server.login(email_address, outlook_password)
server.send_message(msg)
