import smtplib
import datetime

from email.mime.text import MIMEText

dt_now = datetime.datetime.now()

year = dt_now.year
month = dt_now.month
day = dt_now.day

s_year = str(year)
s_month = str(month)
s_day = str(day)

if dt_now.month < 10:
    s_month = '0' + s_month

if dt_now.day < 10:
    s_day = '0' + s_day

email_address = '23ec015@ms.dendai.ac.jp'
outlook_password = '@Kousuke20043711'

body = """質問解決D.B.運営様

お疲れ様です、本日の作業を開始いたしますのでよろしくお願いいたします。

井上航輔
"""

subject = s_year + s_month + s_day + "作業開始連絡井上航輔"

mail_to = 'db_info@risu-kobetsu.co.jp'

msg = MIMEText(body, 'plain')
msg['Subject'] = subject
msg['To'] = mail_to
msg['From'] = email_address

server = smtplib.SMTP('smtp.office365.com', 587)
server.ehlo()
server.starttls()
server.ehlo()

server.login(email_address, outlook_password)
server.send_message(msg)
