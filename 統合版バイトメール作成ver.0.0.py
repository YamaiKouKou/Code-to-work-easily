#統合版バイトメール作成ver0.0

#メール関連のモジュールインポート
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
#アドレス、パスワードが記載されたファイルのディレクトリに移動するためのモジュールインポート
import os
#日付取得のためのモジュールインポート
import datetime
#GUI関連のモジュールインポート
import tkinter as tk
from tkinter import filedialog

#==================================================================================
#日付関連の関数

def GetToday():
    today = datetime.datetime.now()
    return today

def GetShiftDay(today_date):
    shift_days = {}
    week_num = today_date.weekday()
    increase_week_num = 13 - week_num
    shift_days[0] = today_date + datetime.timedelta(increase_week_num)
    for i in range(0, 7, 1):
        shift_days[i+1] = shift_days[0] + datetime.timedelta(i)
        shift_days[i+1] = shift_days[i+1].strftime("%#m/%#d")

    shift_days[0] = shift_days[0].strftime("%Y%m%d")
    return shift_days

#==================================================================================
#メール作成関連の関数

def GetAddressAndPassword():
    os.chdir("C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連")
    with open("Address_and_Password.txt", "r") as f:
        lines = f.readlines()
        AddressAndPassword = lines

    return AddressAndPassword

def CreateTaskStartOrReportEmail(Today, option):
    TaskStartOrReportMessage = MIMEMultipart()
    address_password = GetAddressAndPassword()
    
    message_sender_address = address_password[0].strip()
    message_receiver_address = address_password[2].strip()
    
    Today = Today.strftime("%Y%m%d")
    if option == 1:
        message_subject = Today + " 作業開始連絡 井上航輔"
        message_body = """
質問解決D.B.運営様

お疲れ様です、本日の作業を開始いたしますのでよろしくお願いいたします。

井上航輔
"""
    elif option == 2:
        message_subject = Today + " 作業報告 井上航輔"
        message_body = """
質問解決D.B.運営様

お疲れ様です、本日の作業内容を添付しますので確認をよろしくお願いいたします。

井上航輔
"""
        os.chdir("C://Users//sukes//OneDrive//デスクトップ//バイト関連//1.作業リスト")
        excel = MIMEApplication(open("20240513作業リスト福田次郎改定版.xlsx", 'rb').read())
        excel.add_header('Content-Disposition', 'attachment', filename="20240513作業リスト福田次郎改定版.xlsx")
        TaskStartOrReportMessage.attach(excel)
        
    TaskStartOrReportMessage['Subject'] = message_subject
    TaskStartOrReportMessage['To'] = message_receiver_address
    TaskStartOrReportMessage['From'] = message_sender_address
    
    TaskStartOrReportMessage.attach(MIMEText(message_body, 'plain'))
    
    return TaskStartOrReportMessage

def CreateDesiredShiftEmail(ShiftDayList):
    address_password = GetAddressAndPassword()
    
    message_sender_address = address_password[0].strip()
    message_receiver_address = address_password[2].strip()
    message_subject = ShiftDayList[0] + " 週の作業希望日程日時 井上航輔"
    message_body = """
お疲れ様です、井上航輔です。

{} 週の作業希望日程日時を送付します。

{} : 13:00~18:00
{} : 作業なし
{} : 作業なし
{} : 作業なし
{} : 作業なし
{} : 作業なし
{} : 13:00~18:00
""".format(ShiftDayList[0], ShiftDayList[1], ShiftDayList[2],
           ShiftDayList[3], ShiftDayList[4], ShiftDayList[5],
           ShiftDayList[6], ShiftDayList[7])
    
    DesiredShiftMessage = MIMEMultipart()
    DesiredShiftMessage['Subject'] = message_subject
    DesiredShiftMessage['To'] = message_receiver_address
    DesiredShiftMessage['From'] = message_sender_address
    
    DesiredShiftMessage.attach(MIMEText(message_body, 'plain'))
    
    return DesiredShiftMessage

#==================================================================================
#メール送信するための関数

def SendEmail(EmailDeta, address_and_password):
    
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(address_and_password[0].strip(), address_and_password[1].strip())
    server.send_message(EmailDeta)
    server.close()
    
#==================================================================================
#tkinter関連の関数

class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        
        tk.Tk.__init__(self, *args, **kwargs)

        self.title("統合版メール作成")
        self.geometry("800x600")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
    #--------------------------メインフレーム----------------------------------------
        #メインページフレーム作成
        self.main_frame = tk.Frame()
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        #タイトルラベル作成
        self.titleLabel = tk.Label(self.main_frame, text="なんのメールを送る?", font=('Helvetica', '35'))
        self.titleLabel.pack(anchor='center', pady = 50)
        #各フレームに移動するボタン
        #作業開始フレームへ
        self.changePageButton1 = tk.Button(self.main_frame, text="作業開始メール", command=lambda : self.changePage(self.frame1))
        self.changePageButton1.pack(side = tk.LEFT, padx = 100)
        #作業報告フレームへ
        self.changePageButton2 = tk.Button(self.main_frame, text="作業報告メール", command=lambda : self.changePage(self.frame2))
        self.changePageButton2.pack(side = tk.LEFT, padx = 60)
        #シフト希望メール
        self.changePageButton3 = tk.Button(self.main_frame, text="シフト希望メール", command=lambda : self.changePage(self.frame3))
        self.changePageButton3.pack(side = tk.LEFT, padx = 60)
    #----------------------------作業開始メールフレーム--------------------------------
        #作業開始フレーム作成(frame1)
        self.frame1 = tk.Frame()
        self.frame1.grid(row=0, column=0, sticky="nsew")
        #タイトルフレーム
        self.titleLabel = tk.Label(self.frame1, text="作業開始メールは下のボタンから送れるよ!", font=('Helvetica', '25'))
        self.titleLabel.pack(anchor='center', pady = 40)
        #作業開始メール送信ボタン
        self.Button = tk.Button(self.frame1, text="送信する!", command=lambda: SendEmail(CreateTaskStartOrReportEmail(GetToday(), 1), GetAddressAndPassword()))
        self.Button.pack(pady= 80)
        #作業開始フレームからメインフレームに戻るボタン
        self.back_button = tk.Button(self.frame1, text="メインフレームに戻る", command=lambda : self.changePage(self.main_frame))
        self.back_button.pack(side = tk.BOTTOM, pady = 40)
    #---------------------------作業報告メールフレーム--------------------------------
        #作業報告フレーム作成(frame2)
        self.frame2 = tk.Frame()
        self.frame2.grid(row=0, column=0, sticky="nsew")
        #タイトルフレーム
        self.titleLabel = tk.Label(self.frame2, text="作業報告メールは下のボタンから送れます。", font=('Helvetica', '25'))
        self.titleLabel.pack(pady= 80)
        #作業報告メール送信ボタン
        self.Button = tk.Button(self.frame2, text="送信する!", command=lambda: SendEmail(CreateTaskStartOrReportEmail(GetToday(), 2), GetAddressAndPassword()))
        self.Button.pack(pady= 80)
        #作業開始フレームからメインフレームに戻るボタン
        self.back_button = tk.Button(self.frame2, text="メインフレームに戻る", command=lambda : self.changePage(self.main_frame))
        self.back_button.pack(side = tk.BOTTOM, pady = 40)
    #------------------------------シフト希望メールフレーム------------------------------
        #シフト希望メールフレーム作成(frame3)
        self.frame3 = tk.Frame()
        self.frame3.grid(row=0, column=0, sticky="nsew")
        #タイトルフレーム作成
        self.titleLabel = tk.Label(self.frame3, text="シフト希望メールは以下のボタンより、送信可能.({} 週のシフト希望メールである。)".format(GetShiftDay(GetToday())[0]), font=('Helvetica', '13'))
        self.titleLabel.pack(pady= 40)
        #送信ボタン作成
        self.Button = tk.Button(self.frame3, text="送信!", command=lambda : SendEmail(CreateDesiredShiftEmail(GetShiftDay(GetToday())), GetAddressAndPassword()))
        self.Button.pack(pady=100)
        #メインフレームに戻るボタン
        self.main_backButton = tk.Button(self.frame3, text="メインフレームに戻る", command=lambda : self.changePage(self.main_frame))
        self.main_backButton.pack(side=tk.BOTTOM, pady=40)
    #--------------------------------------------------------------------------------
        
        #メインフレームは最初一番上にする
        self.main_frame.tkraise()
    #指定したフレームを一番上にする
    def changePage(self, page):
        page.tkraise()
#==================================================================================
#MAIN部分

if __name__ == "__main__":
    app = App()
    app.mainloop()
