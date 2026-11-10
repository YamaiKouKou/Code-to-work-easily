#統合版バイトメール作成ver0.1

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
#tkinter関連の関数

class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        
        tk.Tk.__init__(self, *args, **kwargs)

        self.title("統合版メール作成")
        self.geometry("800x600")
        self.resizable(False, False)

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
        self.changePageButton3 = tk.Button(self.main_frame, text="シフト希望メール", command=lambda : self.changePageAndSetDate(self.frame3))
        self.changePageButton3.pack(side = tk.LEFT, padx = 60)
    #----------------------------作業開始メールフレーム--------------------------------
        #作業開始フレーム作成(frame1)
        self.frame1 = tk.Frame()
        self.frame1.grid(row=0, column=0, sticky="nsew")
        #タイトルフレーム
        self.titleLabel = tk.Label(self.frame1, text="作業開始メールは下のボタンから送れるよ!", font=('Helvetica', '25'))
        self.titleLabel.pack(anchor='center', pady = 40)
        #作業開始メール送信ボタン
        self.Button = tk.Button(self.frame1, text="送信する!", command=lambda: self.SendEmail(self.CreateTaskStartOrReportEmail(self.GetToday(), 1), self.GetAddressAndPassword()))
        self.Button.pack(pady=40)
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
        self.Button = tk.Button(self.frame2, text="送信する!", command=lambda: self.SendEmail(self.CreateTaskStartOrReportEmail(self.GetToday(), 2), self.GetAddressAndPassword()))
        self.Button.pack(pady=40)
        #作業開始フレームからメインフレームに戻るボタン
        self.back_button = tk.Button(self.frame2, text="メインフレームに戻る", command=lambda : self.changePage(self.main_frame))
        self.back_button.pack(side = tk.BOTTOM, pady = 40)
    #------------------------------シフト希望メールフレーム------------------------------
        #シフト希望メールフレーム作成(frame3)
        self.frame3 = tk.Frame()
        self.frame3.grid(row=0, column=0, sticky="nsew")
        #タイトルフレーム作成
        self.titleLabel = tk.Label(self.frame3, text="シフト希望メールは以下のボタンより、送信可能.({} 週のシフト希望メールである。)".format(self.GetShiftDay(self.GetToday())[0]), font=('Helvetica', '13'))
        self.titleLabel.pack(pady=30)
        #シフト入力部分作成(日から土)
        #-------------日曜------------------------------------
        self.sun_label = tk.Label(self.frame3)
        self.sun_label.place(x=250, y=93)
        self.entry_sun = tk.Entry(self.frame3)
        self.entry_sun.pack(pady=10)
        #-------------月曜------------------------------------
        self.mon_label = tk.Label(self.frame3)
        self.mon_label.place(x=250, y=133)
        self.entry_mon = tk.Entry(self.frame3)
        self.entry_mon.pack(pady=10)
        #-------------火曜------------------------------------
        self.tue_label = tk.Label(self.frame3)
        self.tue_label.place(x=250, y=173)
        self.entry_tue = tk.Entry(self.frame3)
        self.entry_tue.pack(pady=10)
        #-------------水曜------------------------------------
        self.wed_label = tk.Label(self.frame3)
        self.wed_label.place(x=250, y=213)
        self.entry_wed = tk.Entry(self.frame3)
        self.entry_wed.pack(pady=10)
        #-------------木曜------------------------------------
        self.thu_label = tk.Label(self.frame3)
        self.thu_label.place(x=250, y=253)
        self.entry_thu = tk.Entry(self.frame3)
        self.entry_thu.pack(pady=10)
        #-------------金曜------------------------------------
        self.fri_label = tk.Label(self.frame3)
        self.fri_label.place(x=250, y=293)
        self.entry_fri = tk.Entry(self.frame3)
        self.entry_fri.pack(pady=10)
        #-------------土曜------------------------------------
        self.sat_label = tk.Label(self.frame3)
        self.sat_label.place(x=250, y=333)
        self.entry_sat = tk.Entry(self.frame3)
        self.entry_sat.pack(pady=10)
        #送信ボタン作成
        self.Button = tk.Button(self.frame3, text="送信!", command=lambda : self.SendEmail(self.CreateDesiredShiftEmail(self.GetShiftDay(self.GetToday())), self.GetAddressAndPassword()))
        self.Button.pack(pady=50)
        #メインフレームに戻るボタン
        self.main_backButton = tk.Button(self.frame3, text="メインフレームに戻る", command=lambda : self.changePage(self.main_frame))
        self.main_backButton.pack(side=tk.BOTTOM, pady=10)
    #--------------------------------------------------------------------------------
        
        #メインフレームは最初一番上にする
        self.main_frame.tkraise()
    #指定したフレームを一番上にする
    def changePage(self, page):
        page.tkraise()

    #シフト希望の場合のみに使用するコマンド(シフトの月日もGUI画面上に出力する)
    def changePageAndSetDate(self, page):
        #フレームを一番上に表示する
        page.tkraise()
        #シフト希望の日を取得
        Days = self.GetShiftDay(self.GetToday())
        #シフト希望のテキストを変更(日曜~土曜)
        self.sun_label["text"] = Days[1] + " (日)"
        self.mon_label["text"] = Days[2] + " (月)"
        self.tue_label["text"] = Days[3] + " (火)"
        self.wed_label["text"] = Days[4] + " (水)"
        self.thu_label["text"] = Days[5] + " (木)"
        self.fri_label["text"] = Days[6] + " (金)"
        self.sat_label["text"] = Days[7] + " (土)"

    #シフト希望入力フォームを取得
    def getShiftTime(self):
        ShiftTime = [self.entry_sun.get(), self.entry_mon.get(), self.entry_tue.get(),
                     self.entry_wed.get(), self.entry_thu.get(), self.entry_fri.get(),
                     self.entry_sat.get()]
        return ShiftTime

    #==================================================================================
    #日付関連の関数

    def GetToday(self):
        today = datetime.datetime.now()
        return today

    def GetShiftDay(self, today_date):
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

    def GetAddressAndPassword(self):
        os.chdir("C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連")
        with open("Address_and_Password.txt", "r") as f:
            lines = f.readlines()
            AddressAndPassword = lines

        return AddressAndPassword

    def CreateTaskStartOrReportEmail(self, Today, option):
        TaskStartOrReportMessage = MIMEMultipart()
        address_password = self.GetAddressAndPassword()
    
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
            excel = MIMEApplication(open("20240928福田次郎.xlsx", 'rb').read())
            excel.add_header('Content-Disposition', 'attachment', filename="20240928福田次郎.xlsx")
            TaskStartOrReportMessage.attach(excel)
        
        TaskStartOrReportMessage['Subject'] = message_subject
        TaskStartOrReportMessage['To'] = message_receiver_address
        TaskStartOrReportMessage['From'] = message_sender_address
    
        TaskStartOrReportMessage.attach(MIMEText(message_body, 'plain'))
    
        return TaskStartOrReportMessage

    def CreateDesiredShiftEmail(self, ShiftDayList):
        address_password = self.GetAddressAndPassword()
        Shift = self.getShiftTime()
    
        message_sender_address = address_password[0].strip()
        message_receiver_address = address_password[2].strip()
        message_subject = ShiftDayList[0] + " 週の作業希望日程日時 井上航輔"
        message_body = """
お疲れ様です、井上航輔です。

{} 週の作業希望日程日時を送付します。

{} : {}
{} : {}
{} : {}
{} : {}
{} : {}
{} : {}
{} : {}
""".format(ShiftDayList[0], ShiftDayList[1], Shift[0], ShiftDayList[2], Shift[1],
           ShiftDayList[3], Shift[2], ShiftDayList[4], Shift[3], ShiftDayList[5],
           Shift[4], ShiftDayList[6], Shift[5], ShiftDayList[7], Shift[6])
    
        DesiredShiftMessage = MIMEMultipart()
        DesiredShiftMessage['Subject'] = message_subject
        DesiredShiftMessage['To'] = message_receiver_address
        DesiredShiftMessage['From'] = message_sender_address
    
        DesiredShiftMessage.attach(MIMEText(message_body, 'plain'))
    
        return DesiredShiftMessage

    #================================================================================== 
    #メール送信するための関数

    def SendEmail(self, EmailDeta, address_and_password):
    
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(address_and_password[0].strip(), address_and_password[1].strip())
        server.send_message(EmailDeta)
        server.close()
    
#==================================================================================
#MAIN部分

if __name__ == "__main__":
    app = App()
    app.mainloop()
