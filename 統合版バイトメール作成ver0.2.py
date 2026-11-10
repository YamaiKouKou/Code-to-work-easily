#統合版バイトメール作成ver0.2

#GUI関連のモジュールインポート
import customtkinter as ctk
from customtkinter import filedialog
#メール関連のモジュールインポート
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
#その他
import os
import datetime
import openpyxl

#---------------------------------------------------------------------------------

class App(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        #---エクセルファイル参照(ファイルがない場合は作成)----------------------------------------
        try:
            self.workbook = openpyxl.load_workbook('バイト情報.xlsx')

        except FileNotFoundError:
            self.workbook = openpyxl.Workbook()

            sheet = self.workbook.active
            sheet.title = 'バイト情報詳細'

        finally:
            self.ws = self.workbook['バイト情報詳細']
            self.workbook.save('バイト情報.xlsx')

        #クラス定数
        self.excelfilename = self.ws.cell(column=2, row=1).value
        self.today = datetime.datetime.now().strftime("%Y%m%d")
        self.shiftlist = []
        
        #タイトル、画面の大きさ設定
        self.title("統合版バイトメール作成ver0.2")
        self.geometry("700x500")
        self.resizable(False, False)
        #画面の配置設定
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        #ナビゲーションフレーム作成----------------------------------------------------------------------------------------------------------------
        self.navigation_frame = ctk.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, rowspan=5, sticky="snew")
        self.navigation_frame_label = ctk.CTkLabel(self.navigation_frame, text="MENU", font=ctk.CTkFont(size=20, weight="bold"))
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)
        #業務開始メール作成ボタン作成
        self.StartTaskButton = ctk.CTkButton(self.navigation_frame, height=40, border_spacing=10, text="業務開始メール作成", corner_radius=0, fg_color="transparent", text_color="gray10", hover_color="gray70", command=self.starttaskbutton_event)
        self.StartTaskButton.grid(row=1, column=0, sticky="ew")
        #作業報告メール作成ボタン作成
        self.FinishTaskButton = ctk.CTkButton(self.navigation_frame, height=40, border_spacing=10, text="作業報告メール作成", corner_radius=0, fg_color="transparent", text_color="gray10", hover_color="gray70", command=self.finishtaskbutton_event)
        self.FinishTaskButton.grid(row=2, column=0, sticky="ew")
        #シフト希望メール作成ボタン作成
        self.ShiftDesireButton = ctk.CTkButton(self.navigation_frame, height=40, border_spacing=10, text="シフト希望メール作成", corner_radius=0, fg_color="transparent", text_color="gray10", hover_color="gray70", command=self.shiftdesirebutton_event)
        self.ShiftDesireButton.grid(row=3, column=0, sticky="ew")
        
        #業務開始メールフレーム作成-----------------------------------------------------------------------------------------------------------
        self.StartTaskFrame = ctk.CTkFrame(self, fg_color="transparent")
        self.StartTaskFrame.grid_columnconfigure(1, weight=1)
        
        self.send_button = ctk.CTkButton(self.StartTaskFrame, height=30, width=60, text="送信!", command=lambda:self.SendEmail("start"))
        self.send_button.grid(row=0, column=0, sticky="w", padx=10, pady=(5,0))

        self.sendfrom_label = ctk.CTkLabel(self.StartTaskFrame, text="宛先：質問解決D.B.運営チーム", height=30, width=250, fg_color="gray70")
        self.sendfrom_label.grid(row=1, column=0, sticky="w", padx=10, pady=(5,0))

        self.sendto_label = ctk.CTkLabel(self.StartTaskFrame, text="送信元：23ec015@ms.dendai.ac.jp", height=30, width=250, fg_color="gray70")
        self.sendto_label.grid(row=2, column=0, sticky="w", padx=10, pady=(5,0))

        self.mailbody_label = ctk.CTkLabel(self.StartTaskFrame, fg_color="gray70", height=130, justify="left")
        self.mailbody_label.grid(row=3, column=0, sticky="w", padx=10, pady=(5,0))
        self.mailbody_label.configure(text="件名：{} 作業開始連絡 井上航輔\n\n 質問解決D.B.運営様\n\nお疲れ様です、本日の作業を開始いたしますのでよろしくお願いいたします。\n\n井上航輔".format(datetime.datetime.now().strftime("%Y%m%d")))
        
        #作業報告メールフレーム作成------------------------------------------------------------------------------------------------------------
        self.FinishTaskFrame = ctk.CTkFrame(self, fg_color="transparent")
        
        self.file_and_sendbutton_frame = ctk.CTkFrame(self.FinishTaskFrame, fg_color="transparent")
        self.file_and_sendbutton_frame.grid(row=0, column=0, sticky="w")

        self.send_button = ctk.CTkButton(self.file_and_sendbutton_frame, height=30, width=60, text="送信!", command=lambda:self.SendEmail("finish"))
        self.send_button.grid(row=0, column=0, padx=10, pady=(5,0))

        self.attachment_file_button = ctk.CTkButton(self.file_and_sendbutton_frame, width=180, text="添付するファイルを変更", command=self.selectfile)
        self.attachment_file_button.grid(row=0, column=1, sticky="ew", pady=(5,0))

        self.sendfrom_label = ctk.CTkLabel(self.FinishTaskFrame, text="宛先：質問解決D.B.運営チーム", height=30, width=250, fg_color="gray70")
        self.sendfrom_label.grid(row=1, column=0, sticky="w", padx=10, pady=(5,0))

        self.sendto_label = ctk.CTkLabel(self.FinishTaskFrame, text="送信元：23ec015@ms.dendai.ac.jp", height=30, width=250, fg_color="gray70")
        self.sendto_label.grid(row=2, column=0, sticky="w", padx=10, pady=(5,0))

        self.mailbody_label = ctk.CTkLabel(self.FinishTaskFrame, fg_color="gray70", height=130, justify="left")
        self.mailbody_label.grid(row=3, column=0, sticky="w", padx=10, pady=(5,0))
        self.mailbody_label.configure(text="件名：{} 作業報告 井上航輔\n\n 質問解決D.B.運営様\n\nお疲れ様です、本日の作業内容を添付しますので確認をよろしくお願いいたします。\n\n井上航輔".format(datetime.datetime.now().strftime("%Y%m%d")))

        self.attachment_file_label = ctk.CTkLabel(self.FinishTaskFrame, fg_color="gray70", justify="left")
        self.attachment_file_label.grid(row=4, column=0, sticky="w", padx=10, pady=(5,0))
        
        #シフト希望メールフレーム作成------------------------------------------------------------------------------------------------------------
        self.ShiftDesireFrame = ctk.CTkFrame(self, fg_color="transparent")
        self.ShiftDesireFrame.grid_columnconfigure(1, weight=1)

        self.desire_and_sendbutton_frame = ctk.CTkFrame(self.ShiftDesireFrame, fg_color="transparent")
        self.desire_and_sendbutton_frame.grid(row=0, column=0, sticky="w")

        self.send_button = ctk.CTkButton(self.desire_and_sendbutton_frame, height=30, width=60, text="送信!")
        self.send_button.grid(row=0, column=0, padx=10, pady=(5,0))

        self.desire_button = ctk.CTkButton(self.desire_and_sendbutton_frame, width=180, text="シフト希望登録フォーム", command=self.formbutton_event)
        self.desire_button.grid(row=0, column=1, sticky="ew", pady=(5,0))

        self.sendfrom_label = ctk.CTkLabel(self.ShiftDesireFrame, text="宛先：質問解決D.B.運営チーム", height=30, width=250, fg_color="gray70")
        self.sendfrom_label.grid(row=1, column=0, sticky="w", padx=10, pady=(5,0))

        self.sendto_label = ctk.CTkLabel(self.ShiftDesireFrame, text="送信元：23ec015@ms.dendai.ac.jp", height=30, width=250, fg_color="gray70")
        self.sendto_label.grid(row=2, column=0, sticky="w", padx=10, pady=(5,0))

        self.mailbody_label = ctk.CTkLabel(self.ShiftDesireFrame, fg_color="gray70", height=130, justify="left")
        self.mailbody_label.grid(row=3, column=0, sticky="w", padx=10, pady=(5,0))
        self.mailbody_label.configure(text="件名：{} 週の作業希望日程日時 井上航輔\n\n お疲れ様です、井上航輔です。\n\n{} 週の作業希望日程日時を送付します。".format(self.today, self.today))

        self.desire_label = ctk.CTkLabel(self.ShiftDesireFrame, fg_color="gray70", justify="left")
        self.desire_label.grid(row=4, column=0, sticky="w", padx=10, pady=(5,0))

        #シフト希望フォームフレーム作成------------------------------------------------------------------------------------------------------------
        self.DesireFormFrame = ctk.CTkFrame(self, fg_color="transparent")

        #self.WeekSegmentButton = ctk.CTkSegmentedButton(self.DesireFormFrame, values=["1つ前週", "デフォルト週", "1つ後週"], command=self.SegCallback)
        self.WeekSegmentButton = ctk.CTkSegmentedButton(self.DesireFormFrame, values=["1つ前週", "デフォルト週", "1つ後週"])
        self.WeekSegmentButton.grid(row=0, column=0, padx=10, pady=(5,0))
        

        self.ReturnDesireButton = ctk.CTkButton(self.DesireFormFrame, text="保存して戻る", command=self.returndesireframe)
        self.ReturnDesireButton.grid(row=1, column=0, sticky="w", padx=10, pady=(5,0))



        #デフォルト設定------------------------------------------------------------------------------------------------------------------------
        #最初は作業開始メールフレームに設定
        self.StartTaskFrame.grid(row=0, column=1, sticky="snew")
        self.StartTaskButton.configure(fg_color="gray75")
        #作業報告の添付ファイル設定
        self.attachment_file_label.configure(text="添付ファイル："if self.excelfilename is None else "添付ファイル：" + os.path.split(self.excelfilename)[1])
        #作業希望日程の設定
        self.desire_label.configure(text="希望日程" if self.shiftlist is None else "希望日程" + "\n".join(self.shiftlist))
        #セグメントの初期設定は"デフォルト週"
        self.WeekSegmentButton.set("デフォルト週")
        

    #作業報告における、ファイル選択ダイアログ表示関数
    def selectfile(self):
        types = [('Excel ファイル', '*.xlsx')]
        dirc = 'C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連\\1.作業リスト' 
        filename = filedialog.askopenfilename(filetypes=types, initialdir=dirc)
        #excelファイルに指定されたファイルをデフォルト化する関数
        self.excelfilename = filename
        
        self.attachment_file_label.configure(text="添付ファイル：" + os.path.basename(self.excelfilename))
        
    #ボタンの色, フレーム変更関数
    def select_by_name(self, name):
        self.StartTaskButton.configure(fg_color="gray75" if name == "start" else "transparent")
        self.FinishTaskButton.configure(fg_color="gray75" if name == "finish" else "transparent")
        self.ShiftDesireButton.configure(fg_color="gray75" if name == "desire" or name == "form" else "transparent")
        
        if name == "start":
            self.StartTaskFrame.grid(row=0, column=1, sticky="snew")
        else:
            self.StartTaskFrame.grid_forget()
            
        if name == "finish":
            self.FinishTaskFrame.grid(row=0, column=1, sticky="snew")
        else:
            self.FinishTaskFrame.grid_forget()
            
        if name == "desire":
            self.ShiftDesireFrame.grid(row=0, column=1, sticky="snew")
        else:
            self.ShiftDesireFrame.grid_forget()

        if name == "form":
            self.DesireFormFrame.grid(row=0, column=1, sticky="snew")
        else:
            self.DesireFormFrame.grid_forget()

    #メニューフレームの各ボタンイベント関数
    def starttaskbutton_event(self):
        self.select_by_name("start")
        
    def finishtaskbutton_event(self):
        self.select_by_name("finish")
        
    def shiftdesirebutton_event(self):
        self.select_by_name("desire")

    def formbutton_event(self):
        self.select_by_name("form")

    def returndesireframe(self):
        self.select_by_name("desire")

    #シフト希望フォームのセグメントにおける関数
    #def SegCallback(self, value):


    #メール作成関数
    def CreateEmail(self, name):
        MailMessage = MIMEMultipart()

        message_sender_address = "23ec015@ms.dendai.ac.jp"
        message_receiver_address = message_sender_address
        #message_receiver_address = "db_info@risu-kobetsu.co.jp"

        if name == "start":
            message_subject = self.today + " 作業開始連絡 井上航輔"
            message_body = "質問解決D.B.運営様\n\nお疲れ様です、本日の作業を開始いたしますのでよろしくお願いいたします。\n\n井上航輔"

        elif name == "finish":
            if self.excelfilename != "":
                message_subject = self.today + " 作業報告 井上航輔"
                message_body = "質問解決D.B.運営様\n\nお疲れ様です、本日の作業内容を添付しますので確認をよろしくお願いいたします。\n\n井上航輔"
                
                excel = MIMEApplication(open(self.excelfilename, 'rb').read())
                excel.add_header('Content-Disposition', 'attachment', filename=os.path.basename(self.excelfilename))
                MailMessage.attach(excel)

                #excelファイルに添付ファイルのパスを記録して、excelファイルを保存
                self.ws.cell(column=2, row=1).value = self.excelfilename
                self.workbook.save('バイト情報.xlsx')
            else:
                print("failed to send E-mail.")

        #elif name == "desire":

        MailMessage['Subject'] = message_subject
        MailMessage['To'] = message_receiver_address
        MailMessage['From'] = message_sender_address

        MailMessage.attach(MIMEText(message_body, 'plain'))

        return MailMessage

    #メール送信関数
    def SendEmail(self, name):
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()

        server.login("23ec015@ms.dendai.ac.jp", "@Kousuke20043711")
        server.send_message(self.CreateEmail(name))
        server.close()
        print("sent E-mail correctly.")
            
if __name__ == "__main__":
        
    app = App()
    app.mainloop()
