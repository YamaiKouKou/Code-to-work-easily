from tkinter import *
import openpyxl
import datetime

dt = datetime.datetime.now()
dat = dt.strftime("%Y/%m/%d")
dt_date = datetime.datetime.strptime(dat, "%Y/%m/%d")
workbook = openpyxl.load_workbook("C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連\\1.作業リスト\\20240928福田次郎.xlsx")
sheet = workbook['福田次郎動画一覧井上さん0705']

class Cell():
    def getTeacher(self, num):
        self.teacher = sheet.cell(row = num-1410, column = 5).value
        return self.teacher

    def getTitle(self,num):
        self.title = sheet.cell(row = num-1410, column = 2).value
        return self.title
    
    def getSlug(self,num):
        self.slug = sheet.cell(row = num-1410, column = 3).value
        return self.slug
    
    def getDate(self,num):
        self.date = sheet.cell(row = num-1410, column = 7).value
        return self.date

cell = Cell()

def FillCell(num):
    for i in range(9, 18, 1):
        sheet.cell(row = num-1410, column = i).value = "〇"
    sheet.cell(row = num-1410, column = 21).number_format = "yyyy/mm/dd"
    sheet.cell(row = num-1410, column = 21).value = dt_date
    sheet.cell(row = num-1410, column = 22).value = "井上航輔"
    workbook.save("C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連\\1.作業リスト\\20240928福田次郎.xlsx")

def PrintCell(num):
    print("動画タイトル：", cell.getTitle(num))
    print("動画スラッグ：", cell.getSlug(num))
    print("講師：", cell.getTeacher(num))
    print("動画公開日：", cell.getDate(num))
    print()
    
root = Tk()
root.title("業務で使うアプリ_ver0.1")
root.geometry("200x200")

entry = Entry(root)
entry.grid()

def print_and_fill():
    entry_value = entry.get()
    int_value = int(entry_value)
    PrintCell(int_value)
    FillCell(int_value)
    
button = Button(root, text = "唯々諾々", command = print_and_fill)
button.grid()

root.mainloop()

