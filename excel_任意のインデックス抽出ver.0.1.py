import openpyxl
from tkinter import *

def GetExcel(cellnum):
    nums = int(cellnum)
    workbook = openpyxl.load_workbook("C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連\\1.作業リスト\\20240301作業リスト福田次郎 - コピー.xlsx")
    sheet = workbook['福田次郎動画一覧']
    class Cell:
        def getTitle(self,num):
            self.title = sheet.cell(row=num+1, column=2).value
            return self.title
        def getUrl(self,num):
            self.url = sheet.cell(row=num+1, column=4).value
            return self.url
        def getSlug(self,num):
            self.slug = sheet.cell(row=num+1, column=3).value
            return self.slug
        def getTeacher(self,num):
            self.teacher = sheet.cell(row=num+1, column=5).value
            return self.teacher
        def getDate(self,num):
            self.date = sheet.cell(row=num+1, column=7).value
            return self.date
    cell = Cell()
    
    print()
    print("動画タイトル:", cell.getTitle(nums))
    print("動画URL:", cell.getUrl(nums))
    print("スラッグ:", cell.getSlug(nums))
    print("講師:", cell.getTeacher(nums))
    print("日付:", cell.getDate(nums))
        
root = Tk()
root.title("業務で使うアプリ!")
root.geometry("200x200")

entry = Entry(root)
entry.grid()

def pushed():
    entry_value = entry.get()
    GetExcel(entry_value)


button = Button(root, text="button", command=pushed)
button.grid()

root.mainloop()
