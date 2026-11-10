import os
import openpyxl

os.chdir("C:\\Users\\sukes\\OneDrive\\デスクトップ\\バイト関連\\1.作業リスト")

workbook = openpyxl.load_workbook('20240301作業リスト福田次郎 - コピー.xlsx')
sheet = workbook['福田次郎動画一覧']

class Cell:
    def getTitle(self, num):
        self.title = sheet.cell(row=num+1, column=2).value
        return self.title
    def getUrl(self, num):
        self.url = sheet.cell(row=num+1, column=4).value
        return self.url
    def getSlug(self, num):
        self.slug = sheet.cell(row=num+1, column=3).value
        return self.slug
    def getTeacher(self, num):
        self.teacher = sheet.cell(row=num+1, column=5).value
        return self.teacher
    def getDate(self, num):
        self.date = sheet.cell(row=num+1, column=7).value
        return self.date

cell = Cell()

num = int(input('入力した数字のマスにある(No.の数字)詳細を知れます。'))

print()
print("動画タイトル:", cell.getTitle(num))
print("動画URL:", cell.getUrl(num))
print("スラッグ:", cell.getSlug(num))
print("講師:", cell.getTeacher(num))
print("日付:", cell.getDate(num))
