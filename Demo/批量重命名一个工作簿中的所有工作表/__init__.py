import xlwings as xw
import os

# app = xw.App(visible=False, add_book=False)
# path = os.getcwd()
# workbook = app.books.open(path + '\\example.xlsx')
# worksheets = workbook.sheets
# for i in range(len(worksheets)):
#     worksheets[i].name = worksheets[i].name.replace('test', '1')
# workbook.save(path + '\\example.xlsx')
# app.quit()

# 批量重命名一个工作簿中的部分工作表
app = xw.App(visible=False, add_book=False)
path = os.getcwd()
workbook = app.books.open(path + '\\example.xlsx')
worksheets = workbook.sheets
for i in range(len(worksheets))[:5]:  # 利用切片来选中工作表
    worksheets[i].name = worksheets[i].name.replace('Sheet1', 'Sheet0')
workbook.save(path + '\\example.xlsx')
app.quit()
