import xlwings as xw
import os

app = xw.App(visible=True, add_book=False)
for x in range(6):
    workbook = app.books.add()
    path = os.getcwd()
    workbook.save(path + f'\\test{x}.xlsx')
    workbook.close()
app.quit()
