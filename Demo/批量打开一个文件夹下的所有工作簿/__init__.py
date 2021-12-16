import os
import xlwings as xw

file_path = os.getcwd()
file_list = os.listdir(file_path)
app = xw.App(visible=True, add_book=False)
for i in file_list:
    if os.path.splitext(i)[1] == '.xlsx' or os.path.splitext(i)[1] == '.xls':
        print(file_path + '\\' + i)
        app.books.open(file_path + '\\' + i)

# 列出文件夹下所有文件和子文件夹的名称
file_path = os.getcwd()
file_list = os.listdir(file_path)
for x in file_list:
    print(x)
