# 创建工作簿
import os
import xlwings as xw

# app = xw.App(visible=True, add_book=False)
# 启动Excel程序窗口，但不新建工作簿
# 参数visible用于设置Excel程序窗口的可见性
# add_book用于设置启动Excel程序窗口后是否新建工作簿，如果为True，表示新建一个工作簿
# workbook = app.books.add()

# 保存工作簿

# path = os.getcwd()
# workbook.save(path + '\\example.xlsx')
# workbook.close()
# app.quit()

# 打开工作簿
# 需要注意的是，指定的工作簿必须真实存在，并且不能处于已打开的状态。
# app = xw.App(visible=True, add_book=False)
# workbook = app.books.open(r'example.xlsx')

# 操控工作表和单元格

# 修改工作簿Sheet1  A1单元格的值位编号
# worksheet = workbook.sheets['Sheet1']
# worksheet.range('A1').value = '编号'

# 在打开工作簿重添加一个工作表
# worksheet = workbook.sheets.add('产品统计表')

''' 
app = xw.App(visible=True, add_book=False)
workbook = app.books.add()
worksheet = workbook.sheets.add('TestSheet')
worksheet.range('A1').value = '编号'
workbook.save(r'test.xlsx')
workbook.close()
app.quit()
'''


