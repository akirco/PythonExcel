# -*- coding:utf-8 -*-
import os
import win32com.client as win32

# 输入目录
inputdir = '\\input'
# 输出目录
outputdir = '\\output'
if not os.path.exists(outputdir):
    os.mkdir(outputdir)

# 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
for parent, dirnames, filenames in os.walk(inputdir):
    for fn in filenames:
        if fn.split('.')[-1] == "xls":
            filedir = os.path.join(parent, fn)
            print(filedir)
            print(fn)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56
            wb.SaveAs((os.path.join(outputdir, fn.replace('xls', 'xlsx'))), FileFormat=51)
            wb.Close()
            excel.Application.Quit()
