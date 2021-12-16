import numpy as np
import xlwings as xw
import pandas as pd
from openpyxl import *
import os.path
import win32com.client as win32
from shutil import copy
from progressbar import *
from tqdm import tqdm
import pathlib


def get_user_list(filename, col_name):
    file_path = os.getcwd()
    input_path = file_path + '\\input'
    app = xw.App(visible=False, add_book=False)
    sel_path = input_path + '\\' + filename + '.xlsx'
    wb = app.books.open(sel_path)
    worksheet = wb.sheets['Sheet0']
    values = worksheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
    user_list = values[col_name]
    wb.close()
    app.quit()
    return user_list


def sel_data():
    list1 = get_user_list('录入数据', '微信名')
    list2 = get_user_list('学员信息', '姓名')
    values = pd.DataFrame(list1)
    df = values[~values['微信名'].isin(list2)]
    print(df)


if __name__ == '__main__':
    sel_data()
