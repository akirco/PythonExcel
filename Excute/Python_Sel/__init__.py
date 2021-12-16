import os
import pandas as pd
import numpy as np
from pathlib import Path
import zipfile
import re
import shutil

from numpy import int64

global final_data
final_data = {}


def unzip():
    path = 'InputData'
    file_list = os.listdir(path)
    for i in file_list:
        file_path = '.\\InputData\\' + i
        f = zipfile.ZipFile(file_path)
        des_dir = file_path[:file_path.index('.zip')][0:-22].replace('(', '')
        for filename in f.namelist():
            result = f.extract(filename, des_dir)
            extracted_path = Path(result)
            extracted_path.rename(str(extracted_path).encode('cp437').decode('gbk'))
        f.close()
        # os.remove(file_path)


def get_user_view_data():
    data_list = []
    num_list = []
    file_numbers = []
    path = '.\\InputData\\'
    sub_path = os.listdir(path)
    dir_list = [path + x for x in sub_path]
    for dir in dir_list:
        if not dir.endswith('.zip'):
            file_list = os.listdir(dir)
            file_number = len(file_list)
            file_numbers.append(file_number)
            for file in file_list:
                if re.search('异常', file) == None:
                    result = file.split('-')[1].split('.')[0]
                    num_list.append(int(result))
                    read_path = dir + '\\' + file
                    data = pd.read_excel(read_path)
                    length = len(data)
                    data_list.append(length)
    return num_list, data_list


def read_template():
    template = '.\\template.xlsx'
    result = pd.read_excel(template, sheet_name=[0, 1, 2])
    num1 = np.array(result[0]['群编码'])
    num2 = np.array(result[1]['群编码'])
    num3 = np.array(result[2]['群编码'])
    return num1, num2, num3


def get_data():
    # 获取筛选的真实数据
    num_list = get_user_view_data()[0]
    data_list = get_user_view_data()[1]
    sel_data = dict(zip(num_list, data_list))
    return num_list, sel_data


def save_data(num_list, sel_data):
    # 模板数据
    # 当模板数据是为小数时使用或nan
    # num1 = [int(str(i).split('.')[0]) for i in read_template()[0] if str(i).split('.')[0] != 'nan']
    num1 = [int(i) for i in read_template()[0]]
    num2 = [int(i) for i in read_template()[1]]
    num3 = [int(i) for i in read_template()[2]]

    fry_list = []
    lxl_list = []
    hhh_list = []
    result1 = list(set(num1).difference(set(num_list)))
    result2 = list(set(num2).difference(set(num_list)))
    result3 = list(set(num3).difference(set(num_list)))

    value_list = [0 for x in range(len(result1))]
    extra_data1 = dict(zip(result1, value_list))
    extra_data2 = dict(zip(result2, value_list))
    extra_data3 = dict(zip(result3, value_list))

    # 没有的数据
    final_data.update(sel_data)
    final_data.update(extra_data1)
    final_data.update(extra_data2)
    final_data.update(extra_data3)

    keys = [i for i in sorted(final_data)]
    values = [final_data[i] for i in sorted(final_data)]
    for i in read_template()[0]:
        fry = get_dict_keys(final_data, int(i))
        fry_list.extend(fry)
    # print(fry_list)
    for j in read_template()[1]:
        lxl = get_dict_keys(final_data, int(j))
        lxl_list.extend(lxl)
    # print(lxl_list)
    for k in read_template()[2]:
        hhh = get_dict_keys(final_data, int(k))
        hhh_list.extend(hhh)
    # print(hhh_list)

    length = [len(num_list)]
    num_list.extend([' ' for x in range(len(keys) - len(num_list))])
    length.extend([' ' for x in range(len(keys) - 1)])
    fry_list.extend(' ' for x in range(len(keys) - len(fry_list)))
    lxl_list.extend(' ' for x in range(len(keys) - len(lxl_list)))
    hhh_list.extend(' ' for x in range(len(keys) - len(hhh_list)))

    df = pd.DataFrame()
    df['群编码'] = keys
    df['观看直播人数'] = values
    df['在跟群编码'] = num_list
    df['总计'] = length
    df['冯茹燕'] = fry_list
    df['刘晓莉'] = lxl_list
    df['韩会会'] = hhh_list
    writer = pd.ExcelWriter('.\data' + '.xlsx')
    df.to_excel(writer, index=False)
    writer.save()


def del_file():
    file_list = os.listdir('.')
    for i in file_list:
        if os.path.splitext(i)[1] != '.py' and os.path.splitext(i)[1] != '.xlsx' and os.path.splitext(i)[1] != '.exe':
            shutil.rmtree(i)
            os.mkdir(i)


def get_dict_keys(dic, values):
    return [v for (k, v) in dic.items() if k == values]


if __name__ == '__main__':
    # res = read_template()
    # print(res)
    unzip()
    res = get_data()
    save_data(res[0], res[1])
    del_file()
