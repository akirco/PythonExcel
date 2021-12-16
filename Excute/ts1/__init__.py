import multiprocessing

import numpy as np
import xlwings as xw
import pandas as pd
from openpyxl import *
import os.path
import win32com.client as win32
from shutil import copy
from datetime import datetime
import time
import shutil
import zipfile
from gooey import Gooey, GooeyParser

""" 

打开指定所有工作簿(待完善)
按照指定格式筛选数据（待完善）
    是否将所有表汇集到一个工作簿上(待确认)
删除指定列数据（已完成）
按指定列排序（尼玛，需要重构数据，分钟&秒全部映射成秒，进行排序）
设置单元格格式（easy）

程序打包（okay）
功能扩展（想屁吃）

# 208 238 行代码决定由id还是用户昵称筛选

筛选完成后删除xls xlsx input目录下文件
output文件夹下文件压缩

"""


# @Gooey(
#     program_name='微赞数据筛选',
# )
def init():
    # parser = GooeyParser(
    #     description="程序执行前须知：\n1.务必确认电脑上有正版Excel\n2.务必检查文件名是否以【报名】或【话题】开头\n3.请确认下载好的文件格式正确，即【.xls】或【.xlsx】(没有中宏病毒)")
    # parser.add_argument('num', metavar='必填参数', help='请输入您要归类的直播(例：1号线，输入1,点击开始):')
    # args = parser.parse_args()
    # num = args.num
    # # 1号线
    res1_list = (
        101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122,
        123,
        124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145,
        146,
        147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 168,
        169,
        170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191,
        192,
        193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214,
        215,
        216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237,
        238,
        239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256, 257, 258, 259, 260,
        261,
        262, 263, 264, 265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283,
        284,
        285, 286, 287, 288, 289, 290, 291, 292, 293, 294, 295, 296, 297, 298, 299, 300, 301, 302, 303, 304, 305, 306,
        307,
        308, 309, 310, 311, 312, 313, 314, 315, 316, 317, 318, 319, 320, 321, 322, 323, 324, 325, 326, 327, 328, 329,
        330,
        331, 332, 333, 334, 335, 336, 337, 338, 339, 340, 341, 342, 343, 344, 345, 346, 347, 348, 349, 350, 351, 352,
        353,
        354, 355, 356, 357, 358, 359, 360, 361, 362, 363, 364, 365, 366, 367, 368, 369, 370, 371, 372, 373, 374, 375,
        376,
        377, 378, 379, 380, 381, 382, 383, 384, 385, 386, 387, 388, 389, 390, 391, 392, 393, 394, 395, 396, 397, 398,
        399,
        400, 401, 402, 403, 404, 405, 406, 407, 408, 409, 410, 411, 412, 413, 414, 415, 416, 417, 418, 419, 420, 421,
        422,
        423, 424, 425, 426, 427, 428, 429, 430, 431, 432, 433, 434, 435, 436, 437, 438, 439, 440, 441, 442, 443, 444,
        445,
        446, 447, 448, 449, 450, 451, 452, 453, 454, 455, 456, 457, 458, 459, 460, 461, 462, 463, 464, 465, 466, 467,
        468,
        469, 470, 471, 472, 473, 474, 475, 476, 477, 478, 479, 480, 481, 482, 483, 484, 485, 486, 487, 488, 489, 490,
        491,
        492, 493, 494, 495, 496, 497, 498, 499, 500, 501, 502, 503, 504, 505, 506, 507, 508, 509, 510, 511, 512, 513,
        514,
        515, 516, 517, 518, 519, 520, 521, 522, 523, 524, 525, 526, 527, 528, 529, 530, 531, 532, 533, 534, 535, 536,
        537,
        538, 539, 540, 541, 542, 543, 544, 545, 546, 547, 548, 549, 550, 551, 552, 553, 554, 555, 556, 557, 558, 559,
        560,
        561, 562, 563, 564, 565, 566, 567, 568, 569, 570, 571, 572, 573, 574, 575, 576, 577, 578, 579, 580, 581, 582,
        583,
        584, 585, 586, 587, 588, 589, 590, 591, 592, 593, 594, 595, 596, 597, 598, 599, 600, 601, 602, 603, 604, 605,
        606,
        607, 608, 609, 610, 611, 612, 613, 614, 615, 616, 617, 618, 619, 620, 621, 622, 623, 624, 625, 626, 627, 628,
        629,
        630, 631, 632, 633, 634, 635, 636, 637, 638, 639, 640, 641, 642, 643, 644, 645, 646, 647, 648, 649
    )
    # # 2号线
    res2_list = (701, 702, 703, 704, 705, 706, 707, 708, 709, 710, 711, 712, 713, 714, 715, 716, 717, 718, 719, 720,
                 721, 722, 723, 724, 725, 726, 727, 728, 729, 730, 731, 732, 733, 734, 735, 736, 737, 738, 739, 740,
                 741, 742, 743, 744, 745, 746, 747, 748)

    res3_list = (
        801, 802, 803, 804, 805, 806, 807, 808, 809, 810, 811, 812, 813, 814, 815, 816, 817, 818, 819, 820, 821, 822,
        823,
        824, 825)
    num = input('请输入您要归类的直播(例：1号线，输入1回车):')
    print('已接受输入数据,正在处理......')
    if num == '1':
        print('已接受输入数据【1】,正在处理1号线数据......')
        return res1_list, num
    elif num == '2':
        print('已接受输入数据【2】,正在处理2号线数据......')
        return res2_list, num
    elif num == '3':
        print('已接受输入数据【3】,正在处理3号线数据......')
        return res3_list, num


def file_exist():
    file_path = os.getcwd() + '\\XLSX\\'
    file_list = os.listdir(file_path)
    return file_list


'''
i.将xls转换为xlsx
'''


def convert_file():
    input_dir = os.getcwd() + '\\XLS'
    out_dir = os.getcwd() + '\\XLSX'
    file_list = os.listdir(input_dir)
    for i in file_list:
        list_path = input_dir + '\\' + i
        part_name = os.path.splitext(i)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(list_path)
        # xlsx: FileFormat=51
        # xls:  FileFormat=56
        wb.SaveAs((os.path.join(out_dir, part_name[0].replace('xls', 'xlsx'))), FileFormat=51)
        wb.Close()
        excel.Application.Quit()


''' 
删除用户数据表中的无用数据
'''


def delete_data():
    file_path = os.getcwd()
    xlsx_path = file_path + '\\XLSX'
    input_path = file_path + '\\input'
    file_list = os.listdir(xlsx_path)
    for i in file_list:
        if i.startswith('话题'):
            del_path = xlsx_path + '\\' + i
            wb = load_workbook(del_path)
            worksheet = wb.active
            worksheet.delete_cols(3)
            worksheet.delete_cols(3)
            worksheet.delete_cols(3)
            worksheet.delete_cols(3)
            worksheet.delete_cols(3)
            worksheet.delete_cols(7)
            worksheet.delete_cols(8)
            worksheet.delete_cols(8)
            worksheet.delete_cols(8)
            wb.save(input_path + '\\用户数据-已筛选.xlsx')


def del_duplicate_rows():
    file_path = os.getcwd()
    xlsx_path = file_path + '\\XLSX'
    file_list = os.listdir(xlsx_path)
    for i in file_list:
        if i.startswith('用户数据'):
            path = xlsx_path + '\\' + i
            print(path)
            data = pd.DataFrame(pd.read_excel(path, 'Sheet0'))
            data = data.drop_duplicates('用户昵称')
            writer = pd.ExcelWriter(path)
            data.to_excel(writer)
            writer.save()


# 将多个工作簿合并到一个工作簿
# filename  需要筛选文件名
# filter_id  通过？id筛选
def select_data_by_id(filter_id):
    file_path = os.getcwd()
    xlsx_path = file_path + '\\XLSX'
    input_path = file_path + '\\input'
    file_list = os.listdir(xlsx_path)
    app = xw.App(visible=False, add_book=False)
    for i in file_list:
        # print("这里是：", i)
        if i.startswith('报名'):
            wb = app.books.open(xlsx_path + '\\' + i)  # 打开所有
            worksheet = wb.sheets
            data = []
            df = pd.DataFrame
            for j in worksheet:
                # print('遍历', j)
                values = j.range('A1').expand().options(df).value
                # print('编号', values['请务必正确填写群主要求的编号'])
                filtered = values[values['门店编码'] == filter_id]  # 可扩展
                if filtered.empty:
                    pass
                if not filtered.empty:
                    data.append(filtered)
            if not data:
                pass
            else:
                new_workbook = xw.books.add()
                new_worksheet = new_workbook.sheets.add('Sheet0')
                new_worksheet.range('A1').value = pd.concat(data, ignore_index=False)
                new_workbook.save(input_path + '\\' + '报名记录统计-' + filter_id + '.xlsx')  # 待修改

            wb.close()
    app.quit()
    app.kill()


def copy_file_to_input():
    file_path = os.getcwd()
    xlsx_path = file_path + '\\XLSX'
    output_path = file_path + '\\input\\报名记录统计-1258.xlsx'
    file_list = os.listdir(xlsx_path)
    # print(file_list)
    for i in file_list:
        if i.startswith('报名'):
            path = xlsx_path + '\\' + i
            rename_worksheet(path, path, 'sheet1', 'Sheet0')
            copy(path, output_path)


''' 
重命名表名
'''


def rename_worksheet(input_path, output_path, origin_name, final_name):
    app = xw.App(visible=False, add_book=False)
    # input_path = os.getcwd() + '\\input\\example.xlsx'
    # output_path = os.getcwd() + '\\output\\example.xlsx'
    # origin_name = 'Sheet1'
    # final_name = 'Sheet0'
    wb = app.books.open(input_path)
    worksheets = wb.sheets
    for x in range(len(worksheets))[:5]:  # 利用切片来选中工作表
        worksheets[x].name = worksheets[x].name.replace(origin_name, final_name)
    wb.save(output_path)
    app.quit()


def get_user_list(sel_id):
    file_path = os.getcwd()
    input_path = file_path + '\\input'
    file_list = os.listdir(input_path)
    for i in file_list:
        # print(i)
        filter_id = os.path.splitext(i)[0].split('-')[1]
        filter_id = filter_id.split('\n')
        app = xw.App(visible=False, add_book=False)
        for j in filter_id:
            if j == sel_id:
                sel_path = input_path + '\\报名记录统计-' + sel_id + '.xlsx'
                # print(sel_path)
                wb = app.books.open(sel_path)
                worksheet = wb.sheets['Sheet0']
                values = worksheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
                user_list = values['用户id']
                # print(user_list)
                # user_list = values['用户id']
                wb.close()
                return user_list
        app.quit()
        app.kill()


def get_user_data(sel_id):
    sel_list = []
    user_list = get_user_list(sel_id)
    total_user_list = []
    dataframe = pd.DataFrame()
    if user_list is None:
        print(sel_id, '查询不到此数据,正在跳过...')
        pass
    else:
        print("正在处理...用户数据-", sel_id + '.xlsx')
        for user in user_list:
            # 取决去微赞数据:id是否有结尾.0
            user = str(user).split('.')[0]
            sel_list.append(user)
        file_path = os.getcwd()
        input_path = file_path + '\\input'
        com_path = input_path + '\\用户数据-已筛选.xlsx'
        final_path = file_path + '\\output\\用户数据-' + sel_id + '.xlsx'
        data = pd.read_excel(com_path)
        # print(data)
        result = np.array(data)
        lens = len(result)
        # print(lens)
        # 取出报名表有的而话题表没有的数据

        if lens >= 0:
            values = pd.DataFrame(result, index=[x for x in range(1, lens + 1)], columns=list('ABCDEFG'))
            dataframe = values[values['A'].isin(sel_list)]
            dataframe = pd.DataFrame(dataframe)
            # print(df)
            if dataframe.empty:
                print('根据报名表id在话题表中查不到指定数据\n原因可能是：您在设置门店编码时使用的是文本框\n为避免出现类似问题，门店编码组件请选择数字！！')
                '''处理报名表中个别数据在话题表中查询不到的情况'''
                print('程序即将转换id的类型进行筛选！')
                for i in result[:, 0:1]:
                    total_user_list.append(i[0])
                ntul = [str(x) for x in total_user_list]
                # print('ntul:', ntul, '\n', len(ntul))
                # print('sel_list:', sel_list, '\n', len(sel_list))

                # print('交集数据：', list(set(ntul).intersection(set(sel_list))))
                new_sel_list = list(set(ntul).intersection(set(sel_list)))
                new_sel_list = [int(y) for y in new_sel_list]
                '''处理报名表中个别数据在话题表中查询不到的情况'''
                # print(values['A'])
                # print(new_sel_list)
                dataframe = values[values['A'].isin(new_sel_list)]
                # print(dataframe)
                dataframe = pd.DataFrame(dataframe)
                if dataframe.empty:
                    print('转换id类型筛选也出现了错误，请检查话题表是否有数据或话题表是否相互对应！！')
                    pass
                # pass
                # df.rename(
                #     columns={'A': '用户昵称', 'B': '性别', 'C': '真是姓名', 'D': '联系号码', 'E': '收集来源', 'F': '用户状态',
                #              'G': 'IP', 'H': '地区', 'I': '首次观看直播时间', 'J': '最近观看直播时间', 'K': '直播观看时长'},
                #     inplace=True)
                dataframe.rename(
                    columns={'A': '用户id', 'B': '用户昵称', 'C': 'IP', 'D': '地区', 'E': '首次观看直播时间', 'F': '最近观看直播时间',
                             'G': '直播观看时长'},
                    inplace=True)
                writer = pd.ExcelWriter(final_path)
                dataframe.to_excel(writer, index=False)
                writer.save()
        else:
            print('程序中断执行：请检查话题数据表中是否有数据！！！')
            exit()


def get_extra_list(id):
    '''
    针对res3_list报名表->门店编码是》数字《
    '''
    res3_list = (
        101.0, 102.0, 103.0, 104.0, 105.0, 106.0, 107.0, 108.0, 109.0, 110.0, 111.0, 112.0, 113.0, 114.0, 115.0, 116.0,
        117.0, 118.0, 119.0, 120.0, 121.0, 122.0, 123.0, 124.0, 125.0, 126.0, 127.0, 128.0, 129.0, 130.0, 131.0, 132.0,
        133.0, 134.0, 135.0, 136.0, 137.0, 138.0, 139.0, 140.0, 141.0, 142.0, 143.0, 144.0, 145.0, 146.0, 147.0, 148.0,
        149.0, 150.0, 151.0, 152.0, 153.0, 154.0, 155.0, 156.0, 157.0, 158.0, 159.0, 160.0, 161.0, 162.0, 163.0, 164.0,
        165.0, 166.0, 167.0, 168.0, 169.0, 170.0, 171.0, 172.0, 173.0, 174.0, 175.0, 176.0, 177.0, 178.0, 179.0, 180.0,
        181.0, 182.0, 183.0, 184.0, 185.0, 186.0, 187.0, 188.0, 189.0, 190.0, 191.0, 192.0, 193.0, 194.0, 195.0, 196.0,
        197.0, 198.0, 199.0, 200.0, 201.0, 202.0, 203.0, 204.0, 205.0, 206.0, 207.0, 208.0, 209.0, 210.0, 211.0, 212.0,
        213.0, 214.0, 215.0, 216.0, 217.0, 218.0, 219.0, 220.0, 221.0, 222.0, 223.0, 224.0, 225.0, 226.0, 227.0, 228.0,
        229.0, 230.0, 231.0, 232.0, 233.0, 234.0, 235.0, 236.0, 237.0, 238.0, 239.0, 240.0, 241.0, 242.0, 243.0, 244.0,
        245.0, 246.0, 247.0, 248.0, 249.0, 250.0, 251.0, 252.0, 253.0, 254.0, 255.0, 256.0, 257.0, 258.0, 259.0, 260.0,
        261.0, 262.0, 263.0, 264.0, 265.0, 266.0, 267.0, 268.0, 269.0, 270.0, 271.0, 272.0, 273.0, 274.0, 275.0, 276.0,
        277.0, 278.0, 279.0, 280.0, 281.0, 282.0, 283.0, 284.0, 285.0, 286.0, 287.0, 288.0, 289.0, 290.0, 291.0, 292.0,
        293.0, 294.0, 295.0, 296.0, 297.0, 298.0, 299.0, 300.0, 301.0, 302.0, 303.0, 304.0, 305.0, 306.0, 307.0, 308.0,
        309.0, 310.0, 311.0, 312.0, 313.0, 314.0, 315.0, 316.0, 317.0, 318.0, 319.0, 320.0, 321.0, 322.0, 323.0, 324.0,
        325.0, 326.0, 327.0, 328.0, 329.0, 330.0, 331.0, 332.0, 333.0, 334.0, 335.0, 336.0, 337.0, 338.0, 339.0, 340.0,
        341.0, 342.0, 343.0, 344.0, 345.0, 346.0, 347.0, 348.0, 349.0, 350.0, 351.0, 352.0, 353.0, 354.0, 355.0, 356.0,
        357.0, 358.0, 359.0, 360.0, 361.0, 362.0, 363.0, 364.0, 365.0, 366.0, 367.0, 368.0, 369.0, 370.0, 371.0, 372.0,
        373.0, 374.0, 375.0, 376.0, 377.0, 378.0, 379.0, 380.0, 381.0, 382.0, 383.0, 384.0, 385.0, 386.0, 387.0, 388.0,
        389.0, 390.0, 391.0, 392.0, 393.0, 394.0, 395.0, 396.0, 397.0, 398.0, 399.0, 400.0, 401.0, 402.0, 403.0, 404.0,
        405.0, 406.0, 407.0, 408.0, 409.0, 410.0, 411.0, 412.0, 413.0, 414.0, 415.0, 416.0, 417.0, 418.0, 419.0, 420.0,
        421.0, 422.0, 423.0, 424.0, 425.0, 426.0, 427.0, 428.0, 429.0, 430.0, 431.0, 432.0, 433.0, 434.0, 435.0, 436.0,
        437.0, 438.0, 439.0, 440.0, 441.0, 442.0, 443.0, 444.0, 445.0, 446.0, 447.0, 448.0, 449.0, 450.0, 451.0, 452.0,
        453.0, 454.0, 455.0, 456.0, 457.0, 458.0, 459.0, 460.0, 461.0, 462.0, 463.0, 464.0, 465.0, 466.0, 467.0, 468.0,
        469.0, 470.0, 471.0, 472.0, 473.0, 474.0, 475.0, 476.0, 477.0, 478.0, 479.0, 480.0, 481.0, 482.0, 483.0, 484.0,
        485.0, 486.0, 487.0, 488.0, 489.0, 490.0, 491.0, 492.0, 493.0, 494.0, 495.0, 496.0, 497.0, 498.0, 499.0, 500.0,
        501.0, 502.0, 503.0, 504.0, 505.0, 506.0, 507.0, 508.0, 509.0, 510.0, 511.0, 512.0, 513.0, 514.0, 515.0, 516.0,
        517.0, 518.0, 519.0, 520.0, 521.0, 522.0, 523.0, 524.0, 525.0, 526.0, 527.0, 528.0, 529.0, 530.0, 531.0, 532.0,
        533.0, 534.0, 535.0, 536.0, 537.0, 538.0, 539.0, 540.0, 541.0, 542.0, 543.0, 544.0, 545.0, 546.0, 547.0, 548.0,
        549.0, 550.0, 551.0, 552.0, 553.0, 554.0, 555.0, 556.0, 557.0, 558.0, 559.0, 560.0, 561.0, 562.0, 563.0, 564.0,
        565.0, 566.0, 567.0, 568.0, 569.0, 570.0, 571.0, 572.0, 573.0, 574.0, 575.0, 576.0, 577.0, 578.0, 579.0, 580.0,
        581.0, 582.0, 583.0, 584.0, 585.0, 586.0, 587.0, 588.0, 589.0, 590.0, 591.0, 592.0, 593.0, 594.0, 595.0, 596.0,
        597.0, 598.0, 599.0, 600.0, 601.0, 602.0, 603.0, 604.0, 605.0, 606.0, 607.0, 608.0, 609.0, 610.0, 611.0, 612.0,
        613.0, 614.0, 615.0, 616.0, 617.0, 618.0, 619.0, 620.0, 621.0, 622.0, 623.0, 624.0, 625.0, 626.0, 627.0, 628.0,
        629.0, 630.0, 631.0, 632.0, 633.0, 634.0, 635.0, 636.0, 637.0, 638.0, 639.0, 640.0, 641.0, 642.0, 643.0, 644.0,
        645.0, 646.0, 647.0, 648.0, 649.0, 701.0, 702.0, 703.0, 704.0, 705.0, 706.0, 707.0, 708.0, 709.0, 710.0, 711.0,
        712.0, 713.0, 714.0, 715.0, 716.0, 717.0, 718.0, 719.0, 720.0, 721.0, 722.0, 723.0, 724.0, 725.0, 726.0, 727.0,
        728.0, 729.0, 730.0, 731.0, 732.0, 733.0, 734.0, 735.0, 736.0, 737.0, 738.0, 739.0, 740.0, 741.0, 742.0, 743.0,
        744.0, 745.0, 746.0, 747.0, 748.0, 749.0, 801.0, 802.0, 803.0, 804.0, 805.0, 806.0, 807.0, 808.0, 809.0, 810.0,
        811.0, 812.0, 813.0, 814.0, 815.0, 816.0, 817.0, 818.0, 819.0, 820.0, 821.0, 822.0, 823.0, 824.0, 825.0)

    '''
    针对res3_list报名表->门店编码是》文本框《
    '''
    # res3_list = [str(x) for x in res3_list]
    file_path = os.getcwd()
    xlsx_path = file_path + '\\XLSX'
    file_list = os.listdir(xlsx_path)
    extra_path = file_path + '\\extra\\报名记录-错误-' + id + '.xlsx'
    for i in file_list:
        if i.startswith('报名'):
            sel_path = xlsx_path + '\\' + i
            data = pd.read_excel(sel_path)
            result = np.array(data)
            # print(result)
            lens = len(result)
            # print(lens)
            if lens > 0:
                values = pd.DataFrame(result, index=[x for x in range(1, lens + 1)], columns=list('ABCDEFGH'))
                df = values[~values['H'].isin(res3_list)]
                extra_data_list = df['A']
                extra_num_list = df['H']
                df = pd.DataFrame(df)
                df.rename(
                    columns={'A': '用户id', 'B': '用户昵称', 'C': '话题来源', 'D': '来源', 'E': '姓名', 'F': '电话',
                             'G': '报名时间', 'H': '门店编码'},
                    inplace=True)
                writer = pd.ExcelWriter(extra_path)
                df.to_excel(writer, index=False)
                writer.save()
                return extra_data_list, extra_num_list


def get_extra_data(id):
    sel_list = []
    extra_list = []
    dataframe = pd.DataFrame()
    result = get_extra_list(id)
    if result is None:
        print('请检查[XLS]目录中是否有文件！！！')
        exit()
    else:
        id_list = result[0]
        # print(len(id_list))
        # print(id_list)
        # num_list = result[1]
        # print('num_list:\n', num_list)
        if id_list is None:
            print('无垃圾数据...')
            pass
        else:
            print("正在处理...垃圾数据")
            for i in id_list:
                # id尼玛又变成字符串了
                i = str(i)
                sel_list.append(i)
                # print(i)
                # print(type(i))
            file_path = os.getcwd()
            input_path = file_path + '\\input'
            com_path = input_path + '\\用户数据-已筛选.xlsx'
            extra_path = file_path + '\\extra\\用户数据-错误-' + id + '.xlsx'
            data = pd.read_excel(com_path)
            user_data = np.array(data)
            # print(u_data)
            # num_data = np.array(num_list)
            # print(num_data)
            # print(len(num_data))
            # sel_list = [int(x) for x in sel_list]
            lens = len(user_data)
            if lens >= 0:
                values = pd.DataFrame(user_data, index=[x for x in range(1, lens + 1)], columns=list('ABCDEFG'))
                dataframe = values[values['A'].isin(sel_list)]
                dataframe = pd.DataFrame(dataframe)
                # df.insert(7, 'H', num_data, allow_duplicates=True)
                if dataframe.empty:
                    print('筛选数据类型可能有误，正在转换数据类型并求数据交集')
                    for i in user_data[:, 0:1]:
                        extra_list.append(i[0])
                    ntul = [str(x) for x in extra_list]
                    # print('ntul:', ntul, '\n', len(ntul))
                    # print('sel_list:', sel_list, '\n', len(sel_list))

                    # print('交集数据：', list(set(ntul).intersection(set(sel_list))))
                    new_sel_list = list(set(ntul).intersection(set(sel_list)))
                    new_sel_list = [int(y) for y in new_sel_list]
                    '''处理报名表中个别数据在话题表中查询不到的情况'''
                    # print(values['A'])
                    dataframe = values[values['A'].isin(new_sel_list)]
                    # print(dataframe)
                    dataframe = pd.DataFrame(dataframe)
                    if dataframe.empty:
                        print('转换id类型筛选也出现了错误，请检查话题表是否有数据或话题表是否相互对应！！')
                        exit()
                    # df.rename(
                    #     columns={'A': '用户昵称', 'B': '性别', 'C': '真是姓名', 'D': '联系号码', 'E': '收集来源', 'F': '用户状态',
                    #              'G': 'IP', 'H': '地区', 'I': '首次观看直播时间', 'J': '最近观看直播时间', 'K': '直播观看时长'},
                    #     inplace=True)
                    dataframe.rename(
                        columns={'A': '用户id', 'B': '用户昵称', 'C': 'IP', 'D': '地区', 'E': '首次观看直播时间', 'F': '最近观看直播时间',
                                 'G': '直播观看时长', 'H': '门店编码'},
                        inplace=True)
                    # print(extra_path)
                    writer = pd.ExcelWriter(extra_path)
                    dataframe.to_excel(writer, index=False)
                    writer.save()
            else:
                print('程序中断执行：请检查话题数据表中是否有数据！！！')
                exit()


def del_extra_files():
    print('正在清理处理过程中的额外文件...')
    file_dir = os.getcwd()
    file_list = os.listdir(file_dir)
    for i in file_list:
        if os.path.splitext(i)[1] != '.py' and os.path.splitext(i)[1] != '.zip' and os.path.splitext(i)[
            1] != '.exe' and i != 'extra' and i != '.git' and i != '.gitattributes' and os.path.splitext(i)[1] != '.md':
            shutil.rmtree(i)
            os.mkdir(i)


def compress_files(id):
    compress_dir = '.\\output'
    compressed_dir = compress_dir + '_' + id + '.zip'
    print('正在压缩文件...', compressed_dir)
    z = zipfile.ZipFile(compressed_dir, 'w', zipfile.ZIP_DEFLATED)
    for dirpath, dirnames, filenames in os.walk(compress_dir):
        file_path = dirpath.replace(compress_dir, '')
        file_path = file_path and file_path + os.sep or ''
        for filename in filenames:
            z.write(os.path.join(dirpath, filename), file_path + filename)
        z.close()


'''
def del_extra_data():
    sel_id = get_user_data('72100')
    file_path = os.getcwd()
    input_path = file_path + '\\input'
    out_path = file_path + '\\output'
    del_path = input_path + '\\用户数据-' + sel_id + '.xlsx'
    workbook = load_workbook(del_path)
    worksheet = workbook.active
    worksheet.delete_cols(1)
    worksheet.delete_rows(1)
    workbook.save(out_path + '\\用户数据.xlsx')
'''

if __name__ == '__main__':
    multiprocessing.freeze_support()
    try:
        # print('程序执行前，务必检查文件名是否以【报名】或【话题】开头')
        # print('注：报名即报名表，话题即直播话题数据表，如果开头不是以上两种，修改即可')
        # print('确认完成后，将1号线或2号线的【.xls】文件放入【XLS】目录')
        num_list = init()
        t1 = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        t1 = datetime.strptime(t1, "%Y-%m-%d %H:%M:%S")
        print('当前时间是：', t1)
        # # 转换文件
        flag = file_exist()
        if not flag:
            convert_file()
        # 删除用户数据表中的无用数据
        delete_data()
        # 删除重复行数据
        # del_duplicate_rows()
        # 筛选固定id的数据
        # select_data_by_id('714')
        # get_user_data('714')
        # 获取垃圾数据
        # get_extra_data('1')
        # get_extra_list('1')
        print(num_list[1])
        get_extra_data(num_list[1])
        get_extra_list(num_list[1])
        for i in num_list[0]:
            i = str(i)
            select_data_by_id(i)
            get_user_data(i)
        t2 = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        t2 = datetime.strptime(t2, "%Y-%m-%d %H:%M:%S")
        print('当前时间是：', t2)
        cost = (t2 - t1).seconds
        minutes = int(cost / 60)
        seconds = cost - 60 * (cost // 60)
        if cost >= 60:
            print('恭喜您花费了' + str(minutes) + '分' + str(seconds) + '秒把数据筛选完成了')
        else:
            print('恭喜您花费了' + str(seconds) + '秒把数据筛选完成了')

        # compress_files(1, '.\\output')
        compress_files(num_list[1])
        del_extra_files()
    except KeyboardInterrupt as e:
        pass
