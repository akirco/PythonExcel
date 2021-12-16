import numpy as np
import xlwings as xw
import pandas as pd
from openpyxl import *
import os.path
import win32com.client as win32
from shutil import copy
from progressbar import *
from datetime import datetime
import shutil
import zipfile

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


def init():
    # # 1号线
    res1_list = (601, 602, 603, 604, 605, 606, 607, 608, 609, 610, 611, 612, 613, 614, 615, 616, 617, 618, 619, 620,
                 621, 622, 623, 624, 625, 626, 627, 628, 629, 630, 631, 632, 633, 634, 635, 636, 637, 638, 639, 640,
                 641, 642, 643, 644, 645, 646, 647, 648, 649, 650, 651, 652, 653, 654, 655, 656, 657, 658, 659
                 )

    # # 2号线
    res2_list = (701, 702, 703, 704, 705, 706, 707, 708, 709, 710, 711, 712, 713, 714, 715, 716, 717, 718, 719, 720,
                 721, 722, 723, 724, 725, 726, 727, 728, 729, 730, 731, 732, 733, 734, 735, 736, 737, 738, 739, 740,
                 741, 742)
    num = input('请输入您要归类的直播(例：1号线，输入1回车):')
    print('已接受输入数据,正在处理......')
    if num == '1':
        return res1_list, num
    elif num == '2':
        return res2_list, num


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
                # user_list = values['用户昵称']
                # print(user_list)
                user_list = values['用户id']
                wb.close()
                return user_list
        app.quit()
        app.kill()


def get_user_data(sel_id):
    sel_list = []
    user_list = get_user_list(sel_id)
    if user_list is None:
        print(sel_id, '查询不到此数据,正在跳过...')
        pass
    else:
        print("正在处理...用户数据-", sel_id + '.xlsx')
        for user in user_list:
            # 取决去微赞数据:id是否有结尾.0
            user = str(user).split('.')[0]
            sel_list.append(user)
        # print(sel_list)
        file_path = os.getcwd()
        input_path = file_path + '\\input'
        com_path = input_path + '\\用户数据-已筛选.xlsx'
        final_path = file_path + '\\output\\用户数据-' + sel_id + '.xlsx'
        data = pd.read_excel(com_path)
        # print(data)
        result = np.array(data)
        # print(result)
        lens = len(result)
        # print(lens)
        if lens >= 0:
            values = pd.DataFrame(result, index=[x for x in range(1, lens + 1)], columns=list('ABCDEFG'))
            df = values[values['A'].isin(sel_list)]
            df = pd.DataFrame(df)
            if df.empty:
                print('根据报名表id在话题表中查不到指定数据，请检查报名表是否与话题表对应！程序即将退出！')
                exit()
            else:
                # df.rename(
                #     columns={'A': '用户昵称', 'B': '性别', 'C': '真是姓名', 'D': '联系号码', 'E': '收集来源', 'F': '用户状态',
                #              'G': 'IP', 'H': '地区', 'I': '首次观看直播时间', 'J': '最近观看直播时间', 'K': '直播观看时长'},
                #     inplace=True)
                df.rename(
                    columns={'A': '用户id', 'B': '用户昵称', 'C': 'IP', 'D': '地区', 'E': '首次观看直播时间', 'F': '最近观看直播时间',
                             'G': '直播观看时长'},
                    inplace=True)
                writer = pd.ExcelWriter(final_path)
                df.to_excel(writer, index=False)
                writer.save()
        else:
            print('程序中断执行：请检查话题数据表中是否有数据！！！')
            exit()


def get_extra_list(id):
    res3_list = (
        601.0, 602.0, 603.0, 604.0, 605.0, 606.0, 607.0, 608.0, 609.0, 610.0, 611.0, 612.0, 613.0, 614.0, 615.0, 616.0,
        617.0,
        618.0, 619.0, 620.0,
        621.0, 622.0, 623.0, 624.0, 625.0, 626.0, 627.0, 628.0, 629.0, 630.0, 631.0, 632.0, 633.0, 634.0, 635.0, 636.0,
        637.0,
        638.0, 639.0, 640.0,
        641.0, 642.0, 643.0, 644.0, 645.0, 646.0, 647.0, 648.0, 649.0, 650.0, 651.0, 652.0, 653.0, 654.0, 655.0, 656.0,
        657.0,
        658.0, 659.0, 701.0, 702.0, 703.0, 704.0, 705.0, 706.0, 707.0, 708.0, 709.0, 710.0, 711.0, 712.0, 713.0, 714.0,
        715.0,
        716.0, 717.0, 718.0, 719.0, 720.0, 721.0, 722.0, 723.0, 724.0, 725.0, 726.0, 727.0, 728.0, 729.0, 730.0, 731.0,
        732.0,
        733.0, 734.0, 735.0, 736.0, 737.0, 738.0, 739.0, 740.0, 741.0, 742.0)
    # res3_list = (601, 602, 603, 604, 605, 606, 607, 608, 609, 610, 611, 612, 613, 614, 615, 616, 617, 618, 619, 620,
    #              621, 622, 623, 624, 625, 626, 627, 628, 629, 630, 631, 632, 633, 634, 635, 636, 637, 638, 639, 640,
    #              641, 642, 643, 644, 645, 646, 647, 648, 649, 650, 651, 652, 653, 654, 655, 656, 657, 658, 659, 701,
    #              702, 703, 704, 705, 706, 707, 708, 709, 710, 711, 712, 713, 714, 715, 716, 717, 718, 719, 720,
    #              721, 722, 723, 724, 725, 726, 727, 728, 729, 730, 731, 732, 733, 734, 735, 736, 737, 738, 739, 740,
    #              741, 742
    #              )
    file_path = os.getcwd()
    xlsx_path = file_path + '\\XLSX'
    file_list = os.listdir(xlsx_path)
    extra_path = file_path + '\\extra\\报名记录-错误-' + id + '.xlsx'
    for i in file_list:
        if i.startswith('报名'):
            sel_path = xlsx_path + '\\' + i
            data = pd.read_excel(sel_path)
            result = np.array(data)
            lens = len(result)
            # print(lens)
            if lens > 0:
                values = pd.DataFrame(result, index=[x for x in range(1, lens + 1)], columns=list('ABCDEFGH'))
                df = values[~values['H'].isin(res3_list)]
                extra_data_list = df['A']
                extra_num_list = df['H']
                # print(extra_data_list)
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
            u_data = np.array(data)
            # print(u_data)
            # num_data = np.array(num_list)
            # print(num_data)
            # print(len(num_data))
            lens = len(u_data)
            if lens >= 0:
                ser1 = pd.DataFrame(u_data, index=[x for x in range(1, lens + 1)], columns=list('ABCDEFG'))
                df = ser1[ser1['A'].isin(sel_list)]
                # print(df)
                df = pd.DataFrame(df)
                # df.insert(7, 'H', num_data, allow_duplicates=True)
                if df is None:
                    print('根据报名表id在话题表中查不到指定数据，请检查报名表是否与话题表对应！程序即将退出！')
                    exit()
                else:
                    # df.rename(
                    #     columns={'A': '用户昵称', 'B': '性别', 'C': '真是姓名', 'D': '联系号码', 'E': '收集来源', 'F': '用户状态',
                    #              'G': 'IP', 'H': '地区', 'I': '首次观看直播时间', 'J': '最近观看直播时间', 'K': '直播观看时长'},
                    #     inplace=True)
                    df.rename(
                        columns={'A': '用户id', 'B': '用户昵称', 'C': 'IP', 'D': '地区', 'E': '首次观看直播时间', 'F': '最近观看直播时间',
                                 'G': '直播观看时长', 'H': '门店编码'},
                        inplace=True)
                    # print(extra_path)
                    writer = pd.ExcelWriter(extra_path)
                    df.to_excel(writer, index=False)
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
            1] != '.exe' and i != 'extra' and i != '.git' and i !='.gitattributes' and os.path.splitext(i)[
            1] != '.md':
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
    try:
        print('程序执行前，务必检查文件名是否以【报名】或【话题】开头')
        print('注：报名即报名表，话题即直播话题数据表，如果开头不是以上两种，修改即可')
        print('确认完成后，将1号线或2号线的【.xls】文件放入【XLS】目录')
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
        get_extra_data('1')
        get_extra_list('1')
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
