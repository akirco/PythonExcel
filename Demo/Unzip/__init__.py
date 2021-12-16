from pathlib import Path
import zipfile
import os


#
# def unzip():
#     zip_path = 'InputData'
#     file_list = os.listdir(zip_path)
#     for i in file_list:
#         with zipfile.ZipFile(zip_path + '\\' + i, 'r') as f:
#             print(f)
#             for fn in f.namelist():
#                 extracted_path = Path(f.extract(fn))
#                 extracted_path.rename(fn.encode('cp437').decode('gbk'))
#
#
# unzip()


def new_unzip():
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


new_unzip()

#
# def rename():
#     Path('InputData\(101-653)门店数据历史_20210926091942\╙├╗º╩²╛▌-104.xlsx').rename(
#         'InputData\\104\\╙├╗º╩²╛▌-104.xlsx'.encode('cp437').decode('gbk'))
#
#
# rename()
