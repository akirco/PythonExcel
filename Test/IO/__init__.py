import os

# 1.返回表示当前工作目录的 unicode 字符串。
path = os.getcwd()
print(path)

# 2.列出指定路径下的文件夹包含的文件和子文件夹名称
path = os.getcwd()  # 获取当前目录
file_list = os.listdir(path)
print(file_list)

# 3.分离文件主名和扩展名
path = 'example.xlsx'
separate = os.path.splitext(path)  # 返回一个元组
print(separate, type(separate))

# 重命名文件和文件夹
# rename()函数除了可以重命名文件，还可以修改文件的路径（移动文件）
path = os.getcwd()
old_name = path + '\\ts.txt'  # 文件不存在会报错
new_name = path + '\\ts.txt'
os.rename(old_name, new_name)

# rename 移动文件

path = os.getcwd()
old_name = path + '\\ts.txt'  # 文件不存在会报错
new_name = path + '\\test' + '\\ts.txt'  # 移动目标目录不存在也会报错
print(new_name)
os.rename(old_name, new_name)
