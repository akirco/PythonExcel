import pandas as pd
import numpy as np

""" 
pandas模块是基于NumPy模块的一个开源Python模块，
广泛应用于完成数据快速分析、数据清洗和准备等工作，
它的名字来源于“panel data”（面板数据）。
pandas模块提供了非常直观的数据结构及强大的数据管理和数据处理功能，
某种程度上可以把pandas模块看成Python版的Excel。
如果是利用Anaconda安装的Python，则自带pandas模块，无须单独安装。
与NumPy模块相比，pandas模块更擅长处理二维数据，其主要有Series和DataFrame两种数据结构。
"""
# s = pd.Series(['one', 'two', 'three'])
# print(s)
# print(s[1])
#
# df = pd.DataFrame([[1, 2], [3, 4], [5, 6]])
# print(df)

# 创建DataFrame时自定义列索引和行索引
# df = pd.DataFrame([[1, 2], [3, 4], [5, 6]], columns=['data', 'score'], index=['A', 'B', 'C'])
# print(df)

# df = pd.DataFrame()
# date = [1, 3, 5]
# score = [2, 4, 6]
# df['date'] = date
# df['score'] = score
# print(df)


# 通过字典创建DataFrame
# df = pd.DataFrame({'a': [1, 2, 3], 'b': [4, 5, 6], 'c': [7, 8, 9]}, index=['x', 'y', 'z'])
# print(df)
# 列索引是字典的键名

# 如果想以字典的键名作为行索引，
# 可以用from_dict()函数将字典转换成DataFrame，
# 同时设置参数orient的值为'index'
""" 
参数orient用于指定以字典的键名作为列索引还是行索引，
默认值为'columns'，即以字典的键名作为列索引，
如果设置成'index'，则表示以字典的键名作为行索引。
"""
# df = pd.DataFrame.from_dict({'a': [1, 2, 3], 'b': [4, 5, 6]}, orient='index')
# print(df)

# 通过二维数组创建DataFrame
# npy = np.arange(12).reshape(3, 4)
# df = pd.DataFrame(npy, index=['a', 'b', 'c'], columns=['A', 'B', 'C', 'D'])
# print(df)

# DataFrame索引的修改

# df = pd.DataFrame([[1, 2], [3, 4], [5, 6]], columns=['date', 'score'], index=['A', 'B', 'C'])
# df.index.name = '公司'
# print(df)
# 如果想重命名索引，可以使用rename()函数
# df = df.rename(index={'A': '万科', 'B': '阿里', 'C': '百度'}, columns={'date': '日期', 'score': '分数'})
# print(df)


# 文件的读取和写入
# data = pd.read_excel('example.xlsx', sheet_name=0)
# print(data)
''' 
read_excel()函数还有其他参数，这里简单介绍几个常用参数：
·sheetname用于指定工作表，可以是工作表名称，也可以是数字（默认为0，即第1个工作表）。
·encoding用于指定文件的编码方式，一般设置为UTF-8或GBK编码，以避免中文乱码。
·index_col用于设置索引列。
'''

# 文件写入

data = pd.DataFrame([[1, 2], [3, 4], [5, 6]], columns=['A', 'B'])
data.to_excel('example.xlsx')
