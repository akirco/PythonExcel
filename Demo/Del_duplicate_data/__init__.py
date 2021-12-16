# 导入pandas包并重命名为pd
import pandas as pd

# 读取Excel中Sheet1中的数据
data = pd.DataFrame(pd.read_excel('用户数据-已筛选.xlsx', 'Sheet1'))

# 查看读取数据内容
# print(data)

# 查看是否有重复行
# re_row = data.duplicated()
# print(re_row)

# 查看去除重复行的数据
# no_re_row = data.drop_duplicates()
# print(no_re_row)

# 查看基于[物品]列去除重复行的数据
wp = data.drop_duplicates(['用户昵称'])
print(wp)

# 将去除重复行的数据输出到excel表中
wp.to_excel("用户数据-已筛选.xlsx")
