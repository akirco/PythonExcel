import pandas as pd
import numpy as np

a = np.array([[1, 2, 3], [4, 5, 6], [7, 8, 9]])
print(type(a))
df1 = pd.DataFrame(a, index=['row0', 'row1', 'row2'], columns=list('ABC'))
print(df1)
df2 = df1.copy()

# 删除/选取某列含有特定数值的行
# df1=df1[df1['A'].isin([1])]
# df1[df1['A'].isin([1])] 选取df1中A列包含数字1的行

df1 = df1[~df1['A'].isin([1])]
# 通过~取反，选取不包含数字1的行
print(df1)
