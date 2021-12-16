import numpy as np

# 使用tolist函数将NumPy数组转换成Python列表

b = np.array([1. + 1.j, 3. + 2.j])
b = b.tolist()
print(b)

# astype函数可以在转换数组时指定数据类型

c = np.array([1. + 1.j, 3. + 2.j])
# ComplexWarning: Casting complex values to real discards the imaginary part
#   c.astype(int)
# c.astype(int)
c.astype('complex')
print(c)
