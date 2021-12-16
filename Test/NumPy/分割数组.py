import numpy as np

a = np.array([[0, 1, 2],
              [3, 4, 5],
              [6, 7, 8]])
# 水平分割
result = np.hsplit(a, 3)
print(result)
result = np.split(a, 3, axis=1)
print(result)

# 垂直分割
result = np.vsplit(a, 3)
print(result)
result = np.split(a, 3, axis=0)
print(result)

# 深度分割
c = np.arange(27).reshape(3, 3, 3)
print('深度分割：', c)
