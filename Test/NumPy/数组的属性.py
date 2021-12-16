import numpy as np

b = np.array([[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11],
              [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]])
print(b.ndim)
print(b.size)
print(b.itemsize)
print(b.nbytes)
print(b.resize(6, 4))
print(b.T)

a = np.array([1.j + 1, 2.j + 3])
# 复数实部
print(a.real)
# 复数虚部
print(a.imag)
print(a.dtype)

c = np.arange(4).reshape(2, 2)
print(c)
print(c.flat)
f = c.flat
for i in f:
    print(i)
