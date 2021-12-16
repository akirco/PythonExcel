import numpy as np


def pythonsum(n):
    a = [i for i in range(n)]
    b = [i for i in range(n)]
    c = []
    for i in range(len(a)):
        a[i] = i ** 2
        b[i] = i ** 3
        c.append(a[i] + b[i])
    print(c)
    # return c


def numpysum(n):
    a = np.arange(n) ** 2
    b = np.arange(n) ** 3
    c = a + b
    print(c)
    # return c


def new_arr():
    m = np.array([np.arange(2), np.arange(2)])
    print('m:', m)
    print(m.shape)


def arr():
    a = np.arange(9).reshape(3, 3)
    print(a)
    b = 2 * a
    print(b)
    # 水平组合
    res1 = np.hstack((a, b))
    res2 = np.concatenate((a, b), axis=1)
    print(res2)
    # 垂直组合
    res3 = np.vstack((a, b))
    print(res3)
    res4 = np.concatenate((a, b), axis=0)
    print(res4)

    # 深度组合 将相同的元组作为参数传给dstack函数，即可完成数组的深度组合。
    res5 = np.dstack((a, b))
    print(res5)
    # column_stack函数对于一维数组按列方向进行组和。
    one = np.arange(3)
    two = one * 2
    print(one)
    print(two, end='\n')
    res6 = np.column_stack((one, two))
    print(res6)
    # 而对于二维数组，column_stack与hstack的效果是相同的
    # 行组合
    np.row_stack()
    # 对于二维数组，row_stack与vstack的效果是相同的


if __name__ == '__main__':
    a = np.arange(5)
    print(a.dtype)
    """ 
这是一个包含5个元素的向量，取值分别为0~4的整数。数组的shape属性
返回一个元组（tuple），元组中的元素即为NumPy数组每一个维度上的大小。上面例子中的数组
是一维的，因此元组中只有一个元素
    """
    print(a.shape)
    pythonsum(10)
    numpysum(10)

    new_arr()
    arr()
