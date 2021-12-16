import random


def generate_pwd(m, n):
    range_start = 10 ** (m - 1)
    range_end = (10 ** m) - 1
    arr = [random.randint(range_start, range_end) for x in range(n)]
    for i in arr:
        with open('Password.txt', 'a') as f:
            print('%s' % i, file=f)


if __name__ == '__main__':
    m = input('请输入密码位数：')
    n = input('生成密码个数：')
    generate_pwd(int(m), int(n))
