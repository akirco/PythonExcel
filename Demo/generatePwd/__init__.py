# coding:utf-8
import random
import string


def generate_pwd(length, nums):
    for j in range(nums):
        # 随机生成数字个数
        Ofnum = random.randint(1, length)
        Ofletter = length - Ofnum

        # 选中ofnum个数字
        slcNum = [random.choice(string.digits) for i in range(Ofnum)]

        # 选中ofletter个字母
        slcLetter = [random.choice(string.ascii_letters) for i in range(Ofletter)]

        # 打乱组合
        slcChar = slcLetter + slcNum
        random.shuffle(slcChar)
        print(slcChar)
        # 生成随机密码

        Pwd = ''.join([i for i in slcChar])
        print(Pwd)
        with open('pwd.txt', 'a') as f:
            print('%s' % Pwd, file=f)


if __name__ == '__main__':
    # GetPassword()自定义随机密码长度
    generate_pwd(4, 101)
