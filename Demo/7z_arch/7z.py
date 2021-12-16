import json
import os
import random
import string

md_list = (601, 602, 603, 604, 605, 606, 607, 608, 609, 610, 611, 612, 613, 614, 615, 616, 617, 618, 619, 620,
           621, 622, 623, 624, 625, 626, 627, 628, 629, 630, 631, 632, 633, 634, 635, 636, 637, 638, 639, 640,
           641, 642, 643, 644, 645, 646, 647, 648, 649, 650, 651, 652, 653, 654, 655, 656, 657, 658, 659, 701,
           702, 703, 704, 705, 706, 707, 708, 709, 710, 711, 712, 713, 714, 715, 716, 717, 718, 719, 720,
           721, 722, 723, 724, 725, 726, 727, 728, 729, 730, 731, 732, 733, 734, 735, 736, 737, 738, 739, 740,
           741, 742)


# 使用命令，压缩、加密单个文件
# arch_name 压缩文件名
# pwd 压缩密码
# arch_file 被压缩文件

# 加密压缩单个文件
def encrypt_arch_files(arch_name, pwd, arch_file):
    file_path = os.getcwd()
    loc_7z_path = '.\\7za.exe'
    archive_cmd = loc_7z_path + ' a ' + '.\\arch\\' + arch_name.__str__() + '.zip' + ' -p' + pwd + ' .\\output\\' + arch_file  # 编辑命令行
    print(archive_cmd)
    os.system(archive_cmd)


def get_arch_files():
    loc_arch_path = '.\\output\\'
    arch_file_list = os.listdir(loc_arch_path)
    return arch_file_list


def generate_pwd(length):
    lens = len(md_list)
    pwd_list = []
    for j in range(lens):
        Ofnum = random.randint(1, length)
        Ofletter = length - Ofnum
        slcNum = [random.choice(string.digits) for i in range(Ofnum)]
        slcLetter = [random.choice(string.ascii_letters) for i in range(Ofletter)]
        slcChar = slcLetter + slcNum
        random.shuffle(slcChar)
        Pwd = ''.join([i for i in slcChar])
        pwd_list.append(Pwd)
    dic_pwd = dict(zip(pwd_list, md_list))
    json_pwd = json.dumps(dic_pwd, indent=4, ensure_ascii=False)
    print(json_pwd)
    with open('pwd.json', 'w') as f:
        f.write(json_pwd)


def get_pwd():
    pwd_list = []
    with open('pwd.json', 'r') as f:
        pwd = f.read()
        result = json.loads(pwd)
        # pwd_list.append(pwd)
        # pwd_list = pwd_list[0].split('\n')
        # # print(pwd_list)
        print('result:', result)
        return result


def get_dict_keys(dic, values):
    return [k for (k, v) in dic.items() if v == values]


def get_dict_pwd_file():
    md_nums = []
    res_pwd = get_pwd()
    current_pwd = []
    current_file = []
    print(res_pwd)
    file_list = get_arch_files()
    for i in file_list:
        current_file.append(i)
        md_num = i.split('.')[0].split('-')[1]
        md_nums.append(md_num)
    print(md_nums)
    for x in md_nums:
        pwd = get_dict_keys(res_pwd, int(x))
        current_pwd.extend(pwd)
    print('current_pwd:', current_pwd)
    print('current_file:', current_file)
    dict_pwd = dict(zip(current_pwd, current_file))
    return dict_pwd


if __name__ == '__main__':
    # generate_pwd(8)
    final_pwd_file = get_dict_pwd_file()

    for pwd, arch_item in final_pwd_file.items():
        arch_item_name = os.path.splitext(arch_item)[0]
        print('正在压缩文件:', arch_item_name)
        encrypt_arch_files(arch_item_name, pwd, arch_item)
