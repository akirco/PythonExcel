import os


# 加密压缩单个文件
# loc_arch_path = '..\\output\\'
# loc_out_path = '..\\arch\\'


# arch_name, pwd, arch_file
def encrypt_arch_files(arch_name, pwd, arch_file):
    file_path = os.getcwd()
    loc_winRar_path = file_path + '\\WinRAR\\WinRAR.exe'

    archive_cmd = loc_winRar_path + ' a -p%s -mezl %s %s' % (pwd, arch_name, arch_file)
    print(archive_cmd)
    os.system(archive_cmd)


if __name__ == '__main__':
    encrypt_arch_files('test.zip', '123', '__init__.py')
