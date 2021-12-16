import os

file_path = os.getcwd()
file_list = os.listdir(file_path)

old_book_name = 'ts'
new_book_name = 'example'
for i in file_list:
    if i.startswith('~$') or os.path.splitext(i)[1] == '.py':  # i.endswith('.py') ==os.path.splitext(i)[1] == '.py'
        continue
    new_file = i.replace(old_book_name, new_book_name)
    old_file_path = os.path.join(file_path, i)
    print(old_file_path)
    new_file_path = os.path.join(file_path, new_file)
    print(new_file_path)
    os.rename(old_file_path, new_file_path)
