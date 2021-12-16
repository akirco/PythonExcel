import os
import time
import win32api
import win32con
import win32gui
import json
from pykeyboard import PyKeyboard
from pymouse import PyMouse

k = PyKeyboard()
m = PyMouse()

config_error = "配置文件有错误!!!!!!!"
windows_not_open_error = "请打开应用后在使用!!!!!!"

# 模拟鼠标点击
def mouse_click(x, y):
    win32api.SetCursorPos([x, y])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

# 输入信息
def type_string(context):
    k.type_string(context)

# 查找窗口，并设置位置
def find_window(window_name):
    # 获取句柄
    hwnd = win32gui.FindWindow(None, window_name)
    if hwnd == 0:
        return 0
    # 获取窗口左上角和右下角坐标
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    # 计算高度和宽度
    win_width = right - left
    win_height = bottom - top
    # 设置窗口位置为0,0
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, win_width, win_height, win32con.SWP_SHOWWINDOW)
    return hwnd

def config_application():
    print("正在使用配置模式")
    while True:
        time.sleep(0.1)
        print(f'\r当前鼠标位置：{m.position()}', end="")

def run(times, time_gap, operations):
    # 循环次数
    for i in range(0, times):
        print(f'正在进行第{i+1}次操作')
        # 执行操作
        for operation in operations:
            time.sleep(time_gap)
            operation_type = list(operation.keys())[0]
            # 点击操作
            if operation_type == "click":
                point = operation[operation_type]
                mouse_click(point[0], point[1])
            # 输入操作
            elif operation_type == "type_string":
                type_str_tmp = operation[operation_type]
                type_string(type_str_tmp)
            else:
                print(config_error)
                return

def init() -> None:
    # 直接从json文件中读取数据返回一个python对象
    data = ''
    try:
        f = open('config.json', encoding='utf-8')
        data = json.load(f)
    except:
        print("缺少配置文件!!!!")
        return
    title = data['title']
    if title is None:
        print(config_error)
    hwnd = find_window(title)
    if hwnd == 0:
        print(windows_not_open_error)
        return
    print("请选择需要的功能：\n1. 配置 \n2. 运行")
    input_type = input("> ")

    if input_type == "1":
        config_application()
    elif input_type == "2":
        run(data['time'], data['time_interval'], data['operation'])
    else:
        print("能不能正常输入？？？")

if __name__ == '__main__':
    init()
    os.system("pause")