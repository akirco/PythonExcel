import sys

import PySimpleGUI as sg

sg.theme('LightBlue2')  # Add a touch of color
# All the stuff inside your window.
# sg.change_look_and_feel('Material1')
layout = [[sg.Text('程序执行前，务必检查文件名是否以【报名】或【话题】开头\n'
                   '注：报名即报名表，话题即直播话题数据表，如果开头不是以上两种，修改即可确认完成后，\n'
                   '将1号线或2号线的【.xls】文件放入【XLS】目录')],
          [sg.In(key='dir_name'), sg.FolderBrowse(button_text='打开', initial_folder='.', target='dir_name')],
          [sg.Text('请输入您要筛选的直播编号：(例：1号线则输入【1】)'), sg.InputText(key='-IN-')],
          [sg.Text('您要筛选的是:'), sg.Text(size=(0, 1), key='-OUTPUT-'), sg.Text(key='-msg-')],
          [sg.ProgressBar(100)],
          [sg.Button('Ok'), sg.Button('Cancel')]]

# Create the Window
window = sg.Window(size=(900, 460), title='微赞数据筛选', layout=layout, background_color='white',
                   auto_size_text=True,
                   auto_size_buttons=True,
                   icon='.\\Remote.ico')
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    print(event, values)
    if event == 'Ok':
        print('OK')
        window['-OUTPUT-'].update(values['-IN-'])
        window['-msg-'].update('号线')
    if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
        break
    print(sys.argv)

window.close()
