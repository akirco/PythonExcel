from tkinter import *

root = Tk()
root.geometry('800x450')
root.title('微赞数据筛选')
root.iconbitmap('Remote.ico')
msg = Message(root, text='程序执行前，务必检查文件名是否以【报名】或【话题】开头\n'
                         '注：报名即报名表，话题即直播话题数据表，如果开头不是以上两种，修改即可确认完成后，\n'
                         '将1号线或2号线的【.xls】文件放入【XLS】目录', relief=GROOVE)
msg.place(relx=0.01, y=0.01, relheight=0.5, width=500)
root.mainloop()
