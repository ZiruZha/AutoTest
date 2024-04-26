# -*- coding: utf-8 -*-
# @Time : 2024/4/23 16:26
# @File : myGui.py
# @Software: PyCharm
from tkinter import *
import JoiMaker_V200_lib
import waterRPA_V200_lib
import TongXinDuiDian_V102_lib

titleFontSize = 20
textFontSize = int(2 / 3 * titleFontSize)
relativeHeight = 0.08


def selDeviceType():
    dic = {0: '甲', 1: '乙', 2: '丙'}
    s = "您选了" + dic.get(var.get()) + "项"
    pass


def multiCheck():
    ifR = '1' if CheckVar61.get() == 1 else '0'
    ifP = '1' if CheckVar62.get() == 1 else '0'
    ifL = '1' if CheckVar63.get() == 1 else '0'
    ifD = '1' if CheckVar64.get() == 1 else '0'
    ifY = '1' if CheckVar65.get() == 1 else '0'
    ifK = '1' if CheckVar66.get() == 1 else '0'
    temp = ifR + ifP + ifL + ifD + ifY + ifK
    return int(temp)


def call1():
    try:
        JoiMaker_V200_lib.main(inp2.get(), inp20.get())
    except Exception as e:
        s = '---------------ERROR---------------\n' + str(e) + '\n'
        txt9.insert(END, s)  # 追加显示运算结果
        inp2.delete(0, END)  # 清空输入
        inp20.delete(0, END)  # 清空输入
        # print(e)
    else:
        txt9.insert(END, '打包完成\n')  # 追加显示运算结果
        inp2.delete(0, END)  # 清空输入
        inp20.delete(0, END)  # 清空输入


def call2():
    try:
        waterRPA_V200_lib.main(multiCheck())
    except Exception as e:
        s = '---------------ERROR---------------\n' + str(e) + '\n'
        txt9.insert(END, s)  # 追加显示运算结果
    else:
        txt9.insert(END, '保护功能测试完成\n')  # 追加显示运算结果


def call3():
    try:
        temp = TongXinDuiDian_V102_lib.main(int(inp5.get()))
    except Exception as e:
        s = '---------------ERROR---------------\n' + str(e) + '\n'
        txt9.insert(END, s)  # 追加显示运算结果
    else:
        txt9.insert(END, temp)
        txt9.insert(END, '61850测试完成\n')  # 追加显示运算结果
    pass


if __name__ == '__main__':
    root = Tk()
    root.title('自动测试')
    # 标签类
    lb0 = Label(root,
                # anchor='nw',
                text='为人民服务',
                bg='red',
                fg='black',
                font=('华文行楷', int(titleFontSize * 4 / 3)),
                relief=FLAT)
    lb0 = Label(root,
                # anchor='nw',
                text='欢迎使用自动测试',
                bg='red',
                fg='black',
                font=('微软雅黑', int(titleFontSize * 4 / 3)),
                relief=FLAT)
    lb01 = Label(root,
                 anchor='sw',
                 text='——毛泽东',
                 bg='red',
                 fg='black',
                 font=('华文行楷', textFontSize),
                 relief=FLAT)
    lb1 = Label(root,
                # anchor='nw',
                text='自动打包',
                # bg='black',
                fg='black',
                font=('微软雅黑', titleFontSize),
                relief=FLAT)
    # lb10 = Label(root,
    #              # anchor='nw',
    #              text='运行结果',
    #              # bg='black',
    #              fg='black',
    #              font=('微软雅黑', titleFontSize),
    #              relief=FLAT)
    # lb11 = Label(root,
    #              # anchor='nw',
    #              # text='运行结果',
    #              bg='black',
    #              # fg='black',
    #              # font=('微软雅黑', titleFontSize),
    #              relief=FLAT)
    lb2 = Label(root,
                # anchor='w',
                text='《工程程序需求统计和交付管理》\n文件中待打包文件起始行序号：',
                # bg='black',
                fg='black',
                font=('微软雅黑', textFontSize),
                relief=FLAT)
    lb20 = Label(root,
                 # anchor='w',
                 text='终止行序号：',
                 # bg='black',
                 fg='black',
                 font=('微软雅黑', textFontSize),
                 relief=FLAT)
    lb4 = Label(root,
                bg='black',
                relief=FLAT)
    lb40 = Label(root,
                 # anchor='nw',
                 text='保护功能测试&61850测试',
                 # bg='black',
                 fg='black',
                 font=('微软雅黑', titleFontSize),
                 relief=FLAT)
    lb5 = Label(root,
                # anchor='w',
                text='装置类型：',
                # bg='black',
                fg='black',
                font=('微软雅黑', textFontSize),
                relief=FLAT)
    lb50 = Label(root,
                 anchor='e',
                 text='传动每',
                 # bg='black',
                 fg='black',
                 font=('微软雅黑', textFontSize),
                 relief=FLAT)
    lb51 = Label(root,
                anchor='w',
                text='条记录一次',
                # bg='black',
                fg='black',
                font=('微软雅黑', textFontSize),
                relief=FLAT)
    lb6 = Label(root,
                # anchor='w',
                text='选配型号：',
                # bg='black',
                fg='black',
                font=('微软雅黑', textFontSize),
                relief=FLAT)
    lb8 = Label(root,
                bg='black',
                relief=FLAT)
    lb70 = Label(root,
                 # anchor='nw',
                 text='61850测试',
                 # bg='black',
                 fg='black',
                 font=('微软雅黑', titleFontSize),
                 relief=FLAT)

    # 输入框类
    inp2 = Entry(root)
    inp20 = Entry(root)
    inp5 = Entry(root)
    # 文本框类
    txt9 = Text(root)
    # 按钮类
    btn3 = Button(root,
                  text='开始打包',
                  font=('微软雅黑', textFontSize),
                  command=call1)
    btn7 = Button(root,
                  text='保护功能测试',
                  font=('微软雅黑', textFontSize),
                  command=call2)
    btn70 = Button(root,
                   text='61850测试',
                   font=('微软雅黑', textFontSize),
                   command=call3)
    # 复选框
    CheckVar61 = IntVar()
    CheckVar62 = IntVar()
    CheckVar63 = IntVar()
    CheckVar64 = IntVar()
    CheckVar65 = IntVar()
    CheckVar66 = IntVar()
    ch61 = Checkbutton(root, text='R', font=('微软雅黑', textFontSize), variable=CheckVar61, onvalue=1, offvalue=0)
    ch62 = Checkbutton(root, text='P', font=('微软雅黑', textFontSize), variable=CheckVar62, onvalue=1, offvalue=0)
    ch63 = Checkbutton(root, text='L', font=('微软雅黑', textFontSize), variable=CheckVar63, onvalue=1, offvalue=0)
    ch64 = Checkbutton(root, text='D', font=('微软雅黑', textFontSize), variable=CheckVar64, onvalue=1, offvalue=0)
    ch65 = Checkbutton(root, text='Y', font=('微软雅黑', textFontSize), variable=CheckVar65, onvalue=1, offvalue=0)
    ch66 = Checkbutton(root, text='K', font=('微软雅黑', textFontSize), variable=CheckVar66, onvalue=1, offvalue=0)
    # 单选按钮
    var = IntVar()
    rd51 = Radiobutton(root, text="G", font=('微软雅黑', textFontSize), variable=var, value=0, command=selDeviceType)
    rd52 = Radiobutton(root, text="DA", font=('微软雅黑', textFontSize), variable=var, value=1, command=selDeviceType)
    rd53 = Radiobutton(root, text="DG", font=('微软雅黑', textFontSize), variable=var, value=2, command=selDeviceType)
    # 第0行
    lb0.place(relx=0, rely=0, relheight=0.1, relwidth=1)
    # lb01.place(relx=0.7, rely=0, relheight=0.1, relwidth=0.5)
    # 第1行
    lb1.place(relx=0.1, rely=0.1, relheight=relativeHeight, relwidth=0.8)
    # lb10.place(relx=0.5, rely=0.1, relheight=relativeHeight, relwidth=0.5)
    # lb11.place(relx=0.5, rely=0.1, relheight=0.9, width=1)
    # 第2行
    lb2.place(relx=0.1, rely=0.2, relheight=relativeHeight, relwidth=0.4)
    inp2.place(relx=0.5, rely=0.2, relheight=relativeHeight, relwidth=0.1)
    lb20.place(relx=0.6, rely=0.2, relheight=relativeHeight, relwidth=0.2)
    inp20.place(relx=0.8, rely=0.2, relheight=relativeHeight, relwidth=0.1)
    # 第3行
    btn3.place(relx=0.4, rely=0.3, relheight=relativeHeight, relwidth=0.2)
    # 第4行
    lb4.place(relx=0, rely=0.39, height=1, relwidth=1)
    lb40.place(relx=0.1, rely=0.4, relheight=relativeHeight, relwidth=0.8)
    # 第5行
    lb5.place(relx=0.3, rely=0.5, relheight=relativeHeight, relwidth=0.1)
    rd51.place(relx=0.4, rely=0.5, relheight=relativeHeight, relwidth=0.1)
    rd52.place(relx=0.5, rely=0.5, relheight=relativeHeight, relwidth=0.1)
    rd53.place(relx=0.6, rely=0.5, relheight=relativeHeight, relwidth=0.1)
    lb50.place(relx=0.7, rely=0.5, relheight=relativeHeight, relwidth=0.1)
    inp5.place(relx=0.8, rely=0.5, relheight=relativeHeight, relwidth=0.05)
    lb51.place(relx=0.85, rely=0.5, relheight=relativeHeight, relwidth=0.15)
    # 第6行
    lb6.place(relx=0.1, rely=0.6, relheight=relativeHeight, relwidth=0.2)
    ch61.place(relx=0.3, rely=0.6, relheight=relativeHeight, relwidth=0.1)
    ch62.place(relx=0.4, rely=0.6, relheight=relativeHeight, relwidth=0.1)
    ch63.place(relx=0.5, rely=0.6, relheight=relativeHeight, relwidth=0.1)
    ch64.place(relx=0.6, rely=0.6, relheight=relativeHeight, relwidth=0.1)
    ch65.place(relx=0.7, rely=0.6, relheight=relativeHeight, relwidth=0.1)
    ch66.place(relx=0.8, rely=0.6, relheight=relativeHeight, relwidth=0.1)
    # 第7行
    btn7.place(relx=0.2, rely=0.7, relheight=relativeHeight, relwidth=0.2)
    btn70.place(relx=0.6, rely=0.7, relheight=relativeHeight, relwidth=0.2)
    # lb70.place(relx=0.1, rely=0.7, relheight=relativeHeight, relwidth=0.8)
    # 第8行
    lb8.place(relx=0, rely=0.79, height=1, relwidth=1)
    # 第9行
    txt9.place(relx=0.05, rely=0.8, relheight=0.15, relwidth=0.9)

    root.geometry('1200x700')  # 这里的乘号不是 * ，而是小写英文字母 x
    root.mainloop()
