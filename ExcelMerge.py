# coding=utf-8
""" 在python编写代码的时候，避免不了会出现或是用到中文，这时候你需要在文件开头加上中文注释。
比如创建一个python list，在代码上面注释上它的用途，如果开头不声明保存编码的格式是什么，
那么它会默认使用#ASKII码保存文件，这时如果你的代码中有中文就会出错了，即使你的中文是包含在注释里面的。
所以加上中文注释很重要。#coding=utf-8或者：#coding=gbk
以u或U开头的字符串表示unicode字符串
 Unicode是书写国际文本的标准方法。如果你想要用非英语写文本,那么你需要有一个支持Unicode的编辑器。
 类似地,Python允许你处理Unicode文本——你只需要在字符串前加上前缀u或U。"""


import os
import socket
import struct
import ttk
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
# import sys
# import time
# import datetime
# 引入外部开发模块
import xlwt
# import xlrd
# from openpyxl import load_workbook
from tkinter import *
import tkFileDialog,  tkMessageBox
import tkinter.filedialog as filedialog
# import glob
file_name = ""
currentPath = os.path.abspath(os.curdir)
reload(sys)
sys.setdefaultencoding('gbk')


def ask_quit():
    sys.exit()

root = Tk()
root.protocol("WM_DELETE_WINDOW", ask_quit)
# 设置窗口的大小宽x高+偏移量
root.geometry('800x600+500+200')
root.title(u'数据表自动合并')

print(u"正在加载数据，请稍等~~~")


def cmd_open_file1():
    r1 = tkFileDialog.askopenfilename(title='打开文件', filetypes=[('Excel', '*.xls *.xlsx *xlsm'), ('All Files', '*')])
    lb.insert(END, u'文件路径为：' + r1)
    global reference_file_name
    reference_file_name = r1


def cmd_open_file2():
    r2 = tkFileDialog.askopenfilename(title='打开文件', filetypes=[('Excel', '*.xls *.xlsx *xlsm'), ('All Files', '*')])
    lb.insert(END, u'文件路径为：' + r2)
    global feedback_file_name
    feedback_file_name = r2


def getdir(filepath=os.getcwd()):
    """
    用于获取目录下的文件列表
    """
    cf = os.listdir(filepath)
    for i in cf:
        lb.insert(END, i)
        root.update_idletasks()
    return cf


# def soft_auth(set_session, apply_day):
#     days = date_session(apply_day)
#     if days > set_session:
#         tkMessageBox.showinfo(title='', message=U'软件已过期，请联系：15712758016')
#         sys.exit()
#     elif days < 0:
#         tkMessageBox.showinfo(title='', message=U'您的电脑系统有问题！')
#         sys.exit()


def connection(s, server_ip):
    # s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect((server_ip, 12306))
    except socket.error, msg:
        tkMessageBox.showinfo(title='', message=U'局域网连接失败，请检查网络连接，错误代码 : ' + str(msg[0]) + ' Message ' + msg[1])
        # raw_input('Press <enter> to close!')
        sys.exit()


def merge_excel():
    """ COOEC not need """
    # soft_auth(120, datetime.datetime(2018, 1, 9))
    """ COOEC need """
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    connection(s, '10.195.8.200')

    excel_df1_all = pd.DataFrame()

    state_entry.delete(0, END)  # 清空entry里面的内容
    lb.delete(0, END)
    root.update_idletasks()
    # 调用filedialog模块的askdirectory()函数去打开文件夹
    global filepath
    filepath = filedialog.askdirectory()
    if filepath:
        state_entry.insert(0, filepath)  # 将选择好的路径加入到entry里面
    print (filepath)
    file_list = getdir(filepath)

    lb.insert(END, U'正在汇总，请稍等~~~')
    root.update_idletasks()

    if e1.get().strip() == '' or e2.get().strip() == '':
        tkMessageBox.showinfo(title='', message=U'请填写需合并第几个Sheet，以及需要跳过行数')
        return

    sheet_no = int(e1.get())-1
    skip_rows = int(e2.get())

    for fileName in file_list:
        excel_df1 = pd.read_excel(filepath + '/' + fileName, sheetname=sheet_no, header=None, skiprows=skip_rows)
        excel_df1_all = pd.concat((excel_df1_all, excel_df1), axis=0)

    # print excel_df1_all.head(5)

    excel_writer = pd.ExcelWriter(U"SPFSummary.xls")
    excel_df1_all.to_excel(excel_writer,  index=False)

    excel_writer.save()
    lb.insert(END, U'汇总完成！')
    root.update_idletasks()

    tkMessageBox.showinfo(title='', message=U'合并完成!')


def cmd_open_dir():
    output_dir = currentPath.decode('gbk')
    # output_dir = unicode(output_dir, 'utf-8')
    os.startfile(output_dir)

"""创建frame作为容器"""
frame_top = ttk.LabelFrame(root, width=750, height=100, text=u'操作步骤介绍：')
frame_top.grid(row=0, column=0, padx=5, pady=5)           # 使用grid设置各个容器位置
frame_top.grid_propagate(0)
frame_middle = ttk.LabelFrame(root, width=750, height=150, text=u'操作过程：')
frame_middle.grid(row=2, column=0, padx=5, pady=5)           # 使用grid设置各个容器位置
frame_middle.grid_propagate(0)
frame_bottom = LabelFrame(width=750, height=350, bg='white', text=u'显示')
frame_bottom.grid(row=4, column=0, pady=3)
frame_bottom.grid_propagate(0)

sb = Scrollbar(frame_bottom)
sb.pack(side=RIGHT, fill=Y)                   # 需要先 将滚动条放置 到一个合适的位置 , 然后开始填充 .
lb = Listbox(frame_bottom, width=105, height=15, yscrollcommand=sb.set)      # 内容 控制滚动条 .
lb.pack(side=LEFT, fill=BOTH)
sb.config(command=lb.yview)  # 滑轮控制内容.

"""创建元素frame_top"""
# 左对齐，文本居中
label1 = Label(frame_top, text=u'1. 自动合并多个excel文件，按列顺序号自动合并，不区分列名\n'
                               u'2. “需合并第几个Sheet”和“需要跳过的行数”必须填写\n\n'
                               u'如有使用问题，请联系张高尉：zhanggw1@cooec.com.cn',wraplength=700,
               justify='left', bg='#5CACEE')
label1.grid(row=0, column=0, padx=3, pady=5)
# sticky=Tkinter.W  当该列中其他行或该行中的其他列的某一个功能拉长这列的宽度或高度时，
# 设定该值可以保证本行保持左对齐，N：北/上对齐  S：南/下对齐  W：西/左对齐  E：东/右对齐

Label(frame_middle, text=U'1、需合并第几个Sheet:').grid(row=0, column=0, padx=3, pady=5)
Label(frame_middle, text=U'2、需要跳过的行数:').grid(row=0, column=2, padx=3, pady=5)
v1 = StringVar()    # 设置变量 .
v2 = StringVar()
e1 = Entry(frame_middle, textvariable=v1)            # 用于储存 输入的内容
e2 = Entry(frame_middle, textvariable=v2)
e1.grid(row=0, column=1, padx=3, pady=5)      # 进行表格式布局 .
e2.grid(row=0, column=3, padx=3, pady=5)

buttonOpenFile = ttk.Button(frame_middle, text=u'A1.合并所有数据表', command=merge_excel)
buttonOpenFile.grid(row=1, column=0, padx=3, pady=5, sticky=W)
state_entry = Entry(frame_middle, width=60)
state_entry.grid(sticky=W + N, row=1, column=1, columnspan=4, padx=5, pady=5)
buttonChooseFile = ttk.Button(frame_middle, text=u'A2.打开汇总文件路径', command=cmd_open_dir)
buttonChooseFile.grid(row=1, column=5, padx=3, pady=5, sticky=W)

root.mainloop()
