# -*- coding: utf-8 -*-
__author__ = 'ouyangmin'
__time__ = '2021/2/14 15:43'


import os    #导入OS模块
import xlwings as xw   #导入xlwings模块
file_path = 'e:\\table\\统计表.xlsx'   #给出工作簿所在的文件夹路径
file_list = os.listdir(file_path)    #列出路径下的所有文件和子文件夹的名称
app = xw.App(visible = True, add_book = False)   #启动Excel程序
for i in file_list:
    if os.path.splitext(i)[1] == '.xlsx':   #判断文件夹下文件的扩展名称是否为“.xlsx”
        app.books.open(file_path + '\\' + i)    #打开工作簿
