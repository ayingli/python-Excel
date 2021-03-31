# -*- coding: utf-8 -*-
__author__ = 'ouyangmin'
__time__ = '2021/2/14 15:43'

import xlwings as xw     #导入xlwings模块，用来操作Excel的

app = xw.App(visible = False, add_book = False)     #启动Excel程序
workbook = app.books.open('e:\\table\\统计表.xlsx')    #打开指定目录下的工作簿
worksheets = workbook.sheets           #获取该工作簿中的所有工作表
for i in range(len(worksheets)):      #遍历获取到的工作表
    worksheets[i].name = worksheets[i].name.replace('2020年', '')    #重命名工作表
workbook.save('e:\\table\\统计表1.xlsx')     #将重命名之后的工作簿重新保存
app.quit()    #退出Excel程序
