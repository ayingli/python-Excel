# -*- coding: utf-8 -*-
__author__ = 'ouyangmin'
__time__ = '2021/2/16 07:00'


import os
import xlwings as xw
file_path = 'e:\\table\\信息表'
file_list = os.listdir(file_path)
old_sheet_name = 'Sheet1'    #需要修改的表名
new_sheet_name = '员工信息'     #修改后的表名
app = xw.App(visible = False, add_book = False)
for i in file_list:
    if i.startswith('~$'):
        continue
    old_file_path = os.path.join(file_path, i)
    workbook = app.books.open(old_file_path)
    for j in workbook.sheets:
        if j.name == old_sheet_name:     #判断工作表名是否为“sheet1”
            j.name = new_sheet_name   #如果是，则重命名工作表
    workbook.save()    #保存工作簿
app.quit()
