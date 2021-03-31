# -*- coding: utf-8 -*-
__author__ = 'ouyangmin'
__time__ = '2021/2/16 07:00'

#如果该代码不能满足您日常工作需求，可发邮件到liyysap@126.com ，说清楚您的需求，我尽力帮您实现哦

import os
import xlwings as xw
file_path = 'e:\\table\\销售表1'
file_list = os.listdir(file_path)
sheet_name = '产品销售区域'      #给出要删除的工作表名称
app = xw.App(visible = False, add_book = False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths = os.path.join(file_path, i)
    workbook = app.books.open(file_paths)
    for j in workbook.sheets:
        if j.name == sheet_name:   #判断工作簿中是否有名为“产品销售区域”的工作表
            j.delete()  #如果有，则删除这张工作表
            break
    workbook.save()
app.quit()
