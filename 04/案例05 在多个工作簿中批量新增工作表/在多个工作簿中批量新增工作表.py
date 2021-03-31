# -*- coding: utf-8 -*-
__author__ = 'ouyangmin'
__time__ = '2021/2/16 07:00'

#如果该代码不能满足您日常工作需求，可发邮件到liyysap@126.com ，说清楚您的需求，我尽力帮您实现哦
import os
import xlwings as xw
file_path = 'e:\\table\\销售表'    #新增工作表的工作簿的所在路径
file_list = os.listdir(file_path)   #列出文件夹下所有文件和子文件夹的名称
sheet_name = '产品销售区域'    #要新增的工作表的名称
app = xw.App(visible = False, add_book = False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths = os.path.join(file_path, i)   #构造新增工作表的工作簿的文件路径
    workbook = app.books.open(file_paths)   #根据路径打开需要新增工作表的工作簿
    sheet_names = [j.name for j in workbook.sheets]   #获取打开的工作簿中所有工作表的名称
    if sheet_name not in sheet_names:   #判断工作簿中是否不存在名为“产品销售区域”的工作表
        workbook.sheets.add(sheet_name)   #如果不存在，则新增工作表“产品销售区域”
        workbook.save()   #保存工作簿
app.quit()
