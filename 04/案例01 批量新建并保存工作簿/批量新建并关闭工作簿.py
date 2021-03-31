# -*- coding: utf-8 -*-
__author__ = 'ouyangmin'
__time__ = '2021/2/14 15:43'

#如果该代码不能满足您日常工作需求，可发邮件到liyysap@126.com ，说清楚您的需求，我尽力帮您实现哦

import xlwings as xw                                   #导入xlwings模块
app = xw.App(visible=True, add_book=False)             #启动Excel程序，但是不新建工作簿
for i in range(20):
    workbook = app.books.add()                         #新建工作簿
    workbook.save(f'e:\\practice\\table{i}.xlsx')           #保存新建的多个工作簿
    workbook.close()                                   #关闭工作簿
app.quit()                                             #退出

