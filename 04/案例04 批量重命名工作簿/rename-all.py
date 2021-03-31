# -*- coding: utf-8 -*-
__author__ = 'ouyangmin'
__time__ = '2021/2/16 07:00'


import os
file_path = 'e:\\table\\产品销售表' 
file_list = os.listdir(file_path)
old_book_name = '销售表'          #给出工作簿名中需要替换的旧的关键字
new_book_name = '分部产品销售表'     #给出工作簿名中要替换为的新关键字
for i in file_list:
    if i.startswith('~$'):     #判断是否有文件名以“~$”开头的临时文件
        continue     #如果有，则跳过这种类型的文件
    new_file = i.replace(old_book_name, new_book_name)    #执行查找和替换，生成新的工作簿名
    old_file_path = os.path.join(file_path, i)     #构造需要重命名工作簿的完整路径
    new_file_path = os.path.join(file_path, new_file)    #构造重命名后工作簿的完整路径
    os.rename(old_file_path, new_file_path)   #执行重命名
