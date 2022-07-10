#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2022/7/10 11:38
# @Author : karinlee
# @FileName : workbook_split.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976
# @github : https://github.com/karinlee1988/
# @gitee : https://gitee.com/karinlee/
# @Personal website : https://karinlee.cn/



import openpyxl

class WorksheetSplitByColumn(object):
    """
    用于按列的内容拆分excel工作薄中的一个工作表，分别另存为多个独立的工作薄/打印
    """

    def __init__(self,path,column_index,title_index):
        self.path = path
        self.column_index = column_index
        self.title_index = title_index
        self.workbook = openpyxl.load_workbook(self.path)
        self.worksheet = self.workbook.active

    def get_worksheet_title(self):
        workbook_template = openpyxl.Workbook()
        worksheet_template = workbook_template.active
        for row in range(1,self.title_index+1):
            for column in (1,self.worksheet.max_column+1):
                worksheet_template.cell(row=row,column=column).value = self.worksheet.cell(row=row,column=column).value

        workbook_template.save('title.xlsx')





if __name__ == '__main__':
    path = '托收单附件明细待拆分打印.xlsx'
    app = WorksheetSplitByColumn(path,0,3)
    app.get_worksheet_title()






