import xlrd,sys,time
import easygui
from datetime import date,datetime


path = easygui.fileopenbox()
wb = xlrd.open_workbook(path) #打开文件  
print(wb.sheet_names())#获取所有表格名字


sheet1 = wb.sheet_by_index(0)#通过索引获取表格
print(dir(xlrd))
print(sheet1)
print(sheet1.name,sheet1.nrows,sheet1.ncols)

rows = sheet1.row_values(2)#获取行内容
cols = sheet1.col_values(3)#获取列内容
print(rows)
print(cols)
input('Press Enter to exit...ssssss')
print(sheet1.cell(1,0).value)#获取表格里的内容，三种方式
print(sheet1.cell_value(1,0))
print(sheet1.row(1)[0].value)
input('Press Enter to exit...2')


