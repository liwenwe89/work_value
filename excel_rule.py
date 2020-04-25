
import xlrd
import xlwt
import easygui
import datetime
import time



path = easygui.fileopenbox()
print(path)


#excle 路径
#path=r"D:\gitRepo\work_value\temp2.xlsx"


#打开文件
excel_book = xlrd.open_workbook(path) 

#获取第2个sheet页面，第一个sheet为目录页
excel_table_1 = excel_book.sheet_by_index(0)

#获取有效行数
sheet1_nrows = excel_table_1.nrows
sheet1_ncols = excel_table_1.ncols
print(sheet1_nrows)
print(sheet1_ncols)

#list_val = [[] for i in range(sheet1_nrows)]

temp_i = 0 #当前计数
temp_j = 0 #当前计数

list_val_single = [ str,str,str,str,str ]#项目名称，项目经理，本周进度，下周计划，风险
list_val = [list_val_single for i in range(20)] #最大二十个项目
count = 0
#每一行都保存为一个列表
for i in range(0,sheet1_nrows):
    for j in range(0,sheet1_ncols):
        value = excel_table_1.cell(i,j).value
        if (value == r"项目名称"):
            list_val[count][0] = excel_table_1.cell(i,j+1).value #+1为项目名称
            count = count+1
        if(value == r"汇报人"):
            list_val[count-1][1] = excel_table_1.cell(i,j+1).value
        if(value == r"本周进度"):
            list_val[count-1][2] = excel_table_1.cell(i+1,j).value
        if(value == r"下周计划"):
            list_val[count-1][3] = excel_table_1.cell(i+1,j).value
        if(value == r"处理中"):
            list_val[count-1][4] = list_val[count-1][4]+"\n问题描述:" + excel_table_1.cell(i,j-6).value + "\n问题描述："+excel_table_1.cell(i,j-3).valu
        print("%d,%d\n",i,j)
        print("%s",value)

ccc =0
