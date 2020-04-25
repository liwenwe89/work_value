
import xlrd
import xlwt
import easygui
import datetime
import time
import openpyxl

path = easygui.fileopenbox()
print(path)


#excle 路径
#path=r"D:\gitRepo\work_value\temp2.xlsx"


#打开文件
excel_book = xlrd.open_workbook(path) 

#获取第2个sheet页面，第一个sheet为目录页
excel_table_1 = excel_book.sheet_by_index(1)

#获取有效行数
sheet1_nrows = excel_table_1.nrows
sheet1_ncols = excel_table_1.ncols
print(sheet1_nrows)
print(sheet1_ncols)

#list_val = [[] for i in range(sheet1_nrows)]

temp_i = 0 #当前计数
temp_j = 0 #当前计数

list_val_single = ["","","","","",""]#项目名称，项目经理，本周进度，下周计划，风险
list_val = [list_val_single for i in range(100)] #最大二十个项目
list_link = [str for i in range(100)]
count = 0
#每一行都保存为一个列表
for i in range(0,sheet1_nrows):
    for j in range(0,sheet1_ncols):
        value = excel_table_1.cell(i,j).value
        list_val[count] = ["","","","","",""]
        if (value == r"项目名称"):
            list_val[count][0] = excel_table_1.cell(i,j+1).value #+1为项目名称
            list_link[count] = "B"+str(i+1)
            count = count+1
        if(value == r"汇报人"):
            list_val[count-1][1] = excel_table_1.cell(i,j+1).value
        if(value == r"本周进度"):
            list_val[count-1][2] = excel_table_1.cell(i+1,j).value
        if(value == r"下周计划"):
            list_val[count-1][3] = excel_table_1.cell(i+1,j).value
        if(value == r"处理中"):
            list_val[count-1][4] = list_val[count-1][4]+"\n 问题描述:" + str(excel_table_1.cell(i,j-6).value) + "\n 本周进展："+ str(excel_table_1.cell(i,j-3).value)
        print("%d,%d\n",i,j)
        print("%s",value)

###  以上获取信息
workbook_0 = openpyxl.load_workbook(path)
worksheet_0= workbook_0.worksheets[0]
worksheet_1= workbook_0.worksheets[1]
print(worksheet_0)
rows=worksheet_0.max_row #最大行列
columns=worksheet_0.max_column
print(rows,columns)  #32 13
for i in range (1,rows):
    for j in range(1,columns):
        worksheet_0.cell(i,j,"") #清空
    

title=['项目名称', '项目经理', '本周进度', '下周计划', '风险']
for i in range(1,1+len(title)):
    worksheet_0.cell(1, i,title[i-1])
for i in range(2,2+count):
    for j in range(1,1+len(list_val_single)):    
        worksheet_0.cell(i, j,list_val[i-2][j-1])
        if (j == 1):
            str_temp = r"#sheet0!" + str(list_link[i-0])
            worksheet_0.cell(i, j,list_val[i-2][j-1]).hyperlink = str_temp
workbook_0.save(path)
