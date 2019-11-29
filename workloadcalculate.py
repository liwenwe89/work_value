# -*- encoding: utf-8 -*-
##
#  @file workloadcommunicate.py
#  @brief 工作量统计
#  @author chendan5
#  @date 2019-9-5
#  @version 1.0
#  @warning 本文件的相关警告信息
#  @History 修改历史记录列表

import xlwt, xlrd
import openpyxl
from openpyxl import  Workbook
from openpyxl  import load_workbook
from openpyxl.styles import Font, colors, Alignment
from xlutils.copy import copy as xl_copy
import copy
import sys
import os
import datetime
import time
from datetime import date,datetime
import tkinter
from tkinter import filedialog


#打开文件
root = tkinter.Tk()
root.withdraw()
filename = filedialog.askopenfilename(initialdir=os.path.dirname(__file__), title="请选择excel文件",filetypes=(("excel文件", "*.xlsx"or"*.xls"), ("所有文件", "*.*")))

#读取excel数据
#ReadExcel = xlrd.open_workbook(r'工作量计算简化版.xlsx')
ReadExcel = xlrd.open_workbook(filename)
ReadExcelSheet=ReadExcel.sheet_by_index(0)
nrow = ReadExcelSheet.nrows  # 行数
ncol = ReadExcelSheet.ncols
print(nrow, ncol)

pmname = []
pmstart = []
pmend = []
pmcomplexity = []

if (ReadExcelSheet.cell_value(0, ncol-1) != '资源'):
    print("计算失败，请核对模板")
    os.system('pause')

for i in range(nrow):
    cell1 = ReadExcelSheet.cell_value(i, ncol-1)
    cell2 = ReadExcelSheet.cell_value(i, 15)
    if ("项目" in cell1 and cell2 == '开发中') or ("项目" in cell1 and cell2 == '已完成'):
        pmnamedata = ReadExcelSheet.cell_value(i,7)
        pmstartdata = str(ReadExcelSheet.cell_value(i,10))
        pmcomplexitydata = ReadExcelSheet.cell_value(i,19)
        timearrayshiji = str(ReadExcelSheet.cell_value(i,14))
        timearraysjihua = str(ReadExcelSheet.cell_value(i, 13))
        if  len(timearrayshiji) == 0:
            if len(timearraysjihua) == 0:#预研项目只有样机时间
                pmenddata = ReadExcelSheet.cell_value(i, 12)#发布时间是样机时间
            else:
                pmenddata = ReadExcelSheet.cell_value(i, 13)#发布时间是计划发布时间
        else:
            pmenddata = ReadExcelSheet.cell_value(i, 14)#发布时间是实际发布时间
        pmname.append(pmnamedata)
        pmstart.append(pmstartdata)
        pmend.append(pmenddata)
        pmcomplexity.append(pmcomplexitydata)

#写入新的excel数据
WriteExcel = xlwt.Workbook(encoding='utf-8') #创建Workbook，相当于创建Excel
# 创建sheet，Sheet1为表的名字，cell_overwrite_ok为是否覆盖单元格
WriteExcelsheet = WriteExcel.add_sheet(u'中间过程', cell_overwrite_ok=True)

#向表中添加数据
WriteExcelsheet.write(0, 0, '项目经理')
WriteExcelsheet.write(0, 1, '启动时间')
WriteExcelsheet.write(0, 2, '实际发布时间')
WriteExcelsheet.write(0, 3, '复杂度')
WriteExcelsheet.write(0, 4, '工作量')
WriteExcelsheet.write(0, 5, '换算后工作量')

for j in range(len(pmname)):
    WriteExcelsheet.write(j+1, 0, pmname[j])
    WriteExcelsheet.write(j+1, 1, pmstart[j])
    WriteExcelsheet.write(j+1, 2, pmend[j])
    WriteExcelsheet.write(j+1, 3, pmcomplexity[j])

#设置时间
ReadExcelSheet2=ReadExcel.sheet_by_index(1)
if (ReadExcelSheet2.cell_value(0, 0) != '开始时间') or (ReadExcelSheet2.cell_value(0, 1) != '结束时间'):
    print("计算失败，请核对模板")
    os.system('pause')

dt1 = ReadExcelSheet2.cell_value(1,0)
dt2 = ReadExcelSheet2.cell_value(1,1)

#dt1=43647.0 #'2019/7/1'
#dt2=43738.0 #'2019/9/30'

#dt3 = datetime.date(datetime.strptime(d1,'%Y/%m/%d'))
#dt4 = datetime.date(datetime.strptime(d2,'%Y-%m-%d'))
row = len(pmname)
print(row)

#计算工作量
for i in range(0,row):
    if( float(pmend[i]) < dt1):
        circle = 0
        WriteExcelsheet.write(i+1, 4, circle)
        WriteExcelsheet.write(i+1, 5, circle * pmcomplexity[i])
        continue
    elif (float(pmstart[i]) <= dt1) and (float(pmend[i]) <= dt2):#结束时间在范围内
        circle = float(pmend[i]) - dt1
        WriteExcelsheet.write(i+1, 4, circle)
        WriteExcelsheet.write(i+1, 5, circle * pmcomplexity[i])
        continue
    elif (float(pmstart[i]) <= dt1) and (float(pmend[i]) >= dt2):#超过范围
        circle = dt2 - dt1
        WriteExcelsheet.write(i+1, 4, circle)
        WriteExcelsheet.write(i + 1, 5, circle * pmcomplexity[i])
        continue
    elif (float(pmstart[i]) >= dt1) and (float(pmend[i]) <= dt2):#开始时间在范围内
        circle = float(pmend[i]) - float(pmstart[i])
        WriteExcelsheet.write(i+1, 4, circle)
        WriteExcelsheet.write(i + 1, 5, circle * pmcomplexity[i])
        continue
    elif (float(pmstart[i]) >= dt1) and (float(pmstart[i]) <= dt2) and (float(pmend[i]) >= dt2):
        circle = dt2 - float(pmstart[i])
        WriteExcelsheet.write(i+1, 4, circle)
        WriteExcelsheet.write(i+1, 5, circle * pmcomplexity[i])
        continue
    else:
        circle = 0
        WriteExcelsheet.write(i+1, 4, circle)
        WriteExcelsheet.write(i+1, 5, circle * pmcomplexity[i])

WriteExcel.save('test.xls')

#合并同项目经理
#读取excel数据
wb = xlrd.open_workbook(r'test.xls')
sheet=wb.sheet_by_index(0)
nnrow = sheet.nrows  # 行数
nncol = sheet.ncols
chendan = 0
chenwenjia = 0
wanghaili = 0
feixiaowei = 0
lihongyu = 0
wangleiqiang = 0
chenpeng = 0
jiangyannan = 0
mamengjuan = 0
lijunqiang = 0
panyu= 0
wuhuan = 0
liwenwei = 0
liuguosheng = 0
WriteExcelsheet = WriteExcel.add_sheet(u'最终结果',cell_overwrite_ok=True)

#向表中添加数据
WriteExcelsheet.write(0, 0, '项目经理')
WriteExcelsheet.write(0, 1, '总工作量')

list = ['陈丹','陈文佳','王海利','费晓伟','郦泓宇','王雷强','陈朋','姜彦男','马梦娟','李俊强','潘宇','武焕','李文伟','刘国胜']
for i in range(len(list)):
    WriteExcelsheet.write(i+1,0,list[i])

for i in range(nnrow):
    cell3 = sheet.cell_value(i, 0)
    cell4 = sheet.cell_value(i, 5)
    if ("陈丹" in cell3):
        chendan = chendan + cell4
        WriteExcelsheet.write(1, 1, chendan)
    if("陈文佳" in cell3):
        chenwenjia = chenwenjia + cell4
        WriteExcelsheet.write(2, 1, chenwenjia)
    if("王海利" in cell3):
        wanghaili = wanghaili + cell4
        WriteExcelsheet.write(3, 1, wanghaili)
    if("费晓伟" in cell3):
        feixiaowei = feixiaowei + cell4
        WriteExcelsheet.write(4, 1, feixiaowei)
    if ("郦泓宇" in cell3):
        lihongyu = lihongyu + cell4
        WriteExcelsheet.write(5, 1, lihongyu)
    if ("王雷强" in cell3):
        wangleiqiang = wangleiqiang + cell4
        WriteExcelsheet.write(6, 1, wangleiqiang)
    if ("陈朋" in cell3):
        chenpeng = chenpeng + cell4
        WriteExcelsheet.write(7, 1, chenpeng)
    if ("姜彦男" in cell3):
        jiangyannan = jiangyannan + cell4
        WriteExcelsheet.write(8, 1, jiangyannan)
    if ("马梦娟" in cell3):
        mamengjuan = mamengjuan + cell4
        WriteExcelsheet.write(9, 1, mamengjuan)
    if ("李俊强" in cell3):
        lijunqiang = lijunqiang + cell4
        WriteExcelsheet.write(10, 1, lijunqiang)
    if ("潘宇" in cell3):
        panyu = panyu + cell4
        WriteExcelsheet.write(11, 1, panyu)
    if ("武焕" in cell3):
        wuhuan = wuhuan + cell4
        WriteExcelsheet.write(12, 1, wuhuan)
    if ("李文伟" in cell3):
        liwenwei = liwenwei + cell4
        WriteExcelsheet.write(13, 1, liwenwei)
    if ("刘国胜" in cell3):
        liuguosheng = liuguosheng + cell4
        WriteExcelsheet.write(14, 1, liuguosheng)

WriteExcel.save('test.xls')
print("计算完成")
os.system('pause')
