# -*- coding: utf-8 -*-

import xlrd
import xlwt
import easygui
import datetime
import time



path = easygui.fileopenbox()
print(path)
'''
while(input_str_break != "100"):
    input_str_break = input("打开失败\n")
    print("输入了 %s" %(input_str_break))
'''

# excle 路径
#path=r"D:\gitRepo\work_value\temp2.xlsx"

#sheet名字
sheet_name_str = "2019年"

#计算时间，开始年月日
startdate_str = "2019-07-01"
startdate_time = time.strptime(startdate_str,"%Y-%m-%d")
startdate_datetime = datetime.datetime(startdate_time[0],startdate_time[1],startdate_time[2])
start_days = (startdate_datetime- datetime.datetime(1899,12,31)).days+1  # 差值 比excle计算出来少1 excel是第几天，这里计算是差几天，所以要从19891231开始计算并最后加上1
 

enddate_str ="2019-09-30"
enddate_time= time.strptime(enddate_str,"%Y-%m-%d")
enddate_datetime =datetime.datetime(enddate_time[0],enddate_time[1],enddate_time[2])
enddate_days =( enddate_datetime- datetime.datetime(1899,12,31)).days+1  # 差值 比excle计算出来少1 excel是第几天，这里计算是差几天，所以要从19891231开始计算并最后加上1
#可以用于计算
print("days + %f"  %(enddate_datetime-startdate_datetime).days +"\n")

#打开文件
excel_book = xlrd.open_workbook(path) 

#获取第一个sheet页面
excel_table_1 = excel_book.sheet_by_index(0)
print(excel_table_1.name)

#获取有效行数
sheet1_nrows = excel_table_1.nrows
sheet1_ncols = excel_table_1.ncols

list_val = [[] for i in range(sheet1_nrows)]
print(sheet1_nrows)

#每一行都保存为一个列表
for i in range(0,sheet1_nrows):
    list_val[i] = excel_table_1.row(i)

excel_sheet_1 = excel_table_1
#第二行为基本信息
name_col = 0 #x姓名
pro_start_clo = 0 #开始时间
pro_preprod_clo = 0 #样机
pro_end_clo = 0 #结束时间
pro_end_rea_clo = 0 #实际结束时间
pro_hard_index_clo = 0 #难度系数
pro_quality_clo = 0 #质量系数
pro_status_clo = 0 #状态栏

for i in range(0,sheet1_ncols):
    if(excel_table_1.cell(1,i).ctype != 1):
#        print(i)
        continue
    list_val[1][i] = excel_table_1.cell(1,i).value
#    print(list_val[1][i] )
    if(r"项目经理" in list_val[1][i] and ((r"产品" in list_val[1][i]) != True)):
        name_col = i # 从0开始 所以减1 
    if(r"启动" in list_val[1][i]):
        pro_start_clo = i 
    if(r"计划" in list_val[1][i]):
        pro_end_clo = i 
    if(r"样机时间" in list_val[1][i] and ((r"计划" in list_val[1][i]) != True)):
        pro_preprod_clo = i 
    if("调整" in list_val[1][i]):
        pro_end_rea_clo = i 
    if("难度" in list_val[1][i]):
        pro_hard_index_clo = i 
    if("质量" in list_val[1][i]):
        pro_quality_clo = i 
    if("状态" in list_val[1][i]):
        pro_status_clo = i 

print(name_col , #x姓名
pro_start_clo , #开始时间
pro_preprod_clo,#样机时间点
pro_end_clo , #结束时间
pro_end_rea_clo , #实际结束时间
pro_hard_index_clo , #难度系数
pro_quality_clo , #质量系数
pro_status_clo ,) #状态栏)

name_list =[]
value_list=[]
for i in range(100):
    value_list.append(0)
for i in range(2,sheet1_nrows):
    
     # 判断时间 当前开始和最后的时间
    startdate_pro = 0 #用于计算本季度的时间
    startdate_pro_real = 0 #实际的时间
    if(excel_table_1.cell(i,pro_start_clo).ctype != 0):
        startdate_pro_real = max(startdate_pro_real, excel_table_1.cell(i,pro_start_clo).value)
    else:
        startdate_pro_real = 0
        startdate_pro = 0
    if(startdate_pro_real >= enddate_days):
        startdate_pro = 0
    else:
        startdate_pro = startdate_pro_real

    if(startdate_pro_real <= start_days):
        startdate_pro = start_days
#    else:
#        startdate_pro = startdate_pro_real
    
    '''
        以上赋值,加入判断
        假设开始时间小于季度最晚时间，则初始时间作为0，0表示后续不做任何处理
        假设开始时间早于季度时间，本季度从季度时间开始算
    '''

    enddate_pro = 0 
    enddata_pro_real = 0  #实际的结束 时间
    if(excel_table_1.cell(i,pro_preprod_clo).ctype != 0):
        enddata_pro_real = max(enddata_pro_real, excel_table_1.cell(i,pro_preprod_clo).value)
    if(excel_table_1.cell(i,pro_end_clo).ctype != 0):
        enddata_pro_real = max(enddata_pro_real, excel_table_1.cell(i,pro_end_clo).value)
    if(excel_table_1.cell(i,pro_end_rea_clo).ctype != 0):
        enddata_pro_real = max(enddata_pro_real, excel_table_1.cell(i,pro_end_rea_clo).value)
    if(excel_table_1.cell(i,pro_preprod_clo).ctype ==0 and excel_table_1.cell(i,pro_end_clo).ctype == 0 and excel_table_1.cell(i,pro_end_rea_clo).ctype == 0):
        enddata_pro_real = 0
    '''
    print("excel_table_1.cell(i,pro_preprod_clo).ctype,excel_table_1.cell(i,pro_preprod_clo).value")
    print(excel_table_1.cell(i,pro_preprod_clo).ctype,excel_table_1.cell(i,pro_preprod_clo).value)
    print("excel_table_1.cell(i,pro_end_rea_clo).ctype,excel_table_1.cell(i,pro_end_rea_clo).value")
    print(excel_table_1.cell(i,pro_end_rea_clo).ctype,excel_table_1.cell(i,pro_end_rea_clo).value)
    print("excel_table_1.cell(i,pro_end_clo).ctype,excel_table_1.cell(i,pro_end_clo).value")
    print(excel_table_1.cell(i,pro_end_clo).ctype,excel_table_1.cell(i,pro_end_clo).value)
    print("enddata_pro_real")
    print(enddata_pro_real)
    '''
    if(enddata_pro_real <= start_days):
        enddate_pro = 0
    else:
        enddate_pro = enddata_pro_real

    if(enddata_pro_real >= enddate_days ):
        enddate_pro = enddate_days
  
    '''     
        以上完成几个判断
        （1） 三个都为0 则默认为0
        （2） 否则则找一个最大值
        （3） 如果最大的结束时间小于开始时间，则赋值0 ，0表示不做任何处理
        （4） 如果结束时间大于季度末时间，则计算的时候本季度末作为时间。
    '''


    
    # 形成名字列表
    
    hard_index = 0
    if(excel_table_1.cell(i,pro_hard_index_clo).ctype == 0):
        hard_index=1.0
    else :
        hard_index=excel_table_1.cell(i,pro_hard_index_clo).value
    
    quality_index =  0
    if(excel_table_1.cell(i,pro_quality_clo).ctype == 0):
        quality_index=1.0
    else :
        quality_index=excel_table_1.cell(i,pro_quality_clo).value

    if(enddate_pro != 0 and startdate_pro!=0):#本季度的不为0 ，则进行计算
    # 首先要计算当前获得的所有值，本季度的结束时间，减去开始时间，
        temp1  =  (enddate_pro-startdate_pro_real)*hard_index*quality_index
        temp2  =  (startdate_pro - startdate_pro_real)*hard_index*1.0 # 默认此前获得的时间质量系数都为1
        # 计算此前的值，两者相减
        value_stage = temp1 -  temp2 # 计算值
        print("%d val_stage %f" %(i,value_stage))
    else:
        value_stage = 0
    
    #减去此前获得的值

    #以上计算完成，开始进行统计导入数据

    #找到在列表中名字的序号，用value列表进行赋值相加
    '''不在在名字列表中红，并且不为空，则加入name_list '''
    if((excel_table_1.cell(i,name_col).value in name_list) != True  and excel_table_1.cell(i,name_col).value != ''):
        name_list.append(excel_table_1.cell(i,name_col).value)

#    print(name_list)
    if((excel_table_1.cell(i,name_col).value in name_list) == True):
        temp_j = name_list.index(excel_table_1.cell(i,name_col).value)
        value_list[temp_j] += value_stage

workbook = xlwt.Workbook(encoding = 'utf-8')
# 创建一个worksheet
worksheet = workbook.add_sheet("贡献值",cell_overwrite_ok=True)

for i in range(len(name_list)):    
    print(name_list[i]+": %f" %(value_list[i]))
    worksheet.write(i,1,name_list[i])
    worksheet.write(i,2,value_list[i])
temp_j = len(name_list)
worksheet.write(temp_j,1,'开始日期')
worksheet.write(temp_j+1,1,'结束日期')

dateStyle = xlwt.XFStyle()
dateStyle.num_format_str="yyyy/mm/dd"

worksheet.write(len(name_list),2,startdate_datetime,dateStyle)
worksheet.write(len(name_list)+1,2,enddate_datetime,dateStyle)

workbook.save("贡献值.xls")


