# coding=utf-8
import xlrd
import xlsxwriter
import datetime

Workbook_In = xlrd.open_workbook("data.xlsx")  # 打开数据表格
Sheet_In = Workbook_In.sheet_by_name("Sheet1")  # 注意sheet的名称
Title = Sheet_In.row(0)  # 表头
Order_Number = Sheet_In.nrows  # 订单个数
Machine_Number = 1  # 机器个数
Order_List = []  # 订单列表
Machine_Name = []  # 机器名称列表

#读取数据表
for i in range(1,Order_Number):
    data = Sheet_In.row(i)
    dic = {}
    for j in range(Sheet_In.ncols):
        cell = data[j]
        if j == 5 and cell.ctype == 2:
            val = xlrd.xldate_as_datetime(cell.value, 0)
        else:
            val = cell.value
        dic[Title[j].value] = val
        if j == 0:
            if i == 1:
                Machine_Name.append(val)
            elif val != Machine_Name[Machine_Number-1]:
                Machine_Name.append(val)
                Machine_Number = Machine_Number + 1
    # print(dic)
    Order_List.append(dic)

Order_Number = Order_Number - 1  # 除去表头行
# print(Order_List)
# print(Machine_Name)
# print(Machine_Number)
# print(Order_Number)

#比较时间先后函数
def Compare (Date_1 , Date_2):
    Val_1 = Date_1.year * 13 * 32 + Date_1.month * 32 + Date_1.day
    Val_2 = Date_2.year * 13 * 32 + Date_2.month * 32 + Date_2.day
    # print(Val_1,"+",Val_2)
    return Val_1 < Val_2
    
Find = 0
Cost = []  # 各订单总耗时

#寻找最早开始时间
for i in range(Order_Number):
    Now_Order = Order_List[i]
    Cost.append(Now_Order["总量（万个）"] / Now_Order["效率（万个/班）"])
    if(Now_Order["开始日期"] != ""):
        if(Find == 0):
            First_Day = Now_Order["开始日期"]
            Find = 1
        elif Compare(Now_Order["开始日期"] , First_Day) == 1:
            First_Day = Now_Order["开始日期"]

#时间自增函数
def Next_Day (Dat):
    Y = Dat.year
    M = Dat.month
    D = Dat.day
    R = 0
    if Y % 400 == 0:
        R = 1
    elif Y % 4 == 0 and Y % 100 != 0:
        R = 1
    if M==2:
        if R == 1 and D == 29:
            M = 3
            D = 1
        elif R == 0 and D == 28:
            M = 3
            D = 1
        else:
            D = D + 1
    elif M == 1 or M == 3 or M == 5 or M == 7 or M == 8 or M == 10 or M == 12:
        if D == 31:
            D = 1
            if M == 12:
                Y = Y + 1
                M = 1
            else:
                M = M + 1
        else:
            D = D + 1
    else:
        if D == 30:
            M = M + 1
            D = 1
        else:
            D = D + 1
    return datetime.datetime(Y,M,D)
    # print(New_Date,"    ",New_Date.strftime("%w"))
    # if New_Date.strftime("%w") == 0:
    #     return Next_Day(New_Date)
    # else:
    #     return New_Date
    
Now_Date = First_Day
Last_Day = First_Day
Schedule_Dic = {}  # 各订单的时间和机器安排

# for i in range(300):
#     print(Now_Date)
#     Now_Date = Next_Day(Now_Date)

for i in range(Machine_Number):
    Find  = 0
    for j in range(Order_Number):
        Now_Order = Order_List[j]
        if Now_Order["机器"] == Machine_Name[i]:
            lst = []
            if Find == 0:
                Find = 1
                Left = 3
                if Now_Order["开始日期"] == "":
                    Now_Date = First_Day
                else:
                    Now_Date = Now_Order["开始日期"]
            elif Now_Order["开始日期"] != "" and Compare(Now_Date,Now_Order["开始日期"]) == 1:
                Now_Date = Now_Order["开始日期"]
                Left = 3
            if Now_Date.strftime("%w") == 0:
                Now_Date = Next_Day(Now_Date)
                Left = 3
            while Cost[j] != 0:
                if Cost[j] >= Left:
                    Cost[j] = Cost[j] - Left
                    lst.append({"日期" : Now_Date , "机器" : Machine_Name[i] , "数量" : Left * Now_Order["效率（万个/班）"]})
                    Now_Date = Next_Day(Now_Date)
                    if Now_Date.strftime("%w") == "0":
                        # print(Now_Date)
                        Now_Date = Next_Day(Now_Date)
                        # print(Now_Date)
                    Left = 3
                else:
                    Left = Left - Cost[j]
                    lst.append({"日期" : Now_Date , "机器" : Machine_Name[i] , "数量" : Cost[j] * Now_Order["效率（万个/班）"]})
                    Cost[j] = 0
            Schedule_Dic[j] = lst
    if Compare(Last_Day , Now_Date) == 1:
        Last_Day = Now_Date

# print(First_Day)
# print(Last_Day)
# for i in range(Order_Number):
#     print(Schedule_Dic[i])
#     print("---------------------")

Yea = First_Day.strftime("%Y年")
Fir = First_Day.strftime("（%m.%d")
Las = Last_Day.strftime("~%m.%d）")
Titlename = Yea + "生产计划预排" + Fir + Las
Filename = Titlename + ".xlsx"

Workbook_Out = xlsxwriter.Workbook(Filename)  # 建立输出表格
Sheet_Out = Workbook_Out.add_worksheet("排产")  # 建立Sheet

#初始化格式
Title_Style = Workbook_Out.add_format({
    'bold' : True ,  # 是否加粗
    'font' : "华文中宋" ,  # 字体设置
    'font_size' : 16 ,  # 文字大小设置
    'align' : "center" ,  # 水平位置设置:居中
	'valign' : "vcenter" ,  # 垂直位置设置:居中
    'font_color' : "black" ,  # 文字颜色设置
    'fg_color' : "white" ,  # 单元格背景颜色设置
    'text_wrap' : True ,  # 是否自动换行
    'border' : 5  # 框线宽度
})

Machine_Style = Workbook_Out.add_format({
    'bold' : True ,
    'font' : "宋体" ,
    'font_size' : 10 ,
    'align' : "center" ,
	'valign' : "vcenter" ,
    'font_color' : "black" ,
    'fg_color' : "#D8E4BC" ,
    'text_wrap' : True ,
    'border' : 1
})

Date_Style = Workbook_Out.add_format({
    'bold' : False ,
    'font' : "宋体" ,
    'font_size' : 10 ,
    'align' : "center" ,
	'valign' : "vcenter" ,
    'font_color' : "black" ,
    'fg_color' : "white" ,
    'text_wrap' : True ,
    'border' : 1
})

Date_Style_Sunday = Workbook_Out.add_format({
    'bold' : False ,
    'font' : "宋体" ,
    'font_size' : 10 ,
    'align' : "center" ,
	'valign' : "vcenter" ,
    'font_color' : "black" ,
    'fg_color' : "yellow" ,
    'text_wrap' : True ,
    'border' : 1
})

Normal_Style = Workbook_Out.add_format({
    'bold' : False ,
    'font' : "宋体" ,
    'font_size' : 10 ,
	'valign' : "vcenter" ,
    'font_color' : "black" ,
    'fg_color' : "white" ,
    'text_wrap' : True ,
    'border' : 1
})

Red_Style = Workbook_Out.add_format({
    'bold' : False ,
    'font' : "宋体" ,
    'font_size' : 10 ,
	'valign' : "vcenter" ,
    'font_color' : "red" ,
    'fg_color' : "white" ,
    'text_wrap' : True ,
    'border' : 1
})

Light_Line = Workbook_Out.add_format({'border' : 1})

Heavy_Line = Workbook_Out.add_format({'border' : 5})

Yellow = Workbook_Out.add_format({
    'fg_color' : "yellow" ,
    'border' : 1
})

Remain = Workbook_Out.add_format({
    'border' : 1 ,
    'text_wrap' : True
})

#写表头
Sheet_Out.set_row(0 , 40)
Sheet_Out.merge_range(0 , 0 , 0 , 3 * Machine_Number , "Merged Cells")
Sheet_Out.write(0 , 0 , Titlename , Title_Style)  # 第1行写标题
Sheet_Out.set_row(1 , 25)
Sheet_Out.set_column(0 , 0 , 8)
Sheet_Out.write(1 , 0 , "机台" , Machine_Style)  # 第2行写机器名称
for i in range(Machine_Number):
    Sheet_Out.write(0 , 3 * i + 1 , "" , Heavy_Line)
    Sheet_Out.write(0 , 3 * i + 2 , "" , Heavy_Line)
    Sheet_Out.write(0 , 3 * i + 3 , "" , Heavy_Line)
    Sheet_Out.write(1 , 3 * i + 2 , "" , Light_Line)
    Sheet_Out.write(1 , 3 * i + 3 , "" , Light_Line)
    Sheet_Out.merge_range(1 , 3 * i + 1 , 1 , 3 * i + 3 , "Merged Cells")
    Sheet_Out.write(1 , 3 * i + 1 , Machine_Name[i] , Machine_Style)
    Sheet_Out.set_column(3 * i + 1 , 3 * i + 2 , 15)
    Sheet_Out.set_column(3 * i + 3 , 3 * i + 3 , 8)

Sheet_Out.write(2 , 0 , "日期" , Date_Style)  # 表头各项信息栏
for i in range(Machine_Number):
    Sheet_Out.write(2 , 3 * i + 1 , "品牌" , Date_Style)
    Sheet_Out.write(2 , 3 * i + 2 , "工单号\n（备注）" , Date_Style)
    Sheet_Out.write(2 , 3 * i + 3 , "日产量（万印）" , Date_Style)

def Process (str):
    return '"' + str + '"'

def Add (now_order , str , S , Count):
    if Count == 0:
        if now_order["是否标红"] == 1:
            return S + ",Red_Style," + Process(str)
        else:
            return S + ",Normal_Style," + Process(str)
    else:
        if now_order["是否标红"] == 1:
            return S + ",Red_Style," + Process(" +" + str)
        else:
            return S + ",Normal_Style," + Process(" +" + str)

#写入生产计划
Now_Date = First_Day
cnt = 0
while(True):
    cnt = cnt + 1
    Form_Date = Now_Date.strftime("%m月%d日")
    if Now_Date.strftime("%w") == "0":
        Sheet_Out.write(cnt + 2 , 0 , Form_Date , Date_Style_Sunday)
    else:
        Sheet_Out.write(cnt + 2 , 0 , Form_Date , Date_Style)
    for i in range(Machine_Number):
        S1 = "(cnt+2,3*i+1"
        S2 = "(cnt+2,3*i+2"
        S3 = "(cnt+2,3*i+3"
        Count1 = 0
        Count2 = 0
        Count3 = 0
        R1 = 0
        R2 = 0
        R3 = 0
        for j in range(Order_Number):
            Now_Order = Order_List[j]
            if Now_Order["机器"] == Machine_Name[i]:
                lst = Schedule_Dic[j]
                for Dic in lst:
                    if Dic["日期"] == Now_Date:
                        S1 = Add(Now_Order , Now_Order["产品名称"] , S1 , Count1)
                        if Count1 == 0:
                            str1 = Now_Order["产品名称"]
                            if Now_Order["是否标红"] == 1:
                                R1 = 1
                        Count1 = Count1 +1
                        if Now_Order["工单号"] != "":
                            S2 = Add(Now_Order , Now_Order["工单号"] , S2 , Count2)
                            if Count2 == 0:
                                str2 = Now_Order["工单号"]
                                if Now_Order["是否标红"] == 1:
                                    R2 = 1
                            Count2 = Count2 +1
                        if Now_Order["产品名称"] != "换牌" and Now_Order["产品名称"] != "月保" and Now_Order["产品名称"] != "周保":
                            S3 = Add(Now_Order , str(round(Dic["数量"],2)) , S3 , Count3)
                            if Count3 == 0:
                                str3 = str(round(Dic["数量"],2))
                                if Now_Order["是否标红"] == 1:
                                    R3 = 1
                            Count3 = Count3 +1
        if Count1 == 1:
            if R1 == 1:
                Sheet_Out.write(cnt + 2 , 3 * i + 1 , str1 , Red_Style)
            else:
                Sheet_Out.write(cnt + 2 , 3 * i + 1 , str1 , Normal_Style)
        else:
            if Count1 == 0:
                S1 = "Sheet_Out.write" + S1 + ",'',Normal_Style"
                exec(S1+")")
            if Count1 > 1:
                S1 = "Sheet_Out.write_rich_string" + S1
                exec(S1+",Remain)")
        if Count2 == 1:
            if R2 == 1:
                Sheet_Out.write(cnt + 2 , 3 * i + 2 , str2 , Red_Style)
            else:
                Sheet_Out.write(cnt + 2 , 3 * i + 2 , str2 , Normal_Style)
        else:
            if Count2 == 0:
                S2 = "Sheet_Out.write" + S2 + ",'',Normal_Style"
                exec(S2+")")
            if Count2 > 1:
                S2 = "Sheet_Out.write_rich_string" + S2
                exec(S2+",Remain)")
        if Count3 == 1:
            if R3 == 1:
                Sheet_Out.write(cnt + 2 , 3 * i + 3 , str3 , Red_Style)
            else:
                Sheet_Out.write(cnt + 2 , 3 * i + 3 , str3 , Normal_Style)
        else:
            if Count3 == 0:
                S3 = "Sheet_Out.write" + S3 + ",'',Normal_Style"
                exec(S3+")")
            if Count3 > 1:
                S3 = "Sheet_Out.write_rich_string" + S3
                exec(S3+",Remain)")
    if Now_Date.strftime("%w") == "0":
        for i in range(Machine_Number):
            Sheet_Out.write(cnt + 2 , 3 * i + 1 , "" , Yellow)
            Sheet_Out.write(cnt + 2 , 3 * i + 2 , "" , Yellow)
            Sheet_Out.write(cnt + 2 , 3 * i + 3 , "" , Yellow)
    if Now_Date == Last_Day:
        break
    Now_Date = Next_Day(Now_Date)

Workbook_Out.close()  # 保存文件

# Made by 中国飞鱼