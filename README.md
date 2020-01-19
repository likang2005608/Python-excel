# Python-excel
Python 读写excel

在学习自动化框架过程中，其中涉及到参数化，需要用excel进行读取数据，学习一下，进行更多的了解 
读取excel可用：xlrd模块 
写入excel可用：xlwt模块 
读取excel表基础操作如下：

import os
from xlrd import open_workbook

file_path = os.path.abspath('.')+r'\data\baidu.xlsx'
print(file_path)
class Excel_Example(object):
    def read_excel(self):
        workbook = open_workbook(file_path)  #打开文件
        #获取所有sheet
        """
        以list形式展示：['1', '搜索']
        """
        print(workbook.sheet_names())
        #根据sheet索引或者名称获取sheet内容，索引从0开始
        sheet1 = workbook.sheet_by_index(0)
        sheet2 = workbook.sheet_by_name("搜索")
        #获取sheet的名称、行数 nrows、列数ncols
        print(sheet1.name,sheet1.nrows,sheet1.ncols)
        print(sheet2.name, sheet2.nrows, sheet2.ncols)

        #获得整行的值，这里从第二行开始输出的值
        for row in range(1,sheet1.nrows):
            print(sheet1.row_values(row))
        #获得整列的值,都以list的方式输出
        for col in range(0,sheet1.ncols):
            print(sheet1.col_values(col))
         #获取单元格内容
            print(sheet1.cell(1,0).value)
            print(sheet1.cell_value(1,0))
            print(sheet1.row(1)[0].value)
        #获取单元格数据类型
            print(sheet1.cell(1,0).ctype)

excel_example = Excel_Example()
excel_example.read_excel()

#以上大概包括如下方法：
打开excel：workbook = open_workbook(file_path)
获得excel表所有sheet页名称：workbook.sheet_names() 以list方式存储
获得某一页信息：sheet1 = workbook.sheet_by_index(0)
             sheet2 = workbook.sheet_by_name("搜索")
可通过sheet的索引（从0开始）、名称获取，返回一个对象
获取sheet的名称、行数 nrows、列数ncols：
sheet1.name,sheet1.nrows,sheet1.ncols
获取某一行的值：sheet1.row_values(row)
获取某一列的值：sheet1.col_values(col)  返回的也为一个list
获取单元格内容
获取单元格内容类型，其中类型输出的数字具体表示类型为：
ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error 

excel写操作： 
主要涉及的方法包括： 
addsheet(‘新增sheet名字’) 
write（行，列，rsheet.cell_value(row,col)值，style） 
save（文件名） 
简单的示例：

import xlrd,xlwt

rpath = r'E:\\score.xlsx'
wpath = r'E:\\output.xlsx'
rbook=xlrd.open_workbook(rpath)
rsheet = rbook.sheet_by_index(0)
nc = rsheet.ncols
rsheet.put_cell(0,nc,xlrd.XL_CELL_TEXT,'总分',None)#添加单元格数据
for row in range(1,rsheet.nrows):
    row_value = rsheet.row_values(row)
    row_score = row_value[1] + row_value[2]
    #row_score = sum(rsheet.row_values(row,1))  
      rsheet.put_cell(row,nc,xlrd.XL_CELL_NUMBER,row_score,None)
wbook = xlwt.Workbook()#创建一个写入excel对象
wsheet = wbook.add_sheet(rsheet.name)#添加一个sheet
style = xlwt.easyxf('align:vertical center,horizontal center')#设置数据展示格式

for row in range(rsheet.nrows):
    for col in range(rsheet.ncols):    
#将原表的每一行数据写入新表中    wsheet.write(row,col,rsheet.cell_value(row,col),style)
wbook.save(wpath)    #保存    

参考文章： 
http://blog.csdn.net/chengxuyuanyonghu/article/details/54951399 基础知识包括读写 
http://blog.csdn.net/fzch_struggling/article/details/45100937

http://blog.csdn.net/suofiya2008/article/details/6284208 针对excel2007，从sql中写入数据到excel 
接下来需要用到的接着学习，至少不懵逼了
————————————————
版权声明：本文为CSDN博主「DDQ_DQ」的原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接及本声明。
原文链接：https://blog.csdn.net/DDQ_DQ/article/details/78097318
