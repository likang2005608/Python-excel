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
