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
