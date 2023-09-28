import os
import xlrd
import xlsxwriter
from number import department

work=xlsxwriter.Workbook('./output.xls')
# 新建一个sheet
sheet=work.add_worksheet('combine')
x1 = 1;
x2 = 1;

for key in department:
    workbook = xlrd.open_workbook(f'./{key}.xls')
    sheet_name = workbook.sheet_names()

    for file_1 in sheet_name:
        table = workbook.sheet_by_name(file_1)
        rows = table.nrows
        clos = table.ncols


        for i in range(1,rows):
            sheet.write_row('A' + str(x1), table.row_values(i))
            sheet.write('D' + str(x1), f'{key}')
            x1 += 1

    print('正在合并第%d个文件 ' % x2)
    print('已完成 ' + key)
    x2 += 1;

work.close()