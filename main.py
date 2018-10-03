# -*- coding: utf-8 -*-
import xlrd

xls = xlrd.open_workbook('imi2018.xls')
sheet_names = xls.sheet_names()

print('Sheet Names', sheet_names)
sheet = xls.sheet_by_name(('4 курс_ИТ').encode('utf-8'))

row_to_day = {
    4: 'Понедельник',
    10: 'Вторник',
    16: 'Среда',
    22: 'Четверг',
    28: 'Пятница',
    34: 'Суббота',
}

cell_obj = sheet.cell(2,8)
print(cell_obj.value)

for row_idx in range(4, sheet.nrows):    # Iterate through rows
    cell_obj = sheet.cell(row_idx, 8)  # Get cell object by row, col
    if row_idx in row_to_day:
        print(row_to_day[row_idx])
    print (str((row_idx-4)%6+1) + ') ' + cell_obj.value)