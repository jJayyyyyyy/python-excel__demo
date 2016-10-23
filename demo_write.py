#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# read from <demo_read.xls> and write into <demo_write.xls>

from xlrd import open_workbook
from xlwt import Workbook

wb = open_workbook('demo_read.xls')
sheet_0 = wb.sheets()[0]

book = Workbook()
sheet_new = book.add_sheet('Sheet_new')

for row in range(sheet_0.nrows):
    for col in range(sheet_0.ncols):
        value = sheet_0.cell(row, col).value
        sheet_new.write(row, col, value)

book.save('demo_write.xls')
















# i--->col_A(col_0), j--->col_D(col_3)
# k--->new

# i, j, k, same_index = 0, 0, 0, 0
# for i in range(sheet_0.nrows):
#     same_value = sheet_0.cell(i, 0).value
#     for j in range(same_index, sheet_0.nrows):
#         if sheet_0.cell(j, 3).value == same_value:
#             sheet_new.write(k, 0, same_value)
#             k = k + 1
#             same_index = j + 1
#             break
#         else:
#             pass
#     if same_index==sheet_0.nrows:
#         break


