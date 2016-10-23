#!/usr/bin/env python
# -*- coding: utf-8 -*-

#	1. from xxx import xxx
from xlrd import open_workbook, cellname

#	2. read from demo.xls
wb = open_workbook('demo_read.xls')

#	3. two ways to read the first sheet of workbook
sheet_0 = wb.sheets()[0]
#	sheet_0 = wb.sheet_by_index(0)

#	4. print the name of this sheet, and total rows and cols of this sheet.
sheet_name = sheet_0.name
total_rows = sheet_0.nrows
total_cols = sheet_0.ncols
# print(sheet_name, total_rows, total_cols)

#	5. print the value the cell, (row_num, col_num)
value = sheet_0.cell(3, 1).value
value = sheet_0.cell_value(0, 1)
value = sheet_0.row(3)[0].value
# print(value)

#	6. 整行整列打印
row_0 = sheet_0.row(0)
row_0 = sheet_0.row_values(0)
col_1 = sheet_0.col(1)
col_1 = sheet_0.col_values(1)
# print(col_1)

#	7. 切片
col_0_row_0to4 = sheet_0.col_values(0, 0, 3)
row_0_col_2toEnd = sheet_0.row_slice(0, 2)
row_1_col_1to3 = sheet_0.row_slice(1, 1, 4)
# print(col_0_row_0to4)

#	8. etc. cell_name, output below is <B2>
cell_name = cellname(2, 1)
# print(cell_name)

#	print content of every sheet in the workbook
for sheet in wb.sheets():
	print(sheet.name)
	for row in range(sheet.nrows):
		values = []
		for col in range(sheet.ncols):
			values.append(sheet.cell(row, col).value)
		print(', '.join(values))
	print()

#	print type and value of every cell
for row in range(sheet_0.nrows):
	for col in range(sheet_0.ncols):
		print(sheet_0.cell_type(row, col), sheet_0.cell_value(row, col))
