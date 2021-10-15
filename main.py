#!/usr/bin/python3

# HW2)) bigdata homework for excel

# import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook(filename='student.xlsx')
sheet_ranges = wb['Sheet1']
# print(sheet_ranges['A2'].value)

# --update total score--#
row_num = 2
while True:
    mid = sheet_ranges.cell(row=row_num, column=3).value  # 30%
    fin = sheet_ranges.cell(row=row_num, column=4).value  # 35%
    hw = sheet_ranges.cell(row=row_num, column=5).value  # 34%
    att = sheet_ranges.cell(row=row_num, column=6).value  # 1% (just add)
    if mid is None:
        break

    total = mid * 0.3 + fin * 0.35 + hw * 0.34 + att
    row_num += 1

    sheet_ranges.cell(row=row_num - 1, column=7, value=total)

# print(row_num)  # 12

# --add total score to list--#
score = []
# print(sheet_ranges.cell(row=row2, column=7).value)

row_num = 2
while row_num < 12:  # row_num 2~11
    score.append(sheet_ranges.cell(row=row_num, column=7).value)
    row_num += 1

# score.append(sheet_ranges.cell(row=2, column=7).value)
# score.append(sheet_ranges.cell(row=3, column=7).value)
# score.append(sheet_ranges.cell(row=4, column=7).value)
# score.append(sheet_ranges.cell(row=5, column=7).value)
# score.append(sheet_ranges.cell(row=6, column=7).value)
# score.append(sheet_ranges.cell(row=7, column=7).value)
# score.append(sheet_ranges.cell(row=8, column=7).value)
# score.append(sheet_ranges.cell(row=9, column=7).value)
# score.append(sheet_ranges.cell(row=10, column=7).value)
# score.append(sheet_ranges.cell(row=11, column=7).value)

# --About rank--#
rank = []
cnt = 0
# print(len(score))

score.sort()
# print(score)

# --update grade--#
row_num = 2
while row_num < 12:  # row_num 2~11
    v = score.index(sheet_ranges.cell(row=row_num, column=7).value)

    if v == 9:
        g = 'A+'
    elif v >= 7:
        g = 'A0'
    elif v >= 5:
        g = 'B+'
    elif v >= 3:
        g = 'B0'
    elif v == 2:
        g = 'C+'
    else:
        g = 'C0'

    sheet_ranges.cell(row=row_num, column=8, value=g)
    row_num += 1

# print(row_num)  # row num = 12

wb.save(filename="output.xlsx")  # file_save
