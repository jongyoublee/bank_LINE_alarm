import openpyxl
import math

EXCEL_PATH = '//Ctcnas/운영기획/1. JS폴더/10. 종엽/0. ★계약내용보고 - 2022(연도구분)_종엽이매일최신화.xlsx'
SHEET_NAME = '미수금'

workbook = openpyxl.load_workbook(EXCEL_PATH)
worksheet = workbook[SHEET_NAME]

print(worksheet.title)

wsList = []

for row in worksheet.iter_rows(min_row = 4, max_col = 7, values_only=True):
    if(row[1]==None):
        continue
    else:
        wsList.append((row[1], row[2], row[3], row[6]))

# For Loop 내부에서 Index 확인하는 법
for list in wsList:
    print(wsList.index(list))
