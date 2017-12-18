import pandas as pd

import openpyxl
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
def readfile(address):
    wb = openpyxl.load_workbook(address)
    return wb


wb = readfile("/Users/Sinclair/Dropbox/students_submissions/testsubmission.xlsx")
ws = wb.active
# sheet = wb.get_sheet_by_name('Sheet1')
# print(ws['C9'].value)

def output_certain_row(Row,start, end):
    for column in range(start, end):
        coor = str(Row)+str(column)
        print(ws[coor].value)
    return 0

output_certain_row('E',3,10)

