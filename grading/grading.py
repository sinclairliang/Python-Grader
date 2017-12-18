import pandas as pd

import openpyxl
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
def readfile(address):
    wb = openpyxl.load_workbook(address)
    return wb


wb = readfile("/Users/Sinclair/Dropbox/students_submissions/testsubmission.xlsx")
ws = wb.active


def output_certain_cell(Row, column):
    coor = str(Row)+str(column)
    return (ws[coor].value)
    # return 0

# print(output_certain_cell('E',3))

for x in range (5,10):
    # print(output_certain_cell('E',x))
    this_coor = 'F' + str(x)
    if output_certain_cell('E',x) == None:
        # ws.cell(6, x).value='X values'
        ws[this_coor] = 'ss'
        print(this_coor)
    else:
        ws[this_coor].value = "1"

wb.save('write_example.xlsx')



