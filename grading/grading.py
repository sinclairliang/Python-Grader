import pandas as pd

import openpyxl
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
def readfile(address):
    wb = openpyxl.load_workbook(address)
    return wb


wb = readfile("/Users/Sinclair/Dropbox/students_submissions/Lab2StudentSubmission_test.xlsx")
# wb = readfile("/Users/Sinclair/Dropbox/students_submissions/testsubmission.xlsx")
ws = wb.active


def output_certain_cell(Row, column):
    coor = str(Row)+str(column)
    return (ws[coor].value)



def grant_exist(row_alphabet, assign_points):
    for x in range (7,162):
        # print(output_certain_cell('E',x))
        this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor = this_coor_alpha + str(x)
        student_coor = 'C' + str(x)
        if output_certain_cell(row_alphabet,x) == None:
            # ws.cell(6, x).value='X values'
            ws[this_coor] = '0'
            print("student " + str(ws[student_coor].value) + " has no answer")
        else:
            ws[this_coor].value = assign_points

def grant_include(row_alphabet, countain, assign_points):
    for x in range (7,162):
        # print(output_certain_cell('E',x))
        this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor = this_coor_alpha + str(x)
        student_coor = 'C' + str(x)
        if output_certain_cell(row_alphabet,x) == None:
            ws[this_coor] = '0'
            print("student " + str(ws[student_coor].value) + " has no answer")
        elif str(countain) in str(output_certain_cell(row_alphabet,x)):
            ws[this_coor].value = assign_points
        else:
            ws[this_coor] = '0'
            print("student " + str(ws[student_coor].value) + " has a wrong answer")

grant_exist('E',20)
grant_exist('H',2)
grant_exist('J',1)
grant_include('L','1.1',1)

wb.save('write_example.xlsx')



