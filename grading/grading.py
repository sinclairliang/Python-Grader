import pandas as pd
import time
import sys

import openpyxl
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
def readfile(address):
    wb = openpyxl.load_workbook(address)
    return wb


wb = readfile("/Users/Sinclair/Dropbox/students_submissions/Lab2StudentSubmission_test.xlsx")
# wb = readfile("/Users/Sinclair/Dropbox/students_submissions/testsubmission.xlsx")
ws = wb.active


def output_certain_cell(row, column):
    coor = str(row)+str(column)
    return ws[coor].value

import sys



def grant_exist(row_alphabet, assign_points):
    print("Start grading question: " + str(ws[row_alphabet+str(6)].value))


    for x in range (7,162):
        # print(output_certain_cell('E',x))
        this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor = this_coor_alpha + str(x)
        student_coor = 'C' + str(x)

        sys.stdout.write('\r')
        # the exact output you're looking for:
        sys.stdout.write("[%-20s] %d%%" % ('='*i, 5*i))
        sys.stdout.flush()
        time.sleep(0.2)

        if output_certain_cell(row_alphabet,x) is None:
            ws[this_coor] = '0'
            # print("student " + str(ws[student_coor].value) + " has no answer")
        else:
            ws[this_coor].value = assign_points
    print("***Finish grading question: " + str(ws[row_alphabet+str(6)].value)+'***\n')


def grant_include(row_alphabet, countain, assign_points):
    print("Start grading question: " + str(ws[row_alphabet+str(6)].value))
    for x in range(7, 162):
        # print(output_certain_cell('E',x))
        this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor = this_coor_alpha + str(x)
        student_coor = 'C' + str(x)
        if output_certain_cell(row_alphabet, x) is None:
            ws[this_coor] = '0'
            print("student " + str(ws[student_coor].value) + " has no answer")
        elif str(countain) in str(output_certain_cell(row_alphabet,x)):
            ws[this_coor].value = assign_points
        else:
            ws[this_coor] = '0'
            print("student " + str(ws[student_coor].value) + " has a wrong answer")
    print("***Finish grading question: " + str(ws[row_alphabet+str(6)].value)+'***\n')

# grant_exist('E', 20)
# grant_exist('H', 2)
# grant_exist('J', 1)
# grant_include('L', '1.1', 1)


# wb.save('write_example.xlsx')

for i in range(1,21):
    sys.stdout.write('\r')
    # the exact output you're looking for:
    sys.stdout.write("[%-20s] %d%%" % ('='*i, 5*i))
    sys.stdout.flush()
    time.sleep(0.25)


