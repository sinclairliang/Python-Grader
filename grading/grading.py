import pandas as pd
import time
import sys
import pyprind
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
    n = 162-7
    timesleep = 0.05
    bar = pyprind.ProgBar(n, bar_char='=', monitor=True, update_interval=2,title=str(ws[row_alphabet+str(6)].value))
    for x in range(7,162):
        # print(output_certain_cell('E',x))
        this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor = this_coor_alpha + str(x)
        student_coor = 'C' + str(x)
        time.sleep(timesleep)
        if output_certain_cell(row_alphabet,x) is None:
            ws[this_coor] = '0'
            # print("student " + str(ws[student_coor].value) + " has no answer")
        else:
            ws[this_coor].value = assign_points
        bar.update()
    print("***Finish grading question: " + str(ws[row_alphabet+str(6)].value)+'***\n')
    # print(bar)
    # return 0


def grant_include(row_alphabet, assign_points, countain, countain1=None, countain2=None, countain3=None):
    print("Start grading question: " + str(ws[row_alphabet+str(6)].value))
    n = 162-7
    timesleep = 0.05
    bar = pyprind.ProgBar(n, monitor=True, title=str(ws[row_alphabet+str(6)].value))
    for x in range(7, 162):
        # print(output_certain_cell('E',x))
        this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor = this_coor_alpha + str(x)
        student_coor = 'C' + str(x)
        if output_certain_cell(row_alphabet, x) is None:
            ws[this_coor] = '0'
            # print("student " + str(ws[student_coor].value) + " has no answer")
        elif str(countain) in str(output_certain_cell(row_alphabet,x)):
            ws[this_coor].value = assign_points
        elif str(countain1) in str(output_certain_cell(row_alphabet,x)):
            ws[this_coor].value = assign_points
        elif str(countain2) in str(output_certain_cell(row_alphabet,x)):
            ws[this_coor].value = assign_points
        elif str(countain3) in str(output_certain_cell(row_alphabet,x)):
            ws[this_coor].value = assign_points
        else:
            ws[this_coor] = '0'
            # print("student " + str(ws[student_coor].value) + " has a wrong answer")
        time.sleep(timesleep)
        bar.update()
    print("\n***Finish grading question: " + str(ws[row_alphabet+str(6)].value)+'***\n')
    # return 0



grant_exist('E', 20)
grant_exist('H', 2)
grant_exist('J', 1)
grant_include('L', 1, '1.1',)
grant_include('R', 1, 'No','both','either','Both')


wb.save('write_example.xlsx')

# for i in range(1,21):
#     sys.stdout.write('\r')
#     # the exact output you're looking for:
#     sys.stdout.write("[%-20s] %d%%" % ('='*i, 5*i))
#     sys.stdout.flush()
#     time.sleep(0.25)



# n = 100
# timesleep = 0.05
#
# bar = pyprind.ProgBar(n)
# for i in range(n):
#     time.sleep(timesleep) # your computation here
#     bar.update()
