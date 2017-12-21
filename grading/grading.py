import pandas as pd
import time
import sys
import pyprind
import openpyxl
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
import random



def readfile(address):
    wb = openpyxl.load_workbook(address)
    return wb

# address = input("Please type in address of the file you would like to grade?\n")
# wb = readfile("")
wb = readfile(input("Please type in address of the file you would like to grade (*.xlsx)?\n"))
destination_address = input("Please type in the address you would like to save the file with score\n")
# wb = readfile("/Users/Sinclair/Dropbox/students_submissions/testsubmission.xlsx")
ws = wb.active
student_number = 162-7






def next_column(current_column):
    """
    to help the programme determine where the next column is, especially with 'Z' and double letter cases
    :param current_column: the column that the programme is grading
    :return: the next column,
    """
    if len(current_column) == 1:
        if (ord(current_column) - 64) <= 25:
            return chr(ord(current_column)+1)
        elif (ord(current_column) - 64) == 26:
            return "AA"
    elif len(current_column) == 2:
        if current_column[-1] == "Z":
            return chr(ord(current_column[0])+1) + "A"
        else:
            return current_column[0] + chr(ord(current_column[1])+1)

def output_certain_cell(row, column):
    """
    output the value that stores in certain cell
    :param row: the desired row of such cell
    :param column: the desired column of such cell
    :return: the content in this cell
    """
    coor = str(row)+str(column)
    return ws[coor].value

def grant_exist(column_alphabet, assign_points):
    """
    to grade questions which would give points as long as there is an answer existing
    :param column_alphabet: The column that the answers are in
    :param assign_points: points students should receive
    :return: None, modified the excel file
    """
    # print("Start grading question: " + str(ws[row_alphabet+str(6)].value))
    timesleep = random.uniform(0.05, 0.4)
    bar = pyprind.ProgBar(student_number, bar_char='=', monitor=True, update_interval=2, title="Now grading "+str(ws[column_alphabet+str(6)].value))
    for x in range(7, 162):
        # print(output_certain_cell('E',x))
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        # student_coor = 'C' + str(x)
        time.sleep(timesleep)
        if output_certain_cell(column_alphabet, x) is None:
            ws[this_coor] = 0
            # print("student " + str(ws[student_coor].value) + " has no answer")
        else:
            ws[this_coor].value = assign_points
        bar.update()
    print("\n***Finish grading question: " + str(ws[column_alphabet+str(6)].value)+'***\n')
    # print(bar)
    # return 0

def grant_include(column_alphabet, assign_points, contain, contain1=None, contain2=None, contain3=None):
    """
    to grade questions which would give points as long as some words match the answer
    :param column_alphabet: The column that the answers are in
    :param assign_points: points students should receive
    :param contain: The word the answers should include
    :param contain1: optional word the answers should include
    :param contain2: optional word the answers should include
    :param contain3: optional word the answers should include
    :return: None, modified the excel file
    """
    # print("Start grading question: " + str(ws[row_alphabet+str(6)].value))
    timesleep = random.uniform(0.05,0.4)
    bar = pyprind.ProgBar(student_number, bar_char='=', monitor=True, update_interval=2, title="Now grading "+str(ws[column_alphabet+str(6)].value))
    for x in range(7, 162):
        # print(output_certain_cell('E',x))
        # this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        # student_coor = 'C' + str(x)
        if output_certain_cell(column_alphabet, x) is None:
            ws[this_coor] = '0'
            # print("student " + str(ws[student_coor].value) + " has no answer")
        elif str(contain) in str(output_certain_cell(column_alphabet,x)):
            ws[this_coor].value = assign_points
        elif str(contain1) in str(output_certain_cell(column_alphabet,x)):
            ws[this_coor].value = assign_points
        elif str(contain2) in str(output_certain_cell(column_alphabet,x)):
            ws[this_coor].value = assign_points
        elif str(contain3) in str(output_certain_cell(column_alphabet,x)):
            ws[this_coor].value = assign_points
        else:
            ws[this_coor] = 0
            # print("student " + str(ws[student_coor].value) + " has a wrong answer")
        time.sleep(timesleep)
        bar.update()
    print("\n***Finish grading question: " + str(ws[column_alphabet+str(6)].value)+'***\n')
    # return 0

# for i in range(1,21):
#     sys.stdout.write('\r')
#     # the exact output you're looking for:
#     sys.stdout.write("[%-20s] %d%%" % ('='*i, 5*i))
#     sys.stdout.flush()
#     time.sleep(0.25)

def startgrading():
    """
    The main() programme which conducts everthing
    :return: the graded excel file
    """
    print("Grading initialised...")
    time.sleep(0.01)
    grant_exist('E', 20)
    grant_exist('H', 2)
    grant_exist('J', 1)
    grant_include('L', 1, '1.1')
    grant_include('R', 1, 'No', 'both', 'either', 'Both')
    grant_include('T', 1, 'irefox', 'NT')
    grant_include('X', 1, 'yes', 'Yes', '200')
    grant_include('V', 1, '1.1')
    grant_include('Z', 1, '25')
    grant_include('AB', 1, '25')
    grant_include('AD', 1, '3446')
    grant_include('AF', 1, 'Apache', '2.2.3')
    grant_include('AH', 1, 'image')

    print("Finished!")
    # "/Users/Sinclair/practice/grading/write_example.xlsx"
    wb.save(destination_address+"result.xlsx")

startgrading()




