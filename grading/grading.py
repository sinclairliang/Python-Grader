import time
import sys
import openpyxl
from openpyxl.compat import range
import random


def readfile(address):
    workbook = openpyxl.load_workbook(address)
    return workbook

# address = input("Please type in address of the file you would like to grade?\n")
# wb = readfile("")
# wb = readfile(input("Please type in address of the file you would like to grade (*.xlsx)?\n"))
# destination_address = input("Please type in the address you would like to save the file with score\n")
# wb = readfile("/Users/Sinclair/Dropbox/students_submissions/testsubmission.xlsx")
wb = readfile("/Users/Sinclair/Dropbox/students_submissions/Lab6StudentSubmissions_test.xlsx")
# wb = readfile("/Users/Sinclair/Dropbox/students_submissions/Lab2StudentSubmission_test.xlsx")
ws = wb.active
first_row = int(input("what is the row number of first student?\n"))  # 8
last_row = int(input("what is the row number of last student?\n"))  # 161
student_number = last_row - first_row + 2  # 161-8+2


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
    timesleep = random.uniform(0.05, 0.1)
    # bar = pyprind.ProgBar(student_number, bar_char='=', monitor=True, update_interval=2, title="\nNow grading "+str(ws[column_alphabet+str(6)].value))
    print("\nNow grading: "+str(ws[column_alphabet+str(first_row-2)].value))
    for x in range(first_row-1, last_row+1):
        # print(output_certain_cell('E',x))

        sys.stdout.write('\r')
        # sys.stdout.write("%.2f%%" % (((x+1)/(162))*100))
        # sys.stdout.write("%s %.2f%%" % ("Progress:" (((x+1)/(162))*100)))
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        # student_coor = 'C' + str(x)
        time.sleep(timesleep)
        if output_certain_cell(column_alphabet, x) is None:
            ws[this_coor] = 0
            # print("student " + str(ws[student_coor].value) + " has no answer")
        else:
            ws[this_coor].value = assign_points
        # bar.update()
    # print("\n***Finish grading question: " + str(ws[column_alphabet+str(6)].value)+'***\n')

    print("\n✓")
    # time.sleep(0.01)
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
    timesleep = random.uniform(0.05, 0.2)
    # bar = pyprind.ProgBar(student_number, bar_char='=', monitor=True, update_interval=2, title="\nNow grading "+str(ws[column_alphabet+str(6)].value))
    print("\nNow grading: "+str(ws[column_alphabet+str(first_row-2)].value))
    for x in range(first_row, last_row+1):
        # print(output_certain_cell('E',x))
        # this_coor_alpha = chr(ord(row_alphabet)+1)
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        sys.stdout.write('\r')
        # sys.stdout.write("%.2f%%" % (((x+1)/(162))*100))
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        # student_coor = 'C' + str(x)
        if output_certain_cell(column_alphabet, x) is None:
            ws[this_coor] = '0'
            # print("student " + str(ws[student_coor].value) + " has no answer")
        elif str(contain) in str(output_certain_cell(column_alphabet, x)):
            ws[this_coor].value = assign_points
        elif str(contain1) in str(output_certain_cell(column_alphabet, x)):
            ws[this_coor].value = assign_points
        elif str(contain2) in str(output_certain_cell(column_alphabet, x)):
            ws[this_coor].value = assign_points
        elif str(contain3) in str(output_certain_cell(column_alphabet, x)):
            ws[this_coor].value = assign_points
        else:
            ws[this_coor] = 0
            # print("student " + str(ws[student_coor].value) + " has a wrong answer")
        time.sleep(timesleep)
        # bar.update()
    # print("\n***Finish grading question: " + str(ws[column_alphabet+str(6)].value)+'***\n')
    print("\n✓")
    # return 0


def gradinglab2():

    """
    The main() programme which conducts everything
    :return: the graded excel file
    """
    print("Grading initialised...")
    time.sleep(0.01)
    grant_exist('E', 20)
    # grant_exist('H', 2)
    # grant_exist('J', 1)
    # grant_include('L', 1, '1.1')
    # grant_include('R', 1, 'No', 'both', 'either', 'Both')
    # grant_include('T', 1, 'irefox', 'NT')
    # grant_include('X', 1, 'yes', 'Yes', '200')
    # grant_include('V', 1, '1.1')
    # grant_include('Z', 1, '25')
    # grant_include('AB', 1, '25')
    # grant_include('AD', 1, '3446')
    # grant_include('AF', 1, 'Apache', '2.2.3')
    # grant_include('AH', 1, 'image')
    # grant_include('AM', 1, '5')
    # grant_include('AO', 1, '5')
    # grant_include('AQ', 1, 'GET', 'POST')
    # grant_exist('AS', 1)
    # grant_include('AU', 1, 'educe')
    # grant_include('AY', 1, 'requesting')
    # grant_include('BE', 1, 'Protocol', 'protocol')
    # grant_include('BI', 1, '4')
    # grant_include('BK', 1, '2', '6')
    # grant_include('BM', 1, '5')
    # grant_include('BO', 1, '3')
    # time.sleep(0.01)

    # wb.save(destination_address+"result.xlsx")
    wb.save("result.xlsx")
    print("The result file has been successfully saved!")
    print("Finished!")
    return 0


def gradinglab6():

    """
    The main() programme which conducts everything
    :return: the graded excel file
    """
    print("Grading initialised...")
    time.sleep(0.01)
    grant_include('D', 2, '24')
    grant_include('F', 1, '3')
    grant_include('H', 1, '128.64.8.0')
    time.sleep(0.01)
    wb.save("result.xlsx")
    print("The result file has been successfully saved!")
    print("Finished!")
    return 0


def main():
    # gradinglab2()
    gradinglab6()


main()
