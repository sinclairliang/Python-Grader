import time
import sys
import openpyxl
from openpyxl.compat import range
import random

def readfile(address):
    workbook = openpyxl.load_workbook(address)
    return workbook

wb = readfile(input("Please type in address of the file you would like to grade (*.xlsx)?\n"))
destination_address = input("Please type in the address you would like to save the file with score\n")
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
    timesleep = random.uniform(0.05, 0.1)
    print("\nNow grading: "+str(ws[column_alphabet+str(first_row-2)].value))
    for x in range(first_row-1, last_row+1):
        sys.stdout.write('\r')
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        time.sleep(timesleep)
        if output_certain_cell(column_alphabet, x) is None:
            ws[this_coor] = 0
        else:
            ws[this_coor].value = assign_points
    print("\n✓")


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

    timesleep = random.uniform(0.05, 0.2)
    print("\nNow grading: "+str(ws[column_alphabet+str(first_row-2)].value))
    for x in range(first_row, last_row+1):
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        sys.stdout.write('\r')
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        if output_certain_cell(column_alphabet, x) is None:
            ws[this_coor] = '0'
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
        time.sleep(timesleep)
    print("\n✓")


def gradinglab2():

    """
    The main() programme which conducts everything
    :return: the graded excel file
    """
    print("Grading initialised...")

    time.sleep(0.01)
    grant_exist('E', 20)
    time.sleep(0.01)
    wb.save(destination_address+"result.xlsx")
    return 0


def main():
    start_time = time.time()
    gradinglab2()
    print("The result file has been successfully saved!")
    print("%s %.2fs" % ("Finish grading! Time elapsed:", (time.time() - start_time)))


main()

