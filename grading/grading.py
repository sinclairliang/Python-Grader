# Sinclair Liang

import time
import sys
import openpyxl
from openpyxl.compat import range
import random


try:
    origin_address = input("Please type in address of the file you would like to grade (*.xlsx)?\n")
    workbook = openpyxl.load_workbook(origin_address)
except:
    print("This is not a valid address\n")
    sys.exit()
try:
    destination_address = input("Please type in the address you would like to save the file with score\n"
                                "Leave it blank if you want to store in the same directory.\n")
except:
    destination_address = origin_address
    pass


worksheet = workbook.active
first_row = int(input("what is the row number of first student?\n"))
last_row = int(input("what is the row number of last student?\n"))
student_number = last_row - first_row + 2


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
    return worksheet[coor].value


def grant_exist(column_alphabet, assign_points):
    """
    to grade questions which would give points as long as there is an answer existing
    :param column_alphabet: The column that the answers are in
    :param assign_points: points students should receive
    :return: None, modified the excel file
    """
    timesleep = random.uniform(0.05, 0.1)
    print("\nNow grading: "+str(worksheet[column_alphabet+str(first_row-2)].value))
    for x in range(first_row-1, last_row+1):
        sys.stdout.write('\r')
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        time.sleep(timesleep)
        if output_certain_cell(column_alphabet, x) is None:
            worksheet[this_coor] = 0
        else:
            worksheet[this_coor].value = assign_points
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
    print("\nNow grading: "+str(worksheet[column_alphabet+str(first_row-2)].value))
    for x in range(first_row, last_row+1):
        this_coor_alpha = next_column(column_alphabet)
        this_coor = this_coor_alpha + str(x)
        sys.stdout.write('\r')
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        if output_certain_cell(column_alphabet, x) is None:
            worksheet[this_coor] = '0'
        elif str(contain) in str(output_certain_cell(column_alphabet, x)):
            worksheet[this_coor].value = assign_points
        elif str(contain1) in str(output_certain_cell(column_alphabet, x)):
            worksheet[this_coor].value = assign_points
        elif str(contain2) in str(output_certain_cell(column_alphabet, x)):
            worksheet[this_coor].value = assign_points
        elif str(contain3) in str(output_certain_cell(column_alphabet, x)):
            worksheet[this_coor].value = assign_points
        else:
            worksheet[this_coor] = 0
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
    return 0


def main():
    start_time = time.time()
    gradinglab2()
    workbook.save(destination_address+"result.xlsx")
    # workbook.close()
    print("The result file has been successfully saved!")
    end_time = time.time()
    print("%s %d m %.2f s " % ("Finish grading! Time elapsed:", int((end_time - start_time)/60), (end_time - start_time)%60))

main()

