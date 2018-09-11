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


def grant_exist(column_letter, assign_points):
    """
    to grade questions which would give points as long as there is an answer existing
    :param column_letter: The column that the answers are in
    :param assign_points: points students should receive
    :return: None, modified the excel file
    """
    timesleep = random.uniform(0.05, 0.1)
    print("\nNow grading: "+str(worksheet[column_letter+str(first_row-2)].value))
    for x in range(first_row-1, last_row+1):
        sys.stdout.write('\r')
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        this_coor_alpha = next_column(column_letter)
        this_coor = this_coor_alpha + str(x)
        time.sleep(timesleep)
        if output_certain_cell(column_letter, x) is None:
            worksheet[this_coor] = 0
        else:
            worksheet[this_coor].value = assign_points
    print("\n✓")


def grant_include(column_letter, assign_points, contain, contain1=None, contain2=None, contain3=None):
    """
    to grade questions which would give points as long as some words match the answer
    :param column_letter: The column that the answers are in
    :param assign_points: points students should receive
    :param contain: The word the answers should include
    :param contain1: optional word the answers should include
    :param contain2: optional word the answers should include
    :param contain3: optional word the answers should include
    :return: None, modified the excel file
    """

    timesleep = random.uniform(0.05, 0.2)
    print("\nNow grading: "+str(worksheet[column_letter+str(first_row-2)].value))
    for x in range(first_row, last_row+1):
        this_coor_alpha = next_column(column_letter)
        this_coor = this_coor_alpha + str(x)
        sys.stdout.write('\r')
        sys.stdout.write("%s %.2f%%" % ("Progress:\t", (((x+1)/(last_row+1))*100)))
        sys.stdout.flush()
        if output_certain_cell(column_letter, x) is None:
            worksheet[this_coor] = '0'
        elif str(contain) in str(output_certain_cell(column_letter, x)):
            worksheet[this_coor].value = assign_points
        elif str(contain1) in str(output_certain_cell(column_letter, x)):
            worksheet[this_coor].value = assign_points
        elif str(contain2) in str(output_certain_cell(column_letter, x)):
            worksheet[this_coor].value = assign_points
        elif str(contain3) in str(output_certain_cell(column_letter, x)):
            worksheet[this_coor].value = assign_points
        else:
            worksheet[this_coor] = 0
        time.sleep(timesleep)
    print("\n✓")


def gradingfinalproject():

    """
    The main() programme which conducts everything
    :return: the graded excel file
    """
    print("Grading initialised...")

    time.sleep(0.01)
    grant_include("C", 1, "192.168.0.1")
    grant_include("E", 1, "00:50:0F:B7:64:45")
    grant_include("G", 1, "255.255.255.0", "/24")
    grant_exist("I", 1)
    grant_include("K", 1, "192.168.0")
    grant_include("M", 1, "00:01:C7:33:65:16")
    grant_include("K", 1, "192.168.0")
    grant_include("M", 1, "00:01:C7:33:65:16")
    grant_include("O", 1, "192.168.0")
    grant_include("Q", 1, "00:02:16:17:D7:A2")
    grant_include("S", 1, "192.168.0")
    grant_exist("U", 1)
    grant_include("W", 1, "192.168.0")
    grant_include("Y", 1, "93", "7C")
    grant_include("AE", 1, "Chotchkie")
    grant_include("AG", 1, "all")
    grant_include("AI", 1, "rinter")
    grant_exist("AK", 5)
    grant_exist("AM", 1)
    grant_include("AO", 1, "68.10.20")
    grant_include("AQ", 1, "255.255.255.0", "/24")
    grant_include("AS", 1, "Jans")
    grant_include("AU", 2, "67.10.14")
    grant_include("AW", 1, "255.255.255.0", "/24")
    grant_include("AY", 1, "68.10.20.1")
    grant_include("BA", 1, "Santa", "Cruz")
    grant_include("BC", 1, "255.255.255.0", "/24", "100.0.0.0")
    grant_include("BE", 1, "255.255.255.0", "/24")
    grant_include("BF", 1, "2", "3")
    grant_exist("BI", 1)
    grant_include("BQ", 1, "ICMP")
    grant_include("BS", 1, "68.10.20")
    grant_include("BU", 1, "67.10.14")

    time.sleep(0.01)
    return 0


def main():
    start_time = time.time()
    gradingfinalproject()
    workbook.save(destination_address+"result.xlsx")
    # workbook.close()
    print("The result file has been successfully saved!")
    end_time = time.time()
    print("%s %d m %.2f s " % ("Finish grading! Time elapsed:", int((end_time - start_time)/60), (end_time - start_time)%60))

main()

