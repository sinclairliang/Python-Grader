# Sinclair Liang

import time
import sys
import openpyxl
from openpyxl.compat import range
import random



def banner():
    banner = '''
    _____       _   _                        _____               _           
    |  __ \     | | | |                      / ____|             | |          
    | |__) |   _| |_| |__   ___  _ __ ______| |  __ _ __ __ _  __| | ___ _ __ 
    |  ___/ | | | __| '_ \ / _ \| '_ \______| | |_ | '__/ _` |/ _` |/ _ \ '__|
    | |   | |_| | |_| | | | (_) | | | |     | |__| | | | (_| | (_| |  __/ |   
    |_|    \__, |\__|_| |_|\___/|_| |_|      \_____|_|  \__,_|\__,_|\___|_|   
            __/ |                                                             
           |___/                                                              
   
'''
    return banner


sys.stdout.write(banner())

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
    grant_exist("D",1)
    grant_exist("F",1)
    grant_exist("H",1)
    grant_exist("J",1)
    grant_exist("L",1)
    grant_exist("N",1)
    grant_exist("P",1)
    grant_exist("R",1)
    grant_exist("T",1)
    grant_exist("V",1)
    grant_exist("X",1)
    grant_exist("Z",1)
    grant_exist("AB",1)
    grant_exist("AD",1)
    grant_exist("AF",1)
    grant_exist("AH",1)
    grant_exist("AJ",1)
    grant_exist("AK",1)
    grant_exist("AL",1)
    grant_exist("AN",1)
    grant_exist("AP",1)
    grant_exist("AP",1)
    grant_exist("AU",1)
    grant_include("AW",2, "68.10.20")
    grant_include("AY",1, "255.255.25","24")
    grant_exist("BA",1)
    grant_include("BC",2, "68.10.20")
    grant_include("BE",1, "255.255")
    grant_include("BG",1,"No","no")
    grant_include("BI",1,"68.10.20","24")
    grant_include("BK",1, "anta","ruz")
    grant_include("BM",5, "100.0.0")
    grant_include("BO",2, "255.255.255","24")
    grant_exist("BQ",5)
    grant_include("BS",1, "3")
    grant_include("CA",2,"Fa4/0", "4/0")
    grant_include("CC",1,"ICMP")
    grant_include("CE",5,"68.10.20")
    grant_include("CG",5,"67.10.14")

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


if __name__ == '__main__':
    main()
