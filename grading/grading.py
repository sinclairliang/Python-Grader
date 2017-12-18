import pandas as pd
import openpyxl
xl = ("/Users/Sinclair/Dropbox/students_submissions/Lab2StudentSubmission.xlsx")
wb = openpyxl.load_workbook(xl)
ws = wb.active
sheet = wb.get_sheet_by_name('Sheet1')
print(ws['C9'].value)
# s = pd.ExcelFile(xl)
# df = s.parse("Sheet1")
# print(df.head())
