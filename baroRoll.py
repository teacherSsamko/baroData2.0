import openpyxl
import os
from os import listdir


BASE_DIR = os.path.dirname(os.path.realpath(__file__))
LIST_DIR = os.path.join(BASE_DIR, "roll_list/")

list_files = listdir("roll_list")
print(list_files)

for f in list_files:
    if '$' in f:
        continue
    f = os.path.join(LIST_DIR, f)
    wb = openpyxl.load_workbook(f, read_only=True)

    sheets = wb.get_sheet_names()
    sheet = wb.get_sheet_by_name(sheets[0])

    count_students = sheet.max_row - 1
    print(count_students)
