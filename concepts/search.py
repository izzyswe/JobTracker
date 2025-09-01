from openpyxl import load_workbook
from datetime import datetime
# CONSTANTS
COLNAME = ["id", "Company Name", "Position", "Address", "CV or Resume", "Web Link", "Status", "Date Applied", "Deadline"]
COLUMNS = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i']

WB = load_workbook("Job-Tracker.xlsx")
WS = WB.active


def searchByDate():
    isDateFormatted = True
    dateFormatStartInput = None
    dateFormatEndInput = None

    while isDateFormatted:
        getDateStartInput = input("date start: \n>")
        try:
            dateFormatStartInput = datetime.strptime(getDateStartInput, "%Y%m%d").date()
            break
        except ValueError:
            print("enter valid date (YYYYMMDD)")

    while isDateFormatted:
        getDateEndInput = input("date end: \n>")
        try:
            dateFormatEndInput = datetime.strptime(getDateEndInput, "%Y%m%d").date()
            break
        except ValueError:
            print("enter valid date (YYYYMMDD)")

    print(f"Searched: {dateFormatStartInput} - {dateFormatEndInput} \n")

    for row in WS.iter_rows(min_row=2, values_only=True):
        cellValue = row[7]
        if dateFormatStartInput <= dateFormatEndInput:
            print(f"Date ------- {cellValue}")
            print(f"Matches ----  {row[1], row[2], row[6]} \n")


def searchByID():
    pass
