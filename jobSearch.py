import os
from datetime import datetime
# from openpyxl import load_workbook


class jobSearch:

    colName = ["id", "Company Name", "Position", "Address", "CV or Resume", "Web Link", "Status", "Date Applied", "Deadline"]
    columns = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i']
    upperCol = [item.upper() for item in columns]
    checkOs = os.system('cls') if os.name == 'nt' else os.system('clear')

    def __init__(self, worksheet):
        self.worksheet = worksheet

    def searchByDate(self):
        isDateFormatted = True
        dateFormatStartInput = None
        dateFormatEndInput = None

        while isDateFormatted:
            getDateStartInput = input("date start: \n> ")
            try:
                dateFormatStartInput = datetime.strptime(getDateStartInput, "%Y%m%d").date()
                break
            except ValueError:
                print("enter valid date (YYYYMMDD)")

        while isDateFormatted:
            getDateEndInput = input("date end: \n> ")
            try:
                dateFormatEndInput = datetime.strptime(getDateEndInput, "%Y%m%d").date()
                break
            except ValueError:
                print("enter valid date (YYYYMMDD)")

        print(f"Searched: {dateFormatStartInput} - {dateFormatEndInput} \n")

        for row in self.worksheet.iter_rows(min_row=2, values_only=True):
            cellValue = row[7]
            if dateFormatStartInput <= dateFormatEndInput:
                print(f"Date ------- {cellValue}")
                print(f"Matches ----  {row[1], row[2], row[6]} \n")

    def searchByID(self):
        isIDNumber = True
        getID: int = input("enter ID: \n> ")
        while isIDNumber:
            if not getID.isnumeric():
                print("error: not a numerical value")
                getID: int = input("enter ID: \n> ")
                break
            else:
                print(f"\n Searched ID: {getID} \n")
                break

        for row in self.worksheet.iter_rows(min_row=1, values_only=True):
            cellValue = row[0]
            if cellValue == getID:
                print(f"Data: {cellValue}")
                print(f"Matches: {row[1], row[2]}")

    def searchByLink(self):
        pass

    def searchByColumn(self):
        pass

    def powerSearch(self):
        pass
