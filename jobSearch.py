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
        getID: int = input("enter ID: \n> ").lower()
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
            else:
                print("Not Found")
                break

    def searchByCompany(self):
        isCompanyStr = True
        getCompanyName = input("enter Company Name: \n> ").lower()
        while isCompanyStr:
            if not isinstance(getCompanyName, str):
                print("Not a string value")
                getCompanyName = input("enter Company Name: \n> ").lower()
                break
            else:
                print(f"Searched {getCompanyName} \n\n")
                break

        for row in self.worksheet.iter_rows(min_row=1, values_only=True):
            cellValue = row[1]
            if cellValue == getCompanyName:
                print(f"Data: {cellValue}")
                print(f"Matches: {row[2], row[3]}")
            # else:
            #     print("Not Found")
            #     break

    def searchbyField(self, fieldIdx, fields):
        isFieldType = True

        getFieldInput = input(f"Enter {fields} \n> ")
        while isFieldType:
            if fields == self.colName[1:6] and isinstance(getFieldInput, str):
                print(f"Searched: {getFieldInput} \n\n")
                break
            else:
                print("Not a string value")
                getFieldInput = input("enter Company Name: \n> ").lower()
                break

            if fields == self.colName[1] and isinstance(getFieldInput, int):
                print(f"Searched: {getFieldInput}")
                break

    def searchByLink(self):
        pass


    def powerSearch(self):
        pass
