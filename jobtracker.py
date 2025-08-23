'''
    Automation: Job Tracker
    Github: Izzyswe
    Name: Isaac

    Date: August 18, 2025
'''
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
# import datetime

######## TODO
#### FIXES
# [fixed]   1. fix getJobDetails function
# [started] 2. work on Search JobDetail Function
# 3. work on updateJobDetail

#### FUTURE FEATURE WANTS
# 1. Add GUI
# 2. DRAG N DROP TXT FILE CONVERSION
# 3. WEB SCRAP LIVE CHANGES


# all the neccesary columns
colName = ["id", "Company Name", "Position", "Address", "CV or Resume", "Web Link", "Status", "Date Applied", "Deadline"]
columns = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i']
upperCol = [item.upper() for item in columns]


# [ADDED] KEPT OVERRIDING SO THIS FUNCTION WILL CHECK IF ON EXIST
# OR AUTOGENERATE IF DOES NOT EXIST TO PREVENT THE ORGINAL ISSUE
def checkWorksheet(filename="Job-Tracker.xlsx"):
    if os.path.exists(filename):
        wb = load_workbook(filename)
    else:
        wb = Workbook()
        ws = wb.active
        generateTitles(ws)
    return wb


# [FIXED] THIS CODE WAS WAY MORE SIMPLER THAN PREVIOUS CODE
def getJobDetails(worksheet):
    print("JOB ENTRY")
    rowData = []
    # loop throught the colName list to get input for every column
    for i in colName:
        value = input(f"Enter {i}\n> ")
        # append every value into the rowData array
        rowData.append(value)
    
    # --- DEBUG LINE ---
    print(f"DEBUG: The worksheet's max_row is currently {worksheet.max_row}")
    # ------------------

    # in the worksheet argument, append the rowData in openpyxl ws
    worksheet.append(rowData)
    print("-- Entry Added! --\n")


def searchJobDetail():
    pass


def updateJobDetail():
    pass


# [DONE] GENERATE TITLE AND ALL COLUMN NAMES
def generateTitles(worksheet):
    heading_fill = PatternFill(fill_type='solid', start_color='87f5a4', end_color='87f5a4')
    heading_font = Font(size=16, bold=True)
    for i, n in zip(upperCol, colName):
        cell = i + "2"
        worksheet[cell] = n

    # after filling row 1 with your headers
    for col_idx, header in enumerate(colName, start=1):
        column_letter = get_column_letter(col_idx)
        # set width a bit bigger than the header length
        worksheet.column_dimensions[column_letter].width = len(header) + 2

    worksheet.merge_cells('A1:I1')
    worksheet['A1'] = "Job Tracker"
    worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    worksheet['A1'].font = heading_font
    worksheet['A1'].fill = heading_fill


def mainMenu(worksheet):
    while True:
        print("\n|  JOB TRACKER MENU | \n\n"
              "1. Enter Job Entry \n" +
              "2. Search Job Entry \n" +
              "3. Update Job Entry \n" +
              "4. Quit Program")
        opts = input("> ")
        print()
        if opts == "4":
            print("Exiting program.")
            break

        match opts:
            case "1":
                getJobDetails(worksheet)
            case "2":
                searchJobDetail()
            case "3":
                updateJobDetail()
            case _:
                print("Invalid Input")


# [ADDED] CREATED AN ENTRY POINT FOR ALL CALLS AND CODE TO BE IN
# SINCE ITS NOT CLASS BASED, IM NOT CREATING A MIAN FUNCTION
if __name__ == "__main__":
    try:
        wb = checkWorksheet()
        ws = wb.active
        mainMenu(ws)
    except KeyboardInterrupt:
        print("\n")
    finally:
        # Save the file
        wb.save("Job-Tracker.xlsx")
        print("workload saved")
