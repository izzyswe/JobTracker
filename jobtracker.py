'''
    Automation: Job Tracker
    Github: Izzyswe
    Name: Isaac

    Date: August 18, 2025
'''
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
# import datetime


# all the neccesary columns
colName = ["id", "Company Name", "Position", "Address", "CV or Resume", "Web Link", "Status", "Date Applied", "Deadline"]
columns = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i']
upperCol = [item.upper() for item in columns]


wb = Workbook()
# get active worksheet
ws = wb.active


# [ALMOST DONE] ADD NEW DATA EVERY ROW, HOWEVER, DON'T ERASE WHEN RERUNNING THE PROGRAM
def getJobDetails():
    print("JOB ENTRY")
    jobDetails = {}
    for i in colName:
        jobTitle = input(f"enter {i} \n >")
        jobDetails[i] = jobTitle

    print("\nFinal details:")
    for i in colName:
        print(f"   {i}: {{ {jobDetails[i]} }}")

    # add to excel sheet
    # Grab the next empty row
    next_row = ws.max_row + 1

    # Add the job details to the new row
    for col_idx, col in enumerate(colName, start=1):
        # Use the column index to get the column letter
        column_letter = get_column_letter(col_idx)
        cell = column_letter + str(next_row)
        # Set the value in the cell
        ws[cell] = jobDetails[col]


def searchJobDetail():
    pass


def updateJobDetail():
    pass


# [DONE] GENERATE TITLE AND ALL COLUMN NAMES
def generateTitles():
    heading_fill = PatternFill(fill_type='solid', start_color='87f5a4', end_color='87f5a4')
    heading_font = Font(size=16, bold=True)
    for i, n in zip(upperCol, colName):
        cell = i + "2"
        ws[cell] = n

    # after filling row 1 with your headers
    for col_idx, header in enumerate(colName, start=1):
        column_letter = get_column_letter(col_idx)
        # set width a bit bigger than the header length
        ws.column_dimensions[column_letter].width = len(header) + 2

    ws.merge_cells('A1:I1')
    ws['A1'] = "Job Tracker"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = heading_font
    ws['A1'].fill = heading_fill


def mainMenu():
    while True:
        print("|  JOB TRACKER MENU | \n\n"
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
                getJobDetails()
            case "2":
                searchJobDetail()
            case "3":
                updateJobDetail()
            case _:
                print("Invalid Input")


generateTitles()
mainMenu()

# Save the file
wb.save("Job-Tracker.xlsx")
