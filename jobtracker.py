'''
    Automation: Job Tracker
    Github: Izzyswe
    Name: Isaac

    Date: August 18, 2025
'''
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
# import datetime


# all the neccesary columns
colName = ["id", "Company Name", "Position", "Address", "CV or Resume", "Web Link", "Status", "Date Applied", "Deadline"]
columns = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i']
upperCol = [item.upper() for item in columns]


wb = Workbook()
# get active worksheet
ws = wb.active


def mainMenu():
    generateTitles()
    while True:
        print("1. Enter Job Entry \n" +
              "2. Search Job Entry \n" +
              "3. Update Job Entry")
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


def getJobDetails():
    print("JOB ENTRY")
    jobDetails = {}
    for i in colName:
        jobTitle = input(f"enter {i} \n >")
        jobDetails[i] = jobTitle
        ws[jobDetails]

    print("\nFinal details:")
    for i in colName:
        print(f"   {i}: {{ {jobDetails[i]} }}")


def searchJobDetail():
    pass


def updateJobDetail():
    pass


# [FIXED FUTURE PROBLEM] this has a fixed column with new data in every row
# def generateTitles():
#     for i in range(1, 11):
#         cell = columns[0] + str(i)
#         ws[cell] = "oof"

# [DONE] GENERATE TITLE AND ALL COLUMN NAMES
def generateTitles():
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


getJobDetails()

# Save the file
wb.save("Job-Tracker.xlsx")


# CONCEPT CODE
# for i in rows:
#     if i == 0 :
#         continue
#     for j in columns:
#         print(j, i)

# documentation Code
# # Rows can also be appended
# ws.append([1, 2, 3])
#
# # Python types will automatically be converted
# # ws['A2'] = datetime.datetime.now()
