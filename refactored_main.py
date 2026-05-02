from modules.jpy import main
from modules.data import dataclass
from openpyxl import Workbook, load_workbook
from jobtracker import getJobDetails, searchJobDetail, updateJobDetail, checkWorksheet, mainMenu
import os

@main
def run():
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

