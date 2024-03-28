import os
import warnings

import openpyxl
from os import listdir
from os.path import join


# Made by Emre KAPLAN


# def assertEqualsCell(pathToFolder, SheetName, Cell, expectedValue):
#     """ Asserts that a cell is equal to the expected """
#     for filename in listdir(pathToFolder):
#         full_path = join(pathToFolder, filename)
#         if os.path.isfile(full_path) and filename.endswith('.xlsx'):
#             wb = openpyxl.load_workbook(full_path, read_only=True, data_only=True)
#             counter = 0
#             ws = wb[SheetName]  # Open up Sheet
#             if ws[Cell].value == expectedValue:
#                 return True

def warn(message, color="red"):
    colors = {
        "red": "\033[91m",
        "yellow": "\033[93m",
        "green": "\033[92m",
        "blue": "\033[94m",
        "purple": "\033[95m",
        "cyan": "\033[96m"
    }
    end_color = "\033[0m"

    if color in colors:
        print(colors[color] + message + end_color)
    else:
        print("Invalid color specified. Defaulting to red.")
        print(colors["red"] + message + end_color)


def assertEqualsCells(pathToFolder, SheetName, CellRange, expectedValues, WhitelistedFormulas):
    valueTestPassed = False
    warning_file = open("Warnings.txt", "w")
    grades_file = open("Grades.txt", "w")
    """ Asserts that a range of cells is equal to the expected tuple """
    # Loop to iterate through files
    for filename in listdir(pathToFolder):
        full_path = join(pathToFolder, filename)
        # Check if the file is an Excel file
        student_number = filename.split(".")[0]
        if os.path.isfile(full_path) and filename.endswith('.xlsx'):
            # First VALUE check
            wb_value_only = openpyxl.load_workbook(full_path, read_only=True, data_only=True)
            ws_value_only = wb_value_only[SheetName]
            target_cells = ws_value_only[CellRange]
            grade = 0
            for value_row in target_cells:
                for cell in value_row:
                    if cell.value in expectedValues:
                        # grade += 1
                        valueTestPassed = True
            # Second FORMULA check
            wb_formula = openpyxl.load_workbook(full_path, read_only=True, data_only=False)
            ws = wb_formula[SheetName]  # Open up Sheet
            target_cells = ws[CellRange]  # Cells to be read
            for row2 in target_cells:
                for cell2 in row2:
                    if valueTestPassed:
                        if cell2.value in WhitelistedFormulas:
                            grade += 1
                        else:
                            if type(cell2.value) is not int:
                                warning_file.write(
                                    "WARNING: Expected Value of student {0} is correct but used formula is not in the "
                                    "whitelist | Cell: {1} | Formula Used: {2}\n".format(
                                        student_number, cell2.coordinate, cell2.value))
                    else:
                        grade -= 1

            grades_file.write("Student {0}'s grade is {1}\n".format(student_number, grade))

path = r"C:\Users\Emre K\Desktop\EMREKAPLAN EXCEL GRADE\ExcelFiles"  # Path to the FOLDER that contains excel files

expected = (46, 47, 197)  # List of expected values MUST BE IN THE SAME ORDER AS
# CELLS

whitelist = ("=SUM(D2:D12)", "=SUM(E2:E12)", "=SUM(F2:F12)")

assertEqualsCells(path, "Sheet1", "D13:F13", expected, whitelist)