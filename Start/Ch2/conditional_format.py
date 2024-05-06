# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Apply conditional formatting to a worksheet 

import openpyxl
from openpyxl.formatting import Rule
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle


filename = "FinancialSample.xlsx"

# Load the workbook
workbook = openpyxl.load_workbook(filename)
sheet = workbook["SalesData"]

# define the style to represent the formatting


# create a rule for the condition


# add the rule to the entire sheet


workbook.save("CondFormat.xlsx")
print("Workbook created successfully!")
