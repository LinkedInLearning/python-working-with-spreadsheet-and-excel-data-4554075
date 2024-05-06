# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Manipulate cell content and styling 

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import openpyxl.styles.numbers as opnumstyle
import datetime


# Create a new workbook
wb = Workbook()

# Get the active worksheet and name it "TestSheet"
sheet = wb.active
sheet.title = "First"

# Add some data to the new sheet
sheet["A1"] = "Test Data"
sheet["B1"] = 123.4567
sheet["C1"] = datetime.datetime(2030, 4, 1)

# Inspect the default styles of each cell
print(sheet["A1"].style)
print(sheet["B1"].number_format)
print(sheet["C1"].number_format)

# Use some built-in styles


# Create styles using Fonts and Colors


# Save the workbook
wb.save("StyledCells.xlsx")
