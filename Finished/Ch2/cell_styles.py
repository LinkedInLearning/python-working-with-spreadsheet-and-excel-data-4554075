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
# sheet["A1"].style = "Title"
# sheet["B1"].number_format = opnumstyle.FORMAT_CURRENCY_USD_SIMPLE
# sheet["B1"].style = "Calculation"
# sheet["C1"].number_format= opnumstyle.FORMAT_DATE_DDMMYY
# sheet["C1"].style = "Accent2"

# Create styles using Fonts and Colors
italic_font = Font(italic=True, size=16)
colored_text = Font(name="Courier New", size=20, color="000000FF")
centered_text = Alignment(horizontal="center", vertical="top")
border_side = Side(border_style="mediumDashed")
cell_border = Border(top=border_side, right=border_side,
                     left=border_side, bottom=border_side)

sheet["A1"].font = italic_font
sheet["B1"].font = colored_text
sheet["B1"].alignment = centered_text
sheet["C1"].border = cell_border

sheet.column_dimensions['A'].width = 30
sheet.row_dimensions[1].height = 50

# Save the workbook
wb.save("StyledCells.xlsx")
