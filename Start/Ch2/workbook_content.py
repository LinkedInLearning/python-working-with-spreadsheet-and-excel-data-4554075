# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Manipulate workbook content 

import openpyxl
from openpyxl.comments import Comment
from collections import defaultdict


# Create a new workbook
filename = "FinancialSample.xlsx"

# Load the workbook
wb = openpyxl.load_workbook(filename)

# Get the active worksheet
sheet = wb.active

# Get entire column or row of cells


# Get a range of cells


# iterate over rows and columns


# create a cell with a comment in it


# save the workbook
