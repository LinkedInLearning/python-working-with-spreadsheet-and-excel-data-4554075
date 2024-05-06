# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Add column filters to a sheet

import csv
from openpyxl import Workbook


def read_csv_to_array(filename):
    # define the array that will hold the data
    data = []
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            data.append(row)
    return data


# Read the data into an array of arrays
inventory_data = read_csv_to_array("Inventory.csv")

# Create a new workbook
wb = Workbook()

# Get the active worksheet and name it "TestSheet"
sheet = wb.active
sheet.title = "Inventory"

for row in inventory_data:
    sheet.append(row)

# Add the filters to the columns
filters = sheet.auto_filter
filters.ref = sheet.dimensions

wb.save("Inventory.xlsx")
