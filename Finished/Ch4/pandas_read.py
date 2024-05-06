# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Using Pandas library to read CSV and Excel data

import pandas as pd


# read a CSV file using read_csv
df = pd.read_csv("Inventory.csv")
print(df)

# read just a portion of the file
df = pd.read_csv("Inventory.csv", skiprows=lambda x: x >= 1 and x < 15, nrows=15)
print(df)

# read an Excel file
df = pd.read_excel("FinancialSample.xlsx", usecols="A:E,H", nrows=15)
print(df)
print(df.dtypes)

# Get information about a workbook
file = pd.ExcelFile('FinancialSample.xlsx')
print("Sheet names are:", file.sheet_names)
