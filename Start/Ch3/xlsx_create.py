# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# XlsxWriter basic operations


import xlsxwriter
import datetime

# create the workbook and add a worksheet
workbook = xlsxwriter.Workbook("XlsxBasics.xlsx")
worksheet = workbook.add_worksheet("Test Sheet")

# Use Letter/Row notation
worksheet.write("A1", "Hello world")
# Use Row,Col notation
worksheet.write(1, 0, "Hello world")

# There are specific write() functions for different data types
worksheet.write_number(2, 0, 12345)
worksheet.write_boolean(3, 0, True)
worksheet.write_url(4, 0, 'https://www.python.org/')

# Write a datetime
date_time = datetime.datetime.strptime('2030-07-28', '%Y-%m-%d')
date_format = workbook.add_format({'num_format': 'd mmmm yyyy'})
worksheet.write_datetime(5, 0, date_time, date_format)

# write multiple values into rows and columns
values = ["Good", "Morning", "Excel"]
worksheet.write_row("A6", values)
worksheet.write_column("D1", values)

# set the zoom on the sheet
worksheet.set_zoom(200)

# save the workbook
workbook.close()
