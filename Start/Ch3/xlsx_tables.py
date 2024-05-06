# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# XlsxWriter Excel Tables


import xlsxwriter

# Sample data
data = [
    ["Item Name", "Category", "Quantity", "Wholesale Price", "Consumer Price"],
    ["Apple", "Fruits", 100, 0.50, 0.75],
    ["Banana", "Fruits", 150, 0.35, 0.50],
    ["Orange", "Fruits", 120, 0.45, 0.65],
    ["Grapes", "Fruits", 80, 0.60, 0.85],
    ["Strawberries", "Fruits", 90, 1.20, 1.50]
]

# create the workbook
workbook = xlsxwriter.Workbook('Tables.xlsx')
worksheet = workbook.add_worksheet("Inventory")

fmt_bold = workbook.add_format({'bold': True})
fmt_money = workbook.add_format(
    {'font_color': 'green', 'num_format': '$#,##0.00'})

# write the data into the workbook
worksheet.write_row(0, 0, data[0], fmt_bold)
for row, itemlist in enumerate(data[1:], start=1):
    # worksheet.write_row(row, 0, itemlist)
    worksheet.write(row, 0, itemlist[0])
    worksheet.write(row, 1, itemlist[1], fmt_bold)
    worksheet.write(row, 2, itemlist[2])
    worksheet.write(row, 3, itemlist[3], fmt_money)
    worksheet.write(row, 4, itemlist[4], fmt_money)

# define a table for the worksheet


worksheet.set_zoom(200)
worksheet.autofit()

workbook.close()
