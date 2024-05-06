# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# XlsxWriter formulas and conditional formatting

import csv
import xlsxwriter


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

# create the workbook
workbook = xlsxwriter.Workbook('Conditional.xlsx')
worksheet = workbook.add_worksheet("Inventory")

fmt_bold = workbook.add_format({'bold': True})
fmt_money = workbook.add_format(
    {'font_color': 'green', 'num_format': '$#,##0.00'})
# define the format for the conditional expression
fmt_cond = workbook.add_format({"bg_color": "#AAFFAA", "bold": True})

# write the data into the workbook
worksheet.write_row(0, 0, inventory_data[0], fmt_bold)
# add the new header for the margin
worksheet.write(0, 5, "Margin", fmt_bold)
# add the data to the worksheet
for row, itemlist in enumerate(inventory_data[1:], start=1):
    # worksheet.write_row(row, 0, itemlist)
    worksheet.write(row, 0, itemlist[0])
    worksheet.write(row, 1, itemlist[1], fmt_bold)
    worksheet.write_number(row, 2, int(itemlist[2]))
    worksheet.write_number(row, 3, float(itemlist[3]), fmt_money)
    worksheet.write_number(row, 4, float(itemlist[4]), fmt_money)
    # calculate the row and column for the formula
    worksheet.write_formula(row, 5, f"=E{row+1}-D{row+1}", fmt_money)

# add the conditional formatting
# worksheet.conditional_format(1, 5, len(inventory_data), 5, {
#     "type": "cell",
#     "criteria": ">=",
#     "value": 0.75,
#     "format": fmt_cond
# })
worksheet.conditional_format(1, 0, len(inventory_data), 5,
                             {'type':     'formula',
                              'criteria': '=$F2>=.75',
                              'format':   fmt_cond
                              })

worksheet.set_zoom(150)
worksheet.autofit()

workbook.close()
