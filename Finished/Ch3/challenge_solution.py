# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Read a single CSV file and split it into multiple worksheets

import csv
import xlsxwriter
from collections import defaultdict


# create a dictionary that will map strings to lists of data that
# represent rows to go with that key name
ws_dict = defaultdict(list)
headers = []

filename = "Inventory.csv"

def read_csv_to_array(filename):
    # define the array that will hold the data
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)

        # read the headers, we will need this for each sheet
        global headers
        headers = next(reader)
        # read the data and distribute it to each key in the dictionary
        for row in reader:
            ws_dict[row[1]].append(row)

# Read the data into the dictionary 
read_csv_to_array(filename)

# create the workbook output
workbook = xlsxwriter.Workbook("Inventory.xlsx")
# create a worksheet for each key in the dictionary
ws_names = ws_dict.keys()
for name in ws_names:
    ws = workbook.add_worksheet(name)
    # write the header row to the sheet
    ws.write_row(0,0,headers)

    # now write all the data for that group to the current sheet
    datalist = ws_dict[name]
    for i, row in enumerate(datalist, start=1):
        ws.write_row(i, 0, row)

    # set the zoom and autofit the columns
    ws.set_zoom(200)
    ws.autofit()

# save the workbook
workbook.close()
