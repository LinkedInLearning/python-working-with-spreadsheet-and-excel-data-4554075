# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Solution for CSV Chapter challenge

# import the csv module from the standard library
import csv
from decimal import Decimal


def read_csv_to_array(filename):
    # define the array that will hold the data
    data = []
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            data.append(row)
    return data


def write_array_to_csv(data, filename):
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)


# Read the data into an array of arrays
inventory_data = read_csv_to_array("Inventory.csv")

# Add the new column to the headers
headers = inventory_data[0]
headers.append("Margin")

# calculate margin for each row
datarows = inventory_data[1:]
for row in datarows:
    margin_value = Decimal(row[4]) - Decimal(row[3])
    row.append(margin_value)

# Write data to CSV file
write_array_to_csv(inventory_data, "output.csv")
