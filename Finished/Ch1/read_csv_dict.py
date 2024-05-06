# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Reading CSV file into an dictionary

import csv
import pprint


def read_csv_to_dict(filename):
    data = {}
    with open(filename, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            # Add row data to dictionary with header as key
            data[row[reader.fieldnames[0]]] = row
    return data


# Example usage
inventory_data = read_csv_to_dict("Inventory.csv")

# Accessing data
pprint.pprint(inventory_data)
# This will print the dictionary for the "Apple" row
pprint.pprint(inventory_data["Apple"])
# This will print the price of Apple
pprint.pprint(inventory_data["Apple"]["Consumer Price"])
