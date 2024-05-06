# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Reading CSV file into an array with a filter function

import csv
import pprint


def read_csv_filter_rows(filename, filter_func):
    # array to hold the filtered data result
    filtered_data = []

    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if filter_func(row):
                filtered_data.append(row)
    return filtered_data

# Example filter function (replace with your specific filtering criteria)


def filter_by_category(row, category):
    return row[1] == category


# Example usage (replace "fruits_and_vegetables.csv" with your filename)
filtered_rows = read_csv_filter_rows(
    "Inventory.csv", lambda row: filter_by_category(row, "Fruits"))

# Print filtered data
pprint.pprint(filtered_rows)
