# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Reading CSV file into an array with a filter function

import csv
import pprint

def read_csv_filter_rows(filename):
  # array to hold the filtered data result
  filtered_data = []

  with open(filename, 'r') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
      filtered_data.append(row)
  return filtered_data

# Filter function (replace with your specific filtering criteria)


# Call the read function with a filter function
filtered_rows = read_csv_filter_rows("Inventory.csv")

# Print filtered data
pprint.pprint(filtered_rows)
