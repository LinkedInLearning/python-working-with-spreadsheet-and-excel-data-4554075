# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Writing CSV file from an array

import csv

# Sample data
data = [
    ["Item Name", "Category", "Quantity", "Wholesale Price", "Consumer Price"],
    ["Apple", "Fruits", 100, 0.50, 0.75],
    ["Banana", "Fruits", 150, 0.35, 0.50],
    ["Orange", "Fruits", 120, 0.45, 0.65],
    ["Grapes", "Fruits", 80, 0.60, 0.85],
    ["Strawberries", "Fruits", 90, 1.20, 1.50]
]

# function to write the data


def write_array_to_csv(data, filename):
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)


# Write data to CSV file
write_array_to_csv(data, "output.csv")
