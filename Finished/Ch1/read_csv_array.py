# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Reading CSV file into an array

# import the csv module from the standard library
import csv


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

# Each row in the array is itself an array of values
print(f"Items: {len(inventory_data)}")
print(inventory_data[0])  # This will print the first row (header)
print(inventory_data[1])  # This will print the first row of data
# This will print the name of the item and quantity
print(inventory_data[1][0], inventory_data[1][2])
