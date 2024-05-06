# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Manipulating DataFrame content

import pandas as pd


df = pd.read_csv("Inventory.csv")
print(df.shape)

# Create a new column of data
df["Margin"] = df["Consumer Price"] - df["Wholesale Price"]
print(df.shape)
print(df.head())

# Modify a column in-place
df["Category"] = df["Category"].apply(lambda x: x.upper())
print(df)

# rename a column in-place
df.rename(columns={
    "Wholesale Price":"Wholesale",
    "Consumer Price":"Consumer"
}, inplace=True)
print(df.head())

# Drop a column
df.drop("Margin", inplace=True, axis=1)
print(df.head())
