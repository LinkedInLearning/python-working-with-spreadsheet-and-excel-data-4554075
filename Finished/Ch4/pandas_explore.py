# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Exploring DataFrame data and structure

import pandas as pd
import locale


df = pd.read_excel("FinancialSample.xlsx")

# Set the display options for Pandas
pd.set_option("display.max.columns", None)
pd.set_option("display.precision", 2)

# Examine the structure of the data set
print(df.shape)
print(df.describe())

# View a subset of the data
print(df.head(5))
print(df.tail(5))

# Get the number of values for a given column
print(df["Product"].value_counts())

# Get the unique data values for a column
print(df["Product"].unique())

# Get the min and max values of a column
print(df["Profit"].max())
print(df["Profit"].min())

# use the loc() function to conditionally sum a column
print(df["Profit"].sum())
print(df.loc[df['Product'] == "Carretera", "Profit"].sum())

# Get the sum of the Profit column for each of the Product types
# and format the output as currency in the user's locale
locale.setlocale(locale.LC_ALL, '')
for prod_name in df["Product"].unique():
    total = df.loc[df['Product'] == prod_name, "Profit"].sum()
    print(f"Profits for {prod_name}:", locale.currency(total, grouping=True))
