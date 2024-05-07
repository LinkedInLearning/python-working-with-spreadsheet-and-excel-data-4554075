# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# Challenge solution for working with Pandas data


import pandas as pd


# read the original data set
df_sales = pd.read_excel("FinancialSample.xlsx")

headers = ["Product", "Gross Sales", "Profits"]
# calculate the summary data
summary_data = []
for prod_name in df_sales["Product"].unique():
    sales_total = df_sales.loc[df_sales['Product']
                               == prod_name, "Gross Sales"].sum()
    profit_total = df_sales.loc[df_sales['Product']
                                == prod_name, "Profit"].sum()
    summary_data.append([prod_name, sales_total, profit_total])

# create a new DataFrame with the summary data
df_summary = pd.DataFrame(summary_data, columns=headers)

# write the data to the new sheet
with pd.ExcelWriter("FinancialSample.xlsx", engine="openpyxl", mode='a') as xlw:
    df_summary.to_excel(xlw, sheet_name="Summary", index=False)
