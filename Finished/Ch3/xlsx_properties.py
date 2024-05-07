# LinkedIn Learning Course
# Example file for Python: Working with Excel and Spreadsheet Data by Joe Marini
# XlsxWriter document properties

import xlsxwriter

workbook = xlsxwriter.Workbook("Properties.xlsx")
worksheet = workbook.add_worksheet()

props = {
    "title": "Document Properties Example",
    "subject": "Shows how to use document properties in XlsxWriter",
    "author": "Joe Marini",
    "company": "LinkedIn Learning",
    "manager": "Dr. Heinz Doofenshmirtz",
    "category": "Example spreadsheets",
    "keywords": "Properties, Sample, XlsxWriter",
    "comments": "Created using XlsxWriter as a LinkedIn Learning Example",
}

# set the standard properties
workbook.set_properties(props)

# set some custom properties
workbook.set_custom_property("Checked by", "Perry P")
workbook.set_custom_property("Approved", True)

workbook.close()
