import openpyxl
from openpyxl.styles import Font

# see examples on https://linuxhint.com/manipulating_excel_python/

# Initializing a Workbook
work_book = openpyxl.load_workbook('search_competitor_ads__2021_02_18 -4.xlsx')

print(work_book.sheetnames)

# Navigating to Second Sheet (at index 1)
work_book.active = 1
# Getting Active Sheet
sheet = work_book.active
print(sheet.title)
# Adding Data to ‘A1’ Cell of Second Sheet
#sheet['A1'] = 'ID'
print(sheet['A1'].value)

# Creating style object
style = Font(name='Consolas', size=13, bold=True, italic=False)
a1 = sheet['A1']
a1.font = style
# saving workbook as ‘example.xlsx’

work_book.save('example.xlsx')