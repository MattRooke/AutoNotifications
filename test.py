from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import os
from datetime import datetime as date

os.path.join(
    "%userprofile%/company/company Team Site - Documents/",
    "Backup Check.xlxs",
)


workbook = load_workbook(
    filename=R"C:\Users\technician1\company\company Team Site - Documents\Backup Check.xlsx"
)
ws = workbook.active

for i in ws[8]:
    if i.value == None:
        i.value = date.today().strftime("%d/%m/%Y")
        break

workbook.save(
    filename=R"C:\Users\technician1\company\company Team Site - Documents\Backup Check.xlsx"
)
