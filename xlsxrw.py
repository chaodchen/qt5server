import xlwings as xw
import time
wb = xw.Book("example.xlsx")
sheet = wb.sheets("Sheet1")


print(wb.fullname)
print(sheet.name)

sheet.range('B1').value = "Niha中文o"