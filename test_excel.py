import xlwings as xw


app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = True

wb = app.books.open("test.xlsx")

sheet = wb.sheets[0]
sheet.range('A1').value  = "hahah"
sheet.range('A2').value  = "jja"
sheet.range('B1').value  = "kjaskd"
sheet.range('B2').value  = "kjaskd"

print(sheet.range('A1:B2').value)

# sheet.range('A1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter 

