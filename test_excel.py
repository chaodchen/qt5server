import xlwings as xw


app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = True

wb = app.books.open("test.xlsx")

sheet = wb.sheets[0]
sheet.range('A1').value  = "hahah"
sheet.range('A1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter 

