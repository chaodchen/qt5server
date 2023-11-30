import xlwings as xw
import time

app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = True

wb = app.books.open("test.xlsx")
sheets = []

sheets.append(wb.sheets.add())
sheets.append(wb.sheets.add())
sheets.append(wb.sheets.add())
sheets.append(wb.sheets.add())
sheets.append(wb.sheets.add())
sheets.append(wb.sheets.add())
wb.sheets.add().activate
time.sleep(1)
sheets[-1].range('A1').value  = "hahah"
sheets[-1].range('A2').value  = "jja"
sheets[-1].range('B1').value  = "kjaskd"
sheets[-1].range('B2').value  = "adasd"
# print(sheet.range('A1:B2').value)
# 获取所有的工作表

# 遍历所有工作表并删除
for i, sheet in enumerate(sheets):
    time.sleep(1)
    if sheet is not None:
        sheet.delete()

