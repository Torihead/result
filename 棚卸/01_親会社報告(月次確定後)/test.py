import win32com.client as w32
import datetime as dt


excel = w32.Dispatch("Excel.Application")
excel.Visible = True

today = dt.datetime.now()                               # 本日の日付を取得
last_month = today - dt.timedelta(days=10)              # 先月
formatted_date = last_month.strftime("%Y.%m")           # 現在の日付をYYYY.MM形式で取得
formatted_month = f"{formatted_date}_"                  # "YYYY.MM_"に変換
if today.month == 4:
    formatted_year = str(today.year - 1)
else:
    formatted_year = str(today.year)

wb_3 = excel.Workbooks.Open(fr"\\MC10\share\OA\EXCEL\OUT\NISS_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx")
ws_3 = wb_3.Worksheets("Sheet1")
for i in range(120, 3, -1):
    if ws_3.Cells(3, i).Value != "日清丸紅飼料(株)":
        ws_3.Columns(i).Delete()
