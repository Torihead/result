import subprocess
import pyautogui
import time
import datetime as dt
import jpholiday

today = dt.date.today()
year = today.year
month = today.month
tenth = dt.date(year, month, 10)

# 10日が土日祝日なら、前営業日に変更
while tenth.weekday() >= 5 or jpholiday.is_holiday(tenth):
    tenth -= dt.timedelta(days=1)

# 先月の10日
last_month = today - dt.timedelta(days=20)
print(last_month)
formatted_month = last_month.strftime("%Y.%m")
if today.month == 4:
    formatted_year = str(today.year - 1)
else:
    formatted_year = str(today.year)

# 検尺表 を作成してエクスポート
import win32com.client as win32
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
filepath = r"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\検尺値入力表.xls"
wb = excel.Workbooks.Open(filepath)

wb.Sheets(7).Select()                           # シートのグループ化解除
format_month = last_month.strftime("%Y%m")
wb.Sheets(7).Range("L1").Value = format_month   # 更新のための処理
wb.RefreshAll()
time.sleep(1)

export_list = ["主原料、中間製品タンク", "配合タンク", "製品タンク", "袋物原料", "端量"]
sheet_object = [wb.Sheets(sheet_name) for sheet_name in export_list]
wb.Sheets(export_list).Select()

output_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\08　月次帳\3 検尺表\{formatted_month}_検尺値入力表.pdf"
excel.ActiveSheet.ExportAsFixedFormat(0, output_path)

wb.Close(SaveChanges=True)
excel.Quit()