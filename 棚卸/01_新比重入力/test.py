import subprocess
import pyautogui
import time
import datetime as dt
import jpholiday
import pyperclip
import win32com.client as win32
import os

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


excel = win32.Dispatch("Excel.Application")
excel.Visible = True
filepath = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\04　調整\1比重計算表(原料).xls"
wb = excel.Workbooks.Open(filepath)
ws_seihin = wb.Sheets(2)
ws_genryou = wb.Sheets(4)

last_row = ws_genryou.Cells(ws_genryou.Rows.Count, 1).End(-4162).Row # xlUp

code_list = []

for row in range(2, last_row + 1):
    code = ws_genryou.Cells(row, 1).Value       # 原料コード
    specific = ws_genryou.Cells(row, 6).Value   # 新比重

    code_list.append({
        "code" : round(code),
        "specific" : specific
    })
print(code_list)