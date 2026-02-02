import subprocess
import pyautogui
import time
import datetime as dt
import jpholiday
import pyperclip
import os
import win32com.client as win32com

today = dt.date.today()
year = today.year
month = today.month
tenth = dt.date(year, month, 10)

# 10日が土日祝日なら、前営業日に変更
while tenth.weekday() >= 5 or jpholiday.is_holiday(tenth):
    tenth -= dt.timedelta(days=1)

# 先月の10日
last_month = today - dt.timedelta(days=15)
formatted_month = last_month.strftime("%Y.%m")
if today.month == 4:
    formatted_year = str(today.year - 1)
else:
    formatted_year = str(today.year)


excel = win32com.Dispatch("Excel.Application")
excel.Visible = True
filepath_list = [
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO1.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO2.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO3.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO4.XLS"
    ]

import shutil

before_path = [
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO1.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO2.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO3.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO4.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO4.pdf"
    ]

output_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\07　終了届\3 製造終了届"

for file in before_path:
    shutil.copy(file, output_path)