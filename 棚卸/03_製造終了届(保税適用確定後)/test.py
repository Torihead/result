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

excel = win32com.Dispatch("Excel.Application")
excel.Visible = True
filepath_list = [
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO1.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO2.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO3.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO4.XLS"
    ]

# 終了届1-3を印刷
def print_excel(filepath, fit_to_one_page=False):   # 印刷の関数　(=False は基本的には設定しないの意味)
    wb = excel.Workbooks.Open(filepath)
    if fit_to_one_page:                             # 引数fit_to_one_pageに、リスト[1]の終了届２を入れると条件分岐
        for sheet in wb.Sheets:                     # シートを１ページに集約する
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
    wb.PrintOut()
    wb.Close(SaveChanges=False)

for i, file in enumerate(filepath_list[0:3]):       # 終了届1-3までを印刷
    print_excel(file, fit_to_one_page=(i == 1))     # (i == 1)繰り返し２回目を指定


excel.Quit()