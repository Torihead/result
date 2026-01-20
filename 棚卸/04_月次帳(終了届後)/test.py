import subprocess
import pyautogui
import time
import datetime as dt
import jpholiday
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

import os
import shutil
before_path = r"\\MC10\share\OA\EXCEL\OUT\TANAOROSI_HYO.XLS"
output_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\08　月次帳\4 棚卸表"
shutil.copy(before_path, output_path)
os.rename(fr"{output_path}\TANAOROSI_HYO.XLS", fr"{output_path}\{formatted_year}_TANAOROSI_HYO.XLS")