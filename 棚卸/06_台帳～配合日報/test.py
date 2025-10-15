import subprocess
import pyautogui
import time
import datetime as dt
import jpholiday
import pyperclip

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


time.sleep(2)
pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
time.sleep(2)

# 全倉庫の製品入出庫台帳
pyautogui.press("tab", presses=6, interval=0.2)
pyautogui.press("0")
pyautogui.press("F6")
time.sleep(0.5)
pyautogui.press("enter")
time.sleep(4)