import subprocess
import pyautogui
import time
import datetime as dt
import jpholiday
import win32com.client as win32

today = dt.date.today()
year = today.year
month = today.month
tenth = dt.date(year, month, 10)

# 10日が土日祝日なら、前営業日に変更
while tenth.weekday() >= 5   or jpholiday.is_holiday(tenth):
    tenth -= dt.timedelta(days=1)
    
# 先月の10日
last_month = today - dt.timedelta(days=20)
print(last_month)
formatted_month = last_month.strftime("%Y.%m")
if today.month <= 4:
    formatted_year = str(today.year - 1)
else:
    formatted_year = str(today.year)


# アプリ起動
subprocess.run(["mstsc.exe", r"C:\Users\USER06\Desktop\OAシステム.rdp"])
pyautogui.sleep(5)

# 起動後アプリの対象クリック、ログイン処理
pyautogui.click(x=826, y=448)
pyautogui.write("12", interval=0.1)
pyautogui.press("enter", presses=3, interval=0.5)
time.sleep(2)

# 原料比重
pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(2)

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
filepath = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\04　調整\1比重計算表(原料).xls"
wb = excel.Workbooks.Open(filepath)
ws_seihin = wb.Sheets(2)
ws_genryou = wb.Sheets(4)

last_row = ws_genryou.Cells(ws_genryou.Rows.Count, 1).End(-4162).Row # xlUp

genhizyu_list = []

for row in range(2, last_row + 1):
    code = ws_genryou.Cells(row, 1).Value       # 原料コード
    specific = ws_genryou.Cells(row, 6).Value   # 新比重

    genhizyu_list.append({
        "code" : round(code),
        "specific" : specific
    })
print(genhizyu_list)

for output in genhizyu_list:                        # 新比重を入力
    pyautogui.write(str(output["code"]))
    pyautogui.press("F5")
    pyautogui.press("tab", presses=3, interval=0.1)
    pyautogui.write(str(output["specific"]))
    pyautogui.press("F9")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)

pyautogui.press("F12")
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(0.5)

#製品比重
last_row = ws_seihin.Cells(ws_genryou.Rows.Count, 1).End(-4162).Row # xlUp

seihizyu_list = []

for row in range(2, last_row + 1):
    code = ws_seihin.Cells(row, 1).Value       # 原料コード
    specific = ws_seihin.Cells(row, 6).Value   # 新比重

    seihizyu_list.append({
        "code" : round(code),
        "specific" : specific
    })
print(seihizyu_list)

for output in seihizyu_list:                        # 新比重を入力
    pyautogui.write(str(output["code"]))
    pyautogui.press("F5")
    pyautogui.press("tab", presses=4, interval=0.1)
    pyautogui.write(str(output["specific"]))
    pyautogui.press("F9")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)

pyautogui.press("F12")
