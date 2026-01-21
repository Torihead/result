import pyautogui
import time
import win32com.client as win32
from common_utils import get_date_info
from app_automation import RDPApp, ExcelUtils

# 日付情報を取得
dates = get_date_info()
formatted_month = dates['formatted_month']
formatted_year = dates['formatted_year']

# RDP接続・ログイン
RDPApp.launch_and_login(sleep_time=5)

# 原料比重メニューへ
RDPApp.navigate_tabs(2)
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(2)

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
filepath = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\04　調整\1比重計算表(原料).xls"
wb = excel.Workbooks.Open(filepath)
ws_seihin = wb.Sheets(2)
ws_genryou = wb.Sheets(4)

last_row = ExcelUtils.get_lastrow(ws_genryou)

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
last_row = ExcelUtils.get_lastrow(ws_seihin)

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
