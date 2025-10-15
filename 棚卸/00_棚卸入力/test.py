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
last_month = today - dt.timedelta(days=29)
print(last_month)
formatted_month = last_month.strftime("%Y.%m")
if today.month == 4:
    formatted_year = str(today.year - 1)
else:
    formatted_year = str(today.year)

format_last_month = today.strftime("%Y.%m")
print(format_last_month)
print(formatted_month)

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

# ↓↓ ここから ↓↓
time.sleep(2)


filepath_hukuromono = fr"\\MC10\share\MICHINOK_共有\0.共有書類\原料\袋物\{formatted_month}月袋物在庫表.xlsx"
wb_hukuromono = excel.Workbooks.Open(filepath_hukuromono)
ws_hukuromono = wb_hukuromono.Sheets(2)

# 代替パス
#fallback_path = fr"\\MC10\share\MICHINOK_共有\0.共有書類\原料\袋物\{formatted_month}月袋物在庫表.xlsx"

#if not os.path.exists(filepath_hukuromono):  # 今月ファイルがまだ作成されず、先月のファイルのsheetにまだある場合
#    filepath_hukuromono = fallback_path

#    wb_hukuromono = excel.Workbooks.Open(filepath_hukuromono)
#    ws_hukuromono = wb_hukuromono.Sheets(2)

last_row = ws_hukuromono.Cells(ws_hukuromono.Rows.Count, 1).End(-4162).Row

zaiko_list = []

for row_A, row_F, row_D in zip(range(4, last_row + 1, 4),range(7, last_row + 1, 4), range(4, last_row + 1, 4)):
    code_Value = ws_hukuromono.Range(f"A{row_A}").Value
    quantity_Value = ws_hukuromono.Range(f"F{row_F}").Value
    weight_Value = ws_hukuromono.Range(f"D{row_D}").Value

    zaiko_list.append({
        "code":int(code_Value),
        "quantity":int(quantity_Value),
        "weight":weight_Value
        })
for item in zaiko_list:
    if item["code"] == 3173:                # ヘイ粉の重量 400kg/TB
        item["weight"] = 400.0
    item["weight"] = int(item["weight"] or 0)    # 全て .以下をintで切捨て

for item in zaiko_list:                         # 在庫のキーを追加
    zaiko = item["quantity"] * item["weight"]
    item["zaiko"] = zaiko
print(zaiko_list)

for output in zaiko_list:
    pyautogui.write(str(output["code"]))
    pyautogui.press("enter")
    pyautogui.write(str(output["zaiko"]))
    pyautogui.press("enter")

print("棚卸入力の処理を完了しました。")