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

this_month = today.strftime("%Y.%m")

formad_month = last_month.strftime("%Y.%#m")
print(formad_month)


# アプリ起動
subprocess.run(["mstsc.exe", r"C:\Users\USER06\Desktop\OAシステム.rdp"])
pyautogui.sleep(3)

# 起動後アプリの対象クリック、ログイン処理
pyautogui.click(x=826, y=448)
pyautogui.write("12", interval=0.1)
pyautogui.press("enter", presses=3, interval=0.5)
time.sleep(2)

pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.press("tab", presses=3, interval=0.2)
pyautogui.press("enter")
time.sleep(1)

pyautogui.press("4")
pyautogui.press("F5")
time.sleep(1)

# 端量データをインプット
excel = win32.Dispatch("Excel.Application")
excel.Visible = True
filepath_haryou = r"\\MC10\share\MICHINOK_共有\0.共有書類\対照帳\端量　在庫表\翌営業日製造端量表（新）.xlsm"
wb_haryou = excel.Workbooks.Open(filepath_haryou)
ws_haryou = wb_haryou.Worksheets("FA端量データ")

last_row = ws_haryou.Cells(ws_haryou.Rows.Count, 2).End(-4162).Row # xlUp

fraction_list = []

for row in range(3, last_row + 1):
    code = ws_haryou.Cells(row, 1).Value       # 製品コード
    fraction = ws_haryou.Cells(row, 2).Value   # 端量

    if fraction > 0:
        if (code // 10000) % 10 in [5, 6, 7]:
            try:
                code = int(round(code))
                if code % 10 > 0:
                    code = code - (code % 10)
                fraction_list.append({
                    "code": code,
                    "fraction": round(fraction)
                    })
            except:
                pass
print(fraction_list)

# 4-倉庫在庫 の入力
for output in fraction_list:
    pyautogui.write(str(output["code"]))
    pyautogui.press("enter")
    pyautogui.write(str(output["fraction"]))
    pyautogui.press("enter")
pyautogui.press("F9")
time.sleep(2)
pyautogui.press("enter")
time.sleep(2)
print("--------------------棚卸入力の倉庫在庫が終了しました。")

# 3-原料在庫 の入力
pyautogui.press("3")
pyautogui.press("F5")
time.sleep(1)

filepath_hukuromono = fr"\\MC10\share\MICHINOK_共有\0.共有書類\原料\袋物\{formatted_month}月袋物在庫表.xlsx"
wb_hukuromono = excel.Workbooks.Open(filepath_hukuromono)
ws_hukuromono = wb_hukuromono.worksheets(formad_month)

# 代替パス
#fallback_path = fr"\\MC10\share\MICHINOK_共有\0.共有書類\原料\袋物\{formatted_month}月袋物在庫表.xlsx"

#if not os.path.exists(filepath_hukuromono):  # 今月ファイルがまだ作成されず、先月のファイルのsheetにまだある場合
#    filepath_hukuromono = fallback_path

#    wb_hukuromono = excel.Workbooks.Open(filepath_hukuromono)
#    ws_hukuromono = wb_hukuromono.Sheets(2)

last_row = ws_hukuromono.Cells(ws_hukuromono.Rows.Count, 1).End(-4162).Row

zaiko_list = []

for row_A, row_F, row_D in zip(range(4, last_row + 1, 4),range(7, last_row + 1, 4), range(4, last_row + 1, 4)):     # A列4行目から4行おき、F列7行目から4行おき、D列4行目から4行おき
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

pyautogui.write("1008")     # Fメーズコード
pyautogui.press("enter")
pyautogui.write("2000")     # Fメーズ数量
pyautogui.press("enter")

pyautogui.write("10037")    # ヘイ粉コード
pyautogui.press("enter")

print("棚卸入力の処理を完了しました。ヘイ粉の数量を手入力してください。")