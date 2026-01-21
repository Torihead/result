import pyautogui
import time
import os
import shutil
import win32com.client as win32
import pyperclip
from common_utils import get_date_info
from app_automation import RDPApp, ExcelUtils

# 日付情報を取得
dates = get_date_info()
today = dates['today']
formatted_month = dates['formatted_month']
formatted_year = dates['formatted_year']
last_month = dates['last_month']

# RDP接続・ログイン
RDPApp.launch_and_login(sleep_time=5)

# 棚卸チェックリスト
RDPApp.navigate_tabs(4)
RDPApp.press_multiple_tabs(2)
pyautogui.press("enter")
time.sleep(1)
pyautogui.write(f"{formatted_month}", interval=0.2)
pyautogui.press("tab")

items = [
        {"code": "0", "label": "棚卸チェックリスト(タンク)"},
        {"code": "3", "label": "棚卸チェックリスト(倉庫原料)"},
        {"code": "4", "label": "棚卸チェックリスト(倉庫製品)"}
        ]

for item in items:
    pyautogui.write(item["code"])
    pyautogui.press("F6")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.click(x=780, y=36)            # エクスポートをクリック
    time.sleep(0.5)
    pyautogui.click(x=364, y=76)            # 保存先をクリック

    path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\08　月次帳\2 棚卸チェックリスト"
    pyperclip.copy(path)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("tab", presses=6, interval=0.2)

    name = f"{formatted_month}_{item['label']}"
    pyperclip.copy(name)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter", presses=2, interval=1)
    time.sleep(1)
    pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
    time.sleep(1)

pyautogui.press("F12")      # 棚卸チェックリスト画面終了
print("--------------------棚卸チェックリストの処理が終了しました。")

# 棚卸表のOUT出力
time.sleep(1)
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(0.5)
pyautogui.write(formatted_month)
pyautogui.press("F6")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("F12")

# 製品入出庫表のOUT出力
pyautogui.press("tab", presses=2, interval=0.2)
pyautogui.press("enter")
time.sleep(0.5)
pyautogui.press("tab")
pyautogui.write(formatted_month)
pyautogui.press("tab")
pyautogui.write("1")
pyautogui.press("F6")
pyautogui.press("enter")
time.sleep(4)

pyautogui.click(x=780, y=36)            # エクスポートをクリック
time.sleep(0.5)
pyautogui.click(x=364, y=76)            # 保存先をクリック
path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\08　月次帳\6 製品入出庫表"
pyperclip.copy(path)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("tab", presses=6, interval=0.2)

name = f"{formatted_month}_製品入出庫表"
pyperclip.copy(name)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter", presses=2, interval=1)
time.sleep(2)
pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
pyautogui.press("F12")
pyautogui.press("e")

print("--------------------製品入出庫表の処理が終了しました。")

# 棚卸表のフォルダ移動
before_path = r"\\MC10\share\OA\EXCEL\OUT\TANAOROSI_HYO.XLS"
output_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\08　月次帳\4 棚卸表"
shutil.copy(before_path, output_path)
os.rename(fr"{output_path}\TANAOROSI_HYO.XLS", fr"{output_path}\{formatted_month}_TANAOROSI_HYO.XLS")

# 棚卸表のpdf変換
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
filepath = fr"{output_path}\{formatted_month}_TANAOROSI_HYO.XLS"
wb = excel.Workbooks.Open(filepath)
wb.Sheets(1).Select()
output_pdf_path = fr"{output_path}\{formatted_month}_TANAOROSI_HYO.pdf"
excel.ActiveSheet.ExportAsFixedFormat(0, output_pdf_path)
wb.Close(SaveChanges=True)

print("--------------------棚卸表の処理が終了しました。")

# 検尺表 を作成してエクスポート
filepath = r"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\検尺値入力表.xls"
wb = excel.Workbooks.Open(filepath)

wb.Sheets(7).Select()                           # シートのグループ化解除

format_month = last_month.strftime("%Y%m")      # "YYYYMM"形式
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

print("--------------------検尺値入力表の処理が終了しました。")

print("月次帳の処理が完了しました。")