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

# アプリ起動
subprocess.run(["mstsc.exe", r"C:\Users\USER06\Desktop\OAシステム.rdp"])
pyautogui.sleep(5)

# 起動後アプリの対象クリック、ログイン処理
pyautogui.click(x=826, y=448)
pyautogui.write("12", interval=0.1)
pyautogui.press("enter", presses=3, interval=0.5)
time.sleep(2)

# 製品入出庫台帳
pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.hotkey("ctrl", "pageup")
pyautogui.press("tab")
pyautogui.press("right")
pyautogui.press("down")
pyautogui.press("enter")
time.sleep(0.5)

lastmonth_start = dt.date(month=today.month - 1, year=today.year, day=1)    # 先月の初日を取得

thismonth_start = dt.date(month=today.month, year=today.year, day=1)        # 今月の初日を取得
lastmonth_end = thismonth_start - dt.timedelta(days=1)                      # 今月1日から、-1日することで先月の最終日を取得

format_lastmonth_start = lastmonth_start.strftime("%Y%m%d")
format_lastmonth_end = lastmonth_end.strftime("%Y%m%d")
pyautogui.press("tab", presses=2, interval=0.2)

pyautogui.write(format_lastmonth_start, interval=0.2)       # 開始日の入力
pyautogui.press("enter")
pyautogui.write(format_lastmonth_end, interval=0.2)         # 終了日の入力
pyautogui.press("enter")
pyautogui.write("1")                    # 1ページ目を選択     # 自倉庫のみ
pyautogui.press("F6")
time.sleep(0.5)
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=780, y=36)            # エクスポートをクリック
time.sleep(0.5)
pyautogui.click(x=364, y=76)            # 保存先をクリック

path_ownhouse = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\10　製品入出庫台帳"
pyperclip.copy(path_ownhouse)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("tab", presses=6, interval=0.2)

name_ownhouse = f"{formatted_month}_製品入出庫台帳.pdf"
pyperclip.copy(name_ownhouse)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter", presses=2, interval=1)
time.sleep(2)
pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
time.sleep(1)
pyautogui.press("F12")
time.sleep(0.5)
pyautogui.press("enter")
time.sleep(0.5)

# 全倉庫の製品入出庫台帳
pyautogui.press("tab", presses=2, interval=0.2)

pyautogui.write(format_lastmonth_start, interval=0.2)       # 開始日の入力
pyautogui.press("enter")
pyautogui.write(format_lastmonth_end, interval=0.2)         # 終了日の入力
pyautogui.press("enter")
pyautogui.press("F6")
time.sleep(0.5)
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=780, y=36)            # エクスポートをクリック
time.sleep(0.5)
pyautogui.click(x=364, y=76)            # 保存先をクリック

path_allhouse = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\10　製品入出庫台帳\全倉庫"
pyperclip.copy(path_allhouse)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("tab", presses=6, interval=0.2)

name_allhouse = f"{formatted_month}_ALL_製品入出庫台帳.pdf"
pyperclip.copy(name_allhouse)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter", presses=2, interval=1)
time.sleep(2)
pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
time.sleep(0.5)
pyautogui.press("F12")
time.sleep(0.5)
print("--------------------製品入出庫台帳の処理が終了しました。")

# 原料入出庫台帳
pyautogui.press("down")
pyautogui.press("enter")
time.sleep(0.5)
pyautogui.write(format_lastmonth_start, interval=0.2)       # 開始日の入力
pyautogui.press("enter")
pyautogui.write(format_lastmonth_end, interval=0.2)         # 終了日の入力
pyautogui.press("F6")
time.sleep(0.5)
pyautogui.press("enter", presses=2, interval=0.5)
time.sleep(4)
pyautogui.click(x=780, y=36)            # エクスポートをクリック
time.sleep(1)
pyautogui.click(x=364, y=76)            # 保存先をクリック

path_genryou = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\11　原料入出庫台帳"
pyperclip.copy(path_genryou)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("tab", presses=6, interval=0.2)

name_genryou = f"{formatted_month}_原料入出庫台帳.pdf"
pyperclip.copy(name_genryou)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter", presses=2, interval=1)
time.sleep(1)
pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
pyautogui.press("F12")

print("--------------------原料入出庫台帳の処理が終了しました。")

# 半製品入出庫表、台帳
path_cyukan_daityou = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\13　半製品入出庫"    # 台帳のパス、ファイル名
name_cyukan_daityou = f"{formatted_month}_半製品入出庫台帳.pdf"
path_cyukan_hyou = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\13　半製品入出庫"       # 表のパス、ファイル名
name_cyukan_hyou = f"{formatted_month}_半製品入出庫表.pdf"

hanseihin_list = [
                {"path" : path_cyukan_daityou, "name" : name_cyukan_daityou},
                {"path" : path_cyukan_hyou, "name" : name_cyukan_hyou}
                 ]

for item in hanseihin_list:
    pyautogui.press("down")
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.write(formatted_month, interval=0.2)       # 年月の入力
    pyautogui.press("F6")

    time.sleep(0.5)
    pyautogui.press("enter", presses=2, interval=0.5)
    time.sleep(4)
    pyautogui.click(x=780, y=36)            # エクスポートをクリック
    time.sleep(1)
    pyautogui.click(x=364, y=76)            # 保存先をクリック

    pyperclip.copy(item["path"])            # リストのパスをコピー
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("tab", presses=6, interval=0.2)

    pyperclip.copy(item["name"])           # リストの名前をコピー
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter", presses=2, interval=1)
    time.sleep(1)
    pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
    pyautogui.press("F12")
    time.sleep(1)
print("--------------------半製品入出庫の処理が終了しました。")

# 配合日報・集計
pyautogui.press("down")
pyautogui.press("enter")
time.sleep(0.5)
pyautogui.write(format_lastmonth_start, interval=0.2)       # 開始日の入力
pyautogui.press("enter")
pyautogui.write(format_lastmonth_end, interval=0.2)         # 終了日の入力
pyautogui.press("enter")
pyautogui.press("space")
pyautogui.press("enter")
pyautogui.press("space")
pyautogui.press("F6")
time.sleep(0.5)
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=780, y=36)            # エクスポートをクリック
time.sleep(1)
pyautogui.click(x=364, y=76)            # 保存先をクリック

path_zisseki = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\14　配合日報・集計"
pyperclip.copy(path_zisseki)            # リストのパスをコピー
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("tab", presses=6, interval=0.2)

name_zisseki = f"{formatted_month}_配合日報.pdf"
pyperclip.copy(name_zisseki)           # リストの名前をコピー
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter", presses=2, interval=1)
time.sleep(1)
pyautogui.click(x=1899, y=14)           # 閉じるボタンをクリック
pyautogui.press("F12")
time.sleep(1)
pyautogui.press("e")

print("--------------------配合実績の処理が終了しました。")

print("台帳～配合日報の全処理を完了しました。")