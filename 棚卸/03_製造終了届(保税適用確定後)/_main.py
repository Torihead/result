import pyautogui
import time
import pyperclip
import os
import win32com.client as win32com
from common_utils import get_date_info
from app_automation import RDPApp, ExcelUtils

# 日付情報を取得
dates = get_date_info()
last_month = dates['last_month']
formatted_month = dates['formatted_month']
formatted_year = dates['formatted_year']
tenth = dates['tenth']

# 手差しの印刷用紙をセットするよう促す
pyautogui.alert("製造終了届の処理を開始します。\n最初に、プリンターの手差しトレイにA4用紙をセットしてください。\nセット後、OKを押してください。")

# RDP接続・ログイン
RDPApp.launch_and_login(sleep_time=3)

# 終了届をOUT出力
RDPApp.navigate_tabs(5)
pyautogui.press("tab")
pyautogui.press("right", presses=2, interval=0.2)
pyautogui.press("enter")  # 製造終了届を選択
time.sleep(2)
pyautogui.write(last_month.strftime('%Y%m'), interval=0.1)   # 年月を入力
pyautogui.press("enter")
pyautogui.write(tenth.strftime("%Y%m%d"), interval=0.1) # 年月日を入力
pyautogui.press("F6")
pyautogui.press("enter")       # 終了届を出力
time.sleep(3)
pyautogui.press("F12")
time.sleep(1)

# 法定台帳を法定台帳フォルダにpdf保存
pyautogui.press("down")
pyautogui.press("enter")  # 法定台帳を選択
time.sleep(2)
pyautogui.write(last_month.strftime('%Y%m'), interval=0.1)      # 年月を入力
pyautogui.press("enter")

items = [
    {"code": "1000", "label": "MEIZU", "folder": "1000_メイズ"},
    {"code": "1040", "label": "TOUMITU", "folder": "1040_糖蜜"},
    {"code": "2000", "label": "OOMUGI", "folder": "2000_大麦"}
    ]

for item in items:
    pyautogui.press('backspace', presses=4, interval=0.3)       # 入力されている原料コードを消す
    pyautogui.write(item["code"], interval=0.1)                 # 原料コード
    pyautogui.press("F6")
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.click(x=781, y=37)                                # エクスポートをクリック
    time.sleep(2)
    pyautogui.click(x=364, y=76)                                # 保存先をクリック

    path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\07　終了届\1 法定台帳\{item['folder']}"
    pyperclip.copy(path)
    pyautogui.hotkey("ctrl", "v")                               # item["folder"]のフォルダーパスを張り付ける

    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("tab", presses=6, interval=0.2)

    name = fr"法定台帳_{item['label']}_{formatted_month}"        # item['label']の名前で、pdfファイル保存
    pyperclip.copy(name)
    
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter", presses=2, interval=1)
    time.sleep(2)
    pyautogui.click(x=1899, y=14)   # 閉じるボタン

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

# 終了届4の編集
wb_s4 = excel.Workbooks.Open(filepath_list[3])
ws_last = wb_s4.Sheets(wb_s4.Sheets.Count)      # 最後のsheetの選択

ws_last.Range("A58").Value = "備考に＊印のものについては、下記混用承認申請書に基づくものです。"
ws_last.Range("B59").Value = "混用承認申請書"

output_path = r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO4.pdf"
wb_s4.ExportAsFixedFormat(          # 終了届4をpdf
        Type=0,                     # xlTypePDF
        Filename=output_path,
        Quality=0,                  # xlQualityStandard
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=True
        )

wb_s4.Close(SaveChanges=True)
excel.Quit()

# 終了届4.pdf 両面印刷
pdf_path = r"\\MC10\share\OA\EXCEL\OUT\SEZOSYURYO4.pdf"
os.startfile(pdf_path)
time.sleep(3)

pyautogui.hotkey("alt", "space")    # ウィンドウの最大化
time.sleep(0.5)
pyautogui.press("x")
time.sleep(0.8)

pyautogui.hotkey("ctrl", "p")       # 印刷を選択
time.sleep(1)
pyautogui.press("tab")              # 印刷設定を開く
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("tab", presses=10, interval=0.1) # 両面印刷を選択
pyautogui.press("down") 
pyautogui.hotkey("ctrl", "pageup")               # 給紙タグを選択
pyautogui.hotkey("ctrl", "pageup")
pyautogui.press("tab", presses=5, interval=0.1)  # 給紙トレイの選択
pyautogui.press("up")
pyautogui.press("enter")                         # 設定画面閉じ
time.sleep(1)
pyautogui.press("tab")
pyautogui.hotkey("shift", "tab")
pyautogui.hotkey("shift", "tab")
pyautogui.press("enter")                         # 印刷開始

time.sleep(5)
pyautogui.hotkey("alt", "f4")                    # アプリ終了

# ファイルを棚卸フォルダに移動
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

print("製造終了届の処理が完了しました。")