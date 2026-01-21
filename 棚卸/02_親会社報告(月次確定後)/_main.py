import os
import datetime as dt
import shutil as sh
import pythoncom
from win32com.client import Dispatch
from common_utils import get_date_info

# 割戻表作成の関数の呼び出し
import 割戻表 as wri
wri.main()

# 親会社報告書の関数を呼び出し
import 親会社報告書 as oya
oya.main()

# 日付情報を取得
dates = get_date_info()
today = dt.datetime.now()
formatted_date = dates['formatted_month']
formatted_month = f"{formatted_date}_"
formatted_year = dates['formatted_year']

os.startfile(fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月")

genryouNSK = [
              fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}SYOYUSYA_GENRYO_NSK.XLS",
              fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_NSK.XLS"
              ]
warimodosi = [
              fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI1.XLS",
              fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI2.XLS",
              fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI3.XLS"
              ]
genryouNSK_pdf = [
                  fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_NSK.pdf"
                  ]
warimodosi_pdf = [
                  fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI1.pdf",
                  fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI2.pdf",
                  fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI3.pdf"
                  ]
houkokusyo_xlsx = [
                fr"\\MC10\share\OA\EXCEL\OUT\NICH_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx",
                fr"\\MC10\share\OA\EXCEL\OUT\YUKI_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx",
                fr"\\MC10\share\OA\EXCEL\OUT\NISS_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx"
                ]

for file in genryouNSK:
    sh.copy(file, fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\05　原料入出庫表")
for file in warimodosi:
    sh.copy(file, fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\07　終了届\2 割戻表")
for file in genryouNSK_pdf:
    sh.copy(file, fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\08　月次帳\1 原料入出庫表")
for file in warimodosi_pdf:
    sh.copy(file, fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\08　月次帳\5 割戻表")
for index, file in enumerate(houkokusyo_xlsx):
    if index == 0:
        sh.copy(file, fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\06　親会社月次報告\1　日和")
    elif index == 1:
        sh.copy(file, fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\06　親会社月次報告\2　雪印")
    elif index == 2:
        sh.copy(file, fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\06　親会社月次報告\3　日清")

print("--------------------ファイルの移動を完了しました。")

# ショートカットを作成して、親会社報告のフォルダに張り付ける。
shortcut_folder = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_date}月\06　親会社月次報告"

def create_shortcut(target_path, shortcut_name, shortcut_dir):
    pythoncom.CoInitialize()
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(os.path.join(shortcut_dir, shortcut_name + ".lnk"))
    shortcut.Targetpath = target_path
    shortcut.WorkingDirectory = os.path.dirname(target_path)
    shortcut.save()

for file in genryouNSK:
    filename = os.path.splitext(os.path.basename(file))[0]
    create_shortcut(file, filename, shortcut_folder)

# 在庫証明の関数の呼び出し
import 在庫証明 as zaiko
zaiko.main()

print("--------------------ショートカットの作成を完了しました。")
print("全ての処理を完了しました。")