import os
import datetime as dt
import shutil as sh
import win32com.client as w32

print("--------------------処理を開始します。")

import 証明依頼書
証明依頼書.create_証明依頼書()

before_path_list = [
    r"\\MC10\share\OA\EXCEL\OUT\GMG_UKHRI_DAI.xlsx",
    r"\\MC10\share\OA\EXCEL\OUT\KAKO_HOKOKU1.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\KAKO_HOKOKU4.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\2021.09.29_3_証明依頼書.docx"
    ]

contract = (input("契約日を入力してください(例 2025/4/1): "))
try:
    dt_obj = dt.datetime.strptime(contract, "%Y/%m/%d") # 日付の形式をYYYY/m/dに変換
    formatted_date = dt_obj.strftime("%Y.%m.%d")        # 変換後の形式を保存
except ValueError:
    print("日付の形式が正しくありません。YYYY/m/dの形式で入力してください。")
    exit()

while True:
    partner = input("取引先の番号を入力してください\n \
                  1 工業会, 2 全畜連, 3 丸紅, 4 全農, 5 三井物産 : ")
    try:
        partner = int(partner)  # 入力を整数に変換
        if partner in [1, 2, 3, 4, 5]:
            break
        else:
            print("1-5の数字を入力してください。")
    except ValueError:
        print("無効な入力です。1-5の数字を入力してください。")

filename = f"{formatted_date}_{partner}_"               # ファイル名のフォーマット

after_path_list = [
    fr"\\MC10\share\OA\EXCEL\OUT\{filename}GMG_UKHRI_DAI.xlsx",
    fr"\\MC10\share\OA\EXCEL\OUT\{filename}KAKO_HOKOKU1.XLS",
    fr"\\MC10\share\OA\EXCEL\OUT\{filename}KAKO_HOKOKU4.XLS",
    fr"\\MC10\share\OA\EXCEL\OUT\{filename}証明依頼書.docx"
    ]

for before, after in zip(before_path_list, after_path_list):    # before_path_listとafter_path_listをペア(zip)にして繰り返し処理
    if os.path.exists(before):                                  
        os.rename(before, after)                                # ファイル名を変更
        print(f"{before} を {after} に変更しました。")
    else:
        print(f"{before} が見つかりません。")

os.startfile(r"\\MC10\share\農政_電磁記録帳票")               # 電磁記録帳票フォルダを開く
for file in after_path_list:                                # 作成した完了報告書一式を電磁記録フォルダにコピー
    sh.copy(file, r"\\MC10\share\農政_電磁記録帳票")

print("--------------------報告書の作成を終了しました。")
print("すべての処理が完了しました。")