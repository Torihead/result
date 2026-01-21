import time
import win32com.client as win32
import pyautogui
import shutil
from common_utils import get_date_info
from app_automation import RDPApp, ExcelUtils

# 日付情報を取得
dates = get_date_info()
today = dates['today']
formatted_month = dates['formatted_month']
formatted_year = dates['formatted_year']
last_month = dates['last_month']
format_month = dates['format_month']

# RDP接続・ログイン
RDPApp.launch_and_login(sleep_time=5)

# とうもろこし調査票のタブを選択
RDPApp.navigate_tabs(4)
pyautogui.press("tab")
pyautogui.press("right")
pyautogui.press("enter")
time.sleep(0.5)

# OUT出力してアプリ閉じ
pyautogui.write(formatted_month, interval=0.2)
pyautogui.press("F6")
pyautogui.press("enter")
time.sleep(6)
pyautogui.press("F12")
pyautogui.press("e")

# OUTフォルダ内の調査票を、棚卸フォルダに移動
before_path = r"\\MC10\share\OA\EXCEL\OUT\TOMOROKOSI_TYOSA.XLS"
after_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\09　農水"
shutil.copy(before_path, after_path)

# とうもろこし調査票
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
time.sleep(2)

input_file =  r"c:\Users\USER06\工業会調査票（原本）.xlsx"                                        # 工業会調査票(原本)
wb = excel.Workbooks.Open(input_file)
output_filepath = fr"\\192.168.82.21\share\MICHINOK_共有\2.小島\工業会\{format_month}_02070.xls" # 流通状況調査票
wb.SaveAs(output_filepath, FileFormat=56)                                                      # 工業会用の調査票を、新しく名前つけて保存

source_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\09　農水\TOMOROKOSI_TYOSA.XLS"         # とうもろこし調査票
target_path_R = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\09　農水\流通状況調査票\調査票.xlsx"    # 流通状況調査票
target_path_K = output_filepath                                                                                                                     # 工業会調査票

wb_source = excel.Workbooks.Open(source_path)
wb_R = excel.Workbooks.Open(target_path_R)
wb_K = excel.Workbooks.Open(target_path_K)

# 調査票の入力
format_year = last_month.strftime("%y")
format_month = last_month.strftime("%m")

def copy_value(target_sheet, source_sheet, target_row, source_row, col_pairs):  # 「sheet、列、行ペアリスト」指定で転載処理
    for target_col, source_col in col_pairs:
        target_sheet.Cells(target_row, target_col).Value = source_sheet.Cells(source_row, source_col).Value

for sheet in [wb_R.Sheets(1), wb_K.Sheets(1)]:  # 年月の入力
    sheet.Range("B2").Value = format_year
    sheet.Range("D2").Value = format_month

# 生産量の合計 tが切捨てされない場合の処理
wb_genryou = excel.Workbooks.Open(fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{formatted_year}年度\{formatted_month}月\05　原料入出庫表\{formatted_month}_GENRYO_NSK.XLS")
ws_genryou = wb_genryou.Sheets(1)

product_sum_tyosa = int(wb_source.Sheets(1).Range("BY13").Value)                          # 調査票の生産量合計
last_row = wb_genryou.Sheets(1).Cells(wb_genryou.Sheets(1).Rows.Count, 6).End(-4162).Row  # F列最終行の取得
val = wb_genryou.Sheets(1).Range(f"F{last_row}").Value                                    # 原料入出庫表の生産量合計
product_sum_genryou = int(val / 1000)                                                     # 原料の生産量を切捨て処理した値
print(f"調査票の生産量合計: {product_sum_tyosa} , 原料入出庫表の生産量合計: {product_sum_genryou}")

meetfeed_sum = wb_source.Sheets(1).Range("AY13").Value
if (product_sum_tyosa - 1) == product_sum_genryou:
    print(f"生産量の合計に差異があるため、{product_sum_tyosa} → {product_sum_genryou}の値に調整します。")
    meetfeed_sum -= 1
    wb_source.Sheets(1).Range("AY13").Value = meetfeed_sum
else:
    print("生産量の合計に差異はありません。調整は行いません。")

product_sum_tyosa = int(wb_source.Sheets(1).Range("BY13").Value)                     # 数値の再取得
print(f"調整後の肉牛用 合計: {meetfeed_sum}\n調整後の調査票の生産量合計: {product_sum_tyosa}")
wb_genryou.Close(SaveChanges=False)

# sheet1枚目
sheet1_column_paires = [
    (18, 43),   # R列 ← AQ列
    (20, 47),   # T列 ← AU列
    (21, 51),   # U列 ← AY列
    (22, 55),   # V列 ← BC列
    (24, 63),   # X列 ← BK列
    (26, 72),   # Z列 ← BT列
    (14, 27),   # N列 ← AA列
    (15, 32),   # O列 ← AE列
]

# 転載処理
copy_value(wb_R.Sheets(1), wb_source.Sheets(1), 11, 13, sheet1_column_paires)
copy_value(wb_K.Sheets(1), wb_source.Sheets(1), 11, 13, sheet1_column_paires)

copy_value(wb_R.Sheets(1), wb_source.Sheets(1), 12, 17, sheet1_column_paires)
copy_value(wb_K.Sheets(1), wb_source.Sheets(1), 13, 17, sheet1_column_paires)

copy_value(wb_R.Sheets(1), wb_source.Sheets(1), 15, 23, sheet1_column_paires)
copy_value(wb_K.Sheets(1), wb_source.Sheets(1), 16, 23, sheet1_column_paires)

copy_value(wb_R.Sheets(1), wb_source.Sheets(1), 16, 25, sheet1_column_paires)
copy_value(wb_K.Sheets(1), wb_source.Sheets(1), 17, 25, sheet1_column_paires)

copy_value(wb_R.Sheets(1), wb_source.Sheets(1), 17, 27, sheet1_column_paires)
copy_value(wb_K.Sheets(1), wb_source.Sheets(1), 18, 27, sheet1_column_paires)

copy_value(wb_R.Sheets(1), wb_source.Sheets(1), 32, 43, sheet1_column_paires)
#copy_value(wb_K.Sheets(1), wb_source.Sheets(1), 32, 43, sheet1_column_paires)  # 保護されてエラー出るので、下の行で代用
wb_K.Sheets(1).Range("N33").Value = wb_source.Sheets(1).Range("AA43").Value

copy_value(wb_R.Sheets(1), wb_source.Sheets(1), 32, 43, sheet1_column_paires)
#copy_value(wb_K.Sheets(1), wb_source.Sheets(1), 33, 43, sheet1_column_paires)  # 保護されてエラー出るので、下の行で代用
wb_K.Sheets(1).Range("O33").Value = wb_source.Sheets(1).Range("AE43").Value

print("--------------------Sheet1の処理が終了しました。")

# sheet2枚目
sheet2_column_paires = [
    (13, 17),    # M列 ← Q列
    (14, 18),    # N列 ← R列
    (16, 20),    # P列 ← T列
    (18, 22),    # R列 ← V列
    (24, 28),    # X列 ← AB列
]
copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 11, 11, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 11, 11, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 14, 14, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 14, 14, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 19, 19, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 19, 19, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 20, 20, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 20, 20, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 29, 29, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 29, 29, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 32, 32, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 32, 32, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 34, 34, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 34, 34, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 36, 36, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 36, 36, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 34, 34, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 34, 34, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 45, 45, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 45, 45, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 46, 46, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 46, 46, sheet2_column_paires)

copy_value(wb_R.Sheets(2), wb_source.Sheets(2), 48, 48, sheet2_column_paires)
copy_value(wb_K.Sheets(2), wb_source.Sheets(2), 48, 48, sheet2_column_paires)

print("--------------------Sheet2の処理が終了しました。")

# sheet3枚目
sheet3_column_paires = [
    (14, 18),    # N列 ← R列
    (24, 28),    # X列 ← AB列
]
copy_value(wb_R.Sheets(3), wb_source.Sheets(3), 9, 11, sheet3_column_paires)
copy_value(wb_K.Sheets(3), wb_source.Sheets(3), 9, 11, sheet3_column_paires)

print("--------------------Sheet3の処理が終了しました。")

# sheet4枚目
sheet4_column_paires = [
    (15, 38),    # O列 ← AL列
    (16, 43),    # P列 ← AQ列
    (17, 47),    # Q列 ← AU列
    (18, 52),    # R列 ← AZ列
    (20, 60),    # T列 ← BH列
    (22, 69),    # V列 ← BQ列
]
copy_value(wb_R.Sheets(4), wb_source.Sheets(4), 13, 13, sheet4_column_paires)
copy_value(wb_K.Sheets(4), wb_source.Sheets(4), 13, 13, sheet4_column_paires)

copy_value(wb_R.Sheets(4), wb_source.Sheets(4), 14, 14, sheet4_column_paires)
copy_value(wb_K.Sheets(4), wb_source.Sheets(4), 14, 14, sheet4_column_paires)

copy_value(wb_R.Sheets(4), wb_source.Sheets(4), 15, 15, sheet4_column_paires)
copy_value(wb_K.Sheets(4), wb_source.Sheets(4), 15, 15, sheet4_column_paires)

copy_value(wb_R.Sheets(4), wb_source.Sheets(4), 16, 16, sheet4_column_paires)
copy_value(wb_K.Sheets(4), wb_source.Sheets(4), 16, 16, sheet4_column_paires)

print("--------------------Sheet4の処理が終了しました。")

wb_source.Close(SaveChanges=True)
wb_R.Close(SaveChanges=True)
wb_K.Close(SaveChanges=True)
#wb_K.SaveAs(target_path_K, FileFormat=56)

excel.Quit()


print("とうもろこし調査票の全工程を完了しました。")