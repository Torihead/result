import time
import win32com.client as w32
import csv
from datetime import datetime

# 処理と印刷を行う関数
def process_and_print(date_str):
    try:
        file_path = r"\\MC10\share\MICHINOK_共有\0.共有書類\対照帳\端量　在庫表\翌営業日製造端量表（新）.xlsm"

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False

        wb = excel.Workbooks.Open(file_path)
        ws_TB = wb.Worksheets("翌日製造使用端量表")
        ws_PB = wb.Worksheets("翌営業日PB製造出荷予定表")

        formatdate = date_str.strftime('%m月%d日')
        ws_TB.Range("H2").Value = formatdate       # 日付を入力
        time.sleep(2)

        wb.RefreshAll()                           # 全体更新
        time.sleep(6)

        ws_TB.PrintOut()                          # 印刷
        ws_PB.PrintOut()

        wb.Close(SaveChanges=True)
        excel.Quit()
        time.sleep(2)

    except Exception as e:
        print(f"エラーが発生しました: {e}")

# CSVファイルから、次の営業日を取得
def main():
    csv_file_path = r"C:\Users\USER06\Desktop\Auto_code\16時印刷\出勤日カレンダー.csv"
    try:
        today = datetime.now().date()
        next_working = None

        with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row['working'] == '1':
                    d = datetime.strptime(row['date'], '%Y/%m/%d').date()

                    # 本日より後の日付をリスト追加
                    if d > today and (next_working is None or d < next_working):
                        next_working = d

        if next_working:
            print(f"次の営業日: {next_working.strftime('%Y/%m/%d')}")
            process_and_print(next_working)
        else:
            print("次の営業日が見つかりません。")


    except Exception as e:
        print(f"エラーが発生しました : {e}")

if __name__ == "__main__":
    main()