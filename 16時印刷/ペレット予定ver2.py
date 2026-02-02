import win32com.client as w32
import time
import csv
from datetime import datetime

# ペレット予定を印刷する関数
def process_and_print(data_str):
    try:
        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False

        date_obj = datetime.strptime(data_str, "%Y/%m/%d")

        manth = date_obj.strftime('%Y.%m')

        file_path = fr"\\MC10\share\MICHINOK_共有\0.共有書類\ペレット予測\{manth} みちのくペレット在庫.xlsm"
        workbook = excel.Workbooks.Open(file_path)
        time.sleep(1)

        worksheet = workbook.Worksheets("注文")

        worksheet.PrintOut()
        time.sleep(1)

        workbook.Close(SaveChanges=False)
        excel.Quit()
        time.sleep(2)

    except Exception as e:
        print(f"エラーが発生しました: {e}")

def main():
    csv_file_path = r"C:\Users\USER06\Desktop\Auto_code\16時印刷\出勤日カレンダー.csv"
    try:
        today = datetime.now().date()
        next_working = None

        with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row['working'] == '1':
                    d = datetime.strptime(row['date'], "%Y/%m/%d").date()

                    # 本日より後の日付をリスト追加
                    if d > today and (next_working is None or d < next_working):
                        next_working = d

        if next_working:
            date_str = next_working.strftime("%Y/%m/%d")
            print(f"次の営業日: {date_str}")
            process_and_print(date_str)
        else:
            print("今後の営業日が見つかりません。")
                
    except FileNotFoundError:
        print(f"CSVファイルが見つかりません: {csv_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()