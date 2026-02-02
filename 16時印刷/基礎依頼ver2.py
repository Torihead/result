import win32com.client as w32
import time
import csv
from datetime import datetime

# 指定された日付文字列を使用して、処理と印刷を行う関数
def process_and_print(date_str):
    try:
        date_obj = datetime.strptime(date_str, "%Y/%m/%d")
        formatted_date = date_obj.strftime("%Y.%m.%d")
        new_file_name = f"基礎依頼票_{formatted_date}.xlsm"
        new_file_path = fr"\\MC10\share\MICHINOK_共有\（仮）\基礎依頼\基礎依頼票_BaukUp\{new_file_name}"

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False

        file_path = r"\\MC10\share\MICHINOK_共有\（仮）\基礎依頼\基礎依頼票.xlsm"
        workbook = excel.Workbooks.Open(file_path)
        time.sleep(3)

        worksheet = workbook.Worksheets("製造基礎依頼予定表")
        time.sleep(1)
        workbook.RefreshAll()
        time.sleep(10)

        workbook.SaveAs(new_file_path)
        print(f"保存した基礎依頼ファイル ➡  {new_file_path}")
        time.sleep(2)

        rng = worksheet.Range("E5:E50")
        rng.Copy()
        rng.PasteSpecial(Paste=-4163)
        worksheet.Range("E5").AutoFilter(Field:=4, Criteria1:="<>昼", Operator:=1)

        worksheet.PrintOut()
        time.sleep(3)

        workbook.Close(SaveChanges=False)
        excel.Quit()
        time.sleep(2)

    except Exception as e:
        print(f"エラーが発生しました: {e}")

# 指定されたCSVファイルから、次の営業日の日付を取得し、処理を実行
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