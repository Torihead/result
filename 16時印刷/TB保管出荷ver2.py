import time
import win32com.client as w32
import csv
from datetime import datetime, date

def get_next_two_working_days(csv_file_path: str):
    """ CSVファイルから2営業日を取得する関数 """
    today = date.today()
    target_dates = []
    with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 必須列が揃っているかチェック（不正行の安全対策）
            if 'date' not in row or 'working' not in row:
                continue
            if row['working'] == '1':
                try:
                    d = datetime.strptime(row['date'], "%Y/%m/%d").date()
                except ValueError:
                    # フォーマット不正行はスキップ
                    continue
                if d > today:
                    target_dates.append(d)

    target_dates.sort()
    # 最初の2営業日を返す
    return target_dates[:2]

def process_and_print(next_date_str: str, second_date_str: str | None = None):
    """ 取得した日付文字列を使用して、処理と印刷を行う関数
    second_date_str は存在しない場合 None のままでOK """

    try:
        # 次の勤務日（必須）
        next_date_obj = datetime.strptime(next_date_str, "%Y/%m/%d")
        formatted_date = next_date_obj.strftime("%Y/%m/%d")

        # その次の勤務日（存在する場合）
        formatted_second_date = None
        if second_date_str:
            second_date_obj = datetime.strptime(second_date_str, "%Y/%m/%d")
            formatted_second_date = second_date_obj.strftime("%Y/%m/%d")

        # Excel操作
        excel = w32.Dispatch("Excel.Application")
        excel.Visible = True

        file_path = r"\\MC10\share\MICHINOK_共有\0.共有書類\対照帳\端量　在庫表\TB保管出荷.xlsx"
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Worksheets("Sheet1")
        time.sleep(2)

        ws.Range("G1").Value = formatted_date
        if formatted_second_date:
            ws.Range("J1").Value = formatted_second_date
        wb.RefreshAll()
        time.sleep(4)

        ws.Range("A1").AutoFilter(10, "<>")
        time.sleep(1)

        ws.PrintOut()
        time.sleep(3)
        wb.Close(SaveChanges=True)
        excel.Quit()
        time.sleep(3)

    except Exception as e:
        print(f"エラーが発生しました: {e}")

def main():
    csv_file_path = r"C:\Users\USER06\Desktop\Auto_code\16時印刷\出勤日カレンダー.csv"
    try:
        next_two = get_next_two_working_days(csv_file_path)

        if not next_two:
            print("次の営業日が見つかりません。")
            return
        
        # 次の営業日
        next_working = next_two[0]
        next_working_str = next_working.strftime("%Y/%m/%d")

        # その次の営業日
        second_working_str = None
        if len(next_two) >= 2:
            second_working = next_two[1]
            second_working_str = second_working.strftime("%Y/%m/%d")

        print(f"次の営業日: {next_working_str}")
        if second_working_str:
            print(f"その次の営業日: {second_working_str}")
        
        # 処理を引き渡し
        process_and_print(next_working_str, second_working_str)


    except FileNotFoundError:
        print(f"CSVファイルが見つかりません: {csv_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()