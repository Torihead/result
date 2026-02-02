import win32com.client as w32
import pyautogui
import time
import csv
from datetime import datetime, timedelta

# 製品入出庫予定の処理と印刷をする関数
def process_and_print(date_str):
    try:
        pyautogui.click(x=700, y=302)                             # 受付業務タグをクリック
        time.sleep(0.5)
        pyautogui.click(x=761, y=448)                             # 製造入出庫予定照会をクリック
        time.sleep(1)
        pyautogui.press("F5")                                     # 表示ボタン
        time.sleep(0.7)
        pyautogui.press("Tab", presses=2, interval=0.1)
        pyautogui.press("Enter")                                  # OUT出力
        time.sleep(1)

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False
        file_path = r"\\MC10\share\OA\EXCEL\OUT\12_SEIHIN_NSK_YOTEI.XLS"
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Worksheets("製品入出庫定照会")
        time.sleep(1)

        # 通常は3列を削除。翌日が休業なら、6列を削除する
        today = datetime.now().date()
        if date_str == today + timedelta(days=1):    # 翌営業日 == 明日なら通常
            ws.Columns["F:H"].Delete()
        else:
            ws.Columns["F:K"].Delete()

        #フィルター設定
        ws.Range("D8").AutoFilter(Field:=4, Criteria1:="<>BB", Operator:=1)
        ws.Range("D8").AutoFilter(Field:=5, Criteria1:="<>臨港倉庫", Operator:=1)
        ws.Range("G8").AutoFilter(Field:=7, Criteria1:="<>0", Operator:=1)

        #レイアウト、印刷
        xlup = -4162                                            # セルの最終行
        last_row = ws.Cells(ws.Rows.Count, 12).End(xlup).Row    # L列の最終行を取得

        ws.PageSetup.PrintArea = f"A7:L{last_row}"              # 印刷範囲を選択

        ps = ws.PageSetup
        ps.Orientation      = 2                                 # 横方向に印刷
        ps.Zoom             = False                             # Zoom(%)無効
        ps.FitToPagesWide   = 1                                 # ヨコを１ページに
        ps.FitToPagesTall   = 1                                 # タテを１ページに

        ws.PrintOut()                                   # コピー部数
        time.sleep(1)
        ws.PrintOut()

        wb.Close(SaveChanges=True)
        excel.Quit()

    except Exception as e:
        print(f"エラーが発生しました: {e}")

# CSVファイルから、次の営業日を取得する関数
def main():
    try:
        csv_file_path = r"C:\Users\USER06\Desktop\Auto_code\16時印刷\出勤日カレンダー.csv"
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