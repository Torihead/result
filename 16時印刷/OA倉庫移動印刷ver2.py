import pyautogui
from datetime import datetime
import csv

def process_and_print(date_str):
    try:
        # 受付業務タグをクリック
        pyautogui.click(x=700, y=302)
        # 倉庫移動を開く
        pyautogui.click(x=830, y=414)
        pyautogui.sleep(1)
        # 印刷内容を設定
        pyautogui.write(date_str.strftime('%Y%m%d'), interval=0.1)
        pyautogui.press("enter")
        pyautogui.write(date_str.strftime('%Y%m%d'), interval=0.1)
        pyautogui.press("enter", presses=5, interval=0.3)
        pyautogui.sleep(1)
        # 画面印刷をクリック
        pyautogui.click(x=1437, y=859)
        pyautogui.sleep(3)
        # 印刷プレビュー内の操作
        pyautogui.click(x=65, y=37)
        pyautogui.press("enter")
        pyautogui.sleep(3)
        pyautogui.click(x=1897, y=13)
        pyautogui.press("F12")

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

                    # 本日より後の日付にリスト追加
                    if d > today and (next_working is None or d < next_working):
                        next_working = d

        if next_working:
            print(f"次の営業日: {next_working.strftime('%Y/%m/%d')}")
            process_and_print(next_working)
        else:
            print("今後の営業日が見つかりません。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()