import pyautogui
import csv
from datetime import datetime

# 週間受注の印刷する関数
def process_and_print(date_str):
    try:
        # 受付業務タグをクリック
        pyautogui.click(x=700, y=302)
        # 週間受注を開く
        pyautogui.click(x=641, y=381)
        pyautogui.sleep(1)
        pyautogui.write(date_str, interval=0.1)
        pyautogui.press("enter", presses=4, interval=0.6)
        # 画面印刷の画面
        pyautogui.sleep(2)
        pyautogui.click(x=65, y=37)
        pyautogui.sleep(1)
        pyautogui.click(x=1137, y=280)
        pyautogui.sleep(2)
        # 印刷プロパティを開いた状態

        # ページレイアウト (2 in 1)
        pyautogui.click(x=1103, y=585)
        pyautogui.sleep(0.5)
        pyautogui.click(x=1101, y=654)
        pyautogui.sleep(0.5)
        # 両面印刷
        pyautogui.click(x=1101, y=654)
        pyautogui.sleep(0.5)
        pyautogui.click(x=1090, y=728)
        pyautogui.sleep(0.5)
        # 給紙設定
        pyautogui.click(x=950, y=296)
        pyautogui.sleep(0.5)
        pyautogui.click(x=1252, y=526)
        pyautogui.sleep(1)

        # 印刷開始
        pyautogui.press("enter", presses=2, interval=0.5)
        pyautogui.sleep(3)

        # 画面閉じ
        pyautogui.click(x=1897, y=13)
        pyautogui.press("F12")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

# 次の営業日を取得するメイン関数
def main():
    try:
        csv_file_path = r"C:\Users\USER06\Desktop\Auto_code\16時印刷\出勤日カレンダー.csv"
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
            date_str = next_working.strftime("%Y/%m/%d")
            print(f"次の営業日: {date_str}")
            process_and_print(date_str)
        else:
            print("今後の営業日が見つかりません。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()