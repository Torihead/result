import pyautogui
import csv
from datetime import datetime

# 出荷予定の印刷する関数
def process_and_print(date_str):
    try:
        # 受付業務タグをクリック
        pyautogui.click(x=700, y=302)
        # 製品出荷予定表を開く
        pyautogui.click(x=618, y=451)
        pyautogui.sleep(0.5)
        # 印刷内容の設定
        pyautogui.write("1", interval=0.1)
        pyautogui.press("enter")
        pyautogui.write(date_str, interval=0.1)
        pyautogui.press("enter")
        pyautogui.write(date_str, interval=0.1)
        pyautogui.press("enter", presses=3, interval=0.2)
        pyautogui.press("left", presses=2, interval=0.2)
        pyautogui.press("enter")
        pyautogui.press("space")

        # 印刷開始
        pyautogui.press("F6")
        pyautogui.press("left")
        pyautogui.press("enter")
        pyautogui.sleep(5)
        # 画面終了
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
            date_str = next_working.strftime("%Y/%m/%d")
            print(f"次の営業日: {date_str}")
            process_and_print(date_str)
        else:
            print("今後の営業日が見つかりません。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
