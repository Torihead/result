import pyautogui
import datetime as dt
import jpholiday as hd

import get_next_workday as gnw
today = dt.datetime.today()
next_workday = gnw.get_next_weekday(today)

def main():
    # 受付業務タグをクリック
    pyautogui.click(x=700, y=302)
    # 週間受注を開く
    pyautogui.click(x=641, y=381)
    pyautogui.sleep(1)
    pyautogui.write(next_workday.strftime('%Y%m%d'), interval=0.1)
    pyautogui.press("enter", presses=4, interval=0.6)
    # 画面印刷の画面
    pyautogui.sleep(2)
    pyautogui.click(x=65, y=37)
    pyautogui.sleep(1)
    pyautogui.click(x=1137, y=280)
    pyautogui.sleep(1)
    # 印刷プロパティを開いた状態

    # ページレイアウト (2 in 1)
    pyautogui.click(x=1103, y=585)
    pyautogui.click(x=1101, y=654)
    # 両面印刷
    pyautogui.click(x=1101, y=654)
    pyautogui.click(x=1090, y=728)
    # 給紙設定
    pyautogui.click(x=950, y=296)
    pyautogui.click(x=1252, y=526)
    pyautogui.sleep(1)

    # 印刷開始
    pyautogui.press("enter", presses=2, interval=0.5)
    pyautogui.sleep(3)

    # 画面閉じ
    pyautogui.click(x=1897, y=13)
    pyautogui.press("F12")
if __name__ == "__main__":
    main()