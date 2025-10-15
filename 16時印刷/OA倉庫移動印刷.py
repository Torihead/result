import pyautogui
import datetime as dt
import jpholiday as hd

import get_next_workday as gnw
today = dt.datetime.today()
next_workday = gnw.get_next_weekday(today)

def main():
    # 受付業務タグをクリック
    pyautogui.click(x=700, y=302)
    # 倉庫移動を開く
    pyautogui.click(x=830, y=414)
    pyautogui.sleep(1)
    # 印刷内容を設定
    pyautogui.write(next_workday.strftime('%Y%m%d'), interval=0.1)
    pyautogui.press("enter")
    pyautogui.write(next_workday.strftime('%Y%m%d'), interval=0.1)
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
if __name__ == "__main__":
    main()