import pyautogui
import datetime as dt

import get_next_workday as gnw
today = dt.datetime.today()
next_workday = gnw.get_next_weekday(today)

# 出荷予定の印刷
def main():
    # 受付業務タグをクリック
    pyautogui.click(x=700, y=302)
    # 製品出荷予定表を開く
    pyautogui.click(x=618, y=451)
    pyautogui.sleep(0.5)
    # 印刷内容の設定
    pyautogui.write("1", interval=0.1)
    pyautogui.press("enter")
    pyautogui.write(next_workday.strftime('%Y%m%d'), interval=0.1)
    pyautogui.press("enter")
    pyautogui.write(next_workday.strftime('%Y%m%d'), interval=0.1)
    pyautogui.press("enter", presses=3, interval=0.2)
    pyautogui.press("left", presses=2, interval=0.2)
    pyautogui.press("enter")
    pyautogui.press("space")
    pyautogui.press("enter", presses=3, interval=0.2)
    pyautogui.sleep(0.8)
    # 印刷開始
    pyautogui.press("left")
    pyautogui.press("enter")
    pyautogui.sleep(5)
    # 画面終了
    pyautogui.press("F12")
if __name__ == "__main__":
    main()
