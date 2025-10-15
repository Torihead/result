import subprocess
import pyautogui
import time

# rdp起動Class
def main():
    subprocess.run(["mstsc.exe", r"C:\Users\USER06\Desktop\OAシステム.rdp"])
    # 起動時間 待機
    pyautogui.sleep(4)

    # 起動後アプリの対象クリック、ログイン処理
    pyautogui.click(x=826, y=448)
    pyautogui.write("12", interval=0.1)
    pyautogui.press("enter", presses=3, interval=0.5)
    time.sleep(2)
