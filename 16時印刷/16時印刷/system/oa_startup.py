"""
OA起動処理
"""
import subprocess
import pyautogui
import time
import config


def main() -> None:
    """OAシステムを起動してログイン"""
    try:
        print("OA起動処理開始")
        
        subprocess.run(["mstsc.exe", config.RDP_FILE])

        # 起動時間 待機
        time.sleep(config.WAIT_OA_START)

        # 起動後アプリの対象クリック、ログイン処理
        pyautogui.click(x=config.OA_CLICK_X, y=config.OA_CLICK_Y)
        pyautogui.write("12", interval=0.1)
        pyautogui.press("enter", presses=3, interval=0.5)
        time.sleep(2)

        print("OA起動処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


if __name__ == "__main__":
    main()
