"""
週間受注の印刷処理
"""
import pyautogui
import time
import workday_utils
import print_utils
import config


def process_and_print() -> None:
    """週間受注を印刷"""
    try:
        next_date = workday_utils.get_next_working_day()
        if not next_date:
            print("次の営業日が見つかりません。")
            return

        date_str = next_date.strftime('%Y%m%d')
        print(f"週間受注印刷処理開始: {date_str}")

        # 受付業務タグをクリック
        print_utils.click_uketsukegimu()

        # 週間受注を開く
        print_utils.safe_click(config.SHUKAN_JUSCHU_X, config.SHUKAN_JUSCHU_Y, 
                                sleep_after=config.WAIT_MEDIUM)
        print_utils.safe_write(date_str, interval=0.1)
        print_utils.safe_press("enter", presses=4, interval=0.6)

        # 画面印刷の画面
        time.sleep(2)
        print_utils.safe_click(65, 37, sleep_after=config.WAIT_MEDIUM)
        print_utils.safe_click(1137, 280, sleep_after=2)

        # ページレイアウト (2 in 1)
        print_utils.safe_click(1103, 585, sleep_after=0.5)
        print_utils.safe_click(1101, 654, sleep_after=0.5)

        # 両面印刷
        print_utils.safe_click(1101, 654, sleep_after=0.5)
        print_utils.safe_click(1090, 728, sleep_after=0.5)

        # 給紙設定
        print_utils.safe_click(950, 296, sleep_after=0.5)
        print_utils.safe_click(1252, 526, sleep_after=1)

        # 印刷開始
        print_utils.safe_press("enter", presses=2, interval=0.5)
        time.sleep(3)

        # 画面閉じ
        print_utils.close_screen()
        print_utils.print_preview_close()

        print("週間受注の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
