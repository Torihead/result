"""
倉庫移動の印刷処理
"""
import pyautogui
import time
import workday_utils
import print_utils
import config


def process_and_print() -> None:
    """倉庫移動を印刷"""
    try:
        next_date = workday_utils.get_next_working_day()
        if not next_date:
            print("次の営業日が見つかりません。")
            return

        print(f"倉庫移動印刷処理開始: {next_date}")

        # 受付業務タグをクリック
        print_utils.click_uketsukegimu()

        # 倉庫移動を開く
        print_utils.safe_click(config.SOUKO_IDOU_X, config.SOUKO_IDOU_Y, 
                                sleep_after=1)

        # 印刷内容を設定
        date_str = next_date.strftime('%Y%m%d')
        print_utils.safe_write(date_str, interval=0.1)
        print_utils.safe_press("enter")
        print_utils.safe_write(date_str, interval=0.1)
        print_utils.safe_press("enter", presses=5, interval=0.3)
        time.sleep(1)

        # 画面印刷をクリック
        print_utils.safe_click(1437, 859, sleep_after=3)

        # 印刷プレビュー内の操作
        print_utils.safe_click(65, 37, sleep_after=config.WAIT_MEDIUM)
        print_utils.safe_press("enter")
        time.sleep(3)

        print_utils.close_screen()
        print_utils.print_preview_close()

        print("倉庫移動の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
