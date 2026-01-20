"""
出荷予定の印刷処理
"""
import pyautogui
import time
import workday_utils
import print_utils
import config


def process_and_print() -> None:
    """出荷予定を印刷"""
    try:
        next_date = workday_utils.get_next_working_day()
        if not next_date:
            print("次の営業日が見つかりません。")
            return

        date_str = next_date.strftime("%Y/%m/%d")
        print(f"出荷予定印刷処理開始: {date_str}")

        # 受付業務タグをクリック
        print_utils.click_uketsukegimu()

        # 製品出荷予定表を開く
        print_utils.safe_click(config.SEHIN_SHOKKA_X, config.SEHIN_SHOKKA_Y, 
                                sleep_after=0.5)

        # 印刷内容の設定
        print_utils.safe_write("1", interval=0.1)
        print_utils.safe_press("enter")
        print_utils.safe_write(date_str, interval=0.1)
        print_utils.safe_press("enter")
        print_utils.safe_write(date_str, interval=0.1)
        print_utils.safe_press("enter", presses=3, interval=0.2)
        print_utils.safe_press("left", presses=2, interval=0.2)
        print_utils.safe_press("enter")
        print_utils.safe_press("space")

        # 印刷開始
        print_utils.safe_press("F6")
        print_utils.safe_press("left")
        print_utils.safe_press("enter")
        time.sleep(5)

        # 画面終了
        print_utils.print_preview_close()

        print("出荷予定の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
