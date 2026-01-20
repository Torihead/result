"""
翌営業日端量表の処理と印刷
"""
import win32com.client as w32
import time
import workday_utils
import config


def process_and_print() -> None:
    """翌営業日端量表を処理して印刷"""
    try:
        next_date = workday_utils.get_next_working_day()
        if not next_date:
            print("次の営業日が見つかりません。")
            return

        print(f"翌営業日端量表処理開始: {next_date}")

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False

        wb = excel.Workbooks.Open(config.HARYOU_FILE)
        ws_TB = wb.Worksheets("翌日製造使用端量表")
        ws_PB = wb.Worksheets("翌営業日PB製造出荷予定表")

        formatdate = next_date.strftime('%m月%d日')
        ws_TB.Range("H2").Value = formatdate
        time.sleep(2)

        wb.RefreshAll()
        time.sleep(6)

        ws_TB.PrintOut()
        ws_PB.PrintOut()

        wb.Close(SaveChanges=True)
        excel.Quit()
        time.sleep(2)

        print("翌営業日端量表の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
