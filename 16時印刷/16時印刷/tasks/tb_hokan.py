"""
TB保管出荷の処理と印刷
"""
import win32com.client as w32
import time
from datetime import date
import workday_utils
import config


def process_and_print() -> None:
    """TB保管出荷を処理して印刷"""
    try:
        next_two = workday_utils.get_next_n_working_days(2)

        if not next_two:
            print("次の営業日が見つかりません。")
            return

        next_date = next_two[0]
        second_date = next_two[1] if len(next_two) >= 2 else None

        print(f"TB保管出荷処理開始: {next_date}")

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = True

        wb = excel.Workbooks.Open(config.TB_FILE)
        ws = wb.Worksheets("Sheet1")
        time.sleep(2)

        ws.Range("G1").Value = next_date.strftime("%Y/%m/%d")
        if second_date:
            ws.Range("J1").Value = second_date.strftime("%Y/%m/%d")

        wb.RefreshAll()
        time.sleep(4)

        ws.Range("A1").AutoFilter(10, "<>")
        time.sleep(1)

        ws.PrintOut()
        time.sleep(3)
        wb.Close(SaveChanges=True)
        excel.Quit()
        time.sleep(3)

        print("TB保管出荷の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
