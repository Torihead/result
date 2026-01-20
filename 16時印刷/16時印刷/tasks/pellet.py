"""
ペレット予定の処理と印刷
"""
import win32com.client as w32
import time
from datetime import datetime
import workday_utils
import config


def process_and_print() -> None:
    """ペレット予定を処理して印刷"""
    try:
        next_date = workday_utils.get_next_working_day()
        if not next_date:
            print("次の営業日が見つかりません。")
            return

        date_str = next_date.strftime("%Y/%m/%d")
        date_obj = datetime.strptime(date_str, "%Y/%m/%d")
        manth = date_obj.strftime('%Y.%m')

        print(f"ペレット予定処理開始: {date_str}")

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False

        file_path = f"{config.PELLET_BASE}\\{manth} みちのくペレット在庫.xlsm"
        workbook = excel.Workbooks.Open(file_path)
        time.sleep(1)

        worksheet = workbook.Worksheets("注文")
        worksheet.PrintOut()
        time.sleep(1)

        workbook.Close(SaveChanges=False)
        excel.Quit()
        time.sleep(2)

        print("ペレット予定の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
