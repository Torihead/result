"""
製品入出庫予定の処理と印刷
"""
import win32com.client as w32
import pyautogui
import time
from datetime import datetime, timedelta
import workday_utils
import print_utils
import config


def process_and_print() -> None:
    """製品入出庫予定を処理して印刷"""
    try:
        next_date = workday_utils.get_next_working_day()
        if not next_date:
            print("次の営業日が見つかりません。")
            return

        print(f"製品入出庫予定処理開始: {next_date}")

        print_utils.click_uketsukegimu()
        time.sleep(0.5)

        # 製造入出庫予定照会をクリック
        print_utils.safe_click(config.SEIZO_INYUSHUKO_X, config.SEIZO_INYUSHUKO_Y, 
                                sleep_after=1)
        print_utils.safe_press("F5", sleep_after=0.7)
        print_utils.safe_press("Tab", presses=2, interval=0.1)
        print_utils.safe_press("enter", sleep_after=1)

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(config.OUT_FILE)
        ws = wb.Worksheets("製品入出庫定照会")
        time.sleep(1)

        # 通常は3列を削除。翌日が休業なら、6列を削除する
        today = datetime.now().date()
        if next_date == today + timedelta(days=1):  # 翌営業日 == 明日なら通常
            ws.Columns["F:H"].Delete()
        else:
            ws.Columns["F:K"].Delete()

        # フィルター設定
        ws.Range("D8").AutoFilter(Field=4, Criteria1="<>BB", Operator=1)
        ws.Range("D8").AutoFilter(Field=5, Criteria1="<>臨港倉庫", Operator=1)
        ws.Range("G8").AutoFilter(Field=7, Criteria1="<>0", Operator=1)

        # レイアウト、印刷
        xlup = -4162
        last_row = ws.Cells(ws.Rows.Count, 12).End(xlup).Row

        ws.PageSetup.PrintArea = f"A7:L{last_row}"

        ps = ws.PageSetup
        ps.Orientation = 2
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1

        ws.PrintOut()
        time.sleep(1)
        ws.PrintOut()

        wb.Close(SaveChanges=True)
        excel.Quit()

        print("製品入出庫予定の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
