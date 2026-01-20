"""
基礎依頼票の処理と印刷
"""
import win32com.client as w32
import time
from datetime import datetime
import workday_utils
import config


def process_and_print() -> None:
    """基礎依頼票を処理して印刷"""
    try:
        next_date = workday_utils.get_next_working_day()
        if not next_date:
            print("次の営業日が見つかりません。")
            return

        date_str = next_date.strftime("%Y/%m/%d")
        date_obj = datetime.strptime(date_str, "%Y/%m/%d")
        formatted_date = date_obj.strftime("%Y.%m.%d")
        new_file_name = f"基礎依頼票_{formatted_date}.xlsm"
        new_file_path = f"{config.KISO_IRAI_BACKUP}\\{new_file_name}"

        print(f"基礎依頼処理開始: {date_str}")

        excel = w32.Dispatch("Excel.Application")
        excel.Visible = False

        workbook = excel.Workbooks.Open(config.KISO_IRAI_FILE)
        time.sleep(3)

        worksheet = workbook.Worksheets("製造基礎依頼予定表")
        time.sleep(1)
        workbook.RefreshAll()
        time.sleep(config.WAIT_EXCEL_REFRESH)

        workbook.SaveAs(new_file_path)
        print(f"保存した基礎依頼ファイル ➡  {new_file_path}")
        time.sleep(2)

        rng = worksheet.Range("E5:E50")
        rng.Copy()
        rng.PasteSpecial(Paste=-4163)
        worksheet.Range("E5").AutoFilter(Field=4, Criteria1="<>昼", Operator=1)

        worksheet.PrintOut()
        time.sleep(3)

        workbook.Close(SaveChanges=False)
        excel.Quit()
        time.sleep(2)

        print("基礎依頼の処理を完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise


def main() -> None:
    """メイン関数"""
    process_and_print()


if __name__ == "__main__":
    main()
