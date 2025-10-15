import time
import win32com.client as w32
import get_next_workday as gnw

def main():
    file_path = r"\\MC10\share\MICHINOK_共有\0.共有書類\対照帳\端量　在庫表\翌営業日製造端量表（新）.xlsm"
    formadd_next_workday = f"{gnw.next_workday.month}月{gnw.next_workday.day}日"

    excel = w32.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(file_path)
    ws_TB = wb.Worksheets("翌日製造使用端量表")
    ws_PB = wb.Worksheets("翌営業日PB製造出荷予定表")

    ws_TB.Range("H2").Value = formadd_next_workday       # 日付を入力
    time.sleep(2)

    wb.RefreshAll()                                     # 全体更新
    time.sleep(5)

    ws_TB.PrintOut()                                    # 印刷
    ws_PB.PrintOut()

    wb.Close(SaveChanges=True)
    excel.Quit()
    time.sleep(2)
if __name__ == "__main__":
    main()