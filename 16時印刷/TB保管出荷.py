import os
import time
import win32com.client as w32
import get_next_workday as gnw

#os.system("taskkill /F /IM excel.exe")                # Excelのタスク終了(非常時)

def main():
    formatted_date = f"{gnw.next_workday.year}年{gnw.next_workday.month}月{gnw.next_workday.day}日"
    formatted_second_date = f"{gnw.second_next_workday.year}年{gnw.second_next_workday.month}月{gnw.second_next_workday.day}日"

    #excel = w32.gencache.EnsureDispatch("Excel.Application") # 事前バインディングによる起動(RefreshAll対応)
    excel = w32.Dispatch("Excel.Application")
    excel.Visible = False                                    # バックグラウンド処理

    file_path = r"\\MC10\share\MICHINOK_共有\0.共有書類\対照帳\端量　在庫表\TB保管出荷.xlsx"
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Worksheets("Sheet1")
    time.sleep(2)

    ws.Range("G1").Value = formatted_date
    ws.Range("J1").Value = formatted_second_date

    wb.RefreshAll()
    time.sleep(3)

    #ws.Range("L2").AutoFilter(Field:=10, Criteria1:="<>", Operator:=1)   # Field は、インデックス番号(左からの何番目にあるか)
    ws.Range("A1").AutoFilter(10, "<>")
    time.sleep(1)

    ws.PrintOut()
    time.sleep(2)

    wb.Close(SaveChanges=False)
    excel.Quit()
    time.sleep(2)
if __name__ == "__main__":
    main()