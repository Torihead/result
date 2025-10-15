import win32com.client as w32
import get_next_workday as gnw
import time

def main():
    excel = w32.Dispatch("Excel.Application")       # Excelアプリ起動
    excel.Visible = False                           # ファイルをバックグラウンド処理(False)

    month = gnw.today.strftime('%Y.%m')     # 今月を2025.07で取得

    file_path = fr"\\MC10\share\MICHINOK_共有\0.共有書類\ペレット予測\{month} みちのくペレット在庫.xlsm"
    workbook = excel.Workbooks.Open(file_path)      # Workbook指定
    time.sleep(1)

    worksheet = workbook.Worksheets("注文")         # Worksheet指定
    
    worksheet.PrintOut()                            # 指定sheetをクイック印刷
    time.sleep(1)
    
    workbook.Close(SaveChanges=False)               # ファイルを閉じる
    excel.Quit()                                    # Excelアプリ終了
    time.sleep(2)
if __name__ == "__main__":
    main()