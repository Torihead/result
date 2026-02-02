import os
import time
import win32com.client as w32
import datetime as dt


#os.system("taskkill /F /IM excel.exe")
excel = w32.Dispatch("Excel.Application")
excel.Visible = True

# 今日を取得
today = dt.date.today()
year = today.year
month = today.month
last_month = month - 1

# 年度を取得
if last_month < 4:
    fiscal_year = year - 1
else:
    fiscal_year = year
year_parts = f"{fiscal_year}年度"
month_parts = f"{year}.{last_month:02d}月"

file_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{year_parts}\{month_parts}\06　親会社月次報告\1　日和\原料在庫証明.xls"
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets(1)
time.sleep(1)

ws.Range("A6").Value = f"{year}/{last_month}/1"
wb.RefreshAll()

output_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{year_parts}\{month_parts}\06　親会社月次報告\1　日和\原料在庫証明1.pdf"
excel.ActiveSheet.ExportAsFixedFormat(0, output_path)

wb.Close(SaveChanges=True)
excel.Quit()