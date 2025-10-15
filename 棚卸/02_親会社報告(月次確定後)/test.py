import win32com.client as w32
import time
import os
import datetime as dt
#os.system('taskkill /f /im EXCEL.EXE')

def main():

    excel = w32.Dispatch("Excel.Application")
    excel.Visible = True
    time.sleep(1)

    today = dt.datetime.now()                               # 本日の日付を取得
    last_month = today - dt.timedelta(days=10)              # 先月
    formatted_date = last_month.strftime("%Y.%m")           # 現在の日付をYYYY.MM形式で取得
    formatted_month = f"{formatted_date}_"                  # "YYYY.MM_"に変換

    filepath = fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}SYOYUSYA_GENRYO_NSK.XLS"
    wb = excel.Workbooks.Open(filepath)
    ws = [wb.Worksheets("日和001"),
        wb.Worksheets("雪種002"),
        wb.Worksheets("日清003")
        ]

    
    # 割戻表1を3社ぶん、新規ブックで作成

    filepath = fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI1.XLS"
    wb_genryo = excel.Workbooks.Open(filepath)
    ws = wb_genryo.Sheets("Sheet1")
    time.sleep(1)

    # 親会社報告書のブック作成
    names = ["NICH", "YUKI", "NISS"]
    filename_map = dict(zip(names, ["日和001", "雪種002", "日清003"]))  # names と sheet名 をペアにして辞書化(dict)
    wb_list = {}

    for name in names:                                         # 親会社報告の新規ブックを作成
        ws.Copy()
        new_wb = excel.ActiveWorkbook
        new_filename = f"{name}_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx"
        save_path = os.path.join(r"\\MC10\share\OA\EXCEL\OUT", new_filename)
        new_wb.SaveAs(save_path)
        wb_list[name] = new_wb      # 新規ブックを辞書へ
    
    print(wb_list)

    src_path = fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}SYOYUSYA_GENRYO_NSK.XLS"
    src_wb = excel.Workbooks.Open(src_path)

    for name, sheet_name in filename_map.items():       # 辞書用メソッドitems()で、name:sheet_nameの要素を扱う
        src_sheet = src_wb.Sheets(sheet_name)   # 親会社001,002,003のsheetを取得
        print(name, ",", sheet_name)

        dest_wb = wb_list[name]
        
        src_sheet.Copy(After=dest_wb.Sheets(dest_wb.Sheets.Count))  # 新規ブックの最後にコピー
    src_wb.Close(SaveChanges=False)
    
    for wb_genryo in wb_list.values():     # 新規ブックを閉じる処理
        wb_genryo.Close(SaveChanges=True)

    print("おわり")
    excel.Quit()

if __name__ == "__main__":
    main()
