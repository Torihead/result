import win32com.client as w32
import time
import os
import datetime as dt
#os.system('taskkill /f /im EXCEL.EXE')

def create_親会社報告書():

    excel = w32.Dispatch("Excel.Application")
    excel.Visible = False
    time.sleep(1)

    today = dt.datetime.now()                               # 本日の日付を取得
    last_month = today - dt.timedelta(days=10)              # 先月
    formatted_date = last_month.strftime("%Y.%m")           # 現在の日付をYYYY.MM形式で取得
    formatted_month = f"{formatted_date}_"                  # "YYYY.MM_"に変換
    if today.month == 4:
        formatted_year = str(today.year - 1)
    else:
        formatted_year = str(today.year)

    filepath = fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}SYOYUSYA_GENRYO_NSK.XLS"
    wb = excel.Workbooks.Open(filepath)
    ws = [wb.Worksheets("日和001"),
        wb.Worksheets("雪種002"),
        wb.Worksheets("日清003")
        ]

    # Sheet 3つの表を繰り返し処理
    for i in ws:
        rowcell_B = i.Range("B4").End(-4121).Row
        sum_row = rowcell_B + 1
        
        i.Range(f"B{sum_row}").Value = f"=SUM(B4:B{rowcell_B})"             # B列にSUM関数
        i.Range(f"B{sum_row}").AutoFill(i.Range(f"B{sum_row}:G{sum_row}"))

        last_row_G = i.Range("G4").End(-4121).Row                           # G列の書式をH列にコピー
        i.Range(f"G4:G{last_row_G}").Copy()
        i.Range(f"H4:H{last_row_G}").PasteSpecial(Paste=-4104)
        excel.CutCopyMode = False

        last_row_H = i.Range("H4").End(-4121).Row                           # H列に歩留まりを作成
        i.Range("H4").Value = ""
        i.Range("H5").Value = "=F5/E5"
        i.Range("H5").NumberFormat = "0.00%"
        i.Range("H5").AutoFill(i.Range(f"H5:H{last_row_H}"))

        for rng in range(5, 70):                                            # F列の加工量が0のとき、歩留まり0
            cell_value = i.Range(f"F{rng}").Value
            if cell_value == 0:
                i.Range(f"H{rng}").Value = "0"

        for rng in range(6, sum_row + 2 ):                                  # 表の点線をそろえる(なぜsum_row+2なのか理屈は不明)
            i.Range(f"H{rng}").Borders(8).LineStyle = -4118                 # 上 点線

        border_rng = i.Range(f"B{sum_row}:H{sum_row}")                      # 合計行（sum_row）のB～Hに格子罫線＋下罫線を実線で上書き
        for border in [7, 8, 9, 10, 11, 12]:                    # 格子罫線
            border_rng.Borders(border).LineStyle = 1            # 実線

    # なたね粕の貼り付け
    for n in range(5, rowcell_B):                                       # "日清003"から、なたね粕の行をコピー
        if ws[2].Range(f"A{n}").Value == "87960 Ｃなたね粕":
            ws[2].Range(f"A{n}:H{n}").Copy()

    paste_row = ws[0].Range("A4").End(-4121).Row + 3                    # 再度、最終行を取得する。最初の最終行(rowcell_B=33行目)で固定されるため更新

    ws[0].Range(f"A{paste_row}:H{paste_row}").PasteSpecial(Paste=-4163) # A最終行の3つ下に貼り付け
    excel.CutCopyMode = False
    ws[0].Range(f"H{paste_row}").NumberFormat = "0.00%"
    for border in [7, 8, 9, 10, 11, 12]:                                            # 格子罫線
            ws[0].Range(f"A{paste_row}:H{paste_row}").Borders(border).LineStyle = 1 # 実線

    print("--------------------所有者原料入出庫の編集を完了しました。")

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
    for wb_genryo in wb_list.values():     # 新規ブックを閉じる処理
        wb_genryo.Close(SaveChanges=True)

    src_path = fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}SYOYUSYA_GENRYO_NSK.XLS"
    src_wb = excel.Workbooks.Open(src_path)

    for name, sheet_name in filename_map.items():       # 辞書用メソッドitems()で、name:sheet_nameの要素を扱う
        src_sheet = src_wb.Sheets(sheet_name)
        #src_sheet.Copy()

        dest_wb = wb_list[name]

        src_sheet.Copy(After=dest_wb.Sheets(dest_wb.Sheets.Count))
    src_wb.Close(SaveChanges=False)

    # 親会社報告を編集して保存
    wb_1 = excel.Workbooks.Open(fr"\\MC10\share\OA\EXCEL\OUT\NICH_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx")
    ws_1 = wb_1.Worksheets("Sheet1")
    for i in range(120, 3, -1):                                     # 逆順にループしてシフトズレを防ぐ
        if ws_1.Cells(3, i).Value != "日和産業株式会社":
            ws_1.Columns(i).Delete()                                # 列を削除して左に詰める
    wb_1.Close(SaveChanges=True)

    wb_2 = excel.Workbooks.Open(fr"\\MC10\share\OA\EXCEL\OUT\YUKI_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx")
    ws_2 = wb_2.Worksheets("Sheet1")
    for i in range(120, 3, -1):                                     # 逆順にループしてシフトズレを防ぐ
        if ws_2.Cells(3, i).Value != "雪印種苗株式会社":
            ws_2.Columns(i).Delete()                                # 列を削除して左に詰める
    wb_2.Close(SaveChanges=True)

    wb_3 = excel.Workbooks.Open(fr"\\MC10\share\OA\EXCEL\OUT\NISS_{formatted_month}SYOYUSYA_GENRYO_NSK.xlsx")
    ws_3 = wb_3.Worksheets("Sheet1")
    for i in range(120, 3, -1):                                     # 逆順にループしてシフトズレを防ぐ
        if ws_3.Cells(3, i).Value != "日清丸紅飼料(株)":
            ws_3.Columns(i).Delete()                                # 列を削除して左に詰める
    wb_3.Close(SaveChanges=True)

    excel.Quit()

    print("--------------------親会社報告書の作成を完了しました。")
    print("親会社報告書の処理が完了しました。")