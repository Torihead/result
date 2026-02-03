import os
import time
import win32com.client as win

excel = win.Dispatch("Excel.Application")
excel.Visible = False

file_paths = [
    r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI1.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI2.XLS",
    r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI3.XLS"
]

for file_path in file_paths:                        #ファイルパスのリストを作って、繰り返し処理でコンパクト
    if not os.path.exists(file_path):
        print(f"ファイルが見つかりません: {file_path}")
    
    wb = excel.Workbooks.Open(file_path)
    ws = wb.ActiveSheet
    time.sleep(1)

    last_row = ws.Cells(ws.Rows.Count, 3).End(-4162).Row
    last_column = ws.Cells(last_row, 3).End(-4161).Column
    range_obj = ws.Range(ws.Cells(6, 3), ws.Cells(last_row, last_column))

    range_obj.FormatConditions.Delete()                     # 書式を消去
    cond = range_obj.FormatConditions.Add(1, 3, "0")        # xlValue,xlEqual,"0"を出力
    cond.Font.Color = 16777215                              # 白文字に変更

    # 印刷の編集
    ws.Rows("4:4").ShrinkToFit = True               # 縮小して全体表示

    ws.PageSetup.PrintTitleRows = "$3:$6"           # 行・列の設定
    ws.PageSetup.PrintTitleColumns = "$A:$B"

    pdf_path = file_path.replace(".XLS", ".pdf")
    ws.ExportAsFixedFormat(0, pdf_path)             # PDF出力

    wb.Close(SaveChanges=True)


# 原料入出庫
genryou_path = r"\\MC10\share\OA\EXCEL\OUT\GENRYO_NSK.XLS"
if not os.path.exists(genryou_path):
    print(f"原料入出庫ﾌｧｲﾙが見つかりません: {genryou_path}")

wb = excel.Workbooks.Open(genryou_path)
ws = wb.ActiveSheet
time.sleep(1)

last_row = ws.Cells(ws.Rows.Count, 3).End(-4162).Row
rng = ws.Range(ws.Cells(5, 8), ws.Cells(last_row, 8))
rng.NumberFormatLocal = "0.00%"

ws.PageSetup.PrintTitleRows = "$1:$4"
pdf_path = genryou_path.replace(".XLS", ".pdf")
ws.ExportAsFixedFormat(0, pdf_path)

wb.Close(SaveChanges=True)
excel.Quit()

print("割戻表の処理が完了しました。")
print("原料入出庫の処理が完了しました。")