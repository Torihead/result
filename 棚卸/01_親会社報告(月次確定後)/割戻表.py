import os
import time
import win32com.client as win
import datetime as dt

def create_割戻表():

    excel = win.Dispatch("Excel.Application")
    excel.Visible = True

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

    print("原料入出庫の処理が終了しました。")
    print("--------------------割戻表の処理が完了しました。")

    before_path_list = [r"\\MC10\share\OA\EXCEL\OUT\SYOYUSYA_GENRYO_NSK.XLS",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI1.XLS",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI2.XLS",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI3.XLS",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_NSK.XLS",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI1.pdf",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI2.pdf",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_MODOSI3.pdf",
                        r"\\MC10\share\OA\EXCEL\OUT\GENRYO_NSK.pdf"
                        ]

    today = dt.datetime.now()                               # 本日の日付を取得
    last_month = today - dt.timedelta(days=10)              # 先月
    formatted_date = last_month.strftime("%Y.%m")           # 現在の日付をYYYY.MM形式で取得
    formatted_month = f"{formatted_date}_"                  # "YYYY.MM_"に変換
    if today.month == 4:
        formatted_year = str(today.year - 1)
    else:
        formatted_year = str(today.year)

    after_path_list = [ fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}SYOYUSYA_GENRYO_NSK.XLS",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI1.XLS",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI2.XLS",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI3.XLS",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_NSK.XLS",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI1.pdf",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI2.pdf",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_MODOSI3.pdf",
                        fr"\\MC10\share\OA\EXCEL\OUT\{formatted_month}GENRYO_NSK.pdf"
                        ]

    for before,after in zip(before_path_list, after_path_list):
        if os.path.exists(before):
            os.rename(before, after)
            print(f"{before}を\n{after}に変更しました。")
        else:
            print(f"{before}が見つかりません。")

    print("--------------------ファイルのリネームを完了しました。")
    print("割戻表と原料入出庫の処理が完了しました。")

if __name__ == "__main__":
    create_割戻表() 