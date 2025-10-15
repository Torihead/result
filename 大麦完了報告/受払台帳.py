import time
import win32com.client as w32

def create_受払台帳():
    excel = w32.Dispatch("Excel.Application")
    excel.Visible = True
    filepath = r"\\MC10\share\OA\EXCEL\OUT\GMG_UKHRI_DAI.xlsx"
    wb = excel.Workbooks.Open(filepath)
    ws = wb.Sheets("大麦")
    time.sleep(2)

    # F列の最終行0を確認
    rowcell = ws.Cells(9, 6).End(-4121).Row     # xlDown
    ws.Cells(rowcell, 6).Select()
    if ws.Cells(rowcell, 6).Value != 0:
        print("受払チェック: 大麦の割当がまだのため中断します。")
        wb.Close(SaveChanges=False)
        exit()

    # 日付を消す
    ws.Range("F2").ClearContents()

    # フィルター設定
    try:
        ws.Range("A8:F8").AutoFilter()
    except Exception as e:
        print(f"フィルター失敗:{e}")

    # 9行目でウィンドウ枠の固定
    ws.Range("A9").Select()
    excel.ActiveWindow.FreezePanes = True

    # E9～F列の最下層　0.00表記
    row_max = ws.Cells(ws.Rows.Count, 6).End(-4162).Row     # xlUp
    ws.Range(ws.Cells(9, 5), ws.Cells(row_max, 6)).NumberFormatLocal = "#,##0.00_ ;[赤]-#,##0.00 "

    # 不要な下部のセルの削除
    delete_range = ws.Range(ws.Cells(row_max + 1, 1), ws.Cells(row_max * 2, 6))
    delete_range.Delete(Shift=-4159)    #xlToLeft

    # 月末ロスの計算式を貼り付け
    ws.Range("G3").Value = "月末ロス"
    ws.Range("G4").Formula = '=LEFT(E4,FIND("K",E4)-1)-SUMIF(D:D,G3,E:E)'
    ws.Range("G3:G4").Interior.ColorIndex = 6   # 黄色
    ws.Range("G4").NumberFormatLocal = "#,##0;-#,##0"

    # 月末ロスの値をメッセージ
    loss = round(ws.Range("G4").Value)
    print(f"受払の処理が終わりました。\n\n月末のロスの値は{loss}です。")
    
    # 印刷設定
    ws.Range("G:G").EntireColumn.Hidden = True      # 非表示

    row_num = rowcell
    ws.PageSetup.PrintArea = f"$A$2:$F${row_num}"   # 印刷範囲
    ws.PageSetup.CenterFooter = "&P/&N"             # 1/Nページ
    ws.PageSetup.PrintTitleRows = "$7:$8"           # タイトル行
    #ws.PrintOut()
    #wb.Close(SaveChanges=True)  # 保存して閉じる
    print("--------------------受払台帳の処理が完了しました。")
    # 加工完了4で利用
    return loss

if __name__ == "__main__":
    create_受払台帳()