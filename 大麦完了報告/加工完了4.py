import time
import win32com.client as w32

def create_加工完了4():
    excel = w32.Dispatch("Excel.Application")
    excel.Visible = True
    filepath = r"\\MC10\share\OA\EXCEL\OUT\KAKO_HOKOKU4.XLS"
    wb = excel.Workbooks.Open(filepath)
    ws = wb.Worksheets("レイアウト")
    time.sleep(2)

    # 6行目で表示固定
    ws.Range("A6").Select()
    excel.ActiveWindow.FreezePanes = True

    # セルの結合、中央揃え
    merge_ranges = ["B5:D5", "E5:Y5", "Z5:AB5", "AC5:AF5", "AG5:AI5", "AJ5:AO5", "AP5:AU5"]
    for rng in merge_ranges:
        ws.Range(rng).Merge()
        ws.Range(rng).HorizontalAlignment = -4108   # xlCenter

    # 列幅の調整
    ws.Range("AG:AI").ColumnWidth = 1.88

    # フィルター設定　不可
    #ws.Range("B5:AU5").AutoFilter()

    # 関数の入力
    ws.Range("U3").Formula = '=AG3/J3*100'
    ws.Range("AW5").Value = "※銘柄計※"
    ws.Range("AV6").Formula = '=IF(AG6="","",ROUNDdown(AJ6*AG6/100-AP6,2))'
    ws.Range("AW6").Formula = '=IF(E6="",AW5,E6)'
    ws.Range("AX6").Formula = '=IF(AW6=$AW$5,SUMIF(AW:AW,AW5,AP:AP),"")'
    ws.Range("AY6").Formula = '=IF(AX6="","",AP6-AX6)'
    ws.Range("AV5:AY6").Interior.ColorIndex = 6
    ws.Range("AV6:AY6").AutoFill(Destination=ws.Range("AV6:AY800"))

    # 合計関数を最終行に入力
    aj_rowcell = ws.Range("AJ6").End(-4121)    # xlDown
    ap_rowcell = ws.Range("AP6").End(-4121)
    aj_rowcell.Formula = '=SUMIF($E:$E,"※銘柄計※",AJ:AJ)'
    ap_rowcell.Formula = '=SUMIF($E:$E,"※銘柄計※",AP:AP)'

    # 麦重量をリンクコピー
    ws.Range("AG3").Formula = f"={ap_rowcell.Address}"

    # 月末ロスを受払台帳から呼び出し
    from 受払台帳 import create_受払台帳
    loss_input = create_受払台帳()
    ws.Range("AY2").Value = loss_input

    ros = ws.Range("AY2").Value
    weight = ws.Range("AG3").Value

    # 調整した銘柄(加工完了1で利用)
    name_list = []

    # 条件付き処理
    if ros < weight:
        ws.Range("AZ6").FormulaR1C1 = '=IF(RC[-4]="",IF(RC[-1]>=0.3,IF(R2C51<R3C33,RC[-10]-1,""),""),"")'
        ws.Range("AZ6").AutoFill(Destination=ws.Range("AZ6:AZ800"))

        for i in range(6, 801):
            az_val = ws.Cells(i, 52).Value
            if az_val != "" and az_val is not None:
                ap_rowcell = ws.Cells(i, 42)
                ap_rowcell.Value -= 1
                ap_rowcell.Font.Color = 255
                msg = ws.Range("AW" + str(i - 1)).Value
                name_list.append(msg)

                print(f"{msg}の麦重量を-1しました。")

                weight = ws.Range("AG3").Value
                if ros == weight:
                    break

    # 印刷設定
    ws.Range("AV:AY").EntireColumn.Hidden = True    # 非表示

    last_row = ws.Cells(ws.Rows.Count, 42).End(-4162).Row                   # xlUp
    print_Area = ws.Range(ws.Cells(2, 2), ws.Cells(last_row, 47)).Address   # B2 : AU 最後行
    ws.PageSetup.PrintArea = print_Area                                     # 印刷範囲
    
    ws.PageSetup.CenterFooter = "&P/&N"             # 1/Nページ
    ws.PageSetup.PrintTitleRows = "$5:$5"           # タイトル行
    ws.PageSetup.Zoom = False                       # 拡大縮小印刷を無効化
    ws.PageSetup.FitToPagesWide = False             # 横は自動調整
    ws.PageSetup.FitToPagesTall = False             # 縦は自動調整
    #ws.PrintOut()
    #wb.Close(SaveChanges=True)  # 保存して閉じる
    print("--------------------加工完了4の処理が完了しました。")
    # 加工完了1で利用
    print(name_list)
    return name_list

if __name__ == "__main__":
    create_加工完了4()