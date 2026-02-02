import time
import win32com.client as w32

# リストの中身が合っているか確認
def equal_check(val1, val2):
    if val1 == val2:
        print(f"{val1} == {val2} はTrue")
    else:
        print(f"{val1} == {val2} はFalse")

def create_加工完了1():
    global entries_value    # entries_valueをグローバル変数として宣言

    excel = w32.Dispatch("Excel.Application")
    excel.Visible = True
    filepath = r"\\MC10\share\OA\EXCEL\OUT\KAKO_HOKOKU1.XLS"
    wb = excel.Workbooks.Open(filepath)
    time.sleep(2)

    ws1 = wb.Worksheets("1")
    ws2 = wb.Worksheets("2")
    
    try:
        ws3 = wb.Worksheets("3")
    except Exception:
        ws3 = None

    if ws1.Range("J25").Value != ws1.Range("N25").Value:    # J25とN25の値が等しくない場合
        print(
            "\n加工済みの麦重量が0ではありません。\n"
            "大麦の「受入状況 or 割当入力」を確認してください。"
        )
        print("指定されたマクロはすべてキャンセルになります。")
        return                                              # 処理を中断して終了

    if ws3:
        excel.Application.StatusBar = "sheet3があります。処理を開始します。"
    else:
        excel.Application.StatusBar = "sheet3はありません。処理を省略します"

    # AS4:AY4日付を消す
    ws1.Range("AS4:AY4").ClearContents()
    ws2.Range("AS4:AY4").ClearContents()
    if ws3:
        ws3.Range("AS4:AY4").ClearContents()

    # B17:AX18結合解除
    ws1.Range("B17:AX18").UnMerge()

    # 船とロット情報の欄を作成
    def format_and_label(rng, label):               # フォーマット関数(別名: フォーマッター)
        rng.Merge(True)                 # セル結合
        rng.Interior.ColorIndex = 6     # 黄色
        rng.HorizontalAlignment = -4108 # xlCenter
        rng.Borders.LineStyle   = 1     # xlContinuous  中央揃え
        rng.Value               = label # 関数利用の引数の値を採用

    format_and_label(ws1.Range("Z17:AE18"), "産地")             # (rng, label)の引数
    format_and_label(ws1.Range("AF17:AL18"), "輸入許可番号")     # 船とロット情報の欄を作成
    format_and_label(ws1.Range("AM17:AQ18"), "輸入許可日")
    format_and_label(ws1.Range("AR17:AY18"), "船名")

    # entriesリストを使って繰り替えし処理
    entries = [
        ("Z18", "産地を入力してください: "),
        ("AF18", "輸入許可番号を入力してください: "),
        ("AM18", "輸入許可日を入力してください: "),
        ("AR18", "船名を入力してください: "),
    ]
    entries_value = []                              # entriesリストの値のみを格納するリスト(証明依頼書で使う)
    
    for cell, prompt in entries:                    # cell, promptにentriesリストの値
        Val = input(prompt)    # Val で入力する値を取得
        if Val == "" or Val is None:                # キャンセルされた場合
            print("入力がキャンセルされました。")
            return
        ws1.Range(cell).Value = Val
        entries_value.append(Val)                   # entries_valueリストに追加

    # 小計の入力
    for ws in (ws1, ws2):
        ws.Range("AF75").Formula = "=SUM(AF25:AF74)"
        ws.Range("AM75").Formula = "=SUM(AM25:AM74)"
    if ws3:
        ws3.Range("AF75").Formula = "=SUM(AF25:AF74)"
        ws3.Range("AM75").Formula = "=SUM(AM25:AM74)"
    # 総計の入力
    if ws3:
        ws3.Range("AF77").Formula = "=SUM('1'!AF75, '2'!AF75, '3'!AF75)"
        ws3.Range("AM77").Formula = "=SUM('1'!AM75, '2'!AM75, '3'!AM75)"
    else:
        ws2.Range("AF77").Formula = "=SUM('1'!AF75, '2'!AF75)"
        ws2.Range("AM77").Formula = "=SUM('1'!AM75, '2'!AM75)"

    # 重量を調整した銘柄から-1
    from 加工完了4 import create_加工完了4                # 完了４のreturnの値を代入
    loss_list = create_加工完了4()

    for i in range(25, 75):                             # sheet1の処理
        cell_value = ws1.Range(f"Z{i}").Value           # Z列の値を繰り返し処理で上から順に判別
        if cell_value in loss_list:                     # リストの中身と合致したら、条件開始
            am_cell = ws1.Range(f"AM{i}")               # Z列の銘柄名のAM列の重量を取得
            print(f"{cell_value}の{am_cell}を調整します")
            if isinstance(am_cell.Value, (int, float)): # 整数、浮動小数点 型ならTRUE
                am_cell.Value -= 1
                am_cell.Font.Color = 255                # 値を-1、文字を赤色
                print(f"{am_cell}に調整しました")
            else:
                am_cell.Value = -1
                am_cell.Font.Color = 255
                print(f"{am_cell}に調整しました")
    for i in range(25, 75):                             # sheet2の処理(sheet1同様)
        cell_value = ws2.Range(f"Z{i}").Value
        if cell_value in loss_list:
            am_cell = ws2.Range(f"AM{i}")
            print(f"{cell_value}の{am_cell}を調整します")
            if isinstance(am_cell.Value, (int, float)):
                am_cell.Value -= 1
                am_cell.Font.Color = 255
                print(f"{am_cell}に調整しました")
            else:
                am_cell.Value = -1
                am_cell.Font.Color = 255
                print(f"{am_cell}に調整しました")

    if any(sheet.Name == "3" for sheet in wb.Sheets):       # sheet3がある場合、sheet3の処理を実行
        for i in range(25, 75):                             # sheet3の処理(sheet1同様)
            cell_value = ws3.Range(f"Z{i}").Value
            if cell_value in loss_list:
                am_cell = ws3.Range(f"AM{i}")
                print(f"{cell_value}の{am_cell}を調整します")
                if isinstance(am_cell.Value, (int, float)):
                    am_cell.Value -= 1
                    am_cell.Font.Color = 255
                    print(f"{am_cell}に調整しました")
                else:
                    am_cell.Value = -1
                    am_cell.Font.Color = 255
                    print(f"{am_cell}に調整しました")

    # 月末ロスの調整
    loss = input("欠減の値を入力してください: ")
    ws1.Range("AL5").Value = loss
    ws1.Range("AQ25").Value = "kg" + "\n" + str(loss)
    ws1.Range("AL5").Value = ""       # 入力セルをクリア

    # 印刷設定
    print_list = ["1", "2"]
    if any(sheet.Name == "3" for sheet in wb.Sheets):       # if ws3: と同じ意味
        print_list.append("3")

    for sheet_name in print_list:
        ws = wb.Worksheets(sheet_name)      #sheetのグループ化
        ws.PageSetup.TopMargin = excel.Application.InchesToPoints(0.25)         # 余白を狭い
        ws.PageSetup.BottomMargin = excel.Application.InchesToPoints(0.25)
        ws.PageSetup.LeftMargin = excel.Application.InchesToPoints(0.25)
        ws.PageSetup.RightMargin = excel.Application.InchesToPoints(0.25)
    #wb.Worksheets(print_list).PrintOut()
    #wb.Close(SaveChanges=True)  # 保存して閉じる
    #excel.Application.Quit()     # Excelアプリケーションを終了
    print("--------------------加工完了1の処理が完了しました。")
    return entries_value

if __name__ == "__main__":
    create_加工完了1()
