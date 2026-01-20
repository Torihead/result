import time
import win32com.client as w32
import datetime as dt
import os

def main():

    excel = w32.Dispatch("Excel.Application")
    excel.Visible = True

    # 今日を取得
    today = dt.date.today()
    year = today.year
    month = today.month
    
    # 前月を取得（年をまたぐ場合に対応）
    last_month = month - 1
    last_year = year
    if last_month < 1:
        last_month = 12
        last_year = year - 1

    # 年度を取得（4月が年度開始）
    if month < 4:
        fiscal_year = year - 1
    else:
        fiscal_year = year
    

    # ファイル名のパーツ作成
    year_parts = f"{fiscal_year}年度"
    month_parts = f"{last_year}.{last_month:02d}月"          # 前月の年と月を使用
    print(f"{year_parts}+{month_parts} の在庫証明を作成します。")

    # 日和の在庫証明
    file_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{year_parts}\{month_parts}\06　親会社月次報告\1　日和\原料在庫証明.xls"
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Sheets(1)
    time.sleep(1)

    ws.Range("A6").Value = f"{last_year}/{last_month}/1"
    wb.RefreshAll()
    time.sleep(3)

    output_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{year_parts}\{month_parts}\06　親会社月次報告\1　日和\原料在庫証明.pdf"
    excel.ActiveSheet.ExportAsFixedFormat(0, output_path)

    wb.Close(SaveChanges=True)

    print("日和 在庫証明の作成が終了しました。")

    # 日清の在庫証明
    file_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{year_parts}\{month_parts}\06　親会社月次報告\3　日清\在庫証明.xls"
    wb = excel.Workbooks.Open(file_path)
    time.sleep(1)

    ws_genryou = wb.Worksheets("日清（原料）")
    ws_seihin = wb.Worksheets("日清（製品）")
    ws_kami = wb.Worksheets("日清（紙袋）")

    ws_genryou.Select()                                         # sheetグループ解除
    ws_seihin.AutoFilterMode = False                            # フィルター解除
    ws_kami.AutoFilterMode = False
    ws_genryou.Range("A6").Value = f"{last_year}/{last_month}/1"     # 日付入力

    # データ更新
    wb.RefreshAll()
    time.sleep(3)

    ws_seihin.Range("B11:G11").AutoFilter(Filter:=3, Criteria1:="<>#N/A", Operator:=1)  # フィルター設定
    ws_kami.Range("A11:F11").AutoFilter(Filter:=1, Criteria1:="<>#N/A", Operator:=1)

    # 3つのsheetをグループ指定して、PDF出力
    export_list = ["日清（原料）", "日清（製品）", "日清（紙袋）"]
    sheet_object = [wb.Sheets(sheet_name) for sheet_name in export_list]
    wb.Sheets(export_list).Select()

    output_path = fr"\\MC10\share\MICHINOK_共有\2.小島\終了報告書＆棚卸データ\{year_parts}\{month_parts}\06　親会社月次報告\3　日清\在庫証明.pdf"
    excel.ActiveSheet.ExportAsFixedFormat(0, output_path)

    #wb.Close(SaveChanges=True)
    #excel.Quit()

    print("日清 在庫証明の作成が終了しました。")
    print("--------------------在庫証明の作成が完了しました。")

if __name__ == "__main__":
    main()