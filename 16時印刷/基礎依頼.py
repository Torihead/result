import win32com.client as w32
import get_next_workday as gnw                          # get_next_workday.py 関数呼び出し
import time

def main():
    # 新ファイルの名前変更のためのコード (YYYY.MM.DD 形式)
    formatted_date = gnw.next_workday.strftime("%Y.%m.%d")
    new_file_name = f"基礎依頼票_{formatted_date}.xlsm"
    new_file_path = fr"\\MC10\share\MICHINOK_共有\（仮）\基礎依頼\基礎依頼票_BaukUp\{new_file_name}"

    excel = w32.Dispatch("Excel.Application")               # Excelアプリ起動
    excel.Visible = False                                    # バックグラウンド処理

    file_path = r"\\MC10\share\MICHINOK_共有\（仮）\基礎依頼\基礎依頼票.xlsm"
    workbook = excel.Workbooks.Open(file_path)              # 元ファイルを取得
    time.sleep(3)

    worksheet = workbook.Worksheets("製造基礎依頼予定表")
    time.sleep(1)
    workbook.RefreshAll()                                     # 全体更新
    time.sleep(10)

    # 基礎依頼の保存
    workbook.SaveAs(new_file_path)                           # SaveAs で新しい名前で保存
    print(f"保存した基礎依頼ファイル ➡  {new_file_path}")
    time.sleep(2)

    # 元ファイルのフィルター操作と印刷
    # セルの中身は関数でフィルターできないため、値貼り付けしてから、フィルターで "昼" を除外する
    worksheet.Range("E5:E50").Copy()
    worksheet.Range("E5:E50").PasteSpecial(Paste=-4163)     # VBAのxlPasteValuesと同じ意味
    worksheet.Range("E5").AutoFilter(Field:=4, Criteria1:="<>昼", Operator:=1)# Field は、インデックス番号(左からの何番目にあるか)

    worksheet.PrintOut()
    time.sleep(3)
    workbook.Close(SaveChanges=False)                       # 元ファイルを閉じる
    excel.Quit()
    time.sleep(2)
if __name__ == "__main__":
    main()