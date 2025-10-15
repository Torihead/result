from docx import Document as dox
import shutil as sh
import datetime as dt
import win32com.client as w32

def create_証明依頼書():
    file_path = r"\\MC10\share\農政_電磁記録帳票\丸紅\日清\2021.09.29_契約大麦完了報告書ok\2021.09.29_3_証明依頼書.docx"

    # file_pathのファイルをコピー
    sh.copy(file_path, r"\\MC10\share\OA\EXCEL\OUT")                            # OUTフォルダにコピー
    new_file_path = r"\\MC10\share\OA\EXCEL\OUT\2021.09.29_3_証明依頼書.docx"     # コピー先のパス

    wb = dox(new_file_path)                                     # コピーしたファイルを開く

    today_str = dt.datetime.today().strftime("%Y年%m月%d日")
    from 加工完了1 import create_加工完了1                        # 加工完了1.pyからentries_valueを取得する関数をインポート
    entries_value = create_加工完了1()

    input_items = [("入庫数量", "入庫数量を入力してください(例 1,050,000): "),    # 入力項目のリスト
                ("加工予定", "加工予定の日付を入力してください(例 2025/5/22): "),
                ("加工数量", "加工数量を入力してください(例 1,050,000): ")]
    replace_dict = {"産地" : entries_value[0],                  # 辞書の初期値
                    "輸入許可番号" : entries_value[1],}

    for key, prompt in input_items:                           # replace_dict 辞書の入力
        replace_dict[key] = input(prompt)
    print(replace_dict)

    for para in wb.paragraphs:                                 # Word内のテキストを置き換え
        if "2022年1月26日" in para.text:
            para.text = para.text.replace("2022年1月26日", today_str)
        if "1,050,000" in para.text:
            para.text = para.text.replace("1,050,000", replace_dict["入庫数量"])

    for table in wb.tables:                                     # Word内のテーブルを置き換え
        for row in table.rows:
            for cell in row.cells:
                if "81177672910" in cell.text:
                    cell.text = cell.text.replace("81177672910", replace_dict["輸入許可番号"])
                if "豪州" in cell.text:
                    cell.text = cell.text.replace("豪州", replace_dict["産地"])
                if "1,050,000" in cell.text:
                    cell.text = cell.text.replace("1,050,000", replace_dict["入庫数量"])
                if "2021/10/19" in cell.text:
                    cell.text = cell.text.replace("2021/10/19", replace_dict["加工予定"])
                if "1,049,914" in cell.text:
                    cell.text = cell.text.replace("1,049,914", replace_dict["加工数量"])

    wb.save(new_file_path)
    word = w32.Dispatch("Word.Application")  # Wordアプリケーションを起動
    word.Visible = False
    doc = word.Documents.Open(new_file_path)
    #doc.PrintOut()
    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    create_証明依頼書()