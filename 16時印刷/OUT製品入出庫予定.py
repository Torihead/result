import time
import get_next_workday as gnw
import win32com.client as w32
import pyautogui as pyA

pyA.click(x=700, y=302)                             # 受付業務タグをクリック
time.sleep(0.5)
pyA.click(x=761, y=448)                             # 製造入出庫予定照会をクリック
time.sleep(1)
pyA.press("F5")                                     # 表示ボタン
time.sleep(0.7)
pyA.press("Tab", presses=2, interval=0.1)
pyA.press("Enter")                                  # OUT出力
time.sleep(1)

def main():
    excel = w32.Dispatch("Excel.Application")
    excel.Visible = False
    file_path = r"\\MC10\share\OA\EXCEL\OUT\12_SEIHIN_NSK_YOTEI.XLS"
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Worksheets("製品入出庫定照会")
    time.sleep(1)

    # 普段は、３列を削除
    # 次の日が休みなら、６列を削除する 
    if gnw.next_workday.date() == gnw.today.date() + gnw.dt.timedelta(days=1):
        ws.Columns["F:H"].Delete()
    else:
        ws.Columns["F:K"].Delete()

    #フィルター設定
    ws.Range("D8").AutoFilter(Field:=4, Criteria1:="<>BB", Operator:=1)
    ws.Range("D8").AutoFilter(Field:=5, Criteria1:="<>臨港倉庫", Operator:=1)
    ws.Range("G8").AutoFilter(Field:=7, Criteria1:="<>0", Operator:=1)

    #レイアウト、印刷
    xlup = -4162                                            # セルの最終行
    last_row = ws.Cells(ws.Rows.Count, 12).End(xlup).Row    # L列の最終行を取得

    ws.PageSetup.PrintArea = f"A7:L{last_row}"              # 印刷範囲を選択

    ps = ws.PageSetup
    ps.Orientation      = 2                                 # 横方向に印刷
    ps.Zoom             = False                             # Zoom(%)無効
    ps.FitToPagesWide   = 1                                 # ヨコを１ページに
    ps.FitToPagesTall   = 1                                 # タテを１ページに

    ws.PrintOut()                                   # コピー部数
    time.sleep(1)
    ws.PrintOut()

    wb.Close(SaveChanges=True)
    excel.Quit()

if __name__ == "__main__":
    main()