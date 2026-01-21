"""
棚卸スクリプト用RDP/アプリ操作共通モジュール
RDP接続、アプリ起動、ログイン処理などを集約
"""

import subprocess
import pyautogui
import time


class RDPApp:
    """RDP経由でのアプリ操作を管理するクラス"""
    
    RDP_FILE = r"C:\Users\USER06\Desktop\OAシステム.rdp"
    DEFAULT_CLICK_X = 826
    DEFAULT_CLICK_Y = 448
    DEFAULT_LOGIN_ID = "12"
    DEFAULT_STARTUP_SLEEP = 3
    
    @staticmethod
    def launch_and_login(login_id=DEFAULT_LOGIN_ID, sleep_time=None):
        """
        RDP接続してアプリを起動し、ログイン処理を実行
        
        Args:
            login_id: ログインID（デフォルト："12"）
            sleep_time: RDP接続後の待機時間（秒）。Noneの場合DEFAULT_STARTUP_SLEEP使用
        """
        if sleep_time is None:
            sleep_time = RDPApp.DEFAULT_STARTUP_SLEEP
            
        subprocess.run(["mstsc.exe", RDPApp.RDP_FILE])
        pyautogui.sleep(sleep_time)
        
        # ログイン処理
        pyautogui.click(x=RDPApp.DEFAULT_CLICK_X, y=RDPApp.DEFAULT_CLICK_Y)
        pyautogui.write(login_id, interval=0.1)
        pyautogui.press("enter", presses=3, interval=0.5)
        time.sleep(2)
    
    @staticmethod
    def navigate_tabs(count):
        """Ctrl+PageUpで指定回数タブを遡る"""
        for _ in range(count):
            pyautogui.hotkey("ctrl", "pageup")
    
    @staticmethod
    def select_menu_item(menu_code):
        """メニュー番号を入力してアイテムを選択"""
        pyautogui.write(menu_code)
        pyautogui.press("enter")
        time.sleep(1)
    
    @staticmethod
    def press_multiple_tabs(count, interval=0.1):
        """複数回Tabキーを押す"""
        pyautogui.press("tab", presses=count, interval=interval)
    
    @staticmethod
    def print_excel(filepath, fit_to_one_page=False):
        """
        Excelファイルを開いて印刷（内部用Dispatchクラスを使用）
        
        Args:
            filepath: Excelファイルのパス
            fit_to_one_page: True の場合、1ページに縮小
        """
        import win32com.client as win32
        excel = win32.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(filepath)
        
        if fit_to_one_page:
            for sheet in wb.Sheets:
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
        
        wb.PrintOut()
        wb.Close(SaveChanges=False)
    
    @staticmethod
    def export_dialog_fill_path(path_text):
        """
        エクスポートダイアログで保存先を入力
        
        Args:
            path_text: 入力するパステキスト
        """
        import pyperclip
        pyperclip.copy(path_text)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("enter")
        time.sleep(1)
    
    @staticmethod
    def export_dialog_fill_filename(filename):
        """
        エクスポートダイアログでファイル名を入力
        
        Args:
            filename: 入力するファイル名
        """
        import pyperclip
        pyautogui.press("tab", presses=6, interval=0.2)
        pyperclip.copy(filename)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("enter", presses=2, interval=1)
        time.sleep(1)


class ExcelUtils:
    """Excel操作の共通ユーティリティ"""
    
    @staticmethod
    def get_excel_app(visible=False):
        """Excel アプリケーションを取得"""
        import win32com.client as win32
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = visible
        return excel
    
    @staticmethod
    def get_lastrow(worksheet, column=1):
        """
        ワークシートの最終行を取得
        
        Args:
            worksheet: Excel Worksheet オブジェクト
            column: 調べる列番号（デフォルト：1）
        
        Returns:
            int: 最終行番号
        """
        return worksheet.Cells(worksheet.Rows.Count, column).End(-4162).Row  # xlUp
    
    @staticmethod
    def safe_open_workbook(excel_app, filepath, read_only=False):
        """
        Workbookを安全に開く。読み取り専用フォールバック付き
        
        Args:
            excel_app: Excel.Application オブジェクト
            filepath: ファイルパス
            read_only: 読み取り専用で開くかどうか
        
        Returns:
            Workbook オブジェクト
        """
        try:
            return excel_app.Workbooks.Open(filepath, ReadOnly=read_only)
        except Exception as e:
            if not read_only:
                try:
                    return excel_app.Workbooks.Open(filepath, ReadOnly=True)
                except Exception as e2:
                    raise RuntimeError(
                        f"ファイルを開けませんでした: {filepath}\n"
                        f"通常: {e}\n"
                        f"読み取り専用: {e2}"
                    )
            else:
                raise
    
    @staticmethod
    def copy_cell_value(target_sheet, source_sheet, target_row, source_row, col_pairs):
        """
        複数のセル値をコピー
        
        Args:
            target_sheet: コピー先シート
            source_sheet: コピー元シート
            target_row: コピー先行
            source_row: コピー元行
            col_pairs: [(target_col, source_col), ...] のペアリスト
        """
        for target_col, source_col in col_pairs:
            target_sheet.Cells(target_row, target_col).Value = \
                source_sheet.Cells(source_row, source_col).Value


def close_excel(excel_app):
    """Excel アプリケーションを終了"""
    excel_app.Quit()
