# ==========================================
# ユーティリティ関数集 - utils.py
# ==========================================

import win32com.client as w32
import time
import os
import datetime as dt
from config import FILE_PATHS, LOG_FORMAT

class Logger:
    """ロギングクラス"""
    @staticmethod
    def info(message):
        print(f"[INFO] {message}")
    
    @staticmethod
    def success(message):
        print(f"[✓] {message}")
    
    @staticmethod
    def error(message):
        print(f"[✗] {message}")
    
    @staticmethod
    def warning(message):
        print(f"[⚠] {message}")

class ExcelHelper:
    """Excel操作の共通処理"""
    def __init__(self):
        self.excel = w32.Dispatch("Excel.Application")
        self.excel.Visible = True
        self.workbooks = []
    
    def open_workbook(self, filepath, wait_seconds=2):
        """Excelファイルを開く"""
        try:
            wb = self.excel.Workbooks.Open(filepath)
            self.workbooks.append(wb)
            time.sleep(wait_seconds)
            Logger.success(f"Excelファイルを開きました: {filepath}")
            return wb
        except Exception as e:
            Logger.error(f"ファイルを開けません: {filepath}\n詳細: {e}")
            raise
    
    def close_all(self, save=False):
        """すべてのワークブックを閉じる"""
        try:
            for wb in self.workbooks:
                wb.Close(SaveChanges=save)
            self.excel.Quit()
            Logger.success("Excelを閉じました")
        except Exception as e:
            Logger.error(f"ファイルを閉じられません: {e}")
    
    @staticmethod
    def set_freeze_pane(worksheet, row):
        """行でウィンドウ枠を固定"""
        try:
            worksheet.Range(f"A{row}").Select()
            w32.Dispatch("Excel.Application").ActiveWindow.FreezePanes = True
            Logger.success(f"行{row}でウィンドウ枠を固定しました")
        except Exception as e:
            Logger.error(f"ウィンドウ枠の固定に失敗: {e}")
    
    @staticmethod
    def merge_cells(rng, value, color_index=6, alignment=-4108):
        """セルを結合してフォーマット"""
        try:
            rng.Merge(True)
            rng.Interior.ColorIndex = color_index
            rng.HorizontalAlignment = alignment
            rng.Borders.LineStyle = 1
            rng.Value = value
        except Exception as e:
            Logger.error(f"セル結合に失敗: {e}")

class WordHelper:
    """Word操作の共通処理"""
    @staticmethod
    def replace_in_paragraphs(document, old_text, new_text):
        """段落内のテキストを置き換え"""
        count = 0
        for para in document.paragraphs:
            if old_text in para.text:
                para.text = para.text.replace(old_text, new_text)
                count += 1
        return count
    
    @staticmethod
    def replace_in_tables(document, old_text, new_text):
        """テーブル内のテキストを置き換え"""
        count = 0
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    if old_text in cell.text:
                        cell.text = cell.text.replace(old_text, new_text)
                        count += 1
        return count

class DateHelper:
    """日付処理の共通処理"""
    @staticmethod
    def parse_date(date_str, input_format="%Y/%m/%d"):
        """日付文字列をパース"""
        try:
            dt_obj = dt.datetime.strptime(date_str, input_format)
            return dt_obj
        except ValueError:
            Logger.error(f"日付の形式が正しくありません: {date_str}")
            return None
    
    @staticmethod
    def format_date(dt_obj, output_format="%Y.%m.%d"):
        """日付をフォーマット"""
        return dt_obj.strftime(output_format)

def get_validated_input(prompt, validation_func=None):
    """入力値を取得して検証"""
    while True:
        try:
            value = input(prompt)
            if validation_func:
                if validation_func(value):
                    return value
            else:
                if value and value is not None:
                    return value
            Logger.warning("無効な入力です。もう一度入力してください。")
        except KeyboardInterrupt:
            Logger.error("入力がキャンセルされました。")
            return None

def validate_date(date_str, date_format="%Y/%m/%d"):
    """日付入力を検証"""
    try:
        dt.datetime.strptime(date_str, date_format)
        return True
    except ValueError:
        Logger.error(f"日付の形式が正しくありません。{date_format}で入力してください。")
        return False

def validate_partner(partner_str):
    """パートナー番号を検証"""
    try:
        partner = int(partner_str)
        if partner in range(1, 6):
            return True
        else:
            Logger.error("1-5の数字を入力してください。")
            return False
    except ValueError:
        Logger.error("無効な入力です。1-5の数字を入力してください。")
        return False
