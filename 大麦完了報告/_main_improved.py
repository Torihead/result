# ==========================================
# 改善版 _main.py
# 機能: 大麦完了報告書の自動作成と出力
# ==========================================

import os
import datetime as dt
import shutil as sh
import win32com.client as w32
import glob
from config import FILE_PATHS, PARTNERS
from utils import Logger, DateHelper, ExcelHelper, validate_date, validate_partner, get_validated_input

def get_user_inputs():
    """ユーザー入力を取得"""
    Logger.info("処理を開始します。")
    
    # 契約日の入力
    while True:
        contract = get_validated_input("契約日を入力してください(例 2025/4/1): ")
        if not contract:
            return None
        if validate_date(contract, "%Y/%m/%d"):
            dt_obj = DateHelper.parse_date(contract)
            formatted_date = DateHelper.format_date(dt_obj, "%Y.%m.%d")
            break
    
    # パートナー番号の入力
    while True:
        partner_input = get_validated_input(
            "取引先の番号を入力してください\n 1 工業会, 2 全畜連, 3 丸紅, 4 全農, 5 三井物産: "
        )
        if not partner_input:
            return None
        if validate_partner(partner_input):
            partner = int(partner_input)
            break
    
    filename = f"{formatted_date}_{partner}_"
    Logger.success(f"ファイル名: {filename}")
    
    return {
        "formatted_date": formatted_date,
        "partner": partner,
        "partner_name": PARTNERS[partner],
        "filename": filename
    }

def create_file_paths(filename):
    """ファイルパスを生成"""
    return {
        "before": [
            FILE_PATHS["gmo_dai"],
            FILE_PATHS["kako_hokoku1"],
            FILE_PATHS["kako_hokoku4"],
            FILE_PATHS["shomei"],
        ],
        "after": [
            os.path.join(FILE_PATHS["out_folder"], f"{filename}GMG_UKHRI_DAI.xlsx"),
            os.path.join(FILE_PATHS["out_folder"], f"{filename}KAKO_HOKOKU1.XLS"),
            os.path.join(FILE_PATHS["out_folder"], f"{filename}KAKO_HOKOKU4.XLS"),
            os.path.join(FILE_PATHS["out_folder"], f"{filename}証明依頼書.docx"),
        ]
    }

def close_all_excel():
    """すべてのExcelファイルを閉じる"""
    try:
        excel = w32.Dispatch("Excel.Application")
        while excel.Workbooks.Count > 0:
            excel.Workbooks(1).Close(SaveChanges=False)
        Logger.success("すべてのExcelファイルを閉じました")
    except:
        pass

def print_workbooks(file_list):
    """Excelファイルを印刷"""
    excel = w32.Dispatch("Excel.Application")
    excel.Visible = True
    
    try:
        for filepath in file_list:
            if os.path.exists(filepath):
                Logger.info(f"印刷中: {os.path.basename(filepath)}")
                wb = excel.Workbooks.Open(filepath)
                wb.PrintOut()
                wb.Close(SaveChanges=True)
                Logger.success(f"印刷完了: {os.path.basename(filepath)}")
            else:
                Logger.warning(f"ファイルが見つかりません: {filepath}")
    except Exception as e:
        Logger.error(f"印刷に失敗: {e}")
    finally:
        excel.Quit()

def rename_files(file_pairs):
    """ファイルをリネーム"""
    for before, after in file_pairs:
        try:
            if os.path.exists(before):
                os.rename(before, after)
                Logger.success(f"リネーム完了: {os.path.basename(after)}")
            else:
                Logger.warning(f"ファイルが見つかりません: {before}")
        except Exception as e:
            Logger.error(f"リネーム失敗 {before}: {e}")

def copy_to_archive(file_list, archive_folder):
    """完了報告書を電磁記録フォルダにコピー"""
    try:
        for filepath in file_list:
            if os.path.exists(filepath):
                sh.copy(filepath, archive_folder)
                Logger.success(f"コピー完了: {os.path.basename(filepath)}")
            else:
                Logger.warning(f"ファイルが見つかりません: {filepath}")
    except Exception as e:
        Logger.error(f"コピー失敗: {e}")

def cleanup_old_files():
    """OUTフォルダの古いリネーム済みファイルを削除"""
    try:
        out_folder = FILE_PATHS["out_folder"]
        # 日付形式でリネームされたファイルを検出（20XX.XX.XX_X_*.XLS/XLSX）
        pattern = os.path.join(out_folder, "20*.XLS*")
        old_files = glob.glob(pattern)
        
        if not old_files:
            Logger.info("削除対象のファイルはありません")
            return
        
        for old_file in old_files:
            try:
                os.remove(old_file)
                Logger.success(f"削除: {os.path.basename(old_file)}")
            except Exception as e:
                Logger.warning(f"削除失敗 {os.path.basename(old_file)}: {e}")
    except Exception as e:
        Logger.warning(f"ファイルクリーンアップエラー: {e}")

def main():
    """メイン処理"""
    try:
        Logger.info("--------------------処理を開始します。")
        
        # ステップ0: OUTフォルダの古いファイルをクリーンアップ
        Logger.info("【ステップ0】OUTフォルダの古いファイルをクリーンアップしています...")
        cleanup_old_files()
        
        # ステップ1: 証明依頼書の作成（最初に実行 - この中で加工完了1,4, 受払台帳が呼ばれる）
        Logger.info("【ステップ1】証明依頼書を作成しています...")
        from 証明依頼書 import create_証明依頼書
        create_証明依頼書()
        
        # ステップ2: ユーザー入力を取得
        Logger.info("【ステップ2】ユーザー入力を取得しています...")
        user_data = get_user_inputs()
        if not user_data:
            Logger.error("処理をキャンセルしました。")
            return
        
        # ステップ3: ファイルパスを生成
        file_paths = create_file_paths(user_data["filename"])
        
        # ステップ4: すべてのExcelを閉じる
        Logger.info("【ステップ3】Excelファイルを閉じています...")
        close_all_excel()
        
        # ステップ5: 完了報告書を印刷（リネーム前の元のファイルを印刷）
        Logger.info("【ステップ4】完了報告書を印刷しています...")
        #print_workbooks(file_paths["before"][:3])  # 元のファイル名で印刷
        
        # ステップ6: ファイルをリネーム
        Logger.info("【ステップ5】ファイルをリネームしています...")
        rename_files(zip(file_paths["before"], file_paths["after"]))
        
        # ステップ7: 電磁記録フォルダを開く
        Logger.info("【ステップ6】電磁記録フォルダを開いています...")
        os.startfile(FILE_PATHS["archive_folder"])
        
        # ステップ8: 完了報告書をコピー（リネーム後のファイルをコピー）
        Logger.info("【ステップ7】完了報告書をアーカイブフォルダにコピーしています...")
        copy_to_archive(file_paths["after"], FILE_PATHS["archive_folder"])
        
        Logger.success("--------------------報告書の作成を終了しました。")
        Logger.success("すべての処理が完了しました。")
        
    except Exception as e:
        Logger.error(f"予期しないエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
