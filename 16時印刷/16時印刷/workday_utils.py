"""
営業日ユーティリティ - CSV から営業日を取得する共通関数
"""
import csv
from datetime import datetime, date
from typing import List, Optional
import config


def get_next_working_day(csv_file: str = config.CSV_FILE) -> Optional[date]:
    """
    次の営業日を取得する
    
    Args:
        csv_file: CSVファイルパス
    
    Returns:
        次の営業日（存在しない場合は None）
    """
    try:
        today = datetime.now().date()
        next_working = None

        with open(csv_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get('working') == '1':
                    try:
                        d = datetime.strptime(row['date'], "%Y/%m/%d").date()
                        if d > today and (next_working is None or d < next_working):
                            next_working = d
                    except ValueError:
                        # フォーマット不正行はスキップ
                        continue

        return next_working

    except FileNotFoundError:
        raise FileNotFoundError(f"CSVファイルが見つかりません: {csv_file}")


def get_next_n_working_days(n: int = 2, csv_file: str = config.CSV_FILE) -> List[date]:
    """
    次の n 営業日を取得する
    
    Args:
        n: 取得する営業日数
        csv_file: CSVファイルパス
    
    Returns:
        営業日のリスト（昇順）
    """
    try:
        today = datetime.now().date()
        working_days = []

        with open(csv_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get('working') == '1':
                    try:
                        d = datetime.strptime(row['date'], "%Y/%m/%d").date()
                        if d > today:
                            working_days.append(d)
                    except ValueError:
                        continue

        working_days.sort()
        return working_days[:n]

    except FileNotFoundError:
        raise FileNotFoundError(f"CSVファイルが見つかりません: {csv_file}")
