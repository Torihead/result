"""
棚卸スクリプト共通ユーティリティモジュール
日付計算、月計算などの共通処理を集約
"""

import datetime as dt
import jpholiday


def get_date_info(reference_date=None):
    """
    日付情報を取得して、処理に必要な全ての日付フォーマットを返す
    
    Args:
        reference_date: 基準日（デフォルト：本日）
    
    Returns:
        dict: 以下のキーを含む辞書
            - today: 本日の日付（date）
            - last_month: 先月（date）
            - tenth: 今月10日（営業日調整済み）（date）
            - formatted_month: "YYYY.MM" 形式（先月）
            - formatted_year: "YYYY" 形式（年度）
            - format_month: "YYYYMM" 形式（先月）
            - this_month: "YYYY.MM" 形式（今月）
            - lastmonth_start: 先月初日（date）
            - lastmonth_end: 先月最終日（date）
            - thismonth_start: 今月初日（date）
    """
    today = reference_date or dt.date.today()
    year = today.year
    month = today.month

    # 今月の10日を営業日に調整
    tenth = dt.date(year, month, 10)
    while tenth.weekday() >= 5 or jpholiday.is_holiday(tenth):
        tenth -= dt.timedelta(days=1)

    # 先月の計算
    last_month = today - dt.timedelta(days=20)
    
    # 年度計算（4月始まり）
    if month >= 4:
        fiscal_year = year
    else:
        fiscal_year = year - 1

    return {
        'today': today,
        'last_month': last_month,
        'tenth': tenth,
        'formatted_month': last_month.strftime("%Y.%m"),
        'formatted_year': str(fiscal_year),
        'format_month': last_month.strftime("%Y%m"),
        'this_month': today.strftime("%Y.%m"),
        'format_year': last_month.strftime("%y"),
        'lastmonth_start': get_lastmonth_start(today),
        'lastmonth_end': get_lastmonth_end(today),
        'thismonth_start': dt.date(year=today.year, month=today.month, day=1),
    }


def get_lastmonth_start(reference_date=None):
    """先月初日を取得"""
    today = reference_date or dt.date.today()
    if today.month == 1:
        return dt.date(year=today.year - 1, month=12, day=1)
    else:
        return dt.date(year=today.year, month=today.month - 1, day=1)


def get_lastmonth_end(reference_date=None):
    """先月最終日を取得"""
    today = reference_date or dt.date.today()
    thismonth_start = dt.date(month=today.month, year=today.year, day=1)
    return thismonth_start - dt.timedelta(days=1)


def format_date(date_obj, fmt):
    """
    日付をフォーマットして返す（ショートカット）
    
    Args:
        date_obj: date オブジェクト
        fmt: strftime形式のフォーマット文字列
    
    Returns:
        str: フォーマット済みの日付文字列
    """
    return date_obj.strftime(fmt)
