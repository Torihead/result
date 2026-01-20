# ==========================================
# 共通設定ファイル - config.py
# ==========================================

# ファイルパス設定
FILE_PATHS = {
    "gmo_dai": r"\\MC10\share\OA\EXCEL\OUT\GMG_UKHRI_DAI.xlsx",
    "kako_hokoku1": r"\\MC10\share\OA\EXCEL\OUT\KAKO_HOKOKU1.XLS",
    "kako_hokoku4": r"\\MC10\share\OA\EXCEL\OUT\KAKO_HOKOKU4.XLS",
    "shomei": r"\\MC10\share\OA\EXCEL\OUT\2021.09.29_3_証明依頼書.docx",
    "shomei_template": r"\\MC10\share\農政_電磁記録帳票\丸紅\日清\2021.09.29_契約大麦完了報告書ok\2021.09.29_3_証明依頼書.docx",
    "out_folder": r"\\MC10\share\OA\EXCEL\OUT",
    "archive_folder": r"\\MC10\share\農政_電磁記録帳票",
}

# Excel シート設定
SHEET_NAMES = {
    "kako_hokoku1": ["1", "2", "3"],
    "kako_hokoku4": "レイアウト",
    "gmo_dai": "大麦",
}

# セル設定
CELL_RANGES = {
    "kako_hokoku1": {
        "j25": "J25",
        "n25": "N25",
        "as4_ay4": "AS4:AY4",
        "b17_ax18": "B17:AX18",
        "entries": [
            ("Z18", "産地を入力してください: "),
            ("AF18", "輸入許可番号を入力してください: "),
            ("AM18", "輸入許可日を入力してください: "),
            ("AR18", "船名を入力してください: "),
        ],
    }
}

# 置換文字列辞書
REPLACE_DICT_BASE = {
    "date_format": "2022年1月26日",
    "quantity": "1,050,000",
    "license_num": "81177672910",
    "product": "豪州",
    "schedule": "2021/10/19",
    "scheduled_quantity": "1,049,914",
}

# パートナー情報
PARTNERS = {
    1: "工業会",
    2: "全畜連",
    3: "丸紅",
    4: "全農",
    5: "三井物産",
}

# ロギング設定
LOG_FORMAT = "[{level}] {message}"
