"""
設定ファイル - パス、UI座標、待機時間などを一元管理
"""
from pathlib import Path

# ===== ファイルパス =====
CSV_FILE = r"C:\Users\USER06\Desktop\Auto_code\16時印刷\出勤日カレンダー.csv"
RDP_FILE = r"C:\Users\USER06\Desktop\OAシステム.rdp"

# ===== ネットワークパス =====
SHARE_BASE = r"\\MC10\share"
KISO_IRAI_FILE = rf"{SHARE_BASE}\MICHINOK_共有\（仮）\基礎依頼\基礎依頼票.xlsm"
KISO_IRAI_BACKUP = rf"{SHARE_BASE}\MICHINOK_共有\（仮）\基礎依頼\基礎依頼票_BaukUp"
PELLET_BASE = rf"{SHARE_BASE}\MICHINOK_共有\0.共有書類\ペレット予測"
TB_FILE = rf"{SHARE_BASE}\MICHINOK_共有\0.共有書類\対照帳\端量　在庫表\TB保管出荷.xlsx"
HARYOU_FILE = rf"{SHARE_BASE}\MICHINOK_共有\0.共有書類\対照帳\端量　在庫表\翌営業日製造端量表（新）.xlsm"
OUT_FILE = rf"{SHARE_BASE}\OA\EXCEL\OUT\12_SEIHIN_NSK_YOTEI.XLS"

# ===== UI座標（pyautogui） =====
# OA起動処理
OA_CLICK_X = 826
OA_CLICK_Y = 448

# 受付業務タグ関連
UKETSUKEGIMU_X = 700
UKETSUKEGIMU_Y = 302

# 週間受注関連
SHUKAN_JUSCHU_X = 641
SHUKAN_JUSCHU_Y = 381

# 製品出荷予定表
SEHIN_SHOKKA_X = 618
SEHIN_SHOKKA_Y = 451

# 倉庫移動
SOUKO_IDOU_X = 830
SOUKO_IDOU_Y = 414

# 製造入出庫予定照会
SEIZO_INYUSHUKO_X = 761
SEIZO_INYUSHUKO_Y = 448

# ===== 待機時間（秒） =====
WAIT_OA_START = 4
WAIT_SHORT = 0.5
WAIT_MEDIUM = 1
WAIT_LONG = 2
WAIT_XLWAIT = 3
WAIT_EXCEL_REFRESH = 10
WAIT_FINAL_SLEEP = 15

# ===== その他設定 =====
EXCEL_CLOSE_DELAY = 15  # 最後に Excel を閉じるまでの待機時間
