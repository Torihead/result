# 業務自動化スクリプト集

[![Python](https://img.shields.io/badge/Python-3.13-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Windows](https://img.shields.io/badge/Windows-Automation-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Word](https://img.shields.io/badge/Microsoft-Word-2B579A?logo=microsoftword&logoColor=white)](https://www.microsoft.com/word)

> 製造業における定型業務を自動化するPythonスクリプト集

手作業で行っていた日次・月次の帳票作成・印刷業務を自動化し、**作業時間の大幅削減**と**人的ミスの防止**を実現しました。

---

## 📊 プロジェクト一覧

| プロジェクト | 概要 | 自動化内容 |
|-------------|------|-----------|
| [16時印刷](#1-16時印刷) | 日次帳票の一括印刷 | 9種類の帳票を順次自動印刷 |
| [大麦完了報告](#2-大麦完了報告) | 完了報告書の自動生成 | Excel/Word編集、ファイル管理 |
| [棚卸](#3-棚卸) | 月次棚卸業務の自動化 | 6ステップの帳票作成を自動化 |

---

## 🛠️ 技術スタック

| カテゴリ | 技術 |
|---------|------|
| 言語 | Python 3.13 |
| GUI自動化 | pyautogui |
| Office操作 | win32com.client, python-docx |
| 日付処理 | jpholiday（祝日判定） |
| RDP連携 | mstsc.exe（リモートデスクトップ） |

---

## 📁 プロジェクト詳細

### 1. 16時印刷

毎日16時に実行する**9種類の帳票を自動印刷**するシステム。

#### 処理フロー

```
1. 基礎依頼        → Excelマクロ実行・印刷
2. ペレット予定    → 翌営業日のファイルを自動選択・印刷
3. TB保管出荷      → データ更新・印刷
4. 翌営業日端量表  → マクロ実行・印刷
5. OA起動          → RDP接続・ログイン
6. OA週間受注印刷  → 帳票出力
7. OA出荷予定印刷  → 帳票出力
8. OA倉庫移動印刷  → 帳票出力
9. 製品入出庫予定  → 帳票出力・Excel終了
```

#### アーキテクチャ

```
16時印刷/
├── main.py          # メインスクリプト（タスク実行管理）
├── config.py        # 設定ファイル（パス、座標、待機時間）
├── print_utils.py   # 印刷ユーティリティ
├── workday_utils.py # 営業日計算
├── system/
│   └── oa_startup.py    # RDP接続・ログイン処理
└── tasks/
    ├── kiso_irai.py     # 基礎依頼
    ├── pellet.py        # ペレット予定
    ├── tb_hokan.py      # TB保管出荷
    ├── haryou.py        # 端量表
    ├── oa_shukan.py     # 週間受注
    ├── oa_shokka.py     # 出荷予定
    ├── oa_souko.py      # 倉庫移動
    └── out_sehin.py     # 製品入出庫
```

#### 特徴

- **エラーハンドリング**: 各タスクの成功/失敗を記録し、サマリーを表示
- **営業日判定**: 祝日・休日を考慮した翌営業日の自動計算
- **設定の一元管理**: ファイルパス、UI座標、待機時間を`config.py`で管理

---

### 2. 大麦完了報告

大麦加工の完了報告書を**自動生成・編集**するシステム。

#### 処理フロー

```
Step 0: OUTフォルダのクリーンアップ
    ↓
Step 1: 証明依頼書の作成
    ├─ 加工完了4.py → 月末ロス計算
    ├─ 加工完了1.py → 産地・輸入許可情報
    ├─ 受払台帳.py  → ロス情報取得
    └─ Word文書編集・保存
    ↓
Step 2: ユーザー入力（契約日・取引先）
    ↓
Step 3: Excelファイルを閉じる
    ↓
Step 4: 完了報告書を印刷
    ↓
Step 5: ファイルをリネーム（日付_取引先番号_ファイル名）
    ↓
Step 6: 電磁記録フォルダへコピー
```

#### 特徴

- **Word/Excel連携**: win32com.clientによる自動編集
- **ファイル命名規則**: `YYYY.MM.DD_取引先番号_ファイル名` で統一管理
- **アーカイブ機能**: 処理完了後、自動で保存先フォルダにコピー

---

### 3. 棚卸

月次棚卸に必要な**6ステップの帳票作成を自動化**。

#### 処理フロー

```
月次確定後:
  └─ 00_棚卸入力
  └─ 01_新比重入力
  └─ 02_親会社報告（月次確定後）
        ├─ 割戻表
        ├─ 在庫証明
        └─ 親会社報告書

保税適用確定後:
  └─ 03_製造終了届

製造終了届確認後:
  └─ 04_月次帳

月次バッチ後:
  └─ 05_とうもろこし調査票
  └─ 06_台帳〜配合日報
```

#### 共通モジュール

```python
# app_automation.py - RDP/アプリ操作の共通化

class RDPApp:
    """RDP経由でのアプリ操作を管理"""
    - launch_and_login()    # RDP接続・ログイン
    - navigate_tabs()       # タブ移動
    - select_menu_item()    # メニュー選択
    - print_excel()         # Excel印刷

class ExcelUtils:
    """Excel操作ユーティリティ"""
    - get_lastrow()         # 最終行取得
    - safe_open_workbook()  # 安全なファイルオープン
    - copy_cell_value()     # セル値コピー
```

---

## 💡 工夫したポイント

### 1. 保守性を考慮した設計

```python
# config.py で設定を一元管理
FILE_PATHS = {
    "kiso_irai": r"\\MC10\share\...\基礎依頼票.xlsm",
    "pellet_base": r"\\MC10\share\...\ペレット予測",
}

# UI座標も設定ファイルで管理（環境変更に対応）
OA_CLICK_X = 826
OA_CLICK_Y = 448
```

### 2. 共通処理のモジュール化

```python
# 営業日計算を共通化
def get_next_workday(base_date=None):
    """祝日・土日を除いた翌営業日を取得"""
    ...

# RDP操作を共通化
class RDPApp:
    @staticmethod
    def launch_and_login(login_id="12"):
        ...
```

### 3. エラーハンドリング

```python
def run_task(task_func, task_name: str) -> bool:
    try:
        task_func.main()
        print(f"✓ {task_name} - 完了")
        return True
    except Exception as e:
        print(f"✗ {task_name} - エラー: {e}")
        return False
```

---

## 📈 導入効果

| 指標 | Before | After |
|------|--------|-------|
| 16時印刷作業時間 | 約30分/日 | **約3分/日（自動実行）** |
| 大麦完了報告作成 | 約1時間/回 | **約10分/回** |
| 棚卸帳票作成 | 約2時間/月 | **約20分/月** |
| 人的ミス | 発生あり | **大幅削減** |

---

## 🔧 実行環境

- **OS**: Windows 10/11
- **Python**: 3.8以上
- **必要アプリ**: Microsoft Excel, Microsoft Word
- **ネットワーク**: 社内ファイルサーバーへのアクセス

### 必要ライブラリ

```bash
pip install pyautogui pywin32 python-docx jpholiday pyperclip
```

---

## 📝 ライセンス

MIT

---

**作成者**: Torihead  
**最終更新**: 2026年1月
