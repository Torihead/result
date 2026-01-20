"""
16時印刷メインスクリプト

一連の印刷処理を実行します：
1. 基礎依頼
2. ペレット予定
3. TB保管出荷
4. 翌営業日端量表
5. OA起動
6. OA週間受注印刷
7. OA出荷予定印刷
8. OA倉庫移動印刷
9. 製品入出庫予定

最後に Excel を閉じます。
"""
import time
import os
from pathlib import Path

# ローカルモジュール
from system import oa_startup
from tasks import kiso_irai, pellet, tb_hokan, haryou, oa_shukan, oa_shokka, oa_souko, out_sehin
import config


def run_task(task_func, task_name: str) -> bool:
    """
    タスクを実行してエラーハンドリング
    
    Args:
        task_func: タスク関数の main()
        task_name: タスク名
    
    Returns:
        成功時は True、失敗時は False
    """
    try:
        print(f"\n{'=' * 60}")
        print(f"【開始】 {task_name}")
        print(f"{'=' * 60}")
        task_func.main()
        print(f"✓ {task_name} - 完了")
        return True
    except Exception as e:
        print(f"✗ {task_name} - エラー: {e}")
        return False


def main() -> None:
    """メイン処理"""
    print("\n" + "=" * 60)
    print("16時印刷 - 自動処理開始")
    print("=" * 60)

    results = []

    # 1. 基礎依頼
    results.append(("基礎依頼", run_task(kiso_irai, "基礎依頼")))

    # 2. ペレット予定
    results.append(("ペレット予定", run_task(pellet, "ペレット予定")))

    # 3. TB保管出荷
    results.append(("TB保管出荷", run_task(tb_hokan, "TB保管出荷")))

    # 4. 翌営業日端量表
    results.append(("翌営業日端量表", run_task(haryou, "翌営業日端量表")))

    # 5. OA起動
    results.append(("OA起動", run_task(oa_startup, "OA起動")))

    # 6. OA週間受注印刷
    results.append(("OA週間受注印刷", run_task(oa_shukan, "OA週間受注印刷")))

    # 7. OA出荷予定印刷
    results.append(("OA出荷予定印刷", run_task(oa_shokka, "OA出荷予定印刷")))

    # 8. OA倉庫移動印刷
    results.append(("OA倉庫移動印刷", run_task(oa_souko, "OA倉庫移動印刷")))

    # 9. 製品入出庫予定
    results.append(("製品入出庫予定", run_task(out_sehin, "製品入出庫予定")))

    # 結果サマリー
    print("\n" + "=" * 60)
    print("【処理結果サマリー】")
    print("=" * 60)
    for task_name, success in results:
        status = "✓ 成功" if success else "✗ 失敗"
        print(f"{status} - {task_name}")

    successful = sum(1 for _, s in results if s)
    total = len(results)
    print(f"\n総合: {successful}/{total} 完了")

    # Excel を閉じる
    print(f"\n{config.EXCEL_CLOSE_DELAY}秒待機後、Excel を閉じます...")
    time.sleep(config.EXCEL_CLOSE_DELAY)
    os.system('taskkill /f /im EXCEL.EXE')

    print("\n" + "=" * 60)
    print("全行程の処理を完了しました。")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
