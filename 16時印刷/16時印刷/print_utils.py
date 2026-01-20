"""
印刷ユーティリティ - 共通の印刷・UI操作関数
"""
import time
from typing import Optional
import pyautogui
import config


def safe_click(x: int, y: int, sleep_after: float = config.WAIT_SHORT) -> None:
    """
    安全にクリックして待機
    
    Args:
        x: X座標
        y: Y座標
        sleep_after: クリック後の待機時間
    """
    try:
        pyautogui.click(x=x, y=y)
        time.sleep(sleep_after)
    except Exception as e:
        raise RuntimeError(f"クリック失敗 ({x}, {y}): {e}")


def safe_write(text: str, interval: float = 0.1) -> None:
    """
    安全にテキストを入力
    
    Args:
        text: 入力するテキスト
        interval: 文字間隔
    """
    try:
        pyautogui.write(text, interval=interval)
    except Exception as e:
        raise RuntimeError(f"テキスト入力失敗 ({text}): {e}")


def safe_press(key: str, presses: int = 1, interval: float = 0.1, sleep_after: float = 0) -> None:
    """
    安全にキーを押す
    
    Args:
        key: キー名（'enter', 'space' など）
        presses: 押す回数
        interval: キー間隔
        sleep_after: 押下後の待機時間
    """
    try:
        pyautogui.press(key, presses=presses, interval=interval)
        if sleep_after > 0:
            time.sleep(sleep_after)
    except Exception as e:
        raise RuntimeError(f"キー入力失敗 ({key}): {e}")


def click_and_write(x: int, y: int, text: str, wait_click: float = config.WAIT_SHORT) -> None:
    """
    クリックしてテキストを入力
    
    Args:
        x: X座標
        y: Y座標
        text: 入力するテキスト
        wait_click: クリック後の待機時間
    """
    safe_click(x, y, sleep_after=wait_click)
    safe_write(text)


def click_uketsukegimu() -> None:
    """受付業務タグをクリック"""
    safe_click(config.UKETSUKEGIMU_X, config.UKETSUKEGIMU_Y, sleep_after=config.WAIT_SHORT)


def close_screen() -> None:
    """画面を閉じる（ウィンドウ右上の X をクリック）"""
    safe_click(1897, 13, sleep_after=config.WAIT_SHORT)


def print_preview_close() -> None:
    """印刷プレビューを閉じる（F12キー）"""
    safe_press("F12")
