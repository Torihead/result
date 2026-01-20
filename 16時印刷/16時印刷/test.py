import win32com.client as w32
import get_next_workday as gnw


next_date = gnw.get_next_weekday(gnw.today)                     # 翌日営業日を取得
month = gnw.today.strftime('%Y.%m')     # 今月を2025.07で取得
print(f"今日の月: {month}")