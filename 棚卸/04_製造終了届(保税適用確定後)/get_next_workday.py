import jpholiday as jhd
import datetime as dt

def get_next_weekday(date):
    next_day = date + dt.timedelta(days=1)      # date=1 で7日間に変換
    
    # 翌日が 土日 or 祝日 なら、平日まで繰り返し処理
    while next_day.weekday() >= 5 or jhd.is_holiday(next_day):
        next_day += dt.timedelta(days=1)
    
    return next_day

# 今日の日付を取得
today = dt.datetime.today()
next_workday = get_next_weekday(today)

print(f"次の平日は: {next_workday.strftime('%Y%m%d')}")
