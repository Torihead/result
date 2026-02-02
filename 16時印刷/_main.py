# 基礎依頼
import 基礎依頼ver2 as kiso
kiso.main()
print("基礎依頼の処理を完了しました。")

# ペレット予定
import ペレット予定ver2 as p
p.main()
print("ペレット予定の処理を完了しました。")

# TB出荷予定
import TB保管出荷ver2 as tb
tb.main()
print("TB保管出荷の処理を完了しました。")

# 端量表
import 翌営業日端量表ver2 as haryou
haryou.main()
print("翌営業日端量表の処理を完了しました。")

# OA起動
import OA起動処理 as OAup
OAup.main()

# 出荷予定の印刷
import OA出荷予定印刷ver2 as pSY
pSY.main()
print("OA出荷予定の処理を完了しました。")

# 製品倉庫移動の印刷
import OA倉庫移動印刷ver2 as pSI
pSI.main()
print("OA倉庫移動の処理を完了しました。")

# 週間受注状況の印刷
import OA週間受注印刷ver2 as pSJ
pSJ.main()
print("OA週間受注の処理を完了しました。")

# OUTデータ 製品入出庫予定
import OUT製品入出庫予定ver2 as sny
sny.main()

print("全行程の処理を完了しました。")

import os
import time
time.sleep(15)  # 15秒待機してからExcelを閉じる
os.system('taskkill /f /im EXCEL.EXE')