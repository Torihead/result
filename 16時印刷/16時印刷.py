print("処理を開始します。")

# 基礎依頼
import 基礎依頼 as kiso
kiso.print_基礎依頼()
print("基礎依頼の処理を完了しました。")

# ペレット加工予定
import ペレット予定 as perret
perret.print_ペレット予定()
print("ペレット予定の処理を完了しました。")

# 翌営業日端量表の 製造予定
import 翌営業日端量表 as yeh
yeh.print_翌営業日端量表()
print("翌営業日端量表の処理を完了しました。")

# TB保管出荷
import TB保管出荷 as tbs
tbs.print_TB保管出荷()
print("TB保管出荷の処理を完了しました。")

# OA起動
import OA起動処理 as OAup
OAup.OA_Start()

# 出荷予定の印刷
import OA出荷予定印刷 as pSY
pSY.print_出荷予定()
print("OA出荷予定の処理を完了しました。")

# 製品倉庫移動の印刷
import OA倉庫移動印刷 as pSI
pSI.print_倉庫移動()
print("OA倉庫移動の処理を完了しました。")

# 週間受注状況の印刷
import OA週間受注印刷 as pSJ
pSJ.print_週間受注()
print("OA週間受注の処理を完了しました。")

# OUTデータ 製品入出庫予定
import OUT製品入出庫予定 as sny
sny.print_OUT製品入出庫予定()

print("全行程の処理を完了しました。")