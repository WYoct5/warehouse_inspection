from openpyxl import Workbook

#-------------送り先データの作成--------------

wb = Workbook()
ws = wb.active

ws["A1"] = "問い合わせNO"
ws["B1"] ="送り先名称"
ws["C1"] ="送り先住所"

ws["A2"] ="001234567890"
ws["B2"] ="東京物流センター"
ws["C2"] ="東京都江東区"

ws["A3"] ="009876543210"
ws["B3"] ="大阪物流センター"
ws["C3"] ="大阪府大阪市"

wb.save("destinations.xlsx")

#---------------- 出荷データの作成-------------------

wb = Workbook()
ws = wb.active

ws.append(["送り先名称",	"送り先住所", "品番", "出荷予定数量"])

ws.append(["東京物流センター", "東京都江東区", "PART-A001", 3])
ws.append(["東京物流センター", "東京都江東区", "PART-A002", 2])
ws.append(["大阪物流センター", "大阪府大阪市", "PART-A001", 5])

wb.save("shipments.xlsx")

#-------------商品マスタの作成------------------------

wb = Workbook()
ws = wb.active

ws.append(["JANコード", "品番", "商品名"])
ws.append(["4901234567894", "PART-A001", "商品A"])
ws.append(["4901234567900", "PART-A002", "商品B"])

wb.save("product_master.xlsx")