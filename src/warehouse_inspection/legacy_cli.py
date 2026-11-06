# ===========================================================
# 外部ライブラリ
# ===========================================================


from openpyxl import Workbook, load_workbook


# ===========================================================
# マスタデータ (旧CLI版)
# ===========================================================


jan_master = {
    "4901234567894": "PART-A001",
    "4901234567900": "PART-A002"
}


def load_shipping_plan():
    wb = load_workbook("legacy_data/shipping_plan.xlsx")
    ws = wb.active

    shipping_plan = {}

    for row in ws.iter_rows(min_row = 2, values_only = True):
        destination = row[0]
        product = row[1]
        quantity = row[2]

        if destination not in shipping_plan:
            shipping_plan[destination] = {}

        shipping_plan[destination][product] = {}
        shipping_plan[destination][product]["予定"] = quantity
        shipping_plan[destination][product]["検品済み"] = 0

    return shipping_plan


def load_inspection(items,filename):

    try:
        wb = load_workbook(filename)
    except FileNotFoundError:
        print("保存ファイルがありません")
        return
    
    ws = wb.active

    for row in ws.iter_rows(min_row = 2, values_only = True):
        product = row[0]
        count = row[2]

        print(f"読み込み: {product} → {count}個")

        if product in items:
            items[product]["検品済み"] = count


def save_inspection(items,filename):
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "品番"
    ws["B1"] = "予定数"
    ws["C1"] = "検品済み"

    for row, product in enumerate(items, start=2):
        ws[f"A{row}"] = product
        ws[f"B{row}"] = items[product]["予定"]
        ws[f"C{row}"] = items[product]["検品済み"]

    wb.save(filename)


def select_destination(shipping_plan):
    while True:

        destination = input("出荷先を入力してください")

        if destination not in shipping_plan:
            print("出荷先が登録されていません")

        else:
            items = shipping_plan[destination]

            return destination, items

        


def run_inspection(items, filename, jan_master):

    while True:

        jan = input("JANコードを入力してください (終了はq) : ")


        if jan == "q":
            save_inspection(items,filename)
            break

        if jan not in jan_master:
            print("JANコードが登録されていません")
        else:
            inspect_product(items, jan_master[jan])
            save_inspection(items,filename)


def inspect_product(items, product):
    if product not in items:
        print("登録されていません")
    elif items[product]["検品済み"] >= items[product]["予定"]:
        print("検品数が予定数に達しています")
    else:    
        items[product]["検品済み"] += 1
        print(f"{product}: {items[product]['検品済み']} / {items[product]['予定']}")



if __name__ == "__main__":
    shipping_plan = load_shipping_plan()

    destination, items = select_destination(shipping_plan)

    filename = f"legacy_data/inspection_{destination}.xlsx"

    load_inspection(items, filename)

    run_inspection(items, filename, jan_master)
