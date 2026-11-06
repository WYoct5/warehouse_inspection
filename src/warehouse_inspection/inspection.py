# ===========================================================
# 外部ライブラリ
# ===========================================================


from openpyxl import Workbook, load_workbook


# ===========================================================
# Excel読み込み処理
# ===========================================================


def load_destinations():
    wb = load_workbook("destinations.xlsx")
    ws = wb.active

    destinations = {}

    for row in ws.iter_rows(min_row = 2, values_only = True):
        no = row[0]
        destination = row[1]
        address = row [2]

        destinations[no] = {"送り先名称": destination, "送り先住所": address}

    return destinations


def load_shipments(filename):
    wb = load_workbook(filename)
    ws = wb.active

    shipments = {}

    for row in ws.iter_rows(min_row = 2, values_only = True):
        
        destination = row[0]
        address = row[1]
        product = row[2]
        quantity = row[3]

        if destination not in shipments:
            shipments[destination] = {}

        if address not in shipments[destination]:
            shipments[destination][address] = {}

        if product in shipments[destination][address]:
            raise ValueError(f"データが重複しています → 品番: {product}")

        else:
            shipments[destination][address][product] = quantity

    return shipments   


def load_product_master(filename):

    wb = load_workbook(filename)
    ws = wb.active

    product_master ={}

    for row in ws.iter_rows(min_row = 2, values_only = True):
        
        jan_code = row[0]
        product_number = row[1]
        product = row[2]

        if jan_code not in product_master:

            product_master[jan_code] = {}

            product_master[jan_code] = {
                "品番": product_number,
                "商品名": product
            }

        else:
            raise ValueError(f"JANコードが重複しています: {jan_code}")

    return product_master


# ===========================================================
# 出荷予定の取得処理
# ===========================================================


def get_shipment_by_no(no, destinations, shipments):

    if no not in destinations:
        raise ValueError("未登録のお問い合わせNOです")

    destination = destinations[no]["送り先名称"]
    address = destinations[no]["送り先住所"]


    if destination not in shipments:
        raise ValueError("未登録の送り先です")

    if address not in shipments[destination]:
        raise ValueError("送り先住所が異なります")

    items = shipments[destination][address]

    return items


# ===========================================================
# 検品処理
# ===========================================================


def inspect_product(items, product):
    if product not in items:
        raise ValueError("出荷予定にない商品です")
    elif items[product]["検品済み"] >= items[product]["予定"]:
        raise ValueError("検品数が予定数に達しています")
    else:    
        items[product]["検品済み"] += 1


def is_inspection_complete(items):
    for product in items:
        if items[product]["検品済み"] < items[product]["予定"]:
            return False

    return True

def create_inspection_items(shipment):

    inspection_items= {}

    for product, quantity in shipment.items():

        inspection_items[product] = {}

        inspection_items[product] = {
            "予定": quantity,
            "検品済み": 0
        }

    return inspection_items


def get_product_by_jan(jan_code, product_master):

    if jan_code not in product_master:
        raise ValueError("未登録のJANコードです")

    product = product_master[jan_code]["品番"]

    return product


def inspect_by_jan(jan_code, product_master, items):

    product = get_product_by_jan(jan_code, product_master)
    inspect_product(items, product)

# ===========================================================
# メイン処理
# ===========================================================


if __name__ == "__main__":

    destinations = load_destinations()

    shipments = load_shipments("shipments.xlsx")

    no = input("お問い合わせNOを入力してください")

    try:
        items = get_shipment_by_no(no, destinations, shipments)
        print(items)
    except ValueError as e:
        print(e)
        
