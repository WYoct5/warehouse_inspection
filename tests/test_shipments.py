from warehouse_inspection.inspection import (
    inspect_product,
    load_destinations,
    load_shipments,
    get_shipment_by_no,
)

from warehouse_inspection.legacy_cli import (
    save_inspection,
    load_inspection,
)

from openpyxl import load_workbook, Workbook
import pytest


def test_get_shipment_by_no_valid():

    destinations = load_destinations()

    shipments = load_shipments("shipments.xlsx")

    items = get_shipment_by_no("001234567890", destinations, shipments)

    assert items =={"PART-A001": 3,"PART-A002": 2}

def test_get_shipment_by_no_invalid():

    destinations = load_destinations()

    shipments = load_shipments("shipments.xlsx")

    with pytest.raises(ValueError, match = "未登録のお問い合わせNOです"):
        get_shipment_by_no("999999999999", destinations, shipments)


def test_get_shipment_by_no_address_mismatch():

    destinations = {
        "001234567890": {
            "送り先名称": "東京物流センター",
            "送り先住所": "東京都江東区"
            }
        }

    shipments = {
        "東京物流センター": {
            "東京都港区": {
                "PART-A001": 3
                }
            }
        }

    with pytest.raises(ValueError, match = "送り先住所が異なります"):
        get_shipment_by_no("001234567890", destinations, shipments)


def test_get_shipment_by_no_destination_not_found():

    destinations = {
        "001234567890": {
            "送り先名称": "東京物流センター",
            "送り先住所": "東京都江東区"
        }
    }

    shipments = {
        "大阪物流センター": {
            "大阪府大阪市": {
                "PART-A001": 3
                }
            }
        }

    with pytest.raises(ValueError, match = "未登録の送り先です"):
        get_shipment_by_no("001234567890", destinations, shipments)


def test_load_shipments_duplicate(tmp_path):
    wb = Workbook()
    ws = wb.active

    ws.append(["送り先名称","送り先住所","品番","出荷予定数量"])
    ws.append(["東京物流センター","東京都江東区","PART-A001",3])
    ws.append(["東京物流センター","東京都江東区","PART-A001",2])

    filename = tmp_path / "test_shipments.xlsx"

    wb.save(filename)

    with pytest.raises(ValueError, match = "データが重複しています → 品番: PART-A001"):
        load_shipments(filename)


def test_load_shipments_valid(tmp_path):

    wb = Workbook()
    ws = wb.active

    ws.append(["送り先名称", "送り先住所", "品番", "出荷予定数量"])
    ws.append(["東京物流センター", "東京都江東区", "PART-A001", 3])
    ws.append(["東京物流センター", "東京都江東区", "PART-A002", 2])

    filename = tmp_path / "test_shipments.xlsx"

    wb.save(filename)

    shipments = load_shipments(filename)

    assert shipments == {
        "東京物流センター": {
            "東京都江東区": {
                "PART-A001": 3,
                "PART-A002": 2
                }
            }
        }