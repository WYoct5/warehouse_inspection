from openpyxl import Workbook, load_workbook
from warehouse_inspection.inspection import load_product_master
import pytest

def test_load_product_master_valid(tmp_path):

    wb = Workbook()
    ws = wb.active

    ws.append(["JANコード", "品番", "商品名"])
    ws.append(["4901234567894", "PART-A001", "商品A"])
    ws.append(["4901234567900", "PART-A002", "商品B"])

    filename = tmp_path / "test_product_master.xlsx"

    wb.save(filename)

    product_master = load_product_master(filename)

    assert product_master == {
        "4901234567894": {
            "品番": "PART-A001",
            "商品名": "商品A"
        },
        "4901234567900": {
            "品番": "PART-A002",
            "商品名": "商品B"
        }
    }


def test_load_product_master_duplicate(tmp_path):

    wb = Workbook()
    ws = wb.active

    ws.append(["JANコード", "品番", "商品名"])
    ws.append(["4901234567894", "PART-A001", "商品A"])
    ws.append(["4901234567894", "PART-A002", "商品B"])

    filename = tmp_path / "test_product_master.xlsx"

    wb.save(filename)

    with pytest.raises(ValueError, match = "JANコードが重複しています: 4901234567894"):
        load_product_master(filename)
    