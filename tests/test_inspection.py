from warehouse_inspection.inspection import (
    inspect_product,
    load_destinations,
    load_shipments,
    get_shipment_by_no,
    create_inspection_items,
    get_product_by_jan,
    inspect_by_jan,
    is_inspection_complete,
)

from warehouse_inspection.legacy_cli import (
    save_inspection,
    load_inspection,
)

from openpyxl import load_workbook

import pytest

def test_inspection_count():
    items = {"PART-A001": {"予定": 3, "検品済み": 2}}

    inspect_product(items, "PART-A001")

    assert items["PART-A001"]["検品済み"] == 3


def test_inspection_limit():
    items = {"PART-A001": {"予定": 3, "検品済み": 3}}

    with pytest.raises(ValueError, match = "検品数が予定数に達しています"):
        inspect_product(items, "PART-A001")

    assert items["PART-A001"]["検品済み"] == 3


def test_inspection_unknown():
    items = {"PART-A001": {"予定": 3, "検品済み":2}}

    with pytest.raises(ValueError, match = "出荷予定にない商品です"):
        inspect_product(items, "PART-A999")

    assert items["PART-A001"]["検品済み"] == 2


def test_inspect_product_increases_count():
    items = {
        "PART-A001": {"予定": 3, "検品済み": 1}
    }

    inspect_product(items, "PART-A001")

    assert items["PART-A001"]["検品済み"] == 2


def test_save_inspection(tmp_path):
    items = {
        "PART-A001": {"予定": 3, "検品済み": 2}
    }

    filename = tmp_path / "test_inspection.xlsx"

    save_inspection(items, filename)

    wb = load_workbook(filename)
    ws = wb.active

    assert filename.exists() 

    assert ws["A2"].value == "PART-A001"
    assert ws["B2"].value == 3
    assert ws["C2"].value == 2

def test_load_inspection(tmp_path):
    saved_items = {
        "PART-A001": {"予定": 3, "検品済み": 2}
    }

    filename = tmp_path / "test_inspection.xlsx"

    save_inspection(saved_items, filename)

    items = {
        "PART-A001": {"予定": 3, "検品済み": 0}
    }

    load_inspection(items,filename)

    assert items["PART-A001"]["検品済み"] == 2


def test_create_inspection_items():

    shipment = {
        "PART-A001": 3,
        "PART-A002": 2
    }

    inspection_items = create_inspection_items(shipment)

    assert inspection_items == {
                                "PART-A001": {"予定": 3, "検品済み": 0},
                                "PART-A002": {"予定": 2, "検品済み": 0}
                                }


def test_get_product_by_jan_valid():

    product_master = {
                    "4901234567894": {
                        "品番": "PART-A001",
                        "商品名": "商品A"
                    },
                    "4901234567900": {
                        "品番": "PART-A002",
                        "商品名": "商品B"
                    }
                    }

    jan_code = "4901234567894"

    product = get_product_by_jan(jan_code, product_master)

    assert product == "PART-A001"


def test_get_product_by_jan_invalid():

    product_master = {
                    "4901234567894": {
                        "品番": "PART-A001",
                        "商品名": "商品A"
                    },
                    "4901234567900": {
                        "品番": "PART-A002",
                        "商品名": "商品B"
                    }
                    }

    jan_code = "999999999999"

    with pytest.raises(ValueError, match = "未登録のJANコードです"):
        get_product_by_jan(jan_code, product_master)


def test_inspect_by_jan_valid():

    product_master = {
        "4901234567894": {
            "品番": "PART-A001",
            "商品名": "商品A"
        }
    }

    items = {
        "PART-A001": {"予定": 3, "検品済み": 1}
    }

    jan_code = "4901234567894"

    inspect_by_jan(jan_code, product_master, items)

    assert items == {
                "PART-A001": {"予定": 3, "検品済み": 2}
    }


def test_inspect_by_jan_invalid():

    product_master = {
        "4901234567894": {
            "品番": "PART-A001",
            "商品名": "商品A"
        }
    }

    items = {
        "PART-A001": {"予定": 3, "検品済み": 1}
    }

    jan_code = "9999999999999"

    with pytest.raises(ValueError, match = "未登録のJANコードです"):
        inspect_by_jan(jan_code, product_master, items)

    assert items == {
                "PART-A001": {"予定": 3, "検品済み": 1}
    }


def test_inspect_by_jan_unexpected_product():
        
    product_master = {
        "4901234567894": {
            "品番": "PART-A001",
            "商品名": "商品A"
        }
    }

    items = {
            "PART-B001": {"予定": 3, "検品済み": 1}
    }

    jan_code = "4901234567894"

    with pytest.raises(ValueError, match = "出荷予定にない商品です"):
        inspect_by_jan(jan_code, product_master, items)

    assert items == {
                "PART-B001": {"予定": 3, "検品済み": 1}
    }


def test_inspect_by_jan_over_limit():

    product_master = {
        "4901234567894": {
            "品番": "PART-A001",
            "商品名": "商品A"
        }
    }

    items = {
            "PART-A001": {"予定": 3, "検品済み": 3}
    }

    jan_code = "4901234567894"

    with pytest.raises(ValueError, match = "検品数が予定数に達しています"):
        inspect_by_jan(jan_code, product_master, items)

    assert items == {
                "PART-A001": {"予定": 3, "検品済み": 3}
    }


def test_is_inspection_complete_true():

    items = {
        "PART-A001": {"予定": 3, "検品済み": 3},
        "PART-B001": {"予定": 2, "検品済み": 2}
    }

    assert is_inspection_complete(items)


def test_is_inspection_complete_false():

    items = {
        "PART-A001": {"予定": 3, "検品済み": 3},
        "PART-B001": {"予定": 2, "検品済み": 1}
    }

    assert not is_inspection_complete(items)
