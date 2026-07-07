import pytest
from openpyxl import Workbook

from price_generator.constants import (
    INTERNAL_ORDER_NUMBER_COLUMN_NAME,
    ORDER_WEIGHT_COLUMN_NAME,
    SALES_AMOUNT_COLUMN_NAME,
    SALES_COST_COLUMN_NAME,
)
from price_generator.spreadsheet import import_sales_analysis_to_cache_csv, import_sales_analysis_to_cache_xlsx


def save_workbook(path, header, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def test_import_sales_analysis_xlsx(tmp_path):
    input_path = tmp_path / "sales.xlsx"
    sales_cache = {}
    save_workbook(
        input_path,
        [
            INTERNAL_ORDER_NUMBER_COLUMN_NAME,
            ORDER_WEIGHT_COLUMN_NAME,
            SALES_AMOUNT_COLUMN_NAME,
            SALES_COST_COLUMN_NAME,
        ],
        [
            ["A001", 1.2, 99.9, 60.5],
            ["", 2, 80, 40],
        ],
    )

    summary = import_sales_analysis_to_cache_xlsx(input_path, sales_cache=sales_cache)

    assert summary == [("Sheet1", 1, 0, 1)]
    assert sales_cache == {
        "A001": {
            "order_weight": 1.2,
            "shipping_fee": 2.3,
            "sales_amount": 99.9,
            "sales_cost": 60.5,
            "gross_profit": 37.1,
        }
    }


def test_import_sales_analysis_csv_updates_existing(tmp_path):
    input_path = tmp_path / "sales.csv"
    input_path.write_text(
        "内部订单号,订单重量,销售金额,销售成本\nA001,1.5,100,61\n",
        encoding="utf-8-sig",
    )
    sales_cache = {
        "A001": {
            "order_weight": 1.2,
            "shipping_fee": 2.3,
            "sales_amount": 99.9,
            "sales_cost": 60.5,
            "gross_profit": 37.1,
        }
    }

    summary = import_sales_analysis_to_cache_csv(input_path, sales_cache=sales_cache)

    assert summary == [("CSV", 0, 1, 0)]
    assert sales_cache == {
        "A001": {
            "order_weight": 1.5,
            "shipping_fee": 2.8,
            "sales_amount": 100.0,
            "sales_cost": 61.0,
            "gross_profit": 36.2,
        }
    }


def test_import_sales_analysis_uses_custom_shipping_fees(tmp_path):
    input_path = tmp_path / "sales.xlsx"
    sales_cache = {}
    custom_fees = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
    save_workbook(
        input_path,
        [
            INTERNAL_ORDER_NUMBER_COLUMN_NAME,
            ORDER_WEIGHT_COLUMN_NAME,
            SALES_AMOUNT_COLUMN_NAME,
            SALES_COST_COLUMN_NAME,
        ],
        [["A001", 1.2, 99.9, 60.5]],
    )

    summary = import_sales_analysis_to_cache_xlsx(input_path, sales_cache=sales_cache, shipping_fees=custom_fees)

    assert summary == [("Sheet1", 1, 0, 0)]
    assert sales_cache == {
        "A001": {
            "order_weight": 1.2,
            "shipping_fee": 40,
            "sales_amount": 99.9,
            "sales_cost": 60.5,
            "gross_profit": -0.6,
        }
    }


def test_import_sales_analysis_sets_blank_profit_when_shipping_fee_invalid(tmp_path):
    input_path = tmp_path / "sales.xlsx"
    sales_cache = {}
    save_workbook(
        input_path,
        [
            INTERNAL_ORDER_NUMBER_COLUMN_NAME,
            ORDER_WEIGHT_COLUMN_NAME,
            SALES_AMOUNT_COLUMN_NAME,
            SALES_COST_COLUMN_NAME,
        ],
        [["A001", 0, 99.9, 60.5]],
    )

    summary = import_sales_analysis_to_cache_xlsx(input_path, sales_cache=sales_cache)

    assert summary == [("Sheet1", 1, 0, 0)]
    assert sales_cache == {
        "A001": {
            "order_weight": 0.0,
            "shipping_fee": "请检查",
            "sales_amount": 99.9,
            "sales_cost": 60.5,
            "gross_profit": None,
        }
    }


def test_import_sales_analysis_requires_all_columns(tmp_path):
    input_path = tmp_path / "sales.xlsx"
    save_workbook(input_path, [INTERNAL_ORDER_NUMBER_COLUMN_NAME], [["A001"]])

    with pytest.raises(ValueError, match=ORDER_WEIGHT_COLUMN_NAME):
        import_sales_analysis_to_cache_xlsx(input_path, sales_cache={})
