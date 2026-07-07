from price_generator.constants import (
    AVAILABLE_COLUMN_NAME,
    COST_PRICE_COLUMN_NAME,
    COMBO_WEIGHT_COLUMN_NAME,
    DEFAULT_DELETE_RULES,
    PRODUCT_COLUMN_NAME,
    PRODUCT_TAG_COLUMN_NAME,
    SALES_15_DAYS_COLUMN_NAME,
    SPEC_CODE_COLUMN_NAME,
    STOCK_COLUMN_NAME,
    SUPPLIER_COLUMN_NAME,
    WEIGHT_COLUMN_NAME,
)
from price_generator.headers import build_header_map, find_header_index
from price_generator.rules import should_delete_row


def test_header_aliases_match_combo_weight_and_rule_code():
    assert find_header_index(["组合商品重量"], COMBO_WEIGHT_COLUMN_NAME) == 0
    assert find_header_index(["规则编码"], SPEC_CODE_COLUMN_NAME) == 0


def test_delete_rules_remove_flawed_product_name():
    headers = [PRODUCT_COLUMN_NAME]
    assert should_delete_row(["轻微瑕品"], build_header_map(headers), DEFAULT_DELETE_RULES)


def test_delete_rules_keep_normal_product():
    headers = [
        PRODUCT_COLUMN_NAME,
        COST_PRICE_COLUMN_NAME,
        AVAILABLE_COLUMN_NAME,
        STOCK_COLUMN_NAME,
        PRODUCT_TAG_COLUMN_NAME,
        SUPPLIER_COLUMN_NAME,
        WEIGHT_COLUMN_NAME,
        SALES_15_DAYS_COLUMN_NAME,
    ]
    row = ["正常商品", 8, 120, 20, "2601", "ABC供应商", 0.5, 20]

    assert not should_delete_row(row, build_header_map(headers), DEFAULT_DELETE_RULES)
