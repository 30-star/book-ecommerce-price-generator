import json
import os

from openpyxl import Workbook

from .constants import COST_PRICE_COLUMN_NAME, INSERTED_COLUMN_NAME, PRICE_COLUMN_NAME, PRODUCT_CODE_COLUMN_NAME
from .paths import default_price_cache_path
from .pricing import normalize_price_cache_key


def load_price_cache(path=None):
    cache_path = path or default_price_cache_path()

    if not cache_path or not os.path.exists(cache_path):
        return {}

    with open(cache_path, "r", encoding="utf-8") as cache_file:
        data = json.load(cache_file)

    prices = data.get("prices") if isinstance(data, dict) else None
    if not isinstance(prices, dict):
        raise ValueError("价格库文件格式错误：需要 prices 对象。请清空价格库后重新导入。")

    normalized_cache = {}
    for key, value in prices.items():
        normalized_key = normalize_price_cache_key(key)
        if not normalized_key:
            continue
        if (
            not isinstance(value, dict)
            or "cost_price" not in value
            or "shipping_fee" not in value
            or "price" not in value
        ):
            raise ValueError("价格库文件格式错误：每条记录需要 cost_price、shipping_fee 和 price。请清空价格库后重新导入。")
        normalized_cache[normalized_key] = {
            "cost_price": value["cost_price"],
            "shipping_fee": value["shipping_fee"],
            "price": value["price"],
        }

    return normalized_cache


def save_price_cache(price_cache, path=None):
    cache_path = path or default_price_cache_path()
    normalized_cache = {}
    for key, value in price_cache.items():
        normalized_key = normalize_price_cache_key(key)
        if (
            not normalized_key
            or not isinstance(value, dict)
            or "cost_price" not in value
            or "shipping_fee" not in value
            or "price" not in value
        ):
            continue
        normalized_cache[normalized_key] = {
            "cost_price": value["cost_price"],
            "shipping_fee": value["shipping_fee"],
            "price": value["price"],
        }

    with open(cache_path, "w", encoding="utf-8") as cache_file:
        json.dump({"prices": normalized_cache}, cache_file, ensure_ascii=False, indent=2)


def export_price_cache_xlsx(output_path, price_cache):
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("价格库")
    worksheet.append([PRODUCT_CODE_COLUMN_NAME, COST_PRICE_COLUMN_NAME, INSERTED_COLUMN_NAME, PRICE_COLUMN_NAME])

    for product_code in sorted(price_cache):
        record = price_cache[product_code]
        worksheet.append([product_code, record.get("cost_price"), record.get("shipping_fee"), record.get("price")])

    workbook.save(output_path)
    return len(price_cache)
