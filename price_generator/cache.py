import json
import os

from openpyxl import Workbook

from .constants import PRICE_COLUMN_NAME, PRODUCT_CODE_COLUMN_NAME
from .paths import default_price_cache_path
from .pricing import normalize_price_cache_key


def load_price_cache(path=None):
    cache_path = path or default_price_cache_path()

    if not cache_path or not os.path.exists(cache_path):
        return {}

    with open(cache_path, "r", encoding="utf-8") as cache_file:
        data = json.load(cache_file)

    prices = data.get("prices", data if isinstance(data, dict) else None)
    if not isinstance(prices, dict):
        raise ValueError("售价保存文件格式错误：需要 prices 对象。")

    return {normalize_price_cache_key(key): value for key, value in prices.items() if normalize_price_cache_key(key)}


def save_price_cache(price_cache, path=None):
    cache_path = path or default_price_cache_path()
    normalized_cache = {
        normalize_price_cache_key(key): value
        for key, value in price_cache.items()
        if normalize_price_cache_key(key)
    }

    with open(cache_path, "w", encoding="utf-8") as cache_file:
        json.dump({"prices": normalized_cache}, cache_file, ensure_ascii=False, indent=2)


def export_price_cache_xlsx(output_path, price_cache):
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("价格库")
    worksheet.append([PRODUCT_CODE_COLUMN_NAME, PRICE_COLUMN_NAME])

    for product_code in sorted(price_cache):
        worksheet.append([product_code, price_cache[product_code]])

    workbook.save(output_path)
    return len(price_cache)
