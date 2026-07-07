import json
import os

from .paths import default_sales_analysis_cache_path


def normalize_order_key(value):
    return str(value if value is not None else "").strip()


def load_sales_analysis_cache(path=None):
    cache_path = path or default_sales_analysis_cache_path()

    if not cache_path or not os.path.exists(cache_path):
        return {}

    with open(cache_path, "r", encoding="utf-8") as cache_file:
        data = json.load(cache_file)

    orders = data.get("orders") if isinstance(data, dict) else None
    if not isinstance(orders, dict):
        raise ValueError("销售主题分析价格库格式错误：需要 orders 对象。")

    normalized_cache = {}
    for key, value in orders.items():
        normalized_key = normalize_order_key(key)
        if not normalized_key:
            continue
        if (
            not isinstance(value, dict)
            or "order_weight" not in value
            or "shipping_fee" not in value
            or "sales_amount" not in value
            or "sales_cost" not in value
            or "gross_profit" not in value
        ):
            raise ValueError("销售主题分析价格库格式错误：每条记录需要 order_weight、shipping_fee、sales_amount、sales_cost 和 gross_profit。")
        normalized_cache[normalized_key] = {
            "order_weight": value["order_weight"],
            "shipping_fee": value["shipping_fee"],
            "sales_amount": value["sales_amount"],
            "sales_cost": value["sales_cost"],
            "gross_profit": value["gross_profit"],
        }

    return normalized_cache


def save_sales_analysis_cache(sales_cache, path=None):
    cache_path = path or default_sales_analysis_cache_path()
    normalized_cache = {}

    for key, value in sales_cache.items():
        normalized_key = normalize_order_key(key)
        if (
            not normalized_key
            or not isinstance(value, dict)
            or "order_weight" not in value
            or "shipping_fee" not in value
            or "sales_amount" not in value
            or "sales_cost" not in value
            or "gross_profit" not in value
        ):
            continue
        normalized_cache[normalized_key] = {
            "order_weight": value["order_weight"],
            "shipping_fee": value["shipping_fee"],
            "sales_amount": value["sales_amount"],
            "sales_cost": value["sales_cost"],
            "gross_profit": value["gross_profit"],
        }

    with open(cache_path, "w", encoding="utf-8") as cache_file:
        json.dump({"orders": normalized_cache}, cache_file, ensure_ascii=False, indent=2)
