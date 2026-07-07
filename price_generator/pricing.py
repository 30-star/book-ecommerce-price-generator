from .constants import (
    DEFAULT_HIGH_PRICE_MARGIN,
    DEFAULT_LOW_PRICE_MARGIN,
    DEFAULT_PRICE_THRESHOLD,
    SHIPPING_FEE_RULES,
)
from .headers import get_row_column_value
from .rules import get_row_value, parse_number


def parse_weight(value):
    if value is None:
        return None

    text = str(value).strip().replace(",", "")
    number = ""
    dot_seen = False

    for char in text:
        if char.isdigit():
            number += char
        elif char == "." and not dot_seen:
            number += char
            dot_seen = True
        elif number:
            break

    if not number or number == ".":
        return None

    try:
        return float(number)
    except ValueError:
        return None


def calculate_shipping_fee(value, shipping_fees=None):
    weight = parse_weight(value)
    fees = shipping_fees or [rule[1] for rule in SHIPPING_FEE_RULES]

    if weight is None or weight == 0:
        return "请检查"
    if 0 < weight < 0.25:
        return fees[0]
    if 0.25 <= weight < 0.4:
        return fees[1]
    if 0.4 <= weight < 0.9:
        return fees[2]
    if 0.9 <= weight < 1.4:
        return fees[3]
    if 1.4 <= weight < 1.9:
        return fees[4]
    if 1.9 <= weight < 2.4:
        return fees[5]
    if 2.4 <= weight < 2.9:
        return fees[6]
    if 2.9 <= weight < 3.9:
        return fees[7]
    if 3.9 <= weight < 4.9:
        return fees[8]
    if weight >= 4.9:
        return fees[9]

    return "请检查"


def calculate_price(cost_price_value, shipping_fee_value, price_rules=None):
    cost_price = parse_number(cost_price_value)
    shipping_fee = parse_number(shipping_fee_value)

    if cost_price is None or shipping_fee is None:
        return None

    threshold, low_margin, high_margin = price_rules or (
        DEFAULT_PRICE_THRESHOLD,
        DEFAULT_LOW_PRICE_MARGIN,
        DEFAULT_HIGH_PRICE_MARGIN,
    )
    margin = high_margin if cost_price > threshold else low_margin
    divisor = 1 - margin
    return round((cost_price + shipping_fee) / divisor, 2)


def normalize_match_key(value):
    return str(value if value is not None else "").strip()


def normalize_price_cache_key(value):
    return normalize_match_key(value).upper()


def normalize_spec_code_for_price_lookup(value):
    spec_code = normalize_match_key(value)
    if "$$" in spec_code:
        spec_code = spec_code.split("$$", 1)[0]
    return normalize_price_cache_key(spec_code)


def resolve_price(row, header_map, cost_price_index, shipping_fee, price_rules, price_match_column, price_cache):
    match_key = normalize_price_cache_key(get_row_column_value(row, header_map, price_match_column))

    if match_key and match_key in price_cache:
        return price_cache[match_key], True, False

    cost_price_value = get_row_value(row, cost_price_index)
    price = calculate_price(cost_price_value, shipping_fee, price_rules)

    if match_key and price is not None:
        price_cache[match_key] = price
        return price, False, True

    return price, False, False
