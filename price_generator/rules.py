import json
import os
import re

from .constants import DEFAULT_DELETE_RULES
from .paths import default_rules_path


def is_blank(value):
    return str(value if value is not None else "").strip() == ""


def is_zero_or_blank(value):
    if is_blank(value):
        return True

    try:
        return float(str(value).strip().replace(",", "")) == 0
    except ValueError:
        return False


def parse_number(value):
    if is_blank(value):
        return None

    try:
        return float(str(value).strip().replace(",", ""))
    except ValueError:
        return None


def get_row_value(row, index):
    return row[index] if index != -1 and index < len(row) else None


def text_contains(value, text):
    return text in str(value if value is not None else "")


def text_contains_case_insensitive(value, text):
    return text.lower() in str(value if value is not None else "").lower()


def contains_english_letter(value):
    return any("A" <= char <= "Z" or "a" <= char <= "z" for char in str(value if value is not None else ""))


def has_four_digit_number_less_than(value, threshold):
    text = str(value if value is not None else "")
    return any(int(match) < threshold for match in re.findall(r"\d{4}", text))


def load_delete_rules(path=None):
    rules_path = path or default_rules_path()

    if not rules_path or not os.path.exists(rules_path):
        return DEFAULT_DELETE_RULES

    with open(rules_path, "r", encoding="utf-8") as rules_file:
        data = json.load(rules_file)

    rules = data.get("delete_rules", data if isinstance(data, list) else None)
    if not isinstance(rules, list):
        raise ValueError("规则文件格式错误：需要 delete_rules 列表。")

    return rules


def save_default_rules(path):
    with open(path, "w", encoding="utf-8") as rules_file:
        json.dump({"delete_rules": DEFAULT_DELETE_RULES}, rules_file, ensure_ascii=False, indent=2)


def compare_number(value, operator, expected):
    number = parse_number(value)
    if number is None:
        return False

    if operator == "lt":
        return number < expected
    if operator == "lte":
        return number <= expected
    if operator == "gt":
        return number > expected
    if operator == "gte":
        return number >= expected
    if operator == "eq":
        return number == expected

    raise ValueError(f"不支持的数字比较：{operator}")


def evaluate_condition(row, header_map, condition):
    from .headers import get_row_column_value

    column_name = condition.get("column")
    operator = condition.get("operator")
    expected = condition.get("value")
    value = get_row_column_value(row, header_map, column_name)

    if operator == "blank":
        return is_blank(value)
    if operator == "not_blank":
        return not is_blank(value)
    if operator == "zero_or_blank":
        return is_zero_or_blank(value)
    if operator == "contains":
        return text_contains(value, str(expected))
    if operator == "contains_ci":
        return text_contains_case_insensitive(value, str(expected))
    if operator == "no_english_letter":
        return not contains_english_letter(value)
    if operator == "four_digit_lt":
        return has_four_digit_number_less_than(value, float(expected))
    if operator in {"lt", "lte", "gt", "gte", "eq"}:
        return compare_number(value, operator, float(expected))

    raise ValueError(f"不支持的规则操作：{operator}")


def should_delete_row(row, header_map, delete_rules):
    for rule in delete_rules:
        conditions = rule.get("conditions", [])
        any_conditions = rule.get("any", [])

        if conditions and all(evaluate_condition(row, header_map, condition) for condition in conditions):
            return True
        if any_conditions and any(evaluate_condition(row, header_map, condition) for condition in any_conditions):
            return True

    return False
