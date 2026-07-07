import re

from .constants import COLUMN_ALIASES
from .rules import get_row_value


def normalize_header_name(value):
    text = str(value if value is not None else "").strip()
    text = text.replace("\ufeff", "").replace("\u200b", "").replace("\u200c", "").replace("\u200d", "")
    return re.sub(r"[\s\u00a0\u3000:：()（）\[\]【】{}｛｝_-]+", "", text).lower()


def find_header_index(headers, name):
    normalized = [str(value if value is not None else "").strip() for value in headers]
    compact_headers = [normalize_header_name(value) for value in headers]
    names = [name] + COLUMN_ALIASES.get(name, [])
    compact_names = [normalize_header_name(value) for value in names]

    for index, header in enumerate(normalized):
        for candidate in names:
            if header == candidate:
                return index

    for index, header in enumerate(normalized):
        for candidate in names:
            if candidate in header:
                return index

    for index, header in enumerate(compact_headers):
        for candidate in compact_names:
            if header == candidate or candidate in header:
                return index

    return -1


def build_header_map(headers):
    return {str(value if value is not None else "").strip(): index for index, value in enumerate(headers)}


def get_row_column_value(row, header_map, column_name):
    index = header_map.get(column_name)
    if index is None:
        for header, header_index in header_map.items():
            if column_name in header:
                index = header_index
                break

    return get_row_value(row, index) if index is not None else None
