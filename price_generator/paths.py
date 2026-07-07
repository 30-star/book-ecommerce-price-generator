import os
import sys

from .constants import PRICE_CACHE_FILE_NAME


def app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def resource_path(filename):
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(app_base_dir(), filename)


def default_rules_path():
    return os.path.join(app_base_dir(), "rules.json")


def default_price_cache_path():
    data_root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    data_dir = os.path.join(data_root, "图书电商线上活动价格自动生成器")
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, PRICE_CACHE_FILE_NAME)
