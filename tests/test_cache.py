from openpyxl import load_workbook

from price_generator.cache import export_price_cache_xlsx, load_price_cache, save_price_cache


def test_price_cache_save_load_normalizes_keys(tmp_path):
    cache_path = tmp_path / "price_cache.json"
    save_price_cache({"abc": 12.3, "": 99}, cache_path)

    assert load_price_cache(cache_path) == {"ABC": 12.3}


def test_price_cache_export_xlsx(tmp_path):
    output_path = tmp_path / "cache.xlsx"
    count = export_price_cache_xlsx(output_path, {"ABC": 12.3})

    assert count == 1
    workbook = load_workbook(output_path, read_only=True)
    try:
        rows = list(workbook.active.iter_rows(values_only=True))
    finally:
        workbook.close()

    assert rows == [("商品编码", "售价"), ("ABC", 12.3)]
