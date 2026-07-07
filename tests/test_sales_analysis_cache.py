import pytest

from price_generator.sales_analysis_cache import load_sales_analysis_cache, save_sales_analysis_cache


def test_sales_analysis_cache_save_load(tmp_path):
    cache_path = tmp_path / "sales_analysis_cache.json"
    save_sales_analysis_cache(
        {
            " A001 ": {
                "order_weight": 1.2,
                "shipping_fee": 2.3,
                "sales_amount": 99.9,
                "sales_cost": 60.5,
                "gross_profit": 37.1,
            },
            "": {"order_weight": 9, "shipping_fee": 9, "sales_amount": 9, "sales_cost": 9, "gross_profit": -9},
        },
        cache_path,
    )

    assert load_sales_analysis_cache(cache_path) == {
        "A001": {
            "order_weight": 1.2,
            "shipping_fee": 2.3,
            "sales_amount": 99.9,
            "sales_cost": 60.5,
            "gross_profit": 37.1,
        }
    }


def test_sales_analysis_cache_rejects_invalid_record(tmp_path):
    cache_path = tmp_path / "sales_analysis_cache.json"
    cache_path.write_text(
        '{"orders": {"A001": {"order_weight": 1.2, "shipping_fee": 2.3, "sales_amount": 99.9, "sales_cost": 60.5}}}',
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="gross_profit"):
        load_sales_analysis_cache(cache_path)
