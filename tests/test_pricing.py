import pytest

from price_generator.pricing import (
    calculate_price,
    calculate_shipping_fee,
    normalize_spec_code_for_price_lookup,
)


@pytest.mark.parametrize(
    ("weight", "expected"),
    [
        (0, "请检查"),
        (0.29, 1.3),
        (0.3, 1.5),
        (0.5, 1.9),
        (1, 2.3),
        (1.5, 2.8),
        (2, 3.5),
        (2.5, 3.9),
        (3, 4.3),
        (4, 5.6),
        (5, 13.5),
    ],
)
def test_shipping_fee_weight_boundaries(weight, expected):
    assert calculate_shipping_fee(weight) == expected


def test_price_uses_low_and_high_margin_by_threshold():
    assert calculate_price(4.5, 1.3, (4.5, 0.13, 0.2)) == 6.67
    assert calculate_price(4.51, 1.3, (4.5, 0.13, 0.2)) == 7.26


def test_spec_code_lookup_is_case_insensitive_and_trims_suffix():
    assert normalize_spec_code_for_price_lookup("abc123$$red") == "ABC123"
