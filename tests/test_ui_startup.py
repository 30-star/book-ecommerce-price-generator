from price_generator import ui


def test_ui_imports_default_rules_path():
    assert callable(ui.default_rules_path)
