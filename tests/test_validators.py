import pandas as pd

from src import config, validators


def test_validate_materials_reports_missing_columns():
    df = pd.DataFrame(columns=["Material", "Plant"])
    errors = validators.validate_materials(df)
    assert len(errors) == 1
    for col in config.MATERIALS_REQUIRED_COLUMNS:
        if col not in ("Material", "Plant"):
            assert col in errors[0]


def test_validate_materials_passes_when_complete():
    df = pd.DataFrame(columns=config.MATERIALS_REQUIRED_COLUMNS + ["Extra Column"])
    assert validators.validate_materials(df) == []


def test_validate_sales_order_reports_missing_columns():
    df = pd.DataFrame(columns=["Sales Order"])
    errors = validators.validate_sales_order(df)
    assert len(errors) == 1
    assert "Material" in errors[0]


def test_validate_pricing_only_material_is_hard_required():
    df = pd.DataFrame(columns=["Material", "BDM", "Brand Name"])
    errors, warnings = validators.validate_pricing(df)
    assert errors == []
    assert len(warnings) == 1
    assert "Pack" in warnings[0]
    assert "Sold by Wt" in warnings[0]


def test_validate_pricing_blocks_when_material_missing():
    df = pd.DataFrame(columns=["BDM"])
    errors, warnings = validators.validate_pricing(df)
    assert len(errors) == 1
    assert "Material" in errors[0]


def test_validate_pricing_no_warnings_when_all_present():
    df = pd.DataFrame(columns=config.PRICE_LIST_COLUMNS)
    errors, warnings = validators.validate_pricing(df)
    assert errors == []
    assert warnings == []
