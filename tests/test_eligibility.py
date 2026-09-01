from datetime import date

import pandas as pd

from src.eligibility import (
    eligible_material_universe,
    filter_eligible_inventory,
    filter_food_service_inventory,
    filter_relevant_sales_orders,
)


def _materials_row(**overrides):
    row = {
        "Material": "10019604",
        "Material Description": "OKA CHEESE ORIGINAL 6X190G",
        "Plant": "2910",
        "Plant Name": "TOL Mississauga",
        "Batch": "05641",
        "Storage Location": "1100",
        "Shelf Life Expiration Date": pd.Timestamp("2026-09-01"),
        "Unrestricted Stock": 6,
    }
    row.update(overrides)
    return row


def _pricing_row(**overrides):
    row = {"Material": "10019604", "BDM": "SERENA LEE PR"}
    row.update(overrides)
    return row


def test_excludes_packaging_prefixed_materials():
    materials_df = pd.DataFrame([_materials_row(Material="50016409")])
    pricing_df = pd.DataFrame([_pricing_row(Material="50016409")])
    assert eligible_material_universe(materials_df, pricing_df) == set()


def test_excludes_food_service_bdms():
    materials_df = pd.DataFrame([_materials_row()])
    pricing_df = pd.DataFrame([_pricing_row(BDM="MIGUEL GUTIERREZ")])
    assert eligible_material_universe(materials_df, pricing_df) == set()


def test_excludes_materials_with_no_price_list_row():
    materials_df = pd.DataFrame([_materials_row()])
    pricing_df = pd.DataFrame([_pricing_row(Material="99999999")])
    assert eligible_material_universe(materials_df, pricing_df) == set()


def test_includes_material_that_passes_all_three_rules():
    materials_df = pd.DataFrame([_materials_row()])
    pricing_df = pd.DataFrame([_pricing_row()])
    assert eligible_material_universe(materials_df, pricing_df) == {"10019604"}


def test_normalizes_numeric_material_ids_before_matching():
    materials_df = pd.DataFrame([_materials_row(Material=10019604.0)])
    pricing_df = pd.DataFrame([_pricing_row(Material="10019604")])
    assert eligible_material_universe(materials_df, pricing_df) == {"10019604"}


def test_bdm_lookup_uses_first_match_like_excel_xlookup():
    materials_df = pd.DataFrame([_materials_row()])
    pricing_df = pd.DataFrame(
        [_pricing_row(BDM="SERENA LEE PR"), _pricing_row(BDM="MIGUEL GUTIERREZ")]
    )
    # First Price List row for this material is the eligible BDM -> included.
    assert eligible_material_universe(materials_df, pricing_df) == {"10019604"}


def test_filter_eligible_inventory_applies_date_window_dynamically():
    report_date = date(2026, 9, 1)
    materials_df = pd.DataFrame(
        [
            _materials_row(Batch="TOO_EARLY", **{"Shelf Life Expiration Date": pd.Timestamp("2026-08-31")}),
            _materials_row(Batch="IN_WINDOW", **{"Shelf Life Expiration Date": pd.Timestamp("2026-09-01")}),
            _materials_row(Batch="STILL_IN_WINDOW", **{"Shelf Life Expiration Date": pd.Timestamp("2026-12-03")}),
            _materials_row(Batch="TOO_LATE", **{"Shelf Life Expiration Date": pd.Timestamp("2026-12-04")}),
        ]
    )
    eligible_materials = {"10019604"}

    out = filter_eligible_inventory(materials_df, eligible_materials, report_date)

    assert set(out["Batch"]) == {"IN_WINDOW", "STILL_IN_WINDOW"}


def test_filter_eligible_inventory_uses_a_different_window_for_a_different_report_date():
    # Same batches, only the report date changes -- proves the window is
    # dynamic, not hardcoded to a specific run's date.
    materials_df = pd.DataFrame(
        [
            _materials_row(Batch="A", **{"Shelf Life Expiration Date": pd.Timestamp("2026-10-01")}),
        ]
    )
    eligible_materials = {"10019604"}

    assert len(filter_eligible_inventory(materials_df, eligible_materials, date(2026, 9, 1))) == 1
    assert len(filter_eligible_inventory(materials_df, eligible_materials, date(2026, 11, 1))) == 0


def test_filter_eligible_inventory_excludes_ineligible_materials():
    materials_df = pd.DataFrame([_materials_row(Material="50016409")])
    out = filter_eligible_inventory(materials_df, {"10019604"}, date(2026, 9, 1))
    assert len(out) == 0


def test_filter_relevant_sales_orders_keeps_only_eligible_materials():
    sales_order_df = pd.DataFrame(
        [
            {"Sales Order": "1", "Material": "10019604"},
            {"Sales Order": "2", "Material": "10026066"},
        ]
    )
    out = filter_relevant_sales_orders(sales_order_df, {"10019604"})
    assert list(out["Sales Order"]) == ["1"]


def test_filter_relevant_sales_orders_normalizes_before_matching():
    sales_order_df = pd.DataFrame([{"Sales Order": "1", "Material": "10019604.0"}])
    out = filter_relevant_sales_orders(sales_order_df, {"10019604"})
    assert len(out) == 1


def test_filter_food_service_inventory_keeps_only_excluded_bdms():
    materials_df = pd.DataFrame(
        [_materials_row(Material="10026066", Batch="RANA"), _materials_row(Material="10019604", Batch="RETAIL")]
    )
    pricing_df = pd.DataFrame(
        [_pricing_row(Material="10026066", BDM="MIGUEL GUTIERREZ"), _pricing_row(Material="10019604", BDM="SERENA LEE PR")]
    )
    out = filter_food_service_inventory(materials_df, pricing_df)
    assert list(out["Batch"]) == ["RANA"]


def test_filter_food_service_inventory_ignores_date_window():
    # Rule 3's sheet is "for future use" -- full visibility, no shelf-life
    # window, unlike filter_eligible_inventory.
    materials_df = pd.DataFrame(
        [
            _materials_row(Material="10026066", Batch="LONG_EXPIRED", **{"Shelf Life Expiration Date": pd.Timestamp("2020-01-01")}),
            _materials_row(Material="10026066", Batch="FAR_FUTURE", **{"Shelf Life Expiration Date": pd.Timestamp("2030-01-01")}),
        ]
    )
    pricing_df = pd.DataFrame([_pricing_row(Material="10026066", BDM="Karishma Salian")])
    out = filter_food_service_inventory(materials_df, pricing_df)
    assert set(out["Batch"]) == {"LONG_EXPIRED", "FAR_FUTURE"}


def test_filter_food_service_inventory_excludes_regular_bdms():
    materials_df = pd.DataFrame([_materials_row()])
    pricing_df = pd.DataFrame([_pricing_row()])  # SERENA LEE PR, not excluded
    out = filter_food_service_inventory(materials_df, pricing_df)
    assert len(out) == 0
