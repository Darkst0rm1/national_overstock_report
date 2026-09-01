from datetime import date

import pandas as pd

from src.old_report import build_food_service_report, build_old_report

REPORT_DATE = date(2026, 9, 1)


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


def test_preserves_every_eligible_materials_row_including_duplicates():
    materials_df = pd.DataFrame([_materials_row(), _materials_row()])
    pricing_df = pd.DataFrame([_pricing_row()])

    out = build_old_report(materials_df, pricing_df, {}, REPORT_DATE)
    assert len(out) == 2


def test_drops_rows_for_ineligible_materials():
    materials_df = pd.DataFrame([_materials_row(Material="50016409")])
    pricing_df = pd.DataFrame([_pricing_row(Material="50016409")])

    out = build_old_report(materials_df, pricing_df, {}, REPORT_DATE)
    assert len(out) == 0


def test_drops_rows_outside_the_shelf_life_window():
    materials_df = pd.DataFrame(
        [_materials_row(**{"Shelf Life Expiration Date": pd.Timestamp("2026-01-01")})]
    )
    pricing_df = pd.DataFrame([_pricing_row()])

    out = build_old_report(materials_df, pricing_df, {}, REPORT_DATE)
    assert len(out) == 0


def test_applies_fifo_allocation_via_mapped_plant():
    materials_df = pd.DataFrame(
        [
            _materials_row(Material="10000001", Plant="2925"),  # Lineage Calgary
            _materials_row(Material="10000002", Plant="2920"),  # TOL Calgary
        ]
    )
    pricing_df = pd.DataFrame(
        [_pricing_row(Material="10000001"), _pricing_row(Material="10000002")]
    )
    allocation_totals = {("10000001", "2920"): 3, ("10000002", "2920"): 4}

    out = build_old_report(materials_df, pricing_df, allocation_totals, REPORT_DATE).set_index("Material")

    assert out.loc["10000001", "Allocated"] == 3  # 2925 pools with 2920
    assert out.loc["10000002", "Allocated"] == 4


def test_food_service_report_is_the_complement_of_old_report_by_bdm():
    materials_df = pd.DataFrame(
        [
            _materials_row(Material="10026066", Batch="FOOD_SERVICE"),
            _materials_row(Material="10019604", Batch="RETAIL"),
        ]
    )
    pricing_df = pd.DataFrame(
        [_pricing_row(Material="10026066", BDM="MIGUEL GUTIERREZ"), _pricing_row(Material="10019604")]
    )

    old_report = build_old_report(materials_df, pricing_df, {}, REPORT_DATE)
    food_service = build_food_service_report(materials_df, pricing_df, {})

    assert list(old_report["Batch"]) == ["RETAIL"]
    assert list(food_service["Batch"]) == ["FOOD_SERVICE"]


def test_food_service_report_has_no_date_window():
    materials_df = pd.DataFrame(
        [_materials_row(Material="10026066", **{"Shelf Life Expiration Date": pd.Timestamp("2020-01-01")})]
    )
    pricing_df = pd.DataFrame([_pricing_row(Material="10026066", BDM="Karishma Salian")])

    out = build_food_service_report(materials_df, pricing_df, {})
    assert len(out) == 1
