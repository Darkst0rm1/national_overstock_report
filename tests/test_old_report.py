import pandas as pd

from src.allocation import build_allocation
from src.old_report import ALLOCATION_COLUMN_LETTER_FIELD, build_old_report


def _materials_row(**overrides):
    row = {
        "Material Group Name": "AGROPUR",
        "Material": "10019604",
        "Material Description": "OKA CHEESE ORIGINAL 6X190G",
        "Plant": "2910",
        "Plant Name": "TOL Mississauga",
        "Batch": "05641",
        "Storage Location": "1100",
        "Shelf Life Expiration Date": pd.Timestamp("2026-08-19"),
        "Unrestricted Stock": 6,
    }
    row.update(overrides)
    return row


def test_preserves_every_materials_row_including_duplicates():
    materials_df = pd.DataFrame([_materials_row(), _materials_row()])
    sales_order_df = pd.DataFrame(columns=["Material", "Plant", "Confirmed Quantity (CS)"])
    allocation = build_allocation(sales_order_df)

    out = build_old_report(materials_df, allocation)
    assert len(out) == 2


def test_resolves_allocation_column_letter_via_city_pairing():
    materials_df = pd.DataFrame(
        [
            _materials_row(Material="10000001", Plant="2925"),  # Lineage Calgary
            _materials_row(Material="10000002", Plant="2920"),  # TOL Calgary
        ]
    )
    sales_order_df = pd.DataFrame(
        [["10000001", "2920", 3], ["10000002", "2920", 3]],
        columns=["Material", "Plant", "Confirmed Quantity (CS)"],
    )
    allocation = build_allocation(sales_order_df)
    out = build_old_report(materials_df, allocation)

    letters = out.set_index("Material")[ALLOCATION_COLUMN_LETTER_FIELD]
    assert letters["10000001"] == letters["10000002"]  # both resolve to the 2920 column
