import pandas as pd

from src.allocation import build_allocation, selling_plant_column_letter


def _sales_orders(rows):
    return pd.DataFrame(rows, columns=["Material", "Plant", "Confirmed Quantity (CS)"])


def test_aggregates_by_material_and_plant():
    df = _sales_orders(
        [
            ["10000001", "2910", 5],
            ["10000001", "2910", 3],
            ["10000001", "2920", 2],
            ["10000002", "2920", 7],
        ]
    )
    result = build_allocation(df)

    assert result.materials == ["10000001", "10000002"]
    assert result.plants == ["2910", "2920"]
    assert result.grid["10000001"]["2910"] == 8
    assert result.grid["10000001"]["2920"] == 2
    assert result.grid["10000002"]["2910"] is None  # no rows for this combo -> blank, not 0
    assert result.grid["10000002"]["2920"] == 7


def test_grand_totals():
    df = _sales_orders(
        [
            ["10000001", "2910", 5],
            ["10000001", "2920", 2],
            ["10000002", "2920", 7],
        ]
    )
    result = build_allocation(df)

    assert result.row_totals["10000001"] == 7
    assert result.row_totals["10000002"] == 7
    assert result.col_totals["2910"] == 5
    assert result.col_totals["2920"] == 9
    assert result.grand_total == 14


def test_blank_material_and_plant_sort_last():
    df = _sales_orders(
        [
            ["10000005", "2910", 1],
            [None, "2910", 4],
            ["10000005", None, 9],
        ]
    )
    result = build_allocation(df)

    assert result.materials[-1] == "(blank)"
    assert result.plants[-1] == "(blank)"


def test_existing_zero_confirmed_qty_is_zero_not_blank():
    df = _sales_orders([["10000001", "2910", 0]])
    result = build_allocation(df)
    assert result.grid["10000001"]["2910"] == 0


def test_city_pairing_maps_3pl_plants_to_tol_plant_column():
    # Only TOL plants (2910/2920/2930) take direct orders; Lineage plants
    # (2925/2935) never appear in sales-order data but must resolve to
    # their city's TOL column.
    df = _sales_orders(
        [
            ["10000001", "2910", 1],
            ["10000001", "2920", 1],
            ["10000001", "2930", 1],
        ]
    )
    result = build_allocation(df)

    assert selling_plant_column_letter(result, "2910") == "B"
    assert selling_plant_column_letter(result, "2920") == "C"
    assert selling_plant_column_letter(result, "2930") == "D"
    assert selling_plant_column_letter(result, "2925") == "C"  # Lineage Calgary -> TOL Calgary
    assert selling_plant_column_letter(result, "2935") == "D"  # Lineage Surrey -> TOL Surrey


def test_selling_plant_with_no_order_data_returns_none():
    df = _sales_orders([["10000001", "2910", 1]])
    result = build_allocation(df)
    # 2920 never appears as an order plant in this run, so 2925 (which pairs
    # with 2920) has nothing to look up -- caller should fall back to a
    # literal 0 rather than reference a nonexistent column.
    assert selling_plant_column_letter(result, "2925") is None
