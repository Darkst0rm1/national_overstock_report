import pandas as pd

from src.fifo_allocation import allocate_fifo_by_batch, calculate_available_stock


def _inventory_row(**overrides):
    row = {
        "Material": "10026066",
        "Plant": "2910",
        "Batch": "B1",
        "Shelf Life Expiration Date": pd.Timestamp("2026-09-01"),
        "Unrestricted Stock": 100,
    }
    row.update(overrides)
    return row


def test_calculate_available_stock_is_stock_minus_allocation():
    assert calculate_available_stock(100, 40) == 60


def test_single_batch_receives_allocation_once():
    inventory_df = pd.DataFrame([_inventory_row(**{"Unrestricted Stock": 50})])
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2910"): 20})
    assert list(out["Allocated"]) == [20]
    assert list(out["Available"]) == [30]


def test_multiple_batches_receive_fifo_allocation_not_full_broadcast():
    # Reproduces the reported bug case: material 10026066, plant 2910,
    # batch1=91, batch2=340, total allocation=431 -- consumed once across
    # both batches (earliest expiration first), never applied to each in full.
    inventory_df = pd.DataFrame(
        [
            _inventory_row(Batch="EARLY", **{"Shelf Life Expiration Date": pd.Timestamp("2026-08-25")}, **{"Unrestricted Stock": 91}),
            _inventory_row(Batch="LATE", **{"Shelf Life Expiration Date": pd.Timestamp("2026-09-10")}, **{"Unrestricted Stock": 340}),
        ]
    )
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2910"): 431})
    out = out.set_index("Batch")

    assert out.loc["EARLY", "Allocated"] == 91
    assert out.loc["EARLY", "Available"] == 0
    assert out.loc["LATE", "Allocated"] == 340
    assert out.loc["LATE", "Available"] == 0
    assert out["Allocated"].sum() == 431  # consumed once, not double-counted


def test_allocation_equal_to_total_stock_leaves_zero_available():
    inventory_df = pd.DataFrame([_inventory_row(**{"Unrestricted Stock": 50})])
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2910"): 50})
    assert out.iloc[0]["Allocated"] == 50
    assert out.iloc[0]["Available"] == 0


def test_allocation_below_total_stock_leaves_correct_available():
    inventory_df = pd.DataFrame([_inventory_row(**{"Unrestricted Stock": 50})])
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2910"): 30})
    assert out.iloc[0]["Allocated"] == 30
    assert out.iloc[0]["Available"] == 20


def test_allocation_greater_than_total_stock_does_not_go_negative():
    inventory_df = pd.DataFrame([_inventory_row(**{"Unrestricted Stock": 50})])
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2910"): 999})
    assert out.iloc[0]["Allocated"] == 50
    assert out.iloc[0]["Available"] == 0


def test_no_allocation_entry_defaults_to_zero():
    inventory_df = pd.DataFrame([_inventory_row()])
    out = allocate_fifo_by_batch(inventory_df, {})
    assert out.iloc[0]["Allocated"] == 0
    assert out.iloc[0]["Available"] == 100


def test_unrelated_materials_do_not_affect_the_result():
    inventory_df = pd.DataFrame(
        [_inventory_row(Material="10026066"), _inventory_row(Material="99999999", Batch="B2")]
    )
    out = allocate_fifo_by_batch(
        inventory_df, {("10026066", "2910"): 10, ("11111111", "2910"): 500}
    )
    out = out.set_index("Material")
    assert out.loc["10026066", "Allocated"] == 10
    assert out.loc["99999999", "Allocated"] == 0


def test_plant_2925_pools_with_2920_for_allocation():
    inventory_df = pd.DataFrame([_inventory_row(Plant="2925")])
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2920"): 40})
    assert out.iloc[0]["Allocated"] == 40


def test_plant_2935_pools_with_2930_for_allocation():
    inventory_df = pd.DataFrame([_inventory_row(Plant="2935", **{"Unrestricted Stock": 60})])
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2930"): 60})
    assert out.iloc[0]["Allocated"] == 60


def test_unique_key_uses_the_rows_own_mapped_plant_not_a_constant():
    inventory_df = pd.DataFrame(
        [_inventory_row(Plant="2930"), _inventory_row(Plant="2925", Batch="B2")]
    )
    out = allocate_fifo_by_batch(inventory_df, {})
    keys = dict(zip(out["Plant"], out["Unique Key"]))
    assert keys["2930"] == "10026066_2930"
    assert keys["2925"] == "10026066_2920"  # mapped, not a hardcoded "_2910"


def test_output_row_order_matches_input_order():
    inventory_df = pd.DataFrame(
        [
            _inventory_row(Batch="LATE", **{"Shelf Life Expiration Date": pd.Timestamp("2026-12-01")}),
            _inventory_row(Batch="EARLY", **{"Shelf Life Expiration Date": pd.Timestamp("2026-08-25")}),
        ]
    )
    out = allocate_fifo_by_batch(inventory_df, {("10026066", "2910"): 10})
    assert list(out["Batch"]) == ["LATE", "EARLY"]
