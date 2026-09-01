"""Consumes each (Material, allocation Plant)'s total confirmed order
quantity across that material/plant's inventory batches once, FIFO by
Shelf Life Expiration Date.

Fixes the prior bug of applying the *full* material/plant total to every one
of that material/plant's batches independently, which double- (or n-times-)
counted allocation and could drive Available negative (e.g. material
10026066 at plant 2910: batches of 91 and 340 units both getting the full
431-unit allocation applied, instead of 431 being split 91 then 340 across
them in expiration order).
"""
from __future__ import annotations

import pandas as pd

from src.allocation import map_allocation_plant
from src.normalize import normalize_identifier


def calculate_available_stock(unrestricted_stock: float, batch_allocation: float) -> float:
    """Available = Unrestricted Stock - Batch Allocation. Never negative as
    long as batch_allocation was capped to unrestricted_stock (see
    allocate_fifo_by_batch)."""
    return unrestricted_stock - batch_allocation


def allocate_fifo_by_batch(
    inventory_df: pd.DataFrame, allocation_totals: dict[tuple[str, str], float]
) -> pd.DataFrame:
    """Adds Unique Key, Allocated, and Available columns to inventory_df
    (which must have Material, Plant, Batch, Shelf Life Expiration Date, and
    Unrestricted Stock).

    Each (Material, mapped allocation Plant) key's total confirmed quantity
    in allocation_totals (see allocation.aggregate_allocations) is consumed
    once, batch by batch, sorted by earliest Shelf Life Expiration Date
    first with Batch as a stable tie-breaker -- never applied in full to
    every batch. Row order in the returned DataFrame matches the input
    (the FIFO sort is internal to the consumption pass only).
    """
    work = inventory_df.copy()
    material = work["Material"].map(normalize_identifier)
    alloc_plant = work["Plant"].map(map_allocation_plant)
    sled_sort = pd.to_datetime(work["Shelf Life Expiration Date"], errors="coerce")

    consumption_order = pd.DataFrame(
        {"_material": material, "_alloc_plant": alloc_plant, "_sled_sort": sled_sort, "Batch": work["Batch"]}
    ).sort_values(["_material", "_alloc_plant", "_sled_sort", "Batch"], kind="stable")

    remaining = dict(allocation_totals)
    batch_allocation = pd.Series(0.0, index=work.index)
    for idx in consumption_order.index:
        key = (consumption_order.at[idx, "_material"], consumption_order.at[idx, "_alloc_plant"])
        remaining_for_key = max(remaining.get(key, 0.0), 0.0)
        stock = float(work.at[idx, "Unrestricted Stock"] or 0)
        allocated = min(remaining_for_key, stock)
        batch_allocation.at[idx] = allocated
        remaining[key] = remaining_for_key - allocated

    work["Unique Key"] = material + "_" + alloc_plant
    work["Allocated"] = batch_allocation
    work["Available"] = [
        calculate_available_stock(float(s or 0), a)
        for s, a in zip(work["Unrestricted Stock"], work["Allocated"])
    ]
    return work
