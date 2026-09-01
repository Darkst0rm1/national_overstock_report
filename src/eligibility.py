"""Determines which materials, inventory batches, and sales-order lines are
in scope for the National Overstock Report.

Reverse-engineered against the 2026-08-24 reference files -- see the rule
constants in src/config.py for what each check is and how it was verified.
The eligible MATERIAL universe (packaging prefix / excluded BDM / must have
a Price List row) is derived once from the full, date-unfiltered Materials
export, and both the inventory date-window filter and the sales-order filter
narrow from that same set. This mirrors the reference workbooks' own
relationship between "New Overstock" (377 sales-order lines) and "Old
Report" (271 inventory rows): a material that is only out of its own
shelf-life window still keeps the rest of its sales-order lines in scope,
since eligibility is a materials-catalog concept, not a per-batch one.
"""
from __future__ import annotations

from datetime import date, timedelta

import pandas as pd

from src import config
from src.normalize import normalize_identifier


def _bdm_by_material(pricing_df: pd.DataFrame) -> dict[str, object]:
    """First-match BDM per Material, mirroring Excel XLOOKUP's default
    top-to-bottom search order."""
    materials = pricing_df["Material"].map(normalize_identifier)
    work = pd.DataFrame({"_material": materials, "BDM": pricing_df["BDM"]})
    work = work.dropna(subset=["_material"]).drop_duplicates("_material", keep="first")
    return dict(zip(work["_material"], work["BDM"]))


def eligible_material_universe(materials_df: pd.DataFrame, pricing_df: pd.DataFrame) -> set[str]:
    """Materials that belong in the National Overstock Report at all,
    independent of any specific batch's expiration date."""
    bdm_by_material = _bdm_by_material(pricing_df)
    priced_materials = set(bdm_by_material)

    eligible: set[str] = set()
    for raw in materials_df["Material"]:
        material = normalize_identifier(raw)
        if material is None or material in eligible:
            continue
        if material.startswith(config.PACKAGING_MATERIAL_PREFIX):
            continue
        if material not in priced_materials:
            continue
        if bdm_by_material.get(material) in config.EXCLUDED_FOOD_SERVICE_BDMS:
            continue
        eligible.add(material)
    return eligible


def filter_eligible_inventory(
    materials_df: pd.DataFrame,
    eligible_materials: set[str],
    report_date: date,
) -> pd.DataFrame:
    """Inventory rows whose Material is in scope AND whose Shelf Life
    Expiration Date falls within [report_date, report_date + window] --
    the "New" band the National Overstock Report tracks (rather than
    already-expired or far-future stock)."""
    material = materials_df["Material"].map(normalize_identifier)
    in_scope = material.isin(eligible_materials)

    sled = pd.to_datetime(materials_df["Shelf Life Expiration Date"], errors="coerce")
    window_end = report_date + timedelta(days=config.INVENTORY_WINDOW_DAYS)
    in_window = sled.apply(lambda d: pd.notna(d) and report_date <= d.date() <= window_end)

    return materials_df[in_scope & in_window].copy()


def filter_relevant_sales_orders(
    sales_order_df: pd.DataFrame, eligible_materials: set[str]
) -> pd.DataFrame:
    """Sales-order lines for materials outside the eligible universe must
    not contribute to the Allocation sheet or FIFO consumption."""
    material = sales_order_df["Material"].map(normalize_identifier)
    return sales_order_df[material.isin(eligible_materials)].copy()


def filter_food_service_inventory(
    materials_df: pd.DataFrame, pricing_df: pd.DataFrame
) -> pd.DataFrame:
    """Inventory rows for materials handled by one of the excluded
    food-service BDMs (config.EXCLUDED_FOOD_SERVICE_BDMS) -- the "Foodservice"
    sheet's population, kept for future use. Unlike filter_eligible_inventory
    this applies NO shelf-life date window (full visibility of food-service
    stock, not just what's expiring soon) and no packaging/priced-material
    checks -- it is purely "was this row excluded from the National Overstock
    Report specifically because of the BDM rule."""
    bdm_by_material = _bdm_by_material(pricing_df)
    material = materials_df["Material"].map(normalize_identifier)
    bdm = material.map(bdm_by_material)
    return materials_df[bdm.isin(config.EXCLUDED_FOOD_SERVICE_BDMS)].copy()
