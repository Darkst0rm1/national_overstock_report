"""
Orchestrates the full pipeline: validate -> load -> transform -> write.

This is the only module app.py talks to. Its job is to turn three uploaded
files into either a finished workbook or a list of specific, user-safe
error messages -- it never lets a raw exception (which could contain file
paths or internal details) reach the UI.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date

from openpyxl import load_workbook

from src import data_loader, validators
from src.allocation import aggregate_allocations, build_allocation
from src.eligibility import eligible_material_universe, filter_relevant_sales_orders
from src.excel_writer import build_workbook
from src.old_report import build_food_service_report, build_old_report
from src.overstock import build_overstock
from src.price_list import build_price_list


@dataclass
class ReportResult:
    success: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    workbook_bytes: bytes | None = None


def generate_report(materials_file, sales_order_file, pricing_file, report_date: date) -> ReportResult:
    errors: list[str] = []
    warnings: list[str] = []

    try:
        materials_df = data_loader.load_materials(data_loader.to_buffer(materials_file))
    except Exception:
        return ReportResult(success=False, errors=["Could not read the Materials export. Please confirm it is a valid, unprotected .xlsx file."])

    try:
        sales_order_df = data_loader.load_sales_order(data_loader.to_buffer(sales_order_file))
    except Exception:
        return ReportResult(success=False, errors=["Could not read the Open sales-order export. Please confirm it is a valid, unprotected .xlsx file."])

    try:
        pricing_df = data_loader.load_pricing(data_loader.to_buffer(pricing_file))
    except Exception:
        return ReportResult(success=False, errors=["Could not read the Material/pricing export. Please confirm it is a valid, unprotected .xlsx file."])

    errors += validators.validate_materials(materials_df)
    errors += validators.validate_sales_order(sales_order_df)
    pricing_errors, pricing_warnings = validators.validate_pricing(pricing_df)
    errors += pricing_errors
    warnings += pricing_warnings

    if errors:
        return ReportResult(success=False, errors=errors, warnings=warnings)

    try:
        # Sales orders for materials outside the eligible universe (food-
        # service-only BDMs, packaging SKUs, materials with no Price List
        # entry) must not affect the Allocation sheet or FIFO consumption --
        # see src/eligibility.py. The "Overstock" sheet itself stays the
        # complete, un-filtered export (its own long-standing spec).
        eligible_materials = eligible_material_universe(materials_df, pricing_df)
        relevant_sales_order_df = filter_relevant_sales_orders(sales_order_df, eligible_materials)

        allocation = build_allocation(relevant_sales_order_df)
        allocation_totals = aggregate_allocations(relevant_sales_order_df)
        overstock_df = build_overstock(sales_order_df)
        price_list_df = build_price_list(pricing_df)
        old_report_df = build_old_report(materials_df, pricing_df, allocation_totals, report_date)
        food_service_df = build_food_service_report(materials_df, pricing_df, allocation_totals)

        wb = build_workbook(overstock_df, price_list_df, old_report_df, allocation, food_service_df)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        workbook_bytes = buf.getvalue()

        # Re-open to confirm the workbook is well-formed before handing it
        # to the user -- catches structural mistakes (bad refs, corrupt
        # styles) before they turn into an Excel repair prompt.
        load_workbook(io.BytesIO(workbook_bytes))
    except Exception:
        return ReportResult(
            success=False,
            errors=["An error occurred while building the report. Please double-check the uploaded files match the expected format and try again."],
            warnings=warnings,
        )

    return ReportResult(success=True, errors=[], warnings=warnings, workbook_bytes=workbook_bytes)
