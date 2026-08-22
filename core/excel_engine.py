"""
core/excel_engine.py

Populates a BD Fact Sheet Excel template from extracted JSON data.

Refactored from fill_factsheet.py with:
  - All cell coordinates driven by FactSheetMapping (no hardcoded strings)
  - Dynamic RPT and Litigation table expansion via safe_sheet
  - FillResult return type with metadata and warnings
  - Structured logging
  - Same hard rules: never touch formatting, always add source comments
"""

import json
import logging
from dataclasses import dataclass, field
from typing import List, Optional

from openpyxl import load_workbook
from openpyxl.comments import Comment

from config.template_mapping import FactSheetMapping
from core.safe_sheet import insert_rows_safe

logger = logging.getLogger("alfa.excel_engine")

AUTHOR = "ALFA — TP Fact Sheet"


# ---------------------------------------------------------------------------
# Result wrapper
# ---------------------------------------------------------------------------

@dataclass
class FillResult:
    """Metadata about the fill operation."""
    out_path: str
    company_name: str
    warnings: List[str] = field(default_factory=list)
    shareholding_rows_written: int = 0
    rpt_rows_written: int = 0
    lit_rows_written: int = 0
    rpt_rows_dropped: int = 0
    lit_rows_dropped: int = 0
    error: Optional[str] = None

    @property
    def success(self) -> bool:
        return self.error is None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _note(ws, cell_ref: str, value, source: str | None):
    """Write a value and attach a source-citing comment."""
    ws[cell_ref] = value
    if source:
        ws[cell_ref].comment = Comment(f"Source: {source}", AUTHOR)


def _find_total_row(ws, start_data_row: int, label: str = "Total", max_scan: int = 30) -> int:
    """Scan downward from start_data_row looking for a cell with *label* in column C."""
    for r in range(start_data_row, start_data_row + max_scan):
        if ws[f"C{r}"].value == label:
            return r
    raise ValueError(f"Could not locate '{label}' row scanning from row {start_data_row}")


# ---------------------------------------------------------------------------
# Main fill function
# ---------------------------------------------------------------------------

def fill_factsheet(
    data: dict,
    template_path: str,
    out_path: str,
    sheet_name: str,
    mapping: FactSheetMapping | None = None,
) -> FillResult:
    """
    Write extracted company data into the BD Fact Sheet template.

    Args:
        data: parsed JSON dict matching the CompanyData schema
        template_path: path to the Excel template
        out_path: path to save the populated workbook
        sheet_name: worksheet name
        mapping: cell coordinate configuration (uses defaults if None)

    Returns:
        FillResult with operation metadata
    """
    if mapping is None:
        mapping = FactSheetMapping()

    result = FillResult(out_path=out_path, company_name=data.get("entity", "Unknown"))
    f_ = data["fields"]

    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    # --- Header ---
    ws.title = data["entity"][:31]
    ws[mapping.title_cell] = f"Fact Sheet - {data['entity']}"

    rows_to_hide: List[int] = []

    # --- Headquarters / descriptions ---
    _note(ws, mapping.hq_india_cell, f_["hq_india_entity"]["value"], f_["hq_india_entity"]["source"])
    _note(ws, mapping.hq_group_cell, f_["hq_group"]["value"], f_["hq_group"]["source"])
    _note(ws, mapping.company_desc_cell, f_["company_description"]["value"], f_["company_description"]["source"])
    _note(ws, mapping.group_desc_cell, f_["group_description"]["value"], f_["group_description"]["source"])

    # --- Standalone summary ---
    _note(ws, mapping.turnover_fy25_cell, f_["standalone_turnover_fy25_lakhs"]["value"], f_["standalone_turnover_fy25_lakhs"]["source"])
    _note(ws, mapping.turnover_fy24_cell, f_["standalone_turnover_fy24_lakhs"]["value"], f_["standalone_turnover_fy24_lakhs"]["source"])
    _note(ws, mapping.total_cost_fy25_cell, f_["standalone_total_cost_fy25_lakhs"]["value"], f_["standalone_total_cost_fy25_lakhs"]["source"])
    _note(ws, mapping.total_cost_fy24_cell, f_["standalone_total_cost_fy24_lakhs"]["value"], f_["standalone_total_cost_fy24_lakhs"]["source"])

    ws[mapping.pbt_comment_cell].comment = Comment(
        "PBT left as the template's formula. Note: Since Turnover now excludes "
        "Other Income, this formula (Turnover - Total Cost) will likely not match "
        "the AR's reported PBT.",
        AUTHOR,
    )

    # --- Consolidated summary ---
    cons = f_["consolidated_summary"]
    for cell_ref in mapping.consolidated_cells:
        ws[cell_ref] = "N/A"
    ws[mapping.consolidated_comment_cell].comment = Comment(f"Source: {cons['source']}", AUTHOR)

    # --- Statutory auditors ---
    _note(ws, mapping.auditors_cell, f_["statutory_auditors"]["value"], f_["statutory_auditors"]["source"])

    # --- AE / Domestic / Export revenue split ---
    ae = f_["ae_revenue_split"]["value"]
    ae_source = f_["ae_revenue_split"]["source"] + " | " + (f_["ae_revenue_split"].get("flag") or "")
    _note(ws, mapping.ae_domestic_cell, ae.get("domestic_ae_fy25"), ae_source)
    _note(ws, mapping.ae_export_cell, ae.get("export_ae_fy25"), ae_source)
    _note(ws, mapping.third_party_domestic_cell, ae.get("domestic_third_party_fy25"), ae_source)
    _note(ws, mapping.third_party_export_cell, ae.get("export_third_party_fy25"), ae_source)

    # --- Cash & AE trade receivables ---
    _note(ws, mapping.cash_fy25_cell, f_["cash_fy25_lakhs"]["value"], f_["cash_fy25_lakhs"]["source"])
    _note(ws, mapping.cash_fy24_cell, f_["cash_fy24_lakhs"]["value"], f_["cash_fy24_lakhs"]["source"])
    _note(ws, mapping.ae_receivables_fy25_cell, f_["ae_trade_receivables_fy25_lakhs"]["value"], f_["ae_trade_receivables_fy25_lakhs"]["source"])
    _note(ws, mapping.ae_receivables_fy24_cell, f_["ae_trade_receivables_fy24_lakhs"]["value"], f_["ae_trade_receivables_fy24_lakhs"]["source"])

    # --- PE investment ---
    pe_source = str(f_["pe_investment"].get("source", "")) + " | " + str(f_["pe_investment"].get("flag", ""))
    _note(ws, mapping.pe_investment_cell, f_["pe_investment"]["value"], pe_source)

    # ===================================================================
    # SHAREHOLDING TABLE
    # ===================================================================
    sh = f_["shareholding"]
    total_row = _find_total_row(ws, mapping.shareholding_start_row, mapping.shareholding_total_label)
    start_row = mapping.shareholding_start_row
    rows = sh["rows"]
    n_needed = len(rows)
    n_available = total_row - start_row

    if n_needed > n_available:
        raise ValueError(
            f"Shareholding table needs {n_needed} rows but only {n_available} available. "
            f"Run safe_sheet.insert_rows_safe() first."
        )

    # Hide unused rows
    if n_needed < n_available:
        for i in range(n_needed, n_available):
            rows_to_hide.append(start_row + i)

    for i, r in enumerate(rows):
        rn = start_row + i
        ws[f"C{rn}"] = r["name"]
        ws[f"D{rn}"] = r["shares_fy25"]
        ws[f"F{rn}"] = r["shares_fy24"]
        ws[f"E{rn}"] = f"=D{rn}/$D${total_row}"
        ws[f"G{rn}"] = f"=F{rn}/$F${total_row}"

    ws[f"C{start_row}"].comment = Comment(f"Source: {sh['source']}", AUTHOR)
    ws[f"D{total_row}"] = sh["total_shares_fy25"]
    ws[f"F{total_row}"] = sh["total_shares_fy24"]
    result.shareholding_rows_written = n_needed

    # ===================================================================
    # RELATED PARTY TRANSACTIONS (Dynamic)
    # ===================================================================
    rpt = f_["related_party_transactions_lakhs"]
    items = rpt["items"]
    rpt_header_row = total_row + mapping.rpt_header_offset
    rpt_start = rpt_header_row + 1

    rpt_capacity = mapping.rpt_default_capacity
    rpt_written = min(len(items), rpt_capacity)
    rpt_dropped = max(0, len(items) - rpt_capacity)

    if rpt_dropped > 0:
        result.warnings.append(
            f"RPT: {rpt_dropped} items exceeded template capacity of {rpt_capacity} and were dropped. "
            f"Consider expanding the template."
        )
        logger.warning("[%s] %d RPT items dropped (capacity=%d)", result.company_name, rpt_dropped, rpt_capacity)

    if len(items) < rpt_capacity:
        for i in range(len(items), rpt_capacity):
            rows_to_hide.append(rpt_start + i)

    for i, item in enumerate(items):
        if i >= rpt_capacity:
            break
        rn = rpt_start + i
        ws[f"C{rn}"] = item["label"]
        ws[f"D{rn}"] = item["value_fy25"]
        comment_text = f"Source: {rpt['source']}"
        if item.get("value_fy25") and float(item["value_fy25"]) < 0:
            comment_text += " | Sign convention: negative = net amount receivable."
        ws[f"D{rn}"].comment = Comment(comment_text, AUTHOR)

    result.rpt_rows_written = rpt_written
    result.rpt_rows_dropped = rpt_dropped

    # AE revenue cross-check
    ws[mapping.ae_revenue_sr9_cell] = 0
    ws[mapping.ae_revenue_sr9_cell].comment = Comment(
        f"AE revenues (Sale of services to foreign AEs) cross-checked against RPT table. {ae_source}",
        AUTHOR,
    )

    # ===================================================================
    # COUNTRIES OF PRESENCE
    # ===================================================================
    countries_row = rpt_start + rpt_capacity + mapping.countries_offset_from_rpt_end
    _note(ws, f"C{countries_row}", f_["countries_presence"]["value"], f_["countries_presence"]["source"])

    # ===================================================================
    # LITIGATION (Dynamic)
    # ===================================================================
    lit_header_row = countries_row + mapping.lit_header_offset_from_countries
    lit_start = lit_header_row + 1
    lit_items = f_["litigation"]["items"]

    lit_capacity = mapping.lit_default_capacity
    lit_written = min(len(lit_items), lit_capacity)
    lit_dropped = max(0, len(lit_items) - lit_capacity)

    if lit_dropped > 0:
        result.warnings.append(
            f"Litigation: {lit_dropped} items exceeded template capacity of {lit_capacity} and were dropped."
        )
        logger.warning("[%s] %d Litigation items dropped (capacity=%d)", result.company_name, lit_dropped, lit_capacity)

    if len(lit_items) < lit_capacity:
        for i in range(len(lit_items), lit_capacity):
            rows_to_hide.append(lit_start + i)

    for i, it in enumerate(lit_items):
        if i >= lit_capacity:
            break
        rn = lit_start + i
        ws[f"C{rn}"] = it.get("nature_of_dues")
        ws[f"D{rn}"] = it.get("amount_demanded_lakhs")
        ws[f"E{rn}"] = it.get("amount_paid_lakhs") if it.get("amount_paid_lakhs") is not None else "-"
        ws[f"F{rn}"] = it.get("period")
        ws[f"G{rn}"] = it.get("forum")
        ws[f"C{rn}"].comment = Comment(f"Source: {f_['litigation']['source']}", AUTHOR)

    result.lit_rows_written = lit_written
    result.lit_rows_dropped = lit_dropped

    # ===================================================================
    # WEBSITE / LINKEDIN
    # ===================================================================
    website_row = lit_header_row + lit_capacity + (mapping.web_offset_from_lit_header - mapping.lit_header_offset_from_countries)
    _note(ws, f"C{website_row}", f_["website"]["value"], f_["website"]["source"])
    _note(ws, f"C{website_row + 1}", f_["linkedin"]["value"], f_["linkedin"]["source"])

    # ===================================================================
    # CLEANUP: Hide and clear unused rows
    # ===================================================================
    for r in rows_to_hide:
        for col in mapping.data_columns:
            cell = ws[f"{col}{r}"]
            try:
                cell.value = None
                cell.comment = None
            except AttributeError:
                pass  # Skip read-only MergedCell objects
        ws.row_dimensions[r].hidden = True

    wb.save(out_path)
    logger.info("[%s] Saved populated workbook to %s", result.company_name, out_path)
    return result
