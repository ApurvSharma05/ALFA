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

import logging
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

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
# Pre-fill table expansion
# ---------------------------------------------------------------------------

def expand_dynamic_tables(
    data: dict,
    template_path: str,
    out_path: str,
    sheet_name: str,
    mapping: FactSheetMapping | None = None,
) -> Tuple[str, int]:
    """
    Pre-expand the RPT and Litigation tables in the template to accommodate
    the actual number of extracted items.  Must be called AFTER shareholding
    expansion (if any) and BEFORE fill_factsheet().

    Unlike shareholding expansion (which is handled in app.py because it
    affects all downstream row offsets), RPT and Litigation expansion is
    isolated here because their offsets are computed dynamically from the
    shareholding total row found at fill time.

    Algorithm
    ---------
    The shareholding Total row moves unpredictably when shareholding is
    expanded, so we open the (already-expanded) workbook, locate the Total
    row, derive the exact RPT and Litigation insertion points, and call
    insert_rows_safe() for each table that needs more capacity.

    Args:
        data:          parsed CompanyData dict (already validated)
        template_path: path to the template (after any shareholding expansion)
        out_path:      path to write the final expanded template to
        sheet_name:    worksheet name inside the workbook
        mapping:       FactSheetMapping (uses defaults if None)

    Returns:
        Tuple of (path_to_use, cumulative_delta) where path_to_use is either
        out_path (if any expansion happened) or template_path (no-op).
    """
    if mapping is None:
        mapping = FactSheetMapping()

    f_ = data["fields"]
    rpt_items = f_["related_party_transactions_lakhs"]["items"]
    lit_items = f_["litigation"]["items"]

    rpt_needed = len(rpt_items)
    lit_needed = len(lit_items)

    if rpt_needed <= mapping.rpt_default_capacity and lit_needed <= mapping.lit_default_capacity:
        # Nothing to expand — pass template through unchanged.
        logger.debug("No dynamic table expansion required (RPT=%d, Lit=%d)", rpt_needed, lit_needed)
        return template_path, 0

    # Open workbook to locate the actual Total row (may have shifted due to
    # prior shareholding expansion).
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    try:
        total_row = _find_total_row(ws, mapping.shareholding_start_row, mapping.shareholding_total_label)
    finally:
        wb.close()

    current_path = template_path
    cumulative_delta = 0

    # ── Expand RPT table ──────────────────────────────────────────────────────
    if rpt_needed > mapping.rpt_default_capacity:
        delta_rpt = rpt_needed - mapping.rpt_default_capacity
        # RPT data starts at total_row + rpt_header_offset + 1.
        # We insert BEFORE the first RPT data row so the header stays in place.
        rpt_data_start = total_row + mapping.rpt_header_offset + 1 + cumulative_delta
        rpt_expanded_path = out_path + ".rpt_expanded.tmp"
        insert_rows_safe(
            src_path=current_path,
            out_path=rpt_expanded_path,
            sheet_name=sheet_name,
            threshold=rpt_data_start,
            delta=delta_rpt,
        )
        current_path = rpt_expanded_path
        cumulative_delta += delta_rpt
        logger.info(
            "RPT table expanded by %d rows (needed %d, had %d)",
            delta_rpt, rpt_needed, mapping.rpt_default_capacity,
        )

    # ── Expand Litigation table ───────────────────────────────────────────────
    if lit_needed > mapping.lit_default_capacity:
        delta_lit = lit_needed - mapping.lit_default_capacity
        # Compute lit_start accounting for all prior deltas:
        # lit_header = countries_row + lit_header_offset_from_countries
        # countries_row = rpt_start + effective_rpt_capacity + countries_offset_from_rpt_end
        effective_rpt_capacity = max(rpt_needed, mapping.rpt_default_capacity)
        rpt_start_final = total_row + mapping.rpt_header_offset + 1
        countries_row = rpt_start_final + effective_rpt_capacity + mapping.countries_offset_from_rpt_end
        lit_header_row = countries_row + mapping.lit_header_offset_from_countries
        lit_data_start = lit_header_row + 1 + cumulative_delta  # add previous deltas
        lit_expanded_path = out_path + ".lit_expanded.tmp"
        insert_rows_safe(
            src_path=current_path,
            out_path=lit_expanded_path,
            sheet_name=sheet_name,
            threshold=lit_data_start,
            delta=delta_lit,
        )
        current_path = lit_expanded_path
        cumulative_delta += delta_lit
        logger.info(
            "Litigation table expanded by %d rows (needed %d, had %d)",
            delta_lit, lit_needed, mapping.lit_default_capacity,
        )

    # Rename last tmp file to the requested out_path
    import os, shutil
    shutil.move(current_path, out_path)
    # Clean up intermediate tmp if it still exists (e.g. only one expansion ran)
    for tmp in [out_path + ".rpt_expanded.tmp", out_path + ".lit_expanded.tmp"]:
        if os.path.exists(tmp):
            os.remove(tmp)

    return out_path, cumulative_delta


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

    # Use the actual capacity available in the (possibly expanded) template.
    # This is the number of rows between rpt_start and the next static section.
    # We no longer cap at a fixed default — any pre-expansion done by
    # expand_dynamic_tables() has already made room for all items.
    rpt_capacity = len(items) if len(items) > 0 else mapping.rpt_default_capacity
    rpt_written = len(items)
    rpt_dropped = 0

    if len(items) < mapping.rpt_default_capacity:
        for i in range(len(items), mapping.rpt_default_capacity):
            rows_to_hide.append(rpt_start + i)

    for i, item in enumerate(items):
        rn = rpt_start + i
        ws[f"C{rn}"] = item["label"]
        ws[f"D{rn}"] = item["value_fy25"]
        comment_text = f"Source: {rpt['source']}"
        value_fy25 = item.get("value_fy25")
        if value_fy25 is not None and float(value_fy25) < 0:
            comment_text += " | Sign convention: negative = net amount receivable."
        ws[f"D{rn}"].comment = Comment(comment_text, AUTHOR)

    result.rpt_rows_written = rpt_written
    result.rpt_rows_dropped = rpt_dropped

    # AE revenue Sr 9: derive from validated extraction data instead of hardcoding 0.
    # export_ae_fy25 represents revenue from services sold to foreign (non-Indian) AEs.
    ae_sr9_value = f_.get("ae_revenue_split", {}).get("value", {}).get("export_ae_fy25", 0) or 0
    ws[mapping.ae_revenue_sr9_cell] = ae_sr9_value
    ws[mapping.ae_revenue_sr9_cell].comment = Comment(
        f"AE revenues (Sale of services to foreign AEs) from ae_revenue_split.export_ae_fy25. {ae_source}",
        AUTHOR,
    )

    # ===================================================================
    # COUNTRIES OF PRESENCE
    # ===================================================================
    countries_row = rpt_start + max(rpt_capacity, mapping.rpt_default_capacity) + mapping.countries_offset_from_rpt_end
    _note(ws, f"C{countries_row}", f_["countries_presence"]["value"], f_["countries_presence"]["source"])

    # ===================================================================
    # LITIGATION (Dynamic)
    # ===================================================================
    lit_header_row = countries_row + mapping.lit_header_offset_from_countries
    lit_start = lit_header_row + 1
    lit_items = f_["litigation"]["items"]

    # No capacity cap — expand_dynamic_tables() has already made room.
    lit_capacity = len(lit_items) if len(lit_items) > 0 else mapping.lit_default_capacity
    lit_written = len(lit_items)
    lit_dropped = 0

    if len(lit_items) < mapping.lit_default_capacity:
        for i in range(len(lit_items), mapping.lit_default_capacity):
            rows_to_hide.append(lit_start + i)

    for i, it in enumerate(lit_items):
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
    website_row = lit_header_row + max(lit_capacity, mapping.lit_default_capacity) + (mapping.web_offset_from_lit_header - mapping.lit_header_offset_from_countries)
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
