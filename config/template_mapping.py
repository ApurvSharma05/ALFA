"""
config/template_mapping.py

Decouples every hardcoded cell coordinate from the fill logic.
If the BD template layout changes, update ONE file — this one.
"""

from dataclasses import dataclass


@dataclass
class FactSheetMapping:
    """
    Maps every BD Fact Sheet field to its Excel cell address or row offset.

    Convention:
        - ``*_cell`` → absolute cell reference like "C3"
        - ``*_row``  → 1-based row number
        - ``*_col``  → column letter
    """

    # --- Header ---
    title_cell: str = "A1"

    # --- Headquarters / Descriptions ---
    hq_india_cell: str = "C3"
    hq_group_cell: str = "C4"
    company_desc_cell: str = "C7"
    group_desc_cell: str = "C9"

    # --- Standalone Summary of Operations ---
    turnover_fy25_cell: str = "D12"
    turnover_fy24_cell: str = "F12"
    total_cost_fy25_cell: str = "D13"
    total_cost_fy24_cell: str = "F13"
    pbt_comment_cell: str = "D14"

    # --- Consolidated Summary of Operations ---
    consolidated_cells: tuple = (
        "D18", "F18",  # Turnover
        "D19", "F19",  # Total cost
        "D20", "F20",  # PBT
        "D21", "F21",  # PAT
    )
    consolidated_comment_cell: str = "D18"

    # --- Statutory Auditors ---
    auditors_cell: str = "C23"

    # --- AE Revenue Split ---
    ae_revenue_sr9_cell: str = "D26"  # AE revenues (sale of services to foreign AEs)
    ae_domestic_cell: str = "D27"
    ae_export_cell: str = "D28"
    third_party_domestic_cell: str = "D29"
    third_party_export_cell: str = "D30"

    # --- Cash & AE Trade Receivables ---
    cash_fy25_cell: str = "D34"
    cash_fy24_cell: str = "F34"
    ae_receivables_fy25_cell: str = "D35"
    ae_receivables_fy24_cell: str = "F35"

    # --- PE Investment ---
    pe_investment_cell: str = "C38"

    # --- Shareholding Table ---
    shareholding_header_row: int = 40
    shareholding_start_row: int = 42
    shareholding_default_capacity: int = 2  # rows available in a fresh template
    shareholding_total_label: str = "Total"
    shareholding_insert_threshold: int = 43  # row at which safe_sheet inserts

    # --- RPT Table (offsets relative to shareholding Total row) ---
    rpt_header_offset: int = 2   # RPT header = total_row + 2
    rpt_data_offset: int = 3     # first RPT data row = total_row + 3
    rpt_default_capacity: int = 7

    # --- Countries ---
    countries_offset_from_rpt_end: int = 1  # countries row = rpt_start + rpt_capacity + 1

    # --- Litigation Table ---
    lit_header_offset_from_countries: int = 2
    lit_default_capacity: int = 2

    # --- Website / LinkedIn ---
    web_offset_from_lit_header: int = 10  # website row = lit_header + lit_capacity + 8
    # LinkedIn is always website_row + 1

    # --- Column letters used for clearing/hiding ---
    data_columns: tuple = ("C", "D", "E", "F", "G", "H", "I")
