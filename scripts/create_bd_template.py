"""
scripts/create_bd_template.py

Generates a professional TP BD Fact Sheet Excel template (.xlsx)
that is fully compatible with ALFA's FactSheetMapping and excel_engine.

Run:  python scripts/create_bd_template.py
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter


# ── Colour palette ──────────────────────────────────────────────────────────
NAVY       = "0F172A"
DARK_BLUE  = "1E293B"
MID_BLUE   = "334155"
LIGHT_BLUE = "DBEAFE"
WHITE      = "FFFFFF"
LIGHT_GRAY = "F1F5F9"
BORDER_CLR = "CBD5E1"

# Fills
fill_header    = PatternFill("solid", fgColor=NAVY)
fill_section   = PatternFill("solid", fgColor=DARK_BLUE)
fill_subheader = PatternFill("solid", fgColor=MID_BLUE)
fill_light     = PatternFill("solid", fgColor=LIGHT_GRAY)
fill_white     = PatternFill("solid", fgColor=WHITE)
fill_accent    = PatternFill("solid", fgColor=LIGHT_BLUE)

# Fonts
font_title   = Font(name="Inter", size=14, bold=True, color=WHITE)
font_section = Font(name="Inter", size=11, bold=True, color=WHITE)
font_sub     = Font(name="Inter", size=10, bold=True, color=WHITE)
font_label   = Font(name="Inter", size=10, bold=True, color="1E293B")
font_normal  = Font(name="Inter", size=10, color="334155")
font_small   = Font(name="Inter", size=9, color="64748B")

# Borders
thin_border = Border(
    left=Side(style="thin", color=BORDER_CLR),
    right=Side(style="thin", color=BORDER_CLR),
    top=Side(style="thin", color=BORDER_CLR),
    bottom=Side(style="thin", color=BORDER_CLR),
)

# Alignments
align_left   = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_right  = Alignment(horizontal="right", vertical="center")


def style_range(ws, row, cols, fill, font, alignment=align_left, border=thin_border):
    """Apply consistent styling to a range of cells in a row."""
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = alignment
        cell.border = border


def section_header(ws, row, text, cols=range(1, 10)):
    """Create a dark section header spanning columns."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    cell = ws.cell(row=row, column=1)
    cell.value = text
    cell.fill = fill_section
    cell.font = font_section
    cell.alignment = align_left
    cell.border = thin_border
    for c in range(2, 10):
        ws.cell(row=row, column=c).fill = fill_section
        ws.cell(row=row, column=c).border = thin_border


def label_value_row(ws, row, label, value_cols=(4,), label_col=3, fy_label=False):
    """Create a label + value row."""
    # Column A-B: empty or merged
    for c in range(1, 3):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill_white
        cell.border = thin_border

    # Column C: label
    cell_c = ws.cell(row=row, column=label_col)
    cell_c.value = label
    cell_c.font = font_label
    cell_c.fill = fill_light
    cell_c.alignment = align_left
    cell_c.border = thin_border

    # Data columns
    for c in range(4, 10):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill_white
        cell.border = thin_border
        cell.font = font_normal
        cell.alignment = align_right
        cell.number_format = '#,##0.00'


def create_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "BD Fact Sheet"

    # ── Column widths ───────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 4
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 14

    # ====================================================================
    # ROW 1: Title
    # ====================================================================
    ws.merge_cells("A1:I1")
    ws.row_dimensions[1].height = 32
    cell = ws["A1"]
    cell.value = "Fact Sheet - [Company Name]"
    cell.fill = fill_header
    cell.font = font_title
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border
    for c in range(2, 10):
        ws.cell(row=1, column=c).fill = fill_header
        ws.cell(row=1, column=c).border = thin_border

    # ====================================================================
    # ROW 2: Section — HEADQUARTERS
    # ====================================================================
    section_header(ws, 2, "HEADQUARTERS & COMPANY INFORMATION")

    # Row 3: HQ India
    label_value_row(ws, 3, "Registered Office (India Entity)")
    # Row 4: HQ Group
    label_value_row(ws, 4, "Group Headquarters")
    # Row 5: Blank spacer
    for c in range(1, 10):
        ws.cell(row=5, column=c).fill = fill_white
        ws.cell(row=5, column=c).border = thin_border
    # Row 6: Section — DESCRIPTIONS
    section_header(ws, 6, "COMPANY & GROUP DESCRIPTION")
    # Row 7: Company Description
    label_value_row(ws, 7, "Company Description")
    ws.row_dimensions[7].height = 48
    # Row 8: spacer
    for c in range(1, 10):
        ws.cell(row=8, column=c).fill = fill_white
        ws.cell(row=8, column=c).border = thin_border
    # Row 9: Group Description
    label_value_row(ws, 9, "Group Description")
    ws.row_dimensions[9].height = 48

    # ====================================================================
    # ROW 10: Spacer
    # ====================================================================
    for c in range(1, 10):
        ws.cell(row=10, column=c).fill = fill_white
        ws.cell(row=10, column=c).border = thin_border

    # ====================================================================
    # ROW 11: Section — STANDALONE SUMMARY OF OPERATIONS
    # ====================================================================
    section_header(ws, 11, "STANDALONE SUMMARY OF OPERATIONS (₹ Lakhs)")

    # Sub-header row with FY labels
    for c in range(1, 10):
        cell = ws.cell(row=11, column=c)

    # Row 12: Turnover is at D12, so row 12 is data.
    # Let's embed the FY labels directly into the section header merge text
    # and rely on the D/F column convention.

    # Row 12: Turnover
    label_value_row(ws, 12, "Revenue from Operations (Turnover)")
    ws.cell(row=12, column=4).value = "FY 2024-25"
    ws.cell(row=12, column=4).font = font_label
    ws.cell(row=12, column=4).alignment = align_center
    ws.cell(row=12, column=6).value = "FY 2023-24"
    ws.cell(row=12, column=6).font = font_label
    ws.cell(row=12, column=6).alignment = align_center

    # Row 13: Total Cost
    label_value_row(ws, 13, "Total Expenses (Total Cost)")

    # Row 14: PBT (formula)
    label_value_row(ws, 14, "PBT (Turnover − Total Cost)")
    ws.cell(row=14, column=4).value = "=D12-D13"
    ws.cell(row=14, column=4).number_format = '#,##0.00'
    ws.cell(row=14, column=6).value = "=F12-F13"
    ws.cell(row=14, column=6).number_format = '#,##0.00'

    # Rows 15-16: spacer
    for r in (15, 16):
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = fill_white
            ws.cell(row=r, column=c).border = thin_border

    # ====================================================================
    # ROW 17: Section — CONSOLIDATED SUMMARY
    # ====================================================================
    section_header(ws, 17, "CONSOLIDATED SUMMARY OF OPERATIONS (₹ Lakhs)")

    # Row 18: Turnover
    label_value_row(ws, 18, "Revenue from Operations (Consolidated)")
    # Row 19: Total Cost
    label_value_row(ws, 19, "Total Expenses (Consolidated)")
    # Row 20: PBT
    label_value_row(ws, 20, "PBT (Consolidated)")
    ws.cell(row=20, column=4).value = "=D18-D19"
    ws.cell(row=20, column=6).value = "=F18-F19"
    # Row 21: PAT
    label_value_row(ws, 21, "PAT (Consolidated)")

    # Row 22: spacer
    for c in range(1, 10):
        ws.cell(row=22, column=c).fill = fill_white
        ws.cell(row=22, column=c).border = thin_border

    # ====================================================================
    # ROW 23: Statutory Auditors
    # ====================================================================
    section_header(ws, 22, "STATUTORY AUDITORS")
    label_value_row(ws, 23, "Auditor Name / Firm")

    # Row 24: spacer
    for c in range(1, 10):
        ws.cell(row=24, column=c).fill = fill_white
        ws.cell(row=24, column=c).border = thin_border

    # ====================================================================
    # ROW 25: Section — AE / REVENUE SPLIT
    # ====================================================================
    section_header(ws, 25, "AE / DOMESTIC / EXPORT REVENUE SPLIT (₹ Lakhs)")

    # Row 26: AE Revenue (Sr 9)
    label_value_row(ws, 26, "AE Revenues (Sale of Services to Foreign AEs)")
    # Row 27: AE Domestic
    label_value_row(ws, 27, "AE — Domestic")
    # Row 28: AE Export
    label_value_row(ws, 28, "AE — Export")
    # Row 29: Third Party Domestic
    label_value_row(ws, 29, "Third Party — Domestic")
    # Row 30: Third Party Export
    label_value_row(ws, 30, "Third Party — Export")

    # Row 31-32: spacer
    for r in (31, 32):
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = fill_white
            ws.cell(row=r, column=c).border = thin_border

    # ====================================================================
    # ROW 33: Section — CASH & RECEIVABLES
    # ====================================================================
    section_header(ws, 33, "CASH & AE TRADE RECEIVABLES (₹ Lakhs)")

    # Row 34: Cash
    label_value_row(ws, 34, "Cash & Cash Equivalents")
    ws.cell(row=34, column=4).value = "FY 2024-25"
    ws.cell(row=34, column=4).font = font_label
    ws.cell(row=34, column=4).alignment = align_center
    ws.cell(row=34, column=6).value = "FY 2023-24"
    ws.cell(row=34, column=6).font = font_label
    ws.cell(row=34, column=6).alignment = align_center

    # Row 35: AE Trade Receivables
    label_value_row(ws, 35, "AE Trade Receivables")

    # Row 36-37: spacer
    for r in (36, 37):
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = fill_white
            ws.cell(row=r, column=c).border = thin_border

    # ====================================================================
    # ROW 38: PE INVESTMENT
    # ====================================================================
    section_header(ws, 37, "PE / VC INVESTMENT")
    label_value_row(ws, 38, "Private Equity / Venture Capital Investment")

    # Row 39: spacer
    for c in range(1, 10):
        ws.cell(row=39, column=c).fill = fill_white
        ws.cell(row=39, column=c).border = thin_border

    # ====================================================================
    # ROW 40-43: SHAREHOLDING TABLE
    # ====================================================================
    section_header(ws, 40, "SHAREHOLDING PATTERN")

    # Sub-header row 41
    headers_sh = ["", "", "Shareholder Name", "Shares FY25", "% FY25", "Shares FY24", "% FY24", "", ""]
    for c, h in enumerate(headers_sh, start=1):
        cell = ws.cell(row=41, column=c)
        cell.value = h
        cell.fill = fill_subheader
        cell.font = font_sub
        cell.alignment = align_center
        cell.border = thin_border

    # Data rows 42-43 (default capacity = 2)
    for r in (42, 43):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_white
            cell.border = thin_border
            cell.font = font_normal
            if c == 3:
                cell.alignment = align_left
            elif c in (4, 5, 6, 7):
                cell.alignment = align_right
                cell.number_format = '#,##0'

    # Row 44: Total row
    ws.cell(row=44, column=3).value = "Total"
    for c in range(1, 10):
        cell = ws.cell(row=44, column=c)
        cell.fill = fill_accent
        cell.font = font_label
        cell.border = thin_border
        cell.alignment = align_right if c >= 4 else align_left
        if c in (4, 6):
            cell.number_format = '#,##0'
    # % formulas
    ws.cell(row=44, column=5).value = "100%"
    ws.cell(row=44, column=7).value = "100%"

    # Row 45: spacer
    for c in range(1, 10):
        ws.cell(row=45, column=c).fill = fill_white
        ws.cell(row=45, column=c).border = thin_border

    # ====================================================================
    # ROW 46: RELATED PARTY TRANSACTIONS
    # ====================================================================
    section_header(ws, 46, "RELATED PARTY TRANSACTIONS (₹ Lakhs)")

    # RPT data rows 47-53 (capacity = 7)
    rpt_headers = ["", "", "Nature of Transaction", "Amount FY25", "Amount FY24", "", "", "", ""]
    for c, h in enumerate(rpt_headers, start=1):
        cell = ws.cell(row=46, column=c)
        # Already set by section_header, but add sub text in row after

    for r in range(47, 54):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_white if (r % 2 == 1) else fill_light
            cell.border = thin_border
            cell.font = font_normal
            if c == 3:
                cell.alignment = align_left
            elif c in (4, 5, 6):
                cell.alignment = align_right
                cell.number_format = '#,##0.00'

    # Row 54: spacer
    for c in range(1, 10):
        ws.cell(row=54, column=c).fill = fill_white
        ws.cell(row=54, column=c).border = thin_border

    # ====================================================================
    # ROW 55: COUNTRIES OF PRESENCE
    # ====================================================================
    # Countries row is a data row per mapping (C55 gets the countries list)
    # rpt_start = 47, rpt_capacity = 7, countries_row = 47 + 7 + 1 = 55
    label_value_row(ws, 55, "Countries of Presence")

    # Row 56: spacer
    for c in range(1, 10):
        ws.cell(row=56, column=c).fill = fill_white
        ws.cell(row=56, column=c).border = thin_border

    # ====================================================================
    # ROW 57: LITIGATION
    # ====================================================================
    section_header(ws, 57, "LITIGATION / CONTINGENT LIABILITIES")

    # Sub-headers
    lit_headers = ["", "", "Nature of Dues", "Amount Demanded (₹L)", "Amount Paid (₹L)", "Period", "Forum", "", ""]
    for c, h in enumerate(lit_headers, start=1):
        cell = ws.cell(row=57, column=c)
        # Already merged by section_header, rewrite columns individually
    # Actually we need a sub-header row
    for c, h in enumerate(lit_headers, start=1):
        cell = ws.cell(row=57, column=c)

    # Litigation data rows 58-59 (capacity = 2)
    for r in (58, 59):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_white
            cell.border = thin_border
            cell.font = font_normal
            if c == 3:
                cell.alignment = align_left
            elif c in (4, 5):
                cell.alignment = align_right
                cell.number_format = '#,##0.00'
            elif c in (6, 7):
                cell.alignment = align_left

    # Rows 60-65: spacer + extra
    for r in range(60, 66):
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = fill_white
            ws.cell(row=r, column=c).border = thin_border

    # ====================================================================
    # ROW 66-67: WEBSITE / LINKEDIN
    # ====================================================================
    # Per mapping: web_offset_from_lit_header = 10
    # website_row = lit_header + lit_capacity + (web_offset - lit_header_offset)
    # = 57 + 2 + (10 - 2) = 67
    section_header(ws, 66, "WEB PRESENCE")
    label_value_row(ws, 67, "Website")
    label_value_row(ws, 68, "LinkedIn")

    # ====================================================================
    # PRINT SETTINGS
    # ====================================================================
    ws.sheet_properties.pageSetUpPr = None
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"
    ws.print_area = "A1:I68"

    # ====================================================================
    # SAVE
    # ====================================================================
    out_path = "BD_FactSheet_Template.xlsx"
    wb.save(out_path)
    print(f"[OK] Template saved: {out_path}")
    return out_path


if __name__ == "__main__":
    create_template()
