"""
create_bd_template.py
Generates the BD Fact Sheet Excel template precisely matching FactSheetMapping.

Key rule: S() merges A:H on a given row. NEVER write to individual cells
of that same row afterwards. All data / label cells go on SEPARATE rows.

Run:  python scripts/create_bd_template.py
"""
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── palette ──────────────────────────────────────────────────────────────────
NAVY = "1F3864"; STEEL = "2E74B5"; LIGHT = "D6E4F0"; GOLD = "C9A84C"
WHITE = "FFFFFF"; GREY = "F2F2F2"; BC = "BDD7EE"

def fill(h):  return PatternFill("solid", fgColor=h)
def border(c=BC):
    s = Side(style="thin", color=c)
    return Border(left=s, right=s, top=s, bottom=s)
def align(h="left"): return Alignment(horizontal=h, vertical="center", wrap_text=True)
def font(bold=False, color="1F1F1F", size=10):
    return Font(bold=bold, color=color, size=size, name="Calibri")

# ── cell writers ─────────────────────────────────────────────────────────────
def sec(ws, row, text, bg=STEEL, h=16):
    """Full-width section header. Merges A:H. Never write to this row individually."""
    ws.row_dimensions[row].height = h
    c = ws.cell(row=row, column=1, value=text)
    c.fill = fill(bg); c.font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    c.alignment = align(); c.border = border()
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)

def lbl(ws, r, col, text, bg=LIGHT, bold=False):
    c = ws.cell(row=r, column=col, value=text)
    c.fill = fill(bg); c.font = font(bold=bold, color=NAVY)
    c.alignment = align(); c.border = border()

def dat(ws, r, col, v="", bg=WHITE, sp=None):
    c = ws.cell(row=r, column=col, value=v)
    c.fill = fill(bg); c.font = font(); c.alignment = align(); c.border = border()
    if sp:
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+sp-1)

def num(ws, r, col, v="", bg=WHITE):
    c = ws.cell(row=r, column=col, value=v)
    c.fill = fill(bg); c.font = font(); c.alignment = align(h="right")
    c.border = border(); c.number_format = "#,##0.00"

def pct(ws, r, col, f, bg=WHITE):
    c = ws.cell(row=r, column=col, value=f)
    c.fill = fill(bg); c.font = font(); c.alignment = align(h="right")
    c.border = border(); c.number_format = "0.0%"

def chdr(ws, r, col, text, bg=NAVY):
    c = ws.cell(row=r, column=col, value=text)
    c.fill = fill(bg); c.font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    c.alignment = align(h="center"); c.border = border()

def sp(ws, r):
    ws.row_dimensions[r].height = 5
    for col in range(1, 9):
        ws.cell(row=r, column=col).fill = fill("E8F0F8")

# ── workbook ─────────────────────────────────────────────────────────────────
wb = Workbook(); ws = wb.active; ws.title = "Template"
for i, w in enumerate([4, 4, 38, 18, 14, 18, 14, 14], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"

# ── Row 1: main title (A1 = title_cell) ──────────────────────────────────────
ws.row_dimensions[1].height = 28
ws.merge_cells("A1:H1")
c = ws["A1"]
c.value = "TP Business Development Fact Sheet"
c.fill = fill(NAVY); c.font = Font(bold=True, color=GOLD, size=14, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border()

# ── Row 2: S1 header ─────────────────────────────────────────────────────────
sec(ws, 2, "SECTION 1: COMPANY OVERVIEW")

# Rows 3-4: HQ (C3, C4)
ws.row_dimensions[3].height = 18
lbl(ws, 3, 3, "Headquarters — India Entity"); dat(ws, 3, 4, sp=5)

ws.row_dimensions[4].height = 18
lbl(ws, 4, 3, "Headquarters — Group"); dat(ws, 4, 4, sp=5)

sp(ws, 5)

# Row 6: description sub-header (starts at col 3 = inside the sheet, not merged from A)
ws.row_dimensions[6].height = 14
c6 = ws.cell(row=6, column=3, value="DESCRIPTIONS")
c6.fill = fill(GOLD); c6.font = Font(bold=True, color=NAVY, size=10, name="Calibri")
c6.alignment = align(); c6.border = border()
ws.merge_cells(start_row=6, start_column=3, end_row=6, end_column=8)

# Row 7: Company desc (C7)
ws.row_dimensions[7].height = 60
lbl(ws, 7, 3, "Company Description"); dat(ws, 7, 4, sp=5)

sp(ws, 8)

# Row 9: Group desc (C9)
ws.row_dimensions[9].height = 60
lbl(ws, 9, 3, "Group Description", bg=GREY); dat(ws, 9, 4, bg=GREY, sp=5)

# ── Section 2: Standalone (D12, F12, D13, F13, D14) ─────────────────────────
sec(ws, 10, "SECTION 2: STANDALONE SUMMARY OF OPERATIONS  (Rs. Lakhs)")

ws.row_dimensions[11].height = 18
lbl(ws, 11, 3, "Item", bg=NAVY, bold=True)
ws.cell(row=11, column=3).font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
chdr(ws, 11, 4, "FY 2024-25"); pct(ws, 11, 5, "% Change", bg=NAVY)
ws.cell(row=11, column=5).number_format = "General"
ws.cell(row=11, column=5).font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
chdr(ws, 11, 6, "FY 2023-24")

ws.row_dimensions[12].height = 18
lbl(ws, 12, 3, "Revenue from Operations (Turnover)")
num(ws, 12, 4); num(ws, 12, 6)                          # D12, F12
pct(ws, 12, 5, '=IFERROR((D12-F12)/F12,"")', GREY)

ws.row_dimensions[13].height = 18
lbl(ws, 13, 3, "Total Expenses (Total Cost)", bg=GREY)
num(ws, 13, 4, bg=GREY); num(ws, 13, 6, bg=GREY)        # D13, F13
pct(ws, 13, 5, '=IFERROR((D13-F13)/F13,"")', LIGHT)

ws.row_dimensions[14].height = 18
lbl(ws, 14, 3, "PBT  [= Turnover − Total Cost]", bg=LIGHT, bold=True)
for col, f in [(4, "=D12-D13"), (6, "=F12-F13")]:       # D14
    c = ws.cell(row=14, column=col, value=f)
    c.fill = fill(LIGHT); c.font = Font(bold=True, color=STEEL, size=10, name="Calibri")
    c.alignment = align(h="right"); c.border = border(); c.number_format = "#,##0.00"

sp(ws, 15)

# ── Section 3: Consolidated (D18..D21, F18..F21) ─────────────────────────────
sec(ws, 16, "SECTION 3: CONSOLIDATED SUMMARY OF OPERATIONS  (Rs. Lakhs)")

ws.row_dimensions[17].height = 18
lbl(ws, 17, 3, "Item", bg=NAVY, bold=True)
ws.cell(row=17, column=3).font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
chdr(ws, 17, 4, "FY 2024-25"); chdr(ws, 17, 6, "FY 2023-24")

for i, lab in enumerate(["Revenue from Operations (Turnover)", "Total Expenses (Total Cost)",
                          "Profit Before Tax (PBT)", "Profit After Tax (PAT)"]):
    r = 18 + i; bg = WHITE if i % 2 == 0 else GREY
    ws.row_dimensions[r].height = 18
    lbl(ws, r, 3, lab, bg=bg); num(ws, r, 4, bg=bg); num(ws, r, 6, bg=bg)
    pct(ws, r, 5, f'=IFERROR((D{r}-F{r})/F{r},"")', bg)

sp(ws, 22)

# ── Section 4: Statutory Auditors (C23) — standalone label row ───────────────
# No full-width sec() here to avoid merged cell conflict with C23 data.
ws.row_dimensions[23].height = 18
c23s = ws.cell(row=23, column=1, value="SECTION 4: STATUTORY AUDITORS")
c23s.fill = fill(STEEL); c23s.font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
c23s.alignment = align(); c23s.border = border()
ws.merge_cells(start_row=23, start_column=1, end_row=23, end_column=2)
lbl(ws, 23, 3, "Statutory Auditors"); dat(ws, 23, 4, sp=5)          # C23

sp(ws, 24)

# ── Section 5: AE Revenue Split (D26-D30) ────────────────────────────────────
sec(ws, 25, "SECTION 5: AE / DOMESTIC / EXPORT REVENUE SPLIT  (Rs. Lakhs, FY 2024-25)")

ws.row_dimensions[26].height = 18
lbl(ws, 26, 3, "AE Revenues — Sale of Services to Foreign AEs  (Sr 9)", bg=LIGHT, bold=True)
num(ws, 26, 4, bg=LIGHT)                                            # D26

for i, lab in enumerate(["Domestic — AE (to Indian group entities)",
                          "Export — AE (to foreign group entities)",
                          "Domestic — Third Party",
                          "Export — Third Party"]):
    r = 27 + i; bg = WHITE if i % 2 == 0 else GREY
    ws.row_dimensions[r].height = 18
    lbl(ws, r, 3, lab, bg=bg); num(ws, r, 4, bg=bg)                 # D27-D30

ws.row_dimensions[31].height = 18
lbl(ws, 31, 3, "Total Revenue from Operations", bg=LIGHT, bold=True)
c31 = ws.cell(row=31, column=4, value="=SUM(D27:D30)")
c31.fill = fill(LIGHT); c31.font = Font(bold=True, color=NAVY, size=10, name="Calibri")
c31.alignment = align(h="right"); c31.border = border(); c31.number_format = "#,##0.00"

sp(ws, 32)

# ── Section 6: Cash & AE Receivables (D34, F34, D35, F35) ───────────────────
# Rule: sec() merges the entire row. Keep FY sub-labels as text in lbl() cells.
sec(ws, 33, "SECTION 6: CASH AND AE TRADE RECEIVABLES  (Rs. Lakhs)  |  FY 2024-25 → col D  |  FY 2023-24 → col F")

ws.row_dimensions[34].height = 18
lbl(ws, 34, 3, "Cash and Cash Equivalents"); num(ws, 34, 4); num(ws, 34, 6)  # D34, F34

ws.row_dimensions[35].height = 18
lbl(ws, 35, 3, "AE Trade Receivables", bg=GREY)
num(ws, 35, 4, bg=GREY); num(ws, 35, 6, bg=GREY)                   # D35, F35

sp(ws, 36); sp(ws, 37)

# ── Section 7: PE Investment (C38) ───────────────────────────────────────────
# Cannot use sec() on row 38 — must also write C38 on same row.
ws.row_dimensions[38].height = 18
c38h = ws.cell(row=38, column=1, value="S7: PE / FINANCIAL INVESTOR")
c38h.fill = fill(STEEL); c38h.font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
c38h.alignment = align(); c38h.border = border()
ws.merge_cells(start_row=38, start_column=1, end_row=38, end_column=2)
lbl(ws, 38, 3, "PE / Financial Investor Details"); dat(ws, 38, 4, sp=5)  # C38

sp(ws, 39)

# ── Section 8: Shareholding (hdr=40, col-hdr=41, data=42-43, total=44) ───────
sec(ws, 40, "SECTION 8: SHAREHOLDING PATTERN  (No. of shares)")

ws.row_dimensions[41].height = 18
lbl(ws, 41, 3, "Shareholder Category", bg=NAVY, bold=True)
ws.cell(row=41, column=3).font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
chdr(ws, 41, 4, "Shares FY25"); chdr(ws, 41, 5, "% FY25")
chdr(ws, 41, 6, "Shares FY24"); chdr(ws, 41, 7, "% FY24")

for i in range(2):                                                   # rows 42-43
    r = 42 + i; bg = WHITE if i % 2 == 0 else GREY
    ws.row_dimensions[r].height = 18
    lbl(ws, r, 3, "", bg=bg); num(ws, r, 4, bg=bg); num(ws, r, 6, bg=bg)
    pct(ws, r, 5, f'=IFERROR(D{r}/D44,"")', bg)
    pct(ws, r, 7, f'=IFERROR(F{r}/F44,"")', bg)

# Row 44 — MUST contain "Total" in C44 for _find_total_row()
ws.row_dimensions[44].height = 18
c44 = ws.cell(row=44, column=3, value="Total")
c44.fill = fill(LIGHT); c44.font = Font(bold=True, color=NAVY, size=10, name="Calibri")
c44.alignment = align(); c44.border = border()
num(ws, 44, 4, bg=LIGHT); num(ws, 44, 6, bg=LIGHT)                 # D44, F44

sp(ws, 45)

# ── Section 9: RPT (header=46, data rows 47-53, capacity=7) ─────────────────
sec(ws, 46, "SECTION 9: RELATED PARTY TRANSACTIONS  (Rs. Lakhs, FY 2024-25)  |  Nature → col C  |  Amount → col D")

for i in range(7):                                                   # rows 47-53
    r = 47 + i; bg = WHITE if i % 2 == 0 else GREY
    ws.row_dimensions[r].height = 18
    lbl(ws, r, 3, "", bg=bg); num(ws, r, 4, bg=bg)
    c = ws.cell(row=r, column=5, value="+ve=outflow  /  -ve=inflow")
    c.fill = fill(bg); c.font = Font(color="888888", size=8, italic=True, name="Calibri")
    c.alignment = align(); c.border = border()

# ── Countries (row 55 = rpt_start(47) + capacity(7) + offset(1)) ─────────────
sp(ws, 54)
ws.row_dimensions[55].height = 18
lbl(ws, 55, 3, "Countries of Presence"); dat(ws, 55, 4, sp=5)       # C55

# ── Section 10: Litigation (header=57, data=58-59, capacity=2) ───────────────
sp(ws, 56)
sec(ws, 57, "SECTION 10: LITIGATION / CONTINGENT LIABILITIES  (Rs. Lakhs)  |  Cols: Nature | Demanded | Paid | Period | Forum")

for i in range(2):                                                   # rows 58-59
    r = 58 + i; bg = WHITE if i % 2 == 0 else GREY
    ws.row_dimensions[r].height = 18
    lbl(ws, r, 3, "", bg=bg); num(ws, r, 4, bg=bg); num(ws, r, 5, bg=bg)
    dat(ws, r, 6, bg=bg); dat(ws, r, 7, bg=bg)

# ── Section 11: Online Presence (website=67, linkedin=68) ────────────────────
# website_row = lit_header_row(57) + lit_capacity(2) + (web_offset(10) - lit_header_offset(2))
#             = 57 + 2 + 8 = 67
for r in range(60, 66): sp(ws, r)

# Use a partial header at row 66 (leaves 67 free for website data)
ws.row_dimensions[66].height = 14
c66 = ws.cell(row=66, column=1, value="SECTION 11: ONLINE PRESENCE")
c66.fill = fill(STEEL); c66.font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
c66.alignment = align(); c66.border = border()
ws.merge_cells(start_row=66, start_column=1, end_row=66, end_column=8)

ws.row_dimensions[67].height = 18
lbl(ws, 67, 3, "Website"); dat(ws, 67, 4, sp=5)                     # C67

ws.row_dimensions[68].height = 18
lbl(ws, 68, 3, "LinkedIn", bg=GREY); dat(ws, 68, 4, bg=GREY, sp=5) # C68

# ── Footer ────────────────────────────────────────────────────────────────────
sp(ws, 69)
ws.merge_cells("A70:H70")
c = ws["A70"]
c.value = ("All outputs are AI-generated drafts. "
           "Verify all figures against the source Annual Report before use.")
c.fill = fill("FFF2CC")
c.font = Font(bold=True, color="7F6000", size=9, italic=True, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
c.border = border("F0C040")

ws.sheet_properties.tabColor = NAVY
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True

# ── Save ──────────────────────────────────────────────────────────────────────
out_path = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "data", "templates", "BD_Fact_Sheet_Template.xlsx"
)
os.makedirs(os.path.dirname(out_path), exist_ok=True)
wb.save(out_path)
print(f"Template saved to: {out_path}")
