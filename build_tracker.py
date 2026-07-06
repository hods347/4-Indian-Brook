#!/usr/bin/env python3
"""Build the 4 Indian Brook Road house tracker workbook.

Produces an .xlsx designed to be imported into Google Sheets
(File > Import > Upload > Replace spreadsheet). All formulas used
(SUM, SUMIF, HLOOKUP, MIN, IF) are Sheets-compatible.

Tabs:
  Read Me             - how the workbook fits together
  Basis Tracker       - purchase price + closing costs + capital improvements
  2nd Floor Flooring  - first project, pre-filled from the Footprints Floors
                        quote (#25584) and the Floor & Decor cart
  Project Template    - blank copy of the project structure for future work
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "4-Indian-Brook-House-Tracker.xlsx"

# ---------------------------------------------------------------- palette --
NAVY = "1F3864"
NAVY_LIGHT = "2E5395"
GOLD = "BF9000"
GREY_HEAD = "D9D9D9"
GREY_SOFT = "F2F2F2"
GREEN_SOFT = "E2EFDA"
BLUE_SOFT = "DDEBF7"
YELLOW_SOFT = "FFF2CC"

CUR = '"$"#,##0.00'
CUR0 = '"$"#,##0'
PCT = "0%"
DATE_FMT = "mm/dd/yyyy"

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TITLE_FONT = Font(bold=True, size=16, color="FFFFFF")
SUB_FONT = Font(italic=True, size=10, color="595959")
SECTION_FONT = Font(bold=True, size=11, color="FFFFFF")
HEAD_FONT = Font(bold=True, size=10)
LABEL_FONT = Font(bold=True, size=10)
NOTE_FONT = Font(italic=True, size=9, color="595959")
TOTAL_FONT = Font(bold=True, size=11)

# ------------------------------------------------------------ project layout --
# Both project tabs share one deterministic layout so the Basis Tracker can
# point at a known cell. Comparison labels drive the HLOOKUP offsets below.
COMPARISON_LABELS = [
    "Item # / SKU",          # 0
    "unit",                  # 1  (label text comes from cfg)
    "qty",                   # 2
    "Flooring material",     # 3  -- materials block starts here
    "Underlayment & supplies",
    "Stair nose trim",
    "Stair risers",
    "Contractor discount on materials",
    "Est. sales tax (6.25% MA, materials)",
    "Labor — demo & disposal",
    "Labor — install",
    "Labor — stairs (treads & risers)",
    "TOTAL ESTIMATED",
    "Premium vs. cheapest option",
]
# budget categories -> comparison label they pull their number from
BUDGET_CATEGORIES = [
    ("Flooring materials", "Flooring material"),
    ("Underlayment & supplies", "Underlayment & supplies"),
    ("Stair nose trim", "Stair nose trim"),
    ("Stair risers", "Stair risers"),
    ("Contractor discount", "Contractor discount on materials"),
    ("Sales tax", "Est. sales tax (6.25% MA, materials)"),
    ("Labor — demo & disposal", "Labor — demo & disposal"),
    ("Labor — install", "Labor — install"),
    ("Labor — stairs", "Labor — stairs (treads & risers)"),
    ("Permits & fees", None),
    ("Other", None),
]

COMP_SECTION = 12
COMP_HEADER = COMP_SECTION + 1                       # 13
COMP_FIRST = COMP_HEADER + 1                         # 14
COMP_ROW = {lbl: COMP_FIRST + i for i, lbl in enumerate(COMPARISON_LABELS)}
COMP_LAST = COMP_ROW["Premium vs. cheapest option"]  # 27
COMP_TOTAL = COMP_ROW["TOTAL ESTIMATED"]             # 26

BUD_SECTION = COMP_LAST + 2                          # 29
BUD_HEADER = BUD_SECTION + 1                         # 30
BUD_FIRST = BUD_HEADER + 1                           # 31
BUD_LAST_CAT = BUD_FIRST + len(BUDGET_CATEGORIES) - 1  # 41 ("Other")
BUD_CONT = BUD_LAST_CAT + 1                          # 42 contingency
BUD_TOTAL = BUD_CONT + 1                             # 43 total
ACTUAL_TOTAL_CELL = f"C{BUD_TOTAL}"                  # feeds the Basis Tracker

LOG_SECTION = BUD_TOTAL + 3                          # 46
LOG_HEADER = LOG_SECTION + 1                         # 47
LOG_FIRST, LOG_LAST = LOG_HEADER + 1, LOG_HEADER + 100

MATERIAL_FIRST = COMP_ROW["Flooring material"]
MATERIAL_LAST = COMP_ROW["Stair risers"]
DISCOUNT_ROW = COMP_ROW["Contractor discount on materials"]


def title_bar(ws, text, sub, last_col="G"):
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = text
    c.font = TITLE_FONT
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A2:{last_col}2")
    s = ws["A2"]
    s.value = sub
    s.font = SUB_FONT
    ws.row_dimensions[2].height = 16


def section_bar(ws, row, text, last_col="G", color=NAVY_LIGHT):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.font = SECTION_FONT
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def head_cell(ws, row, col, text, fill=GREY_HEAD):
    c = ws.cell(row=row, column=col, value=text)
    c.font = HEAD_FONT
    c.fill = PatternFill("solid", fgColor=fill)
    c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return c


def money(ws, row, col, value, fmt=CUR, bold=False, fill=None):
    c = ws.cell(row=row, column=col, value=value)
    c.number_format = fmt
    c.border = BOX
    if bold:
        c.font = TOTAL_FONT
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    return c


# ============================================================ project tabs ==
def build_project_sheet(ws, cfg):
    """Shared layout for a project tab; row positions come from the module-
    level layout constants so the Basis Tracker link never drifts."""
    ws.sheet_view.showGridLines = False
    title_bar(ws, cfg["title"], cfg["subtitle"])

    # ---- project info -----------------------------------------------------
    section_bar(ws, 4, "PROJECT INFO")
    info = [
        ("Status", cfg["status"]),
        ("Contractor / quote", cfg["contractor"]),
        ("Materials source", cfg["materials_source"]),
        ("Selected option", cfg["selected"]),
        ("Contingency %", 0.10),
        ("Notes", cfg["notes"]),
    ]
    for i, (label, val) in enumerate(info):
        r = 5 + i
        ws.cell(row=r, column=1, value=label).font = LABEL_FONT
        c = ws.cell(row=r, column=2, value=val)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c.alignment = Alignment(vertical="center", wrap_text=(label == "Notes"))
        if label == "Contingency %":
            c.number_format = PCT
        if label == "Selected option":
            c.fill = PatternFill("solid", fgColor=YELLOW_SOFT)
            c.border = BOX
    ws.row_dimensions[10].height = 28

    dv_status = DataValidation(
        type="list",
        formula1='"Planning,Quoting,Approved,In progress,Complete,On hold"',
        allow_blank=True,
    )
    ws.add_data_validation(dv_status)
    dv_status.add("B5")

    dv_option = DataValidation(
        type="list", formula1=f"=$B${COMP_HEADER}:$D${COMP_HEADER}", allow_blank=True
    )
    ws.add_data_validation(dv_option)
    dv_option.add("B8")

    # ---- step 1: option comparison ----------------------------------------
    section_bar(ws, COMP_SECTION,
                "STEP 1 — COMPARE OPTIONS  (estimated all-in cost per option)")
    head_cell(ws, COMP_HEADER, 1, "Cost component")
    for j, name in enumerate(cfg["options"]):
        head_cell(ws, COMP_HEADER, 2 + j, name, fill=BLUE_SOFT)
    head_cell(ws, COMP_HEADER, 5, "Notes")

    disc = cfg["discount_rate"]
    r_unit = COMP_ROW["unit"]
    r_qty = COMP_ROW["qty"]
    m1, m2 = MATERIAL_FIRST, MATERIAL_LAST
    comparison_rows = [
        ("Item # / SKU", cfg["skus"], None, None),
        (cfg["unit_label"], cfg["unit_prices"], CUR, None),
        (cfg["qty_label"], cfg["qtys"], "#,##0", None),
        ("Flooring material", f"={{col}}{r_unit}*{{col}}{r_qty}", CUR, None),
        ("Underlayment & supplies", cfg["underlayment"], CUR, None),
        ("Stair nose trim", cfg["stair_nose"], CUR, None),
        ("Stair risers", cfg["stair_risers"], CUR, None),
        (f"Contractor discount on materials ({disc:.0%})",
         f"=-{disc}*SUM({{col}}{m1}:{{col}}{m2})", CUR, None),
        ("Est. sales tax (6.25% MA, materials)",
         f"=SUM({{col}}{m1}:{{col}}{DISCOUNT_ROW})*0.0625", CUR, None),
        ("Labor — demo & disposal", cfg["labor_demo"], CUR, None),
        ("Labor — install", cfg["labor_install"], CUR, None),
        ("Labor — stairs (treads & risers)", cfg["labor_stairs"], CUR, None),
        ("TOTAL ESTIMATED",
         f"=SUM({{col}}{m1}:{{col}}{COMP_TOTAL - 1})", CUR, GREEN_SOFT),
        ("Premium vs. cheapest option",
         f"={{col}}{COMP_TOTAL}-MIN($B${COMP_TOTAL}:$D${COMP_TOTAL})", CUR, GREY_SOFT),
    ]
    row_notes = cfg["comparison_notes"]
    for i, (label, values, fmt, fill) in enumerate(comparison_rows):
        r = COMP_FIRST + i
        lc = ws.cell(row=r, column=1, value=label)
        lc.border = BOX
        lc.font = TOTAL_FONT if label.startswith("TOTAL") else Font(size=10)
        if fill:
            lc.fill = PatternFill("solid", fgColor=fill)
        for j in range(3):
            col = get_column_letter(2 + j)
            if isinstance(values, str):  # formula pattern
                v = values.format(col=col)
            else:
                v = values[j] if isinstance(values, (list, tuple)) else values
                if v == "":
                    v = None  # true blank, so math formulas treat it as 0
            c = ws.cell(row=r, column=2 + j, value=v)
            c.border = BOX
            c.alignment = Alignment(horizontal="right" if fmt else "center")
            if fmt:
                c.number_format = fmt
            if fill:
                c.fill = PatternFill("solid", fgColor=fill)
            if label.startswith("TOTAL"):
                c.font = TOTAL_FONT
        note = row_notes.get(label.split(" (")[0])
        if note:
            nc = ws.cell(row=r, column=5, value=note)
            nc.font = NOTE_FONT
            nc.border = BOX
            nc.alignment = Alignment(wrap_text=True, vertical="center")
            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        else:
            ws.cell(row=r, column=5).border = BOX

    # ---- step 2: budget vs actual ------------------------------------------
    section_bar(ws, BUD_SECTION,
                "STEP 2 — BUDGET vs. ACTUAL  (budget auto-fills from the selected option)")
    for col, text in [(1, "Category"), (2, "Budget"),
                      (3, "Actual (from cost log)"), (4, "Remaining")]:
        head_cell(ws, BUD_HEADER, col, text)

    lookup = f"$B${COMP_HEADER}:$D${COMP_TOTAL}"
    for i, (label, comp_label) in enumerate(BUDGET_CATEGORIES):
        r = BUD_FIRST + i
        lc = ws.cell(row=r, column=1, value=label)
        lc.border = BOX
        if comp_label is not None:
            offset = COMP_ROW[comp_label] - COMP_HEADER + 1
            budget_formula = (f'=IF($B$8="","",'
                              f'HLOOKUP($B$8,{lookup},{offset},FALSE))')
        else:
            budget_formula = None
        money(ws, r, 2, budget_formula)
        money(ws, r, 3, f"=SUMIF($D${LOG_FIRST}:$D${LOG_LAST},$A{r},"
                        f"$E${LOG_FIRST}:$E${LOG_LAST})")
        money(ws, r, 4, f'=IF(B{r}="","",B{r}-C{r})')
    ws.cell(row=BUD_CONT, column=1, value="Contingency").border = BOX
    money(ws, BUD_CONT, 2, f"=SUM(B{BUD_FIRST}:B{BUD_LAST_CAT})*$B$9")
    money(ws, BUD_CONT, 3, None)
    money(ws, BUD_CONT, 4, f"=B{BUD_CONT}-C{BUD_CONT}")
    tc = ws.cell(row=BUD_TOTAL, column=1, value="TOTAL")
    tc.font = TOTAL_FONT
    tc.border = BOX
    tc.fill = PatternFill("solid", fgColor=GREEN_SOFT)
    money(ws, BUD_TOTAL, 2, f"=SUM(B{BUD_FIRST}:B{BUD_CONT})", bold=True, fill=GREEN_SOFT)
    money(ws, BUD_TOTAL, 3, f"=SUM(C{BUD_FIRST}:C{BUD_CONT})", bold=True, fill=GREEN_SOFT)
    money(ws, BUD_TOTAL, 4, f"=B{BUD_TOTAL}-C{BUD_TOTAL}", bold=True, fill=GREEN_SOFT)
    ws.cell(row=BUD_TOTAL + 1, column=1,
            value=f"The Actual total ({ACTUAL_TOTAL_CELL}) is what feeds the Basis "
                  "Tracker. Log receipts net of discount, or log the discount as a "
                  "negative amount under 'Contractor discount'.").font = NOTE_FONT

    # ---- step 3: cost log ----------------------------------------------------
    section_bar(ws, LOG_SECTION,
                "STEP 3 — COST LOG  (record every payment; Actuals above update automatically)")
    log_heads = ["Date", "Vendor", "Description", "Category", "Amount",
                 "Payment method", "Notes"]
    for col, text in enumerate(log_heads, start=1):
        head_cell(ws, LOG_HEADER, col, text)
    dv_cat = DataValidation(
        type="list", formula1=f"=$A${BUD_FIRST}:$A${BUD_LAST_CAT}", allow_blank=True
    )
    ws.add_data_validation(dv_cat)
    for r in range(LOG_FIRST, LOG_LAST + 1):
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col == 1:
                c.number_format = DATE_FMT
            if col == 5:
                c.number_format = CUR
        dv_cat.add(f"D{r}")

    widths = {"A": 34, "B": 17, "C": 20, "D": 17, "E": 15, "F": 15, "G": 26}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"


FLOORING_CFG = {
    "title": "PROJECT — SECOND FLOOR FLOORING",
    "subtitle": "4 Indian Brook Road, Ashland, MA 01721",
    "status": "Quoting",
    "contractor": "Footprints Floors of Central MA (ACS Custom Solutions Inc.) — Proposal #25584, 6/30/2026. Labor only; 50% deposit ($4,528.80) before work, balance on completion.",
    "materials_source": "Floor & Decor, Waltham MA (cart re-priced 7/2026); 5% materials discount through flooring contractor",
    "selected": None,
    "notes": "Labor quote excludes materials. Subfloor leveling/remediation not included and may add cost once carpet is removed. 3% fee on credit card payments (max $3,000/card per project).",
    "options": ["Sapelo Shore", "Big Sur", "Gunstock Oak"],
    "skus": ["101128890", "101069698", "101068104"],
    "unit_label": "Price per box",
    "qty_label": "Boxes needed",
    "unit_prices": [61.71, 97.08, 75.57],
    "qtys": [66, 60, 64],
    "underlayment": "=13*59.99",
    "stair_nose": "=7*29.99",
    "stair_risers": "=13*20",
    "discount_rate": 0.05,
    "labor_demo": 2642.30,
    "labor_install": 4535.30,
    "labor_stairs": 1880.00,
    "comparison_notes": {
        "Item # / SKU": "All Floor & Decor. Sapelo Shore & Big Sur: waterproof hybrid resilient plank w/ cork pad, 8mm 7\"x51\". Gunstock Oak: waterproof rigid core LVP.",
        "Underlayment & supplies": "Sentinel Protect Plus underlayment — 13 rolls x $59.99 (100 sqft each), same for every option.",
        "Stair nose trim": "7 stair noses x $29.99.",
        "Stair risers": "13 risers x $20.00.",
        "Contractor discount on materials": "5% off all materials through the flooring contractor.",
        "Est. sales tax (6.25% MA, materials)": "Estimate on discounted materials; replace with the actual receipt tax when purchased.",
        "Labor — demo & disposal": "Footprints: demo carpet, remove & return baseboards, trash removal.",
        "Labor — install": "Footprints: install click-lock floating floor.",
        "Labor — stairs (treads & risers)": "Footprints: treads (open 1 side) + risers.",
        "TOTAL ESTIMATED": "Labor $9,057.60 is identical across options, so the spread comes entirely from the plank you pick.",
    },
}

TEMPLATE_CFG = {
    "title": "PROJECT TEMPLATE — (duplicate this tab for each new project)",
    "subtitle": "How to use: right-click the tab > Duplicate. Rename it, fill in options & quotes, "
                "then add one row for it in the Basis Tracker improvements table pointing at the "
                f"new tab's Actual total (cell {ACTUAL_TOTAL_CELL}).",
    "status": "Planning",
    "contractor": "",
    "materials_source": "",
    "selected": None,
    "notes": "",
    "options": ["Option A", "Option B", "Option C"],
    "skus": ["", "", ""],
    "unit_label": "Unit price",
    "qty_label": "Quantity",
    "unit_prices": ["", "", ""],
    "qtys": ["", "", ""],
    "underlayment": "",
    "stair_nose": "",
    "stair_risers": "",
    "discount_rate": 0.0,
    "labor_demo": "",
    "labor_install": "",
    "labor_stairs": "",
    "comparison_notes": {
        "Stair nose trim": "Rename these two rows for whatever extra materials the project needs.",
        "TOTAL ESTIMATED": "Compare options here, pick one in 'Selected option' above, and the budget below fills itself in.",
    },
}


# ============================================================ basis tracker ==
def build_basis(ws):
    ws.sheet_view.showGridLines = False
    title_bar(ws, "COST BASIS TRACKER", "4 Indian Brook Road, Ashland, MA 01721", last_col="F")

    section_bar(ws, 4, "PROPERTY", last_col="F")
    ws.cell(row=5, column=1, value="Purchase price").font = LABEL_FONT
    money(ws, 5, 2, 1060000, fmt=CUR0)
    ws.cell(row=6, column=1, value="Closing date").font = LABEL_FONT
    d = ws.cell(row=6, column=2, value=date(2026, 7, 17))
    d.number_format = DATE_FMT
    d.border = BOX

    section_bar(ws, 8, "1 — ACQUISITION COSTS  (purchase price + closing costs added to basis)", last_col="F")
    head_cell(ws, 9, 1, "Item")
    head_cell(ws, 9, 2, "Amount")
    head_cell(ws, 9, 3, "Notes")
    ws.merge_cells("C9:F9")
    acq_items = [
        ("Purchase price", 1060000, ""),
        ("Attorney / legal fees", None, "From the closing disclosure"),
        ("Owner's title insurance", None, ""),
        ("Recording fees", None, ""),
        ("Survey / plot plan", None, ""),
        ("Transfer taxes paid by buyer (if any)", None, "In MA the seller usually pays deed stamps"),
        ("Other closing costs added to basis", None, ""),
    ]
    for i, (label, amt, note) in enumerate(acq_items):
        r = 10 + i
        lc = ws.cell(row=r, column=1, value=label)
        lc.border = BOX
        money(ws, r, 2, amt, fmt=CUR0 if r == 10 else CUR)
        nc = ws.cell(row=r, column=3, value=note)
        nc.font = NOTE_FONT
        nc.border = BOX
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    r_acq = 17
    c = ws.cell(row=r_acq, column=1, value="Subtotal — acquisition basis")
    c.font = TOTAL_FONT
    c.border = BOX
    c.fill = PatternFill("solid", fgColor=BLUE_SOFT)
    money(ws, r_acq, 2, "=SUM(B10:B16)", bold=True, fill=BLUE_SOFT)

    section_bar(ws, 19, "2 — CAPITAL IMPROVEMENTS  (auto-fed from each project tab's Actual total)", last_col="F")
    heads = ["Date completed", "Project", "Where tracked", "Capital? (Y/N)", "Cost"]
    for col, text in enumerate(heads, start=1):
        head_cell(ws, 20, col, text)
    # first project, wired to the flooring tab's actual total
    ws.cell(row=21, column=1).number_format = DATE_FMT
    ws.cell(row=21, column=2, value="Second floor flooring")
    ws.cell(row=21, column=3, value="Tab: 2nd Floor Flooring")
    ws.cell(row=21, column=4, value="Y")
    money(ws, 21, 5, f"='2nd Floor Flooring'!{ACTUAL_TOTAL_CELL}")
    dv_yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_yn)
    for r in range(21, 41):
        for col in range(1, 6):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col == 1:
                c.number_format = DATE_FMT
            if col == 5 and r > 21:
                c.number_format = CUR
        dv_yn.add(f"D{r}")
    r_imp = 41
    c = ws.cell(row=r_imp, column=1, value="Subtotal — capital improvements")
    c.font = TOTAL_FONT
    c.border = BOX
    c.fill = PatternFill("solid", fgColor=BLUE_SOFT)
    money(ws, r_imp, 5, '=SUMIF(D21:D40,"Y",E21:E40)', bold=True, fill=BLUE_SOFT)

    ws.merge_cells("A43:C43")
    c = ws.cell(row=43, column=1, value="ADJUSTED COST BASIS")
    c.font = Font(bold=True, size=14, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=GOLD)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[43].height = 26
    tc = money(ws, 43, 4, "=B17+E41", fmt=CUR0, bold=True, fill=YELLOW_SOFT)
    tc.font = Font(bold=True, size=14)
    ws.merge_cells("D43:E43")

    notes = [
        "Only capital improvements (things that add value or extend the property's life) add to basis — new floors yes, repainting a room no.",
        "Set Capital? to N for repairs you still want to track; they'll be listed but excluded from the basis total.",
        "Keep receipts and contracts for everything on this sheet. This is a tracking aid, not tax advice.",
    ]
    for i, n in enumerate(notes):
        c = ws.cell(row=45 + i, column=1, value="• " + n)
        c.font = NOTE_FONT
        ws.merge_cells(start_row=45 + i, start_column=1, end_row=45 + i, end_column=6)

    widths = {"A": 36, "B": 24, "C": 24, "D": 15, "E": 15, "F": 24}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"


# ================================================================= read me ==
def build_readme(ws):
    ws.sheet_view.showGridLines = False
    title_bar(ws, "4 INDIAN BROOK ROAD — HOUSE TRACKER", "Purchase & improvement history", last_col="B")
    lines = [
        ("", ""),
        ("What this is", "One workbook that tracks what the house cost and every project that adds to it. Share it in Google Drive so both of you can edit."),
        ("", ""),
        ("Basis Tracker", "Purchase price ($1,060,000, closing 7/17/2026) + closing costs + capital improvements = adjusted cost basis. Each project tab feeds one row of the improvements table."),
        ("2nd Floor Flooring", "First project. Compare the 3 Floor & Decor plank options (incl. underlayment, stair nose & risers, 5% contractor discount, est. tax, and Footprints labor), pick one in 'Selected option' (yellow cell), and the budget fills in automatically. Log payments in the cost log at the bottom — actuals and the Basis Tracker update themselves."),
        ("Project Template", "For the next project: right-click the tab > Duplicate, rename it, fill in your options/quotes, then add a row in the Basis Tracker improvements table pointing at the new tab's cell "
                             f"{ACTUAL_TOTAL_CELL} (its Actual total)."),
        ("", ""),
        ("Sources", "Labor: Footprints Floors of Central MA proposal #25584 (6/30/2026) — $9,057.60, materials excluded. Materials: Floor & Decor Waltham cart (7/2026) — Sapelo Shore / Big Sur / Gunstock Oak + Sentinel underlayment, plus 7 stair noses @ $29.99 and 13 risers @ $20; 5% contractor discount on materials."),
        ("", ""),
        ("Tip", "In Google Sheets everything here — dropdowns, cross-tab formulas, formatting — survives File > Import. Use 'Replace spreadsheet' when importing so tab references stay intact."),
    ]
    r = 3
    for label, text in lines:
        if label:
            ws.cell(row=r, column=1, value=label).font = Font(bold=True, size=11)
            c = ws.cell(row=r, column=2, value=text)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = max(30, 15 * (len(text) // 90 + 1))
        r += 1
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 110


# ==================================================================== main ==
def main():
    wb = Workbook()
    ws_readme = wb.active
    ws_readme.title = "Read Me"
    ws_basis = wb.create_sheet("Basis Tracker")
    ws_floor = wb.create_sheet("2nd Floor Flooring")
    ws_tmpl = wb.create_sheet("Project Template")

    build_readme(ws_readme)
    build_basis(ws_basis)
    build_project_sheet(ws_floor, FLOORING_CFG)
    build_project_sheet(ws_tmpl, TEMPLATE_CFG)

    ws_readme.sheet_properties.tabColor = "808080"
    ws_basis.sheet_properties.tabColor = GOLD
    ws_floor.sheet_properties.tabColor = NAVY
    ws_tmpl.sheet_properties.tabColor = "A6A6A6"

    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
