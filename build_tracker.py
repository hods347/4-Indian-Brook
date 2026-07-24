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
    section_bar(ws, COMP_SECTION, "OPTION COMPARISON")
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
    section_bar(ws, BUD_SECTION, "BUDGET vs. ACTUAL")
    for col, text in [(1, "Category"), (2, "Budget"),
                      (3, "Actual"), (4, "Remaining")]:
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

    # ---- cost log -----------------------------------------------------------
    section_bar(ws, LOG_SECTION, "COST LOG")
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
    "stair_nose": "=7*29.99",
    "stair_risers": "=13*20",
    "discount_rate": 0.05,
    "labor_demo": 2642.30,
    "labor_install": 4535.30,
    "labor_stairs": 1880.00,
    "comparison_notes": {
        "Item # / SKU": "All Floor & Decor. Sapelo Shore & Big Sur: waterproof hybrid resilient plank w/ cork pad, 8mm 7\"x51\". Gunstock Oak: waterproof rigid core LVP. Underlayment attached — no separate order.",
        "Stair nose trim": "7 stair noses x $29.99.",
        "Stair risers": "13 risers x $20.00.",
        "Contractor discount on materials": "5% off all materials through the flooring contractor.",
        "Est. sales tax (6.25% MA, materials)": "Estimate on discounted materials.",
        "Labor — demo & disposal": "Footprints: demo carpet, remove & return baseboards, trash removal.",
        "Labor — install": "Footprints: install click-lock floating floor.",
        "Labor — stairs (treads & risers)": "Footprints: treads (open 1 side) + risers.",
    },
}

TEMPLATE_CFG = {
    "title": "PROJECT TEMPLATE",
    "subtitle": "",
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
    "stair_nose": "",
    "stair_risers": "",
    "discount_rate": 0.0,
    "labor_demo": "",
    "labor_install": "",
    "labor_stairs": "",
    "comparison_notes": {},
}


# ========================================================= project estimates ==
def build_estimates(ws):
    """Preliminary quote log — one row per quote, before a project graduates
    to its own full tab. Count? = Y marks the quote that should roll into the
    totals, so competing quotes for the same project don't double-count."""
    ws.sheet_view.showGridLines = False
    title_bar(ws, "PROJECT ESTIMATES", "4 Indian Brook Road, Ashland, MA 01721", last_col="H")

    ws.cell(row=4, column=1, value="Total estimated — counted quotes").font = LABEL_FONT
    money(ws, 4, 2, '=SUMIFS(E9:E58,G9:G58,"Y")')
    ws.cell(row=5, column=1, value="Total estimated — counted & marked Now").font = LABEL_FONT
    money(ws, 5, 2, '=SUMIFS(E9:E58,G9:G58,"Y",F9:F58,"Now")')

    section_bar(ws, 7, "QUOTES", last_col="H")
    heads = ["Project", "Category", "Contractor / vendor", "Quote date",
             "Est. cost", "Timing", "Count?", "Notes"]
    for col, text in enumerate(heads, start=1):
        head_cell(ws, 8, col, text)

    starters = [
        ("HVAC system", "Mechanical"),
        ("Carpentry work", "Carpentry"),
    ]
    for i, (project, category) in enumerate(starters):
        ws.cell(row=9 + i, column=1, value=project)
        ws.cell(row=9 + i, column=2, value=category)
        ws.cell(row=9 + i, column=7, value="Y")

    dv_timing = DataValidation(
        type="list", formula1='"Now,Later,Undecided,Passed"', allow_blank=True
    )
    dv_count = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_timing)
    ws.add_data_validation(dv_count)
    for r in range(9, 59):
        for col in range(1, 9):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col == 4:
                c.number_format = DATE_FMT
            if col == 5:
                c.number_format = CUR
            if col == 7:
                c.alignment = Alignment(horizontal="center")
        dv_timing.add(f"F{r}")
        dv_count.add(f"G{r}")

    widths = {"A": 28, "B": 16, "C": 26, "D": 13, "E": 14, "F": 12, "G": 9, "H": 36}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A9"


# ================================================================== vendors ==
def build_vendors(ws):
    ws.sheet_view.showGridLines = False
    title_bar(ws, "VENDORS & CONTACTS", "4 Indian Brook Road, Ashland, MA 01721", last_col="H")

    section_bar(ws, 4, "DIRECTORY", last_col="H")
    heads = ["Company", "Trade / service", "Contact", "Phone", "Email",
             "Used for", "Use again?", "Notes"]
    for col, text in enumerate(heads, start=1):
        head_cell(ws, 5, col, text)

    starters = [
        ("Footprints Floors of Central MA", "Flooring install",
         "S. Donohoe", "(508) 422-4545", "sdonohoe@footprintsfloors.com",
         "2nd floor flooring", "", "ACS Custom Solutions Inc.; proposal #25584"),
        ("Floor & Decor — Waltham", "Flooring materials",
         "", "", "", "2nd floor flooring materials", "", ""),
    ]
    for i, row in enumerate(starters):
        for col, val in enumerate(row, start=1):
            ws.cell(row=6 + i, column=col, value=val or None)

    dv_again = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_again)
    for r in range(6, 46):
        for col in range(1, 9):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col == 7:
                c.alignment = Alignment(horizontal="center")
        dv_again.add(f"G{r}")

    widths = {"A": 30, "B": 18, "C": 16, "D": 16, "E": 30, "F": 24, "G": 11, "H": 34}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


# ====================================================== finishes & materials ==
def build_finishes(ws):
    ws.sheet_view.showGridLines = False
    title_bar(ws, "FINISHES & MATERIALS", "4 Indian Brook Road, Ashland, MA 01721", last_col="H")

    section_bar(ws, 4, "RECORD", last_col="H")
    heads = ["Room / area", "Item", "Brand / product", "Color / finish",
             "SKU / code", "Where purchased", "Date", "Notes"]
    for col, text in enumerate(heads, start=1):
        head_cell(ws, 5, col, text)

    starters = [
        ("Second floor", "Flooring — LVP", "", "", "", "Floor & Decor — Waltham", "", ""),
    ]
    for i, row in enumerate(starters):
        for col, val in enumerate(row, start=1):
            ws.cell(row=6 + i, column=col, value=val or None)

    for r in range(6, 56):
        for col in range(1, 9):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col == 7:
                c.number_format = DATE_FMT

    widths = {"A": 20, "B": 20, "C": 24, "D": 20, "E": 16, "F": 24, "G": 13, "H": 34}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


# ====================================================== appliances & systems ==
# (item, location, typical lifespan yrs, service cadence)
APPLIANCE_ITEMS = [
    ("Heating system (furnace / boiler)", "Basement", 20,
     "Annual tune-up; replace filters every 1-3 months (forced air)"),
    ("Central A/C / heat pump", "Exterior + attic", 15,
     "Annual tune-up; keep outdoor unit clear"),
    ("Water heater", "Basement", 12, "Flush tank annually; test T&P valve"),
    ("Refrigerator", "Kitchen", 13, "Vacuum coils annually"),
    ("Range / oven", "Kitchen", 15, ""),
    ("Dishwasher", "Kitchen", 10, "Clean filter quarterly"),
    ("Microwave", "Kitchen", 9, ""),
    ("Washer", "Laundry", 11, "Inspect supply hoses annually"),
    ("Dryer", "Laundry", 13, "Clean vent duct annually; lint trap every load"),
    ("Garbage disposal", "Kitchen", 10, ""),
    ("Sump pump", "Basement", 10, "Test quarterly; check before spring thaw"),
    ("Roof (asphalt shingle)", "Exterior", 25, "Inspect annually and after major storms"),
    ("Gutters & downspouts", "Exterior", 20, "Clean spring and fall"),
    ("Windows", "Whole house", 25, "Check seals/caulk annually"),
    ("Deck / porch", "Exterior", 20, "Reseal or restain every 2-3 years"),
    ("Garage door & opener", "Garage", 12, "Lubricate tracks/rollers annually"),
    ("Smoke / CO detectors", "Whole house", 10,
     "Test monthly; replace batteries annually"),
    ("Septic system (if applicable)", "Exterior", 30,
     "Pump every 2-3 years; Title 5 inspection at sale"),
    ("Irrigation system (if applicable)", "Exterior", 20,
     "Winterize every fall; startup check in spring"),
]


def build_appliances(ws):
    ws.sheet_view.showGridLines = False
    title_bar(ws, "APPLIANCES & SYSTEMS", "4 Indian Brook Road, Ashland, MA 01721", last_col="K")

    section_bar(ws, 4, "INVENTORY", last_col="K")
    heads = ["Item", "Location", "Make & model", "Serial #", "Install date",
             "Warranty expires", "Typical lifespan (yrs)",
             "Suggested replacement", "Service cadence", "Last serviced", "Notes"]
    for col, text in enumerate(heads, start=1):
        head_cell(ws, 5, col, text)

    first, last = 6, 6 + len(APPLIANCE_ITEMS) + 10  # seeded rows + spares
    for i, (item, loc, life, cadence) in enumerate(APPLIANCE_ITEMS):
        r = first + i
        ws.cell(row=r, column=1, value=item)
        ws.cell(row=r, column=2, value=loc)
        ws.cell(row=r, column=7, value=life)
        ws.cell(row=r, column=9, value=cadence or None)

    for r in range(first, last + 1):
        for col in range(1, 12):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col in (5, 6, 8, 10):
                c.number_format = DATE_FMT
            if col == 7:
                c.alignment = Alignment(horizontal="center")
            if col == 9:
                c.alignment = Alignment(wrap_text=True, vertical="center")
        # suggested replacement = install date + typical lifespan
        ws.cell(row=r, column=8,
                value=f'=IF(OR(E{r}="",G{r}=""),"",EDATE(E{r},12*G{r}))')

    widths = {"A": 30, "B": 15, "C": 22, "D": 16, "E": 13, "F": 13, "G": 10,
              "H": 14, "I": 40, "J": 13, "K": 26}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


# ============================================================ basis tracker ==
# Borrower-paid closing costs from the Closing Disclosure (7/13/2026), in the
# order they appear on page 2. "basis" = default flag for whether the item is
# added to cost basis (a starting point to confirm with a tax preparer).
# Seller-paid items (attorney's fees, some couriers, MA deed excise, realtor
# commissions) are excluded — they are not the buyer's costs.
CLOSING_COSTS = [
    ("Loan underwriting fee", 795.00, "N", "Loan cost"),
    ("Appraisal fee", 750.00, "N", "Loan cost"),
    ("Credit report", 319.00, "N", "Loan cost"),
    ("Flood certification", 5.00, "N", "Loan cost"),
    ("Loan Safe report", 6.30, "N", "Loan cost"),
    ("Lender's title insurance", 1988.00, "N", "Loan cost — protects lender"),
    ("Closing protection letter", 25.00, "Y", "Title service"),
    ("Courier fee (FedEx)", 100.00, "Y", "Settlement service"),
    ("Document preparation fee", 100.00, "Y", "Deed / document prep"),
    ("E-recording fee", 19.00, "Y", "Recording"),
    ("Municipal lien certificate", 50.00, "Y", "Title — Town of Ashland"),
    ("Plot plan / survey", 150.00, "Y", "Survey — Boston Survey"),
    ("Settlement fee", 700.00, "Y", "Settlement / closing"),
    ("Title examination", 325.00, "Y", "Title"),
    ("Recording fee — deed", 155.00, "Y", "Deed recording"),
    ("Recording fee — mortgage", 205.00, "N", "Loan cost"),
    ("Recording fee — additional", 120.00, "Y", "Balance of the $480 gov't recording line"),
    ("Owner's title insurance", 3222.00, "Y", "Owner's policy (optional)"),
    ("Homeowner's insurance (12 mo, prepaid)", 3951.00, "N", "Prepaid — not basis"),
    ("Prepaid interest (7/17–8/1)", 2082.75, "N", "Prepaid loan interest"),
    ("Property taxes (3 mo, prepaid)", 2922.51, "N", "Prepaid — not basis"),
    ("Property taxes escrow (2 mo)", 1948.34, "N", "Escrow deposit — not basis"),
]


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

    # ---- closing costs (per the Closing Disclosure) -----------------------
    section_bar(ws, 8, "CLOSING COSTS  (borrower-paid, per Closing Disclosure 7/13/2026)", last_col="F")
    for col, text in [(1, "Item"), (2, "Amount"), (3, "Adds to basis? (Y/N)")]:
        head_cell(ws, 9, col, text)
    head_cell(ws, 9, 4, "Notes")
    ws.merge_cells("D9:F9")

    cc_first = 10
    cc_last = cc_first + len(CLOSING_COSTS) - 1
    dv_basis = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_basis)
    for i, (label, amt, flag, note) in enumerate(CLOSING_COSTS):
        r = cc_first + i
        lc = ws.cell(row=r, column=1, value=label)
        lc.border = BOX
        money(ws, r, 2, amt, fmt=CUR)
        fc = ws.cell(row=r, column=3, value=flag)
        fc.border = BOX
        fc.alignment = Alignment(horizontal="center")
        dv_basis.add(f"C{r}")
        nc = ws.cell(row=r, column=4, value=note)
        nc.font = NOTE_FONT
        nc.border = BOX
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)

    r_cc_all = cc_last + 1        # total closing costs (reconciles to CD)
    c = ws.cell(row=r_cc_all, column=1, value="Total closing costs (borrower-paid)")
    c.border = BOX
    money(ws, r_cc_all, 2, f"=SUM(B{cc_first}:B{cc_last})")
    r_cc_basis = cc_last + 2      # closing costs that add to basis
    c = ws.cell(row=r_cc_basis, column=1, value="Closing costs added to basis")
    c.font = TOTAL_FONT
    c.border = BOX
    c.fill = PatternFill("solid", fgColor=BLUE_SOFT)
    money(ws, r_cc_basis, 2,
          f'=SUMIF(C{cc_first}:C{cc_last},"Y",B{cc_first}:B{cc_last})',
          bold=True, fill=BLUE_SOFT)

    r_acq = cc_last + 4           # acquisition basis = price + basis closing costs
    ws.merge_cells(f"A{r_acq}:A{r_acq}")
    c = ws.cell(row=r_acq, column=1, value="ACQUISITION BASIS  (purchase price + basis closing costs)")
    c.font = TOTAL_FONT
    c.border = BOX
    c.fill = PatternFill("solid", fgColor=BLUE_SOFT)
    money(ws, r_acq, 2, f"=B5+B{r_cc_basis}", bold=True, fill=BLUE_SOFT)

    # ---- capital improvements ---------------------------------------------
    imp_sec = r_acq + 2
    section_bar(ws, imp_sec, "CAPITAL IMPROVEMENTS", last_col="F")
    imp_head = imp_sec + 1
    heads = ["Date completed", "Project", "Where tracked", "Capital? (Y/N)", "Cost"]
    for col, text in enumerate(heads, start=1):
        head_cell(ws, imp_head, col, text)
    imp_first = imp_head + 1
    imp_last = imp_first + 19
    # first project, wired to the flooring tab's actual total
    ws.cell(row=imp_first, column=1).number_format = DATE_FMT
    ws.cell(row=imp_first, column=2, value="Second floor flooring")
    ws.cell(row=imp_first, column=3, value="Tab: 2nd Floor Flooring")
    ws.cell(row=imp_first, column=4, value="Y")
    money(ws, imp_first, 5, f"='2nd Floor Flooring'!{ACTUAL_TOTAL_CELL}")
    dv_yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_yn)
    for r in range(imp_first, imp_last + 1):
        for col in range(1, 6):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col == 1:
                c.number_format = DATE_FMT
            if col == 5 and r > imp_first:
                c.number_format = CUR
        dv_yn.add(f"D{r}")
    r_imp = imp_last + 1
    c = ws.cell(row=r_imp, column=1, value="Subtotal — capital improvements")
    c.font = TOTAL_FONT
    c.border = BOX
    c.fill = PatternFill("solid", fgColor=BLUE_SOFT)
    money(ws, r_imp, 5, f'=SUMIF(D{imp_first}:D{imp_last},"Y",E{imp_first}:E{imp_last})',
          bold=True, fill=BLUE_SOFT)

    # ---- adjusted cost basis ----------------------------------------------
    r_adj = r_imp + 2
    ws.merge_cells(f"A{r_adj}:C{r_adj}")
    c = ws.cell(row=r_adj, column=1, value="ADJUSTED COST BASIS")
    c.font = Font(bold=True, size=14, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=GOLD)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[r_adj].height = 26
    tc = money(ws, r_adj, 4, f"=B{r_acq}+E{r_imp}", fmt=CUR0, bold=True, fill=YELLOW_SOFT)
    tc.font = Font(bold=True, size=14)
    ws.merge_cells(f"D{r_adj}:E{r_adj}")

    note = ws.cell(row=r_adj + 2, column=1,
                   value="Seller-paid items (MA deed excise $4,856.40, realtor commissions "
                         "$63,900, attorney's fees, etc.) are excluded. The Adds-to-basis "
                         "flags are a starting point — confirm with your tax preparer.")
    note.font = NOTE_FONT
    ws.merge_cells(start_row=r_adj + 2, start_column=1, end_row=r_adj + 2, end_column=6)

    widths = {"A": 40, "B": 20, "C": 18, "D": 15, "E": 15, "F": 24}
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
        ("Basis Tracker", "Purchase price ($1,060,000, closing 7/17/2026) + basis-adding closing costs + capital improvements = adjusted cost basis. Closing costs are itemized from the Closing Disclosure with an Adds-to-basis flag; each project tab feeds one row of the improvements table."),
        ("2nd Floor Flooring", "First project. Compare the 3 Floor & Decor plank options (incl. stair nose & risers, 5% contractor discount, est. tax, and Footprints labor), pick one in 'Selected option' (yellow cell), and the budget fills in automatically. Log payments in the cost log at the bottom — actuals and the Basis Tracker update themselves."),
        ("Project Estimates", "Preliminary phase: log each quote you collect (HVAC, carpentry, etc.) with its estimated cost, and mark Timing (Now / Later / Undecided / Passed) to decide what to take on. Set Count? to Y on the one quote per project you'd actually use — competing quotes marked N stay listed but don't double-count in the totals. When a project is a go, give it its own tab from the template."),
        ("Vendors & Contacts", "Directory of contractors, suppliers, and service companies used on the house."),
        ("Finishes & Materials", "Room-by-room record of paint colors, flooring, fixtures, and their SKUs — for touch-ups and matching repairs later."),
        ("Appliances & Systems", "Inventory of the house's equipment with make/model/serial, warranty dates, and service cadence. Enter an install date and the suggested replacement date computes from the typical lifespan."),
        ("Project Template", "For the next project: right-click the tab > Duplicate, rename it, fill in your options/quotes, then add a row in the Basis Tracker improvements table pointing at the new tab's cell "
                             f"{ACTUAL_TOTAL_CELL} (its Actual total)."),
        ("", ""),
        ("Sources", "Labor: Footprints Floors of Central MA proposal #25584 (6/30/2026) — $9,057.60, materials excluded. Materials: Floor & Decor Waltham cart (7/2026) — Sapelo Shore / Big Sur / Gunstock Oak (underlayment attached to plank), plus 7 stair noses @ $29.99 and 13 risers @ $20; 5% contractor discount on materials."),
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
    ws_basis = wb.active
    ws_basis.title = "Basis Tracker"
    ws_floor = wb.create_sheet("2nd Floor Flooring")
    ws_est = wb.create_sheet("Project Estimates")
    ws_tmpl = wb.create_sheet("Project Template")
    ws_vend = wb.create_sheet("Vendors & Contacts")
    ws_fin = wb.create_sheet("Finishes & Materials")
    ws_appl = wb.create_sheet("Appliances & Systems")
    ws_readme = wb.create_sheet("Read Me")

    build_basis(ws_basis)
    build_project_sheet(ws_floor, FLOORING_CFG)
    build_estimates(ws_est)
    build_project_sheet(ws_tmpl, TEMPLATE_CFG)
    build_vendors(ws_vend)
    build_finishes(ws_fin)
    build_appliances(ws_appl)
    build_readme(ws_readme)

    ws_basis.sheet_properties.tabColor = GOLD
    ws_floor.sheet_properties.tabColor = NAVY
    ws_est.sheet_properties.tabColor = NAVY_LIGHT
    ws_tmpl.sheet_properties.tabColor = "A6A6A6"
    ws_vend.sheet_properties.tabColor = "548235"
    ws_fin.sheet_properties.tabColor = "548235"
    ws_appl.sheet_properties.tabColor = "548235"
    ws_readme.sheet_properties.tabColor = "808080"

    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
