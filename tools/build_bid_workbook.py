"""Builds 'New Excel bidding Spreadsheet.xlsx' for Paris Mechanical.

Rebuild of the 'Template-Empty - WORK IN PROGRESS.xlsm' bidding workbook:
full transparent cost (wage + burden, overhead, material incl. PST, subs, GCs)
with margin added on top, plus Budget, Schedule of Values (Knowify phase
naming) and Knowify actuals review tabs.

Usage: python tools/build_bid_workbook.py [output.xlsx] [knowify_job_summary.xlsx] [takeoff.xlsx]
The optional Knowify export prefills the KNOWIFY PHASES and KNOWIFY TIME tabs
(employee names are replaced with 'Employee nn').
"""
import datetime
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "New Excel bidding Spreadsheet.xlsx"
KNOWIFY_SRC = sys.argv[2] if len(sys.argv) > 2 else None
TAKEOFF_SRC = sys.argv[3] if len(sys.argv) > 3 else None

# ---------- styles ----------
FONT = "Arial"
NAVY = "1F3864"
F_IN = Font(name=FONT, size=10, color="0000FF")
F_CALC = Font(name=FONT, size=10, color="000000")
F_LINK = Font(name=FONT, size=10, color="008000")
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=11, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOPLINE = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

CUR = '$#,##0;($#,##0);"-"'
CUR2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
NUM = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
FAC = '0.00"x";(0.00"x");"-"'

wb = Workbook()


def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def put(ws, ref, value, font=F_CALC, fmt=None, fill=None, bold=False, align=None, border=None):
    c = ws[ref]
    c.value = value
    c.font = Font(name=font.name, size=font.size, bold=bold or font.bold, italic=font.italic, color=font.color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def header_row(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 30


def title(ws, text, sub=None):
    put(ws, "A1", text, F_TITLE)
    if sub:
        put(ws, "A2", sub, F_NOTE)


def note(ws, ref, text):
    ws[ref].comment = Comment(text, "Estimating")


def ok_cf(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'],
                                  fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, color="006100", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"CHECK"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, color="9C0006", bold=True)))


def grey_zero(ws, rng, first_cell):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"{first_cell}=0"], font=Font(name=FONT, color="BFBFBF")))


DIVS = [("P", "Plumbing"), ("H", "Hydronic"), ("V", "Ventilation"), ("AC", "Air Conditioning")]
# Phases mirror how Knowify jobs are set up. Material on 03/04 rows rolls into 05 Material Supply.
PHASES = [
    "01 Shop Drawings, Permits & Mobilization",
    "02 Foundation & Underground",
    "03 Rough-in (Labour)",
    "04 Finishing (Labour)",
    "05 Material Supply",
    "06 Fixtures & Equipment Supply",
    "07 Testing, TAB & Commissioning",
    "08 Completion & Final Documents",
    "09 Office, Supervision & General",
]
LEVEL_PHASES = (2, 3)
MAT_SUPPLY = 4
FIX_EQ = 5
CLASSES = ["Apprentice Lvl 1-2", "Apprentice Lvl 3-6", "Apprentice Lvl 7-8", "Journeyman Lvl 1-2", "Foreman & PM"]

LEVELS_438 = [("P1", "Parkade", None), ("L1", "Level 1 (CRU, lobby)", None), ("L2", "Level 2 (parking, garbage)", None)] + \
    [(f"L{n}", f"Level {n}", 8) for n in range(3, 13)] + [("L13", "Level 13 (amenity, live/work)", None), ("ROOF", "Roof", None)]
LEVELS_OLD = [("U/G", "Ground Work", None), ("P-02", "Parking 2", None), ("PRKG", "Parking 1", None),
          ("L-01", "Level 1", 24), ("L-02", "Level 2", 25), ("L-03", "Level 3", 25), ("L-04", "Level 4", 25),
          ("L-05", "Level 5", 24), ("L-06", "Level 6", None), ("L-07", "Level 7", None), ("L-08", "Level 8", None),
          ("L-09", "Level 9", None), ("L-10", "Level 10", None), ("ROOF", "Roof", None)]
LEVELS = LEVELS_438 if TAKEOFF_SRC else LEVELS_OLD
LV_N = 24

# INPUTS row map
LV_FIRST = 40
LV_LAST = LV_FIRST + LV_N - 1          # 63
LV_TOT = LV_LAST + 1                   # 64
PH_FIRST = LV_TOT + 3                  # 67
PH_LAST = PH_FIRST + len(PHASES) - 1   # 75
CG_HDR = PH_LAST + 2                   # 77
CG_FIRST = CG_HDR + 2                  # 79
COGS = [  # company cost breakdown, MF/CM, % of all costs (from estimator, Oct 2026)
    ("Materials", 0.2725, "Material: pipe, fittings, valves (+PST, consumables)"),
    ("Equipment", 0.1593, "Fixtures & Equipment Supply (phase 06 material)"),
    ("Other Job Costs", 0.0305, "GC rows typed 'Other Job Costs' + contingency"),
    ("Labour", 0.2122, "Labour $ + GC rows typed 'Labour'"),
    ("Subs & Safety", 0.1908, "Subtrades + GC rows typed 'Subs & Safety'"),
    ("Admin Labour (expenses)", 0.0954, "Overhead $/hr x hours + GC rows typed 'Admin & Overhead'"),
    ("General Overhead (net of other income)", 0.0451, "(combined with Admin Labour on BID SUMMARY)"),
]
CG_LAST = CG_FIRST + len(COGS) - 1     # 85


def ph_ref(p):
    return f"INPUTS!$B${PH_FIRST + p}"


PH_RANGE = f"INPUTS!$B${PH_FIRST}:$B${PH_LAST}"
DIV_RANGE = "INPUTS!$C$33:$C$36"

# =====================================================================
# READ ME
# =====================================================================
ws = wb.active
ws.title = "READ ME"
setw(ws, {"A": 4, "B": 34, "C": 100})
title(ws, "New Excel Bidding Spreadsheet: How It Works",
      "Rebuilt from 'Template-Empty - WORK IN PROGRESS.xlsm'. Full transparent cost first, margin added on top.")
r = 4
put(ws, f"B{r}", "FLOW", F_SUB); r += 1
flow = [
    ("1. INPUTS", "Project info, durations, PST, consumables, contingency, target margin per division, levels, your typical COGS %."),
    ("2. LABOUR RATES", "Crew mix per trade (same layout as the original PRJ INFO rates block), shown beside Knowify actual rates and mix."),
    ("2b. TAKEOFF / TAKEOFF CHECK", "Paste the takeoff provider's 'Takeoff Summary' sheet into TAKEOFF at A1. TAKEOFF CHECK lists every line: priced, missing, no unit cost, or TBC."),
    ("3. PLUMBING EST / HVAC EST", "Left to right: takeoff qty, unit rates, TRUE COST, optional ADJUSTMENTS, FINAL COST. Same columns on both sheets."),
    ("4. SUBS & GC", "Subtrade quotes and general conditions. Allocate by % to each division, or leave % blank to auto-split by direct cost."),
    ("4b. CONTINGENCY", "2% general allowance on budgeted cost (adjustable) plus specific risk items. Added to cost before the grand total."),
    ("5. BID SUMMARY", "Cost by type and division, then margin (self-perform and subs). Back-check vs company cost breakdown. Holds every tie-out check."),
    ("6. BUDGET", "Cost budget per division and phase. Phases match Knowify, so the field budget and the actuals line up."),
    ("7. SCHEDULE OF VALUES", "Contract split into Knowify-style lines (e.g. 'P - BLDG A - L1 Plumbing Rough-in (Labour)'), with the Knowify cost budget for each line."),
    ("8. KNOWIFY PHASES / TIME", "Paste the 'Actual vs Budget - Phases' and 'Time' sheets from a Knowify Job Summary export, cell A1, exactly as exported."),
    ("9. KNOWIFY REVIEW", "Reads the two paste tabs: margin after overhead, cost mix, actual crew rates and mix, and how each phase ran vs its budget."),
    ("9b. RATE VARIANCE", "Template vs Knowify actual rates vs a saved snapshot, the effect on this bid, and overhead for reference."),
    ("10. AUDIT NOTES", "Formula problems found in the original template, decisions made, and open questions."),
]
for k, v in flow:
    put(ws, f"B{r}", k, F_BOLD)
    put(ws, f"C{r}", v, align=Alignment(wrap_text=True, vertical="top"))
    r += 1
r += 1
put(ws, f"B{r}", "COLOUR LEGEND", F_SUB); r += 1
for k, v, f, fl in [
    ("Blue text, light yellow fill", "Input. Type here.", F_IN, FILL_IN),
    ("Bright yellow fill", "Key assumption. Review on every bid.", F_IN, FILL_KEY),
    ("Black text", "Formula. Do not type over.", F_CALC, None),
    ("Green text", "Link from another tab.", F_LINK, None),
]:
    put(ws, f"B{r}", k, f, fill=fl)
    put(ws, f"C{r}", v)
    r += 1
r += 1
put(ws, f"B{r}", "MARGIN vs MARKUP", F_SUB); r += 1
for line in [
    "This workbook prices with MARGIN: Contract = Total Cost / (1 - Margin %).",
    "Example: cost $100,000 at 15% margin = $117,647 contract, $17,647 gross profit (17.6% markup).",
    "The original template used 15% MARKUP: $100,000 x 1.15 = $115,000, which is only a 13.0% margin.",
]:
    put(ws, f"C{r}", line); r += 1
r += 1
put(ws, f"B{r}", "PHASES (match Knowify)", F_SUB); r += 1
for line in [
    "Tag each estimate row with the phase the WORK happens in. Rough-in and finishing rows carry both labour and material.",
    "Labour on 03 Rough-in and 04 Finishing is split by level on the SOV; their material rolls into 05 Material Supply automatically,",
    "the same way Knowify jobs are set up (level lines are labour only, material supply is its own line).",
]:
    put(ws, f"C{r}", line); r += 1
r += 1
put(ws, f"B{r}", "UPDATING KNOWIFY DATA", F_SUB); r += 1
for i, line in enumerate([
    "In Knowify, open the job > Summary > Project Summary report and download the Excel file.",
    "In that file, select the used cells of 'Actual vs Budget - Phases' (A1 to column V, last row). Copy.",
    "Paste into cell A1 of KNOWIFY PHASES (clear old data first). Do the same for the 'Time' sheet into KNOWIFY TIME.",
    "On KNOWIFY REVIEW, type the job name, contract (Total Amount) and committed cost from the export's Summary sheet.",
    "If a role or phase shows 'Unmapped', add it to the mapping tables on the right of KNOWIFY REVIEW.",
], 1):
    put(ws, f"C{r}", f"{i}. {line}"); r += 1

# =====================================================================
# INPUTS
# =====================================================================
inp = wb.create_sheet("INPUTS")
setw(inp, {"A": 3, "B": 42, "C": 24, "D": 16, "E": 16, "F": 14, "G": 44, "J": 10, "K": 10})
title(inp, "INPUTS", "Blue cells are inputs. Bright yellow = key assumption to review on every bid.")
put(inp, "B3", "PROJECT INFORMATION", F_SUB)
for row, lab, val in [
    (4, "Project Name", "438 West Pender Street" if TAKEOFF_SRC else None), (5, "Quote Number", None),
    (6, "Project Type", "Mixed-use residential (80 suites + CRU)" if TAKEOFF_SRC else "Multi-Family"),
    (7, "Owner", None), (8, "General Contractor", "Marcon" if TAKEOFF_SRC else None), (9, "Architect", None),
    (10, "Mech. Engineer", "Arcadis / Edge Consultants" if TAKEOFF_SRC else None),
    (11, "Site Address", "438 West Pender Street" if TAKEOFF_SRC else "2650 East 41st Ave"), (12, "City", "Vancouver"),
    (13, "Proposal Date", None), (14, "Proposal Rev", None), (15, "Prepared By", "Paris Mechanical Ltd"),
    (16, "Drawings", "Final IFT Rev 5 (Sept 11 2026)" if TAKEOFF_SRC else "Issued For Tender"),
    (17, "Gross Floor Area (sq ft)", None),
]:
    put(inp, f"B{row}", lab)
    put(inp, f"C{row}", val, F_IN, fill=FILL_IN, border=BOX)
inp["C13"].number_format = "yyyy-mm-dd"
inp["C17"].number_format = NUM
put(inp, "B18", "Number of Units")
put(inp, "C18", f"=D{LV_TOT}", fmt=NUM, border=BOX)

put(inp, "B20", "KEY ASSUMPTIONS", F_SUB)
for row, lab, val, fmt, cmt in [
    (21, "P&H project duration (months)", 24, NUM, "From original PRJ INFO H34."),
    (22, "HVAC project duration (months)", 24, NUM, "From original PRJ INFO H43."),
    (23, "PST on materials", 0.07, PCT, "BC PST 7%. From original PRJ INFO H52. Applied to material rows marked Y."),
    (24, "GST (shown on proposal only, not a cost)", 0.05, PCT, "From original PRJ INFO H51. GST is recoverable, so it is not in cost."),
    (25, "Consumables (% of material)", 0.02, PCT, "Original used 2% of material (P&H SUMMARY H46, HVAC SUMMARY D30)."),
    (26, "Contingency allowance % (set on CONTINGENCY tab)", "=CONTINGENCY!$C$4", PCT, "Default 2% of budgeted cost. Change it on the CONTINGENCY tab."),
    (27, "Warranty reserve (% of fixture & equipment supply)", 0.02, PCT, "Original: 2 years x 1% of equipment+fixtures (PRJ SUMMARY rows 75-78)."),
    (28, "Contract rounding (round up to nearest $)", 100, CUR, "Original rounded each division up to the nearest $100."),
    (29, "Holdback (BC Builders Lien Act)", 0.10, PCT, "Used on the SOV for progress billing."),
]:
    put(inp, f"B{row}", lab)
    if isinstance(val, str):
        put(inp, f"C{row}", val, F_LINK, fmt=fmt, border=BOX)
    else:
        put(inp, f"C{row}", val, F_IN, fmt=fmt, fill=FILL_KEY, border=BOX)
    put(inp, f"G{row}", cmt, F_NOTE)

put(inp, "B31", "TARGET MARGIN BY DIVISION", F_SUB)
header_row(inp, 32, ["Division", "Code", "Self-Perform Margin %", "Equivalent Markup %", "Crew", "Subs Margin %"], start_col=2)
put(inp, "J31", "Margin list", F_NOTE)
put(inp, "K31", "Subs list", F_NOTE)
for i, m in enumerate(range(8, 21)):
    put(inp, f"J{32 + i}", m / 100, F_NOTE, fmt="0%")
for i, m in enumerate(range(10, 16)):
    put(inp, f"K{32 + i}", m / 100, F_NOTE, fmt="0%")
dv_m = DataValidation(type="list", formula1="INPUTS!$J$32:$J$44", allow_blank=False)
dv_s = DataValidation(type="list", formula1="INPUTS!$K$32:$K$37", allow_blank=False)
inp.add_data_validation(dv_m)
inp.add_data_validation(dv_s)
dv_m.add("D33:D36")
dv_s.add("G33:G36")
for i, (code, name) in enumerate(DIVS):
    row = 33 + i
    put(inp, f"B{row}", name, F_IN, fill=FILL_IN, border=BOX)
    put(inp, f"C{row}", code, border=BOX, align=Alignment(horizontal="center"))
    put(inp, f"D{row}", 0.15, F_IN, fmt=PCT, fill=FILL_KEY, border=BOX)
    put(inp, f"E{row}", f"=IFERROR(D{row}/(1-D{row}),0)", fmt=PCT, border=BOX)
    put(inp, f"F{row}", "P&H" if code in ("P", "H") else "HVAC", border=BOX, align=Alignment(horizontal="center"))
    put(inp, f"G{row}", 0.12, F_IN, fmt=PCT, fill=FILL_KEY, border=BOX)
put(inp, "H33", "Drop-downs: self-perform 8% to 20% (standard 15%), subs 10% to 15%.", F_NOTE)
put(inp, "H34", "Margin = profit / contract. Self-perform margin applies to labour, overhead,", F_NOTE)
put(inp, "H35", "material, GCs and contingency; subs margin applies to subtrades only.", F_NOTE)
put(inp, "H36", "Division names feed the SOV line names (Knowify uses 'HVAC' for ventilation).", F_NOTE)

put(inp, "B38", "LEVELS (SOV splits rough-in and finishing labour by level)", F_SUB)
header_row(inp, 39, ["Tag (e.g. BLDG A - L1)", "Floor", "Units", "SOV Weight %"], start_col=2)
for i in range(LV_N):
    row = LV_FIRST + i
    tag, floor, units = LEVELS[i] if i < len(LEVELS) else (None, None, None)
    put(inp, f"B{row}", tag, F_IN, fill=FILL_IN, border=BOX)
    put(inp, f"C{row}", floor, F_IN, fill=FILL_IN, border=BOX)
    put(inp, f"D{row}", units, F_IN, fmt=NUM, fill=FILL_IN, border=BOX)
    put(inp, f"E{row}", f"=IFERROR(D{row}/$D${LV_TOT},0)", F_IN, fmt=PCT, fill=FILL_IN, border=BOX)
put(inp, f"B{LV_TOT}", "TOTAL", F_BOLD)
put(inp, f"D{LV_TOT}", f"=SUM(D{LV_FIRST}:D{LV_LAST})", fmt=NUM, bold=True, border=TOPLINE)
put(inp, f"E{LV_TOT}", f"=SUM(E{LV_FIRST}:E{LV_LAST})", fmt=PCT, bold=True, border=TOPLINE)
put(inp, f"F{LV_TOT}", f'=IF(ABS(E{LV_TOT}-1)<0.0001,"OK","CHECK")', bold=True)
ok_cf(inp, f"F{LV_TOT}")
put(inp, f"G{LV_FIRST}", "Units per level from original PRJ INFO C6:C10.", F_NOTE)
put(inp, f"G{LV_FIRST + 1}", "Weight defaults to units share. Type over it to give parking,", F_NOTE)
put(inp, f"G{LV_FIRST + 2}", "U/G or roof a share. Weights must total 100%.", F_NOTE)
put(inp, f"G{LV_FIRST + 3}", "Multi-building jobs: use tags like 'BLDG A - L1' to match Knowify.", F_NOTE)

put(inp, f"B{PH_FIRST - 1}", "PHASES (drop-down list used on every tab; matches Knowify setup)", F_SUB)
for i, ph in enumerate(PHASES):
    put(inp, f"B{PH_FIRST + i}", ph, F_IN, fill=FILL_IN, border=BOX)
put(inp, f"G{PH_FIRST}", "Rename freely but keep 9 rows and keep the order:", F_NOTE)
put(inp, f"G{PH_FIRST + 1}", "03/04 are split by level; 05 collects 03/04 material.", F_NOTE)

put(inp, f"B{CG_HDR}", "COMPANY COST BREAKDOWN (MF/CM, % of all costs): back-check for every bid", F_SUB)
header_row(inp, CG_HDR + 1, ["Category", "% of All Costs", "Compared on BID SUMMARY against"], start_col=2)
for i, (lab, val, maps) in enumerate(COGS):
    put(inp, f"B{CG_FIRST + i}", lab, border=BOX)
    put(inp, f"C{CG_FIRST + i}", val, F_IN, fmt='0.00%', fill=FILL_KEY, border=BOX)
    put(inp, f"D{CG_FIRST + i}", maps, F_NOTE)
put(inp, f"B{CG_LAST + 1}", "TOTAL", F_BOLD)
put(inp, f"C{CG_LAST + 1}", f"=SUM(C{CG_FIRST}:C{CG_LAST})", fmt='0.00%', bold=True, border=TOPLINE)
put(inp, f"D{CG_LAST + 1}", f'=IF(ABS(C{CG_LAST + 1}-1)<=0.01,"OK","CHECK")', bold=True)
ok_cf(inp, f"D{CG_LAST + 1}")
put(inp, f"B{CG_LAST + 2}", "Overhead per $1 of field labour (Admin + General) / Labour", F_BOLD)
put(inp, f"C{CG_LAST + 2}", f"=IFERROR((C{CG_FIRST + 5}+C{CG_FIRST + 6})/C{CG_FIRST + 3},0)", fmt="0.000", bold=True, fill=FILL_TOT, border=BOX)
put(inp, f"G{CG_FIRST}", "From estimator (Oct 2026). Replace with the updated breakdown when ready.", F_NOTE)
put(inp, f"G{CG_FIRST + 1}", "Total is 100.58% as given (rounding); check allows +/-1%.", F_NOTE)
put(inp, f"G{CG_LAST + 2}", "Used on LABOUR RATES to show the overhead $/hr your books imply.", F_NOTE)
inp.freeze_panes = "A3"

# =====================================================================
# Shared range helpers
# =====================================================================
EST_FIRST, EST_LAST = 7, 156
REV = "'KNOWIFY REVIEW'"
KP_FIRST, KP_LAST = 4, 400      # Knowify phase rows (export data starts row 4)
KT_FIRST, KT_LAST = 3, 300      # Knowify time rows (export data starts row 3)
PERF_FIRST = 26
PERF_LAST = PERF_FIRST + len(DIVS) * len(PHASES) - 1
CREW_HDR = PERF_LAST + 3
CREW_FIRST = CREW_HDR + 2
CREW_SUB = CREW_FIRST + 7       # 5 classes, Excluded, Unmapped, then crew subtotal
CREW_ALL = CREW_FIRST + 8       # all roles
CLASS_EXCL = "Office / Safety (in GC)"


def est_rng(sheet, col):
    return f"'{sheet}'!${col}${EST_FIRST}:${col}${EST_LAST}"


def est_sumifs(col, div_expr, phase_expr=None):
    parts = []
    for sh in ("PLUMBING EST", "HVAC EST"):
        s = f"SUMIFS({est_rng(sh, col)},{est_rng(sh, 'B')},{div_expr}"
        if phase_expr:
            s += f",{est_rng(sh, 'C')},{phase_expr}"
        parts.append(s + ")")
    return "+".join(parts)


def est_mat(div_expr, phases, eq):
    """Material (col V) for phases; eq='Y' only flagged rows, '<>Y' unflagged, None all."""
    parts = []
    for sh in ("PLUMBING EST", "HVAC EST"):
        for q in phases:
            s_ = f"SUMIFS({est_rng(sh, 'V')},{est_rng(sh, 'C')},{ph_ref(q)}"
            if div_expr:
                s_ += f",{est_rng(sh, 'B')},{div_expr}"
            if eq:
                s_ += f",{est_rng(sh, 'L')},\"{eq}\""
            parts.append(s_ + ")")
    return "+".join(parts)


def mat_supply(div_expr):
    return est_mat(div_expr, (2, 3, 4), "<>Y")


def mat_equip(div_expr):
    return est_mat(div_expr, (FIX_EQ,), None) + "+" + est_mat(div_expr, (2, 3, 4), "Y")


def rv(col, a=KP_FIRST, b=KP_LAST):
    return f"{REV}!${col}${a}:${col}${b}"


CREW_CLS = f"{REV}!$B${CREW_FIRST}:$B${CREW_FIRST + 4}"
CREW_RATE = f"{REV}!$E${CREW_FIRST}:$E${CREW_FIRST + 4}"
CREW_MIX = f"{REV}!$F${CREW_FIRST}:$F${CREW_FIRST + 4}"
KN_BLENDED = f"{REV}!$E${CREW_SUB}"
KN_BLENDED_ALL = f"{REV}!$E${CREW_ALL}"
OH_RATIO = f"INPUTS!$C${CG_LAST + 2}"

# =====================================================================
# LABOUR RATES
# =====================================================================
lr = wb.create_sheet("LABOUR RATES")
setw(lr, {"A": 3, "B": 40, "C": 14, "D": 12, "E": 14, "F": 12, "G": 13, "H": 12, "I": 13, "J": 11, "K": 16, "L": 3, "M": 70})
title(lr, "LABOUR RATES (crew mix)", "Same layout as the original PRJ INFO rates block. Both trades tie back the same way.")
put(lr, "B2", "UPDATED 2026-06-29", F_BOLD, fill=FILL_KEY)


def crew_block(top, label, hours_formula):
    put(lr, f"B{top}", label, F_SUB)
    put(lr, f"F{top}", "Rate source:", F_BOLD, align=Alignment(horizontal="right"))
    put(lr, f"G{top}", "Knowify actual", F_IN, fill=FILL_KEY, border=BOX)
    dv_src = DataValidation(type="list", formula1='"Template,Knowify actual"', allow_blank=False)
    lr.add_data_validation(dv_src)
    dv_src.add(f"G{top}")
    header_row(lr, top + 1, ["Classification", "Template Rate $/hr", "Template Mix", "Knowify Actual $/hr",
                             "Knowify Actual Mix", "Rate Used", "Mix Used", "Weighted $/hr", "Hours",
                             "Cost by Wage Breakdown"], start_col=2)
    classes = [(CLASSES[0], 53, 0.40), (CLASSES[1], 62.5, 0.30), (CLASSES[2], 66.5, 0.05),
               (CLASSES[3], 73, 0.11), (CLASSES[4], 96, 0.14)]
    first, last = top + 2, top + 6
    t = last + 1
    use_kn = f'AND($G${top}="Knowify actual",ABS($F${t}-1)<0.001)'
    for i, (nm, rate, mix) in enumerate(classes):
        r_ = first + i
        put(lr, f"B{r_}", nm, F_IN, fill=FILL_IN, border=BOX)
        put(lr, f"C{r_}", rate, F_IN, fmt=CUR2, fill=FILL_KEY, border=BOX)
        put(lr, f"D{r_}", mix, F_IN, fmt=PCT, fill=FILL_KEY, border=BOX)
        put(lr, f"E{r_}", f'=IFERROR(INDEX({CREW_RATE},MATCH(B{r_},{CREW_CLS},0)),0)', F_LINK, fmt=CUR2, border=BOX)
        put(lr, f"F{r_}", f'=IFERROR(INDEX({CREW_MIX},MATCH(B{r_},{CREW_CLS},0)),0)', F_LINK, fmt=PCT, border=BOX)
        put(lr, f"G{r_}", f"=IF({use_kn},E{r_},C{r_})", fmt=CUR2, border=BOX)
        put(lr, f"H{r_}", f"=IF({use_kn},F{r_},D{r_})", fmt=PCT, border=BOX)
        put(lr, f"I{r_}", f"=G{r_}*H{r_}", fmt=CUR2, border=BOX)
        put(lr, f"J{r_}", f"=$C${t + 4}*H{r_}", fmt=NUM, border=BOX)
        put(lr, f"K{r_}", f"=J{r_}*G{r_}", fmt=CUR, border=BOX)
    put(lr, f"B{t}", "Weighted average burdened wage", F_BOLD)
    put(lr, f"C{t}", f"=SUMPRODUCT(C{first}:C{last},D{first}:D{last})", fmt=CUR2, bold=True, border=TOPLINE)
    put(lr, f"D{t}", f"=SUM(D{first}:D{last})", fmt=PCT, bold=True, border=TOPLINE)
    put(lr, f"E{t}", f"=SUMPRODUCT(E{first}:E{last},F{first}:F{last})", F_LINK, fmt=CUR2, bold=True, border=TOPLINE)
    put(lr, f"F{t}", f"=SUM(F{first}:F{last})", fmt=PCT, bold=True, border=TOPLINE)
    put(lr, f"H{t}", f"=SUM(H{first}:H{last})", fmt=PCT, bold=True, border=TOPLINE)
    put(lr, f"I{t}", f"=SUM(I{first}:I{last})", fmt=CUR2, bold=True, fill=FILL_TOT, border=TOPLINE)
    put(lr, f"J{t}", f"=SUM(J{first}:J{last})", fmt=NUM, bold=True, border=TOPLINE)
    put(lr, f"K{t}", f"=SUM(K{first}:K{last})", fmt=CUR, bold=True, border=TOPLINE)
    put(lr, f"B{t + 1}", "Overhead $/hr (salaried staff, office, general)")
    put(lr, f"C{t + 1}", 32.71, F_IN, fmt=CUR2, fill=FILL_KEY, border=BOX)
    put(lr, f"D{t + 1}", "Manual entry per bid. Reference figures on RATE VARIANCE.", F_NOTE)
    put(lr, f"B{t + 2}", "FULLY LOADED RATE $/hr", F_BOLD)
    put(lr, f"C{t + 2}", f"=I{t}+C{t + 1}", fmt=CUR2, bold=True, fill=FILL_TOT, border=BOX)
    put(lr, f"B{t + 3}", "Mix used = 100% and hours x rate ties to breakdown")
    put(lr, f"C{t + 3}", f'=IF(AND(ABS(H{t}-1)<0.0001,ABS(K{t}-J{t}*I{t})<1),"OK","CHECK")', bold=True)
    ok_cf(lr, f"C{t + 3}")
    put(lr, f"B{t + 4}", "Estimated hours (from estimate tabs)")
    put(lr, f"C{t + 4}", hours_formula, F_LINK, fmt=NUM, border=BOX)
    for k, txt in enumerate([
        "Template rates and mix: original PRJ INFO D36:E40 (fully burdened).",
        "Knowify actuals: KNOWIFY TIME, all trades, burden included (confirmed).",
        "Office PM, admin and safety coordinator hours are left out of the Knowify mix",
        "because they are priced on SUBS & GC.",
        "Default source is Knowify actual (estimator decision, Oct 2026). Without Knowify data the template is used.",
    ]):
        put(lr, f"M{first + k}", txt, F_NOTE)
    return t


ph_t = crew_block(4, "PLUMBING & HYDRONIC CREW", "=" + est_sumifs("W", '"P"') + "+" + est_sumifs("W", '"H"'))
hv_t = crew_block(19, "HVAC CREW (ventilation & A/C)", "=" + est_sumifs("W", '"V"') + "+" + est_sumifs("W", '"AC"'))
PH_WAGE, PH_OH = f"'LABOUR RATES'!$I${ph_t}", f"'LABOUR RATES'!$C${ph_t + 1}"
HV_WAGE, HV_OH = f"'LABOUR RATES'!$I${hv_t}", f"'LABOUR RATES'!$C${hv_t + 1}"

cmp_r = hv_t + 7
put(lr, f"B{cmp_r}", "RATE COMPARISON (per hour)", F_SUB)
header_row(lr, cmp_r + 1, ["Rate Basis", "$/hr", "", "", "P&H Est. Hours", "Labour (+OH) Cost"], start_col=2)
for i, (lab, f) in enumerate([
    ("Original template line items (Apprentice 1-2 only)", "=C6"),
    ("Original template 'Budget with' (simple average)", "=AVERAGE(C6:C10)"),
    ("Template crew-mix weighted burdened wage", f"=C{ph_t}"),
    ("Rate used for this bid (P&H)", f"=I{ph_t}"),
    ("Fully loaded rate used: wage + overhead input", f"=C{ph_t + 2}"),
    ("Knowify actual crew average (job pasted)", f"={KN_BLENDED}"),
    ("Knowify actual + overhead", f"={KN_BLENDED}+C{ph_t + 1}"),
]):
    r_ = cmp_r + 2 + i
    put(lr, f"B{r_}", lab, border=BOX)
    put(lr, f"C{r_}", f, fmt=CUR2, border=BOX)
    put(lr, f"F{r_}", f"=$C${ph_t + 4}", fmt=NUM, border=BOX)
    put(lr, f"G{r_}", f"=C{r_}*F{r_}", fmt=CUR, border=BOX)
put(lr, f"K{cmp_r + 2}", "If Knowify's cost per hour is well below the template rate, either Knowify", F_NOTE)
put(lr, f"K{cmp_r + 3}", "is missing burden, or the template rates are high. Settle which is true", F_NOTE)
put(lr, f"K{cmp_r + 4}", "before trusting either for bids.", F_NOTE)
lr.freeze_panes = "A3"

# =====================================================================
# ESTIMATE SHEETS
# =====================================================================
EST_HEAD = ["Takeoff Ref", "Div", "Phase", "Description", "Takeoff Qty", "Manual Qty", "Qty Used", "Unit",
            "Material $/Unit", "Hrs/Unit", "PST? (Y/N)", "Equip Supply? (Y/N)",
            "Material $ (incl PST)", "Hours", "Labour $", "Overhead $", "TRUE COST $",
            "Adj Material $", "Adj Hours %", "Adj Hours", "Adjustment Note",
            "Material $", "Hours", "Labour $", "Overhead $", "TOTAL COST $",
            "Knowify Labour Act/Bud", "Knowify Material Act/Bud", "History Flag", "Takeoff Flag"]
TK = "TAKEOFF CHECK"
TK_FIRST, TK_LAST = 2, 600


def tk(col):
    return f"'{TK}'!${col}${TK_FIRST}:${col}${TK_LAST}"


def classify(sec, tag):
    t = tag.upper()
    if sec == "A":
        if t in ("WC", "SH", "L", "SK"):
            return "P", 3, "Y"
        return ("P", 2, "Y") if t == "W/M" else ("P", 2, "")
    if sec == "B":
        if t.startswith("FCU"):
            return "AC", 2, "Y"
        if t == "TSTAT":
            return "AC", 3, "Y"
        if t == "RG-FCU":
            return "AC", 3, ""
        if t in ("ERV-A", "EF-A"):
            return "V", 2, "Y"
        if t == "AP-ERV":
            return "V", 3, ""
        return "V", 2, ""
    if sec == "C":
        if t.startswith("CU") or t in ("ITM", "DSE401B71", "CP"):
            return "AC", 5, "Y"
        if t == "REF-INS":
            return "AC", 2, ""
        return "AC", 2, "Y"
    if sec == "D":
        return "V", 5, "Y"
    if sec == "E":
        if t == "DET":
            return "P", 1, ""
        if t in ("HT", "TP"):
            return "P", 2, ""
        return "P", 5, "Y"
    if sec == "F":
        return "P", 2, "Y"
    if sec == "G":
        return "P", 3, "Y"
    if sec == "H":
        return ("P", 5, "") if t == "FO-TANK" else ("P", 2, "")
    if sec == "I":
        if t.startswith("S-1"):
            return "V", 3, "Y"
        if t == "FLEX-FCU":
            return "AC", 2, ""
        if t == "FSD":
            return "V", 2, "Y"
        return "V", 2, "Y"
    if sec == "J":
        return {"REF-V": ("AC", 2, ""), "DUCT": ("V", 2, "")}.get(t, ("P", 2, ""))
    return "P", 2, ""


def load_takeoff(path):
    """Return raw rows (for the paste tab) and unique estimate rows per Takeoff Ref."""
    src = load_workbook(path, data_only=True)["Takeoff Summary"]
    raw = [[c.value for c in row] for row in src.iter_rows(min_row=1, max_row=src.max_row, max_col=10)]
    items, seen, sec = [], {}, ""
    for row in raw:
        a, b = row[0], row[1]
        if a and not b and isinstance(a, str) and len(a) > 1 and a[1] == ".":
            sec = a[0]
            continue
        if not b or b == "Tag" or not sec:
            continue
        ref = f"{sec}:{b}"
        if ref in seen:
            continue
        div, ph, eq = classify(sec, str(b))
        model = row[3]
        desc = row[2] if model in (None, "—", "-", "?") else f"{row[2]} [{model}]"
        seen[ref] = True
        items.append((ref, div, ph, desc, row[7], eq))
    return raw, items


TAKEOFF_RAW, TAKEOFF_ITEMS = load_takeoff(TAKEOFF_SRC) if TAKEOFF_SRC else ([], [])
if TAKEOFF_ITEMS:
    PL_ITEMS = [(ref, div, ph, desc, unit, None, None, eq, None) for ref, div, ph, desc, unit, eq in TAKEOFF_ITEMS if div in ("P", "H")]
    HV_ITEMS = [(ref, div, ph, desc, unit, None, None, eq, None) for ref, div, ph, desc, unit, eq in TAKEOFF_ITEMS if div in ("V", "AC")]
else:
    PL_ITEMS = [
        (None, "P", 2, "2-Piece Bathroom (WC, LAV)", "ea", 550, 3, "", None),
        (None, "P", 2, "3-Piece Bathroom (WC, LAV, TUB)", "ea", 860, 5, "", None),
        (None, "P", 2, "Domestic Water Insuite Piping", "suite", 800, 12, "", None),
    ]
    HV_ITEMS = [(None, "V", 2, "Distribution Ductwork", "lb", None, None, "", None)]


def build_est(name, items, subtitle):
    ws = wb.create_sheet(name)
    setw(ws, {"A": 15, "B": 5, "C": 30, "D": 48, "E": 10, "F": 10, "G": 9, "H": 6, "I": 12, "J": 9, "K": 7, "L": 8,
              "M": 14, "N": 9, "O": 13, "P": 12, "Q": 14, "R": 12, "S": 9, "T": 9, "U": 22,
              "V": 14, "W": 9, "X": 13, "Y": 12, "Z": 15, "AA": 11, "AB": 11, "AC": 20, "AD": 18})
    title(ws, name, subtitle)
    put(ws, "A3", "TOTALS", F_BOLD)
    for col in "MNOPQRTVWXYZ":
        put(ws, f"{col}3", f"=SUM({col}{EST_FIRST}:{col}{EST_LAST})", fmt=NUM if col in "NTW" else CUR, bold=True,
            fill=FILL_TOT, border=BOX)
    for rng_, lab, fill in (("A5:L5", "QUANTITY & UNIT RATES (inputs)", FILL_IN), ("M5:Q5", "TRUE COST", FILL_TOT),
                            ("R5:U5", "ADJUSTMENTS (optional)", FILL_KEY), ("V5:Z5", "FINAL COST (true cost + adjustments)", FILL_TOT),
                            ("AA5:AD5", "CHECKS", FILL_SEC)):
        ws.merge_cells(rng_)
        c = ws[rng_.split(":")[0]]
        c.value = lab
        c.font = F_BOLD
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
    put(ws, "A4", "Qty comes from TAKEOFF via the Takeoff Ref (Manual Qty overrides it). Labour $ = hours x crew rate; "
                  "Overhead $ = hours x overhead $/hr (LABOUR RATES). Adjustments sit beside true cost so every change stays visible.", F_NOTE)
    header_row(ws, 6, EST_HEAD)
    for i in range(EST_FIRST, EST_LAST + 1):
        idx = i - EST_FIRST
        item = items[idx] if idx < len(items) else None
        vals = {"K": "Y"}
        if item:
            ref, div, ph, desc, unit, mat, hrs, eq, mq = item
            vals.update({"A": ref, "B": div, "C": PHASES[ph], "D": desc, "F": mq, "H": unit, "I": mat, "J": hrs, "L": eq or None})
        for col in "ABCDFHIJKL":
            c = put(ws, f"{col}{i}", vals.get(col), F_IN, fill=FILL_IN, border=BOX)
            if col == "F":
                c.number_format = NUM1
            if col == "I":
                c.number_format = CUR2
            if col == "J":
                c.number_format = '0.00;(0.00);"-"'
        for col in "RSTU":
            c = put(ws, f"{col}{i}", None, F_IN, fill=FILL_KEY if col != "U" else FILL_IN, border=BOX)
            c.number_format = {"R": CUR, "S": PCT, "T": NUM1, "U": "@"}[col]
        put(ws, f"E{i}", f'=IF($A{i}="","",SUMIFS({tk("D")},{tk("B")},$A{i}))', F_LINK, fmt=NUM1, border=BOX)
        put(ws, f"G{i}", f'=IF($F{i}<>"",$F{i},N($E{i}))', fmt=NUM1, border=BOX)
        rate_w = f'IF(OR($B{i}="V",$B{i}="AC"),{HV_WAGE},{PH_WAGE})'
        rate_o = f'IF(OR($B{i}="V",$B{i}="AC"),{HV_OH},{PH_OH})'
        put(ws, f"M{i}", f'=$G{i}*$I{i}*(1+IF($K{i}="N",0,INPUTS!$C$23))', fmt=CUR, border=BOX)
        put(ws, f"N{i}", f"=$G{i}*$J{i}", fmt=NUM1, border=BOX)
        put(ws, f"O{i}", f"=$N{i}*{rate_w}", fmt=CUR, border=BOX)
        put(ws, f"P{i}", f"=$N{i}*{rate_o}", fmt=CUR, border=BOX)
        put(ws, f"Q{i}", f"=$M{i}+$O{i}+$P{i}", fmt=CUR, bold=True, border=BOX)
        put(ws, f"V{i}", f"=$M{i}+$R{i}", fmt=CUR, border=BOX)
        put(ws, f"W{i}", f"=$N{i}*(1+$S{i})+$T{i}", fmt=NUM1, border=BOX)
        put(ws, f"X{i}", f"=$W{i}*{rate_w}", fmt=CUR, border=BOX)
        put(ws, f"Y{i}", f"=$W{i}*{rate_o}", fmt=CUR, border=BOX)
        put(ws, f"Z{i}", f"=$V{i}+$X{i}+$Y{i}", fmt=CUR, bold=True, border=BOX)
        crit_l = f"{rv('AB')},$B{i},{rv('AC')},$C{i},{rv('AD')},1"
        put(ws, f"AA{i}", f'=IF(OR($B{i}="",$C{i}=""),"",IFERROR(SUMIFS({rv("AF")},{crit_l})/SUMIFS({rv("AE")},{crit_l}),""))',
            fmt=FAC, border=BOX)
        mat_ph = f'IF($L{i}="Y",{ph_ref(FIX_EQ)},IF(OR($C{i}={ph_ref(2)},$C{i}={ph_ref(3)}),{ph_ref(MAT_SUPPLY)},$C{i}))'
        crit_m = f"{rv('AB')},$B{i},{rv('AC')},{mat_ph},{rv('AD')},1"
        put(ws, f"AB{i}", f'=IF(OR($B{i}="",$C{i}=""),"",IFERROR(SUMIFS({rv("AH")},{crit_m})/SUMIFS({rv("AG")},{crit_m}),""))',
            fmt=FAC, border=BOX)
        put(ws, f"AC{i}", f'=IF(MAX(N($AA{i}),N($AB{i}))>1.1,"Past job ran over budget","")', border=BOX)
        put(ws, f"AD{i}", f'=IF($A{i}="","",IFERROR(INDEX({tk("F")},MATCH($A{i},{tk("B")},0)),"Ref not in takeoff"))', border=BOX)
    for col in ("AA", "AB"):
        ws.conditional_formatting.add(f"{col}{EST_FIRST}:{col}{EST_LAST}",
                                      FormulaRule(formula=[f'AND(ISNUMBER({col}{EST_FIRST}),{col}{EST_FIRST}>1.1)'],
                                                  fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add(f"AD{EST_FIRST}:AD{EST_LAST}", FormulaRule(
        formula=[f'AND(AD{EST_FIRST}<>"",AD{EST_FIRST}<>"Priced")'], fill=PatternFill("solid", fgColor="FFEB9C")))
    for formula1, col in ((DIV_RANGE, "B"), (PH_RANGE, "C"), ('"Y,N"', "K"), ('"Y,N"', "L")):
        dv = DataValidation(type="list", formula1=formula1, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}{EST_FIRST}:{col}{EST_LAST}")
    note(ws, "A6", "Section letter + tag from the takeoff (e.g. A:WC = in-suite water closets). Leave blank for manual lines.")
    note(ws, "F6", "Type a number to override the takeoff quantity. Leave blank to use the takeoff.")
    note(ws, "L6", "Y = this row's material is equipment or fixtures. It goes to 06 Fixtures & Equipment Supply on BUDGET/SOV "
                   "instead of 05 Material Supply.")
    note(ws, "R6", "Material dollars added (or negative to remove). No PST added on adjustments.")
    note(ws, "S6", "Productivity factor on hours, e.g. 15% = 15% more hours. Use the Knowify history columns as a guide.")
    note(ws, "T6", "Lump hours added (or negative).")
    note(ws, "AA6", "How the same Div + Phase ran vs budget on the last Knowify job pasted. A warning light, not a correction.")
    ws.freeze_panes = "E7"
    ws.auto_filter.ref = f"A6:AD{EST_LAST}"


build_est("PLUMBING EST", PL_ITEMS, "Plumbing (P) and Hydronic (H). Lines generated from the takeoff; add manual lines below.")
build_est("HVAC EST", HV_ITEMS, "Ventilation (V) and Air Conditioning (AC). Identical columns and formulas to PLUMBING EST.")

# =====================================================================
# TAKEOFF (paste) + TAKEOFF CHECK
# =====================================================================
to = wb.create_sheet("TAKEOFF")
setw(to, {"A": 24, "B": 18, "C": 50, "D": 24, "E": 10, "F": 10, "G": 9, "H": 7, "I": 18, "J": 40})
for r_, row in enumerate(TAKEOFF_RAW, 1):
    for c_, v in enumerate(row, 1):
        if v is not None:
            to.cell(row=r_, column=c_, value=v).font = Font(name=FONT, size=10)
tc = wb.create_sheet(TK)
setw(tc, {"A": 8, "B": 22, "C": 54, "D": 10, "E": 7, "F": 22, "G": 16, "H": 3, "I": 70})
header_row(tc, 1, ["Section", "Takeoff Ref", "Description", "Qty", "Unit", "Status", "Estimate Cost $"])
T_ = "TAKEOFF"
for r_ in range(TK_FIRST, TK_LAST + 1):
    a, b = f"{T_}!A{r_}", f"{T_}!B{r_}"
    prev = '""' if r_ == TK_FIRST else f"A{r_ - 1}"
    put(tc, f"A{r_}", f'=IF(AND({a}<>"",{b}="",MID({a},2,1)="."),LEFT({a},1),{prev})')
    put(tc, f"B{r_}", f'=IF(OR({b}="",{b}="Tag",A{r_}=""),"",A{r_}&":"&{b})')
    put(tc, f"C{r_}", f'=IF(B{r_}="","",{T_}!C{r_})')
    put(tc, f"D{r_}", f'=IF(B{r_}="","",IF(ISNUMBER({T_}!G{r_}),{T_}!G{r_},0))', fmt=NUM1)
    put(tc, f"E{r_}", f'=IF(B{r_}="","",{T_}!H{r_})')
    est_cost = "+".join(f"SUMIFS({est_rng(sh, 'Z')},{est_rng(sh, 'A')},B{r_})" for sh in ("PLUMBING EST", "HVAC EST"))
    in_est = "+".join(f"COUNTIF({est_rng(sh, 'A')},B{r_})" for sh in ("PLUMBING EST", "HVAC EST"))
    put(tc, f"G{r_}", f'=IF(B{r_}="","",{est_cost})', fmt=CUR)
    put(tc, f"F{r_}", f'=IF(B{r_}="","",IF(NOT(ISNUMBER({T_}!G{r_})),"TBC in takeoff",IF(({in_est})=0,"MISSING FROM ESTIMATE",'
                      f'IF(D{r_}=0,"Zero qty",IF(G{r_}=0,"NO UNIT COST","Priced")))))')
for txt, fill in (("MISSING FROM ESTIMATE", "FFC7CE"), ("NO UNIT COST", "FFEB9C"), ("TBC in takeoff", "FFEB9C"), ("Priced", "C6EFCE")):
    tc.conditional_formatting.add(f"F{TK_FIRST}:F{TK_LAST}", CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                  fill=PatternFill("solid", fgColor=fill)))
put(tc, "I1", "HOW IT WORKS", F_BOLD)
for k, txt in enumerate([
    "TAKEOFF holds the provider's 'Takeoff Summary' sheet, pasted at cell A1 (columns A to J only).",
    "Each line gets a Takeoff Ref = section letter + tag (e.g. A:WC). Lines with the same ref add together.",
    "Estimate rows pick up the quantity by ref. This list shows whether every takeoff line is in the estimate and priced.",
    "MISSING FROM ESTIMATE: add a row with this ref on PLUMBING EST or HVAC EST.",
    "NO UNIT COST: the row exists but material $/unit and hrs/unit are blank.",
    "TBC in takeoff: the provider has not counted it yet (see their Assumptions & RFIs).",
], 2):
    put(tc, f"I{k}", txt, F_NOTE)
put(tc, "I9", "Lines missing from estimate", F_BOLD)
put(tc, "I10", f'=COUNTIF(F{TK_FIRST}:F{TK_LAST},"MISSING FROM ESTIMATE")', bold=True)
put(tc, "I11", "Lines with no unit cost", F_BOLD)
put(tc, "I12", f'=COUNTIF(F{TK_FIRST}:F{TK_LAST},"NO UNIT COST")', bold=True)
put(tc, "I13", "Lines TBC in takeoff", F_BOLD)
put(tc, "I14", f'=COUNTIF(F{TK_FIRST}:F{TK_LAST},"TBC in takeoff")', bold=True)
grey_zero(tc, f"A{TK_FIRST}:G{TK_LAST}", f'LEN($B{TK_FIRST})')
tc.freeze_panes = "A2"

# =====================================================================
# SUBS & GC
# =====================================================================
sg = wb.create_sheet("SUBS & GC")
setw(sg, {"A": 5, "B": 34, "C": 24, "D": 34, "E": 8, "F": 9, "G": 13, "H": 15, "I": 8, "J": 8, "K": 8, "L": 8,
          "M": 13, "N": 13, "O": 13, "P": 13, "Q": 9})
title(sg, "SUBTRADES & GENERAL CONDITIONS", "Allocation %: enter a split per division, or leave all four blank to auto-split by direct cost.")
put(sg, "H3", "Direct cost by division", F_NOTE, align=Alignment(horizontal="right"))
put(sg, "H4", "Auto split share", F_NOTE, align=Alignment(horizontal="right"))
for i, (code, nm) in enumerate(DIVS):
    col = get_column_letter(9 + i)
    put(sg, f"{col}3", "=" + est_sumifs("Z", f'"{code}"'), F_LINK, fmt=CUR)
    put(sg, f"{col}4", f"=IFERROR({col}3/SUM($I$3:$L$3),0)", fmt=PCT)
SG_HEAD = ["#", "Item", "Company / Basis", "Phase", "Qty", "Unit", "Rate $", "Amount $",
           "% P", "% H", "% V", "% AC", "Plumbing $", "Hydronic $", "Ventilation $", "A/C $", "Check"]
SUB_FIRST, SUB_LAST = 7, 28
GC_FIRST, GC_LAST = 33, 60
SUB_COLS = ["M", "N", "O", "P"]


def alloc_row(r_):
    for i in range(4):
        pc = get_column_letter(9 + i)
        put(sg, f"{SUB_COLS[i]}{r_}", f"=$H{r_}*IF(SUM($I{r_}:$L{r_})=0,{pc}$4,{pc}{r_})", fmt=CUR, border=BOX)
    put(sg, f"Q{r_}", f'=IF($H{r_}=0,"",IF(ABS(SUM(M{r_}:P{r_})-$H{r_})<0.01,"OK","CHECK"))', bold=True, border=BOX)


def alloc_inputs(r_, split):
    for i, code in enumerate(["P", "H", "V", "AC"]):
        put(sg, f"{get_column_letter(9 + i)}{r_}", split.get(code), F_IN, fmt='0%;;""', fill=FILL_IN, border=BOX)


put(sg, "A5", "SUBTRADES (each becomes its own SOV line, as in Knowify)", F_SUB)
header_row(sg, 6, SG_HEAD)
SUBS = [
    ("Insulation and Heat Trace", "Adler", 2, {"P": 1}),
    ("Duct Insulation", None, 2, {"V": 0.5, "AC": 0.5}),
    ("Fire Sprinklers", None, 2, {"P": 1}),
    ("P&H Controls", None, 5, {"P": 0.5, "H": 0.5}),
    ("HVAC Controls", "Olympic Controls", 5, {"V": 0.5, "AC": 0.5}),
    ("TAB and Commissioning", "Western Mechanical", 6, {"V": 1}),
    ("A/C Commissioning", None, 6, {"AC": 1}),
    ("Canning / Piping Plans", "Paris Mechanical", 0, {"P": 1}),
    ("Coring / Scanning", None, 2, {}),
    ("Fixtures Caulking", "Paris Mechanical", 3, {"P": 1}),
    ("Water Balancing", "Western Mechanical", 6, {"P": 1}),
    ("Sump Pumps", None, 5, {"P": 1}),
    ("Chemical Treatment", None, 6, {"H": 1}),
    ("P&H Seismic / Schedule", None, 2, {"P": 0.5, "H": 0.5}),
    ("HVAC Seismic / Schedule", None, 2, {"V": 0.5, "AC": 0.5}),
    ("As-built Drawings", "Paris Mechanical", 7, {}),
    ("O&M Manuals", "Paris Mechanical", 7, {}),
    ("Plumbing Bond", None, 0, {"P": 1}),
    ("HVAC Bond", None, 0, {"V": 0.5, "AC": 0.5}),
    ("Welding", None, 2, {"V": 1}),
    ("Gas Service / Other", None, 2, {}),
]
for k in range(SUB_FIRST, SUB_LAST + 1):
    idx = k - SUB_FIRST
    s = SUBS[idx] if idx < len(SUBS) else None
    put(sg, f"A{k}", idx + 1, border=BOX, align=Alignment(horizontal="center"))
    put(sg, f"B{k}", s[0] if s else None, F_IN, fill=FILL_IN, border=BOX)
    put(sg, f"C{k}", s[1] if s else None, F_IN, fill=FILL_IN, border=BOX)
    put(sg, f"D{k}", PHASES[s[2]] if s else None, F_IN, fill=FILL_IN, border=BOX)
    put(sg, f"E{k}", 1, F_IN, fmt=NUM1, fill=FILL_IN, border=BOX)
    put(sg, f"F{k}", "LS", F_IN, fill=FILL_IN, border=BOX)
    put(sg, f"G{k}", None, F_IN, fmt=CUR, fill=FILL_IN, border=BOX)
    put(sg, f"H{k}", f"=E{k}*G{k}", fmt=CUR, border=BOX)
    alloc_inputs(k, s[3] if s else {})
    alloc_row(k)
put(sg, f"B{SUB_LAST + 1}", "SUBTRADES TOTAL", F_BOLD)
for col in "HMNOP":
    put(sg, f"{col}{SUB_LAST + 1}", f"=SUM({col}{SUB_FIRST}:{col}{SUB_LAST})", fmt=CUR, bold=True, fill=FILL_TOT, border=TOPLINE)

put(sg, f"A{GC_FIRST - 2}", "GENERAL CONDITIONS (Qty x Rate)", F_SUB)
header_row(sg, GC_FIRST - 1, SG_HEAD)
MAXDUR = "=MAX(INPUTS!$C$21,INPUTS!$C$22)"
GC = [
    ("Supervision - Plumbing", "=INPUTS!$C$21", "months", None, 8, {"P": 1}),
    ("Supervision - Hydronic", "=INPUTS!$C$21", "months", None, 8, {"H": 1}),
    ("Supervision - Ventilation", "=INPUTS!$C$22", "months", None, 8, {"V": 1}),
    ("Supervision - A/C", "=INPUTS!$C$22", "months", None, 8, {"AC": 1}),
    ("Office Project Manager (dedicated to this job)", MAXDUR, "months", 7500, 8, {}),
    ("Health & Safety", MAXDUR, "months", None, 8, {}),
    ("Bonus - Plumbing", 1, "LS", None, 8, {"P": 1}),
    ("Bonus - Hydronic", 1, "LS", None, 8, {"H": 1}),
    ("Bonus - Ventilation", 1, "LS", None, 8, {"V": 1}),
    ("Bonus - A/C", 1, "LS", None, 8, {"AC": 1}),
    ("Bonus - Office", 1, "LS", None, 8, {}),
    ("Shop Drawings", 1, "LS", None, 0, {}),
    ("Plumbing & Gas Permit", 1, "LS", None, 0, {"P": 1}),
    ("Prints", 1, "LS", None, 0, {}),
    ("Rentals", 4, "months", None, 8, {}),
    ("Trailer", 12, "months", None, 8, {}),
    ("Legal Fees", MAXDUR, "months", None, 8, {}),
    ("Freight", MAXDUR, "months", None, 8, {}),
]
for code, nm in DIVS:
    GC.append((f"Warranty Reserve - {nm}", 1, "LS",
               f'=ROUNDUP(INPUTS!$C$27*({mat_equip(chr(34) + code + chr(34))}),-2)', 7, {code: 1}))
for k in range(GC_FIRST, GC_LAST + 1):
    idx = k - GC_FIRST
    g = GC[idx] if idx < len(GC) else None
    put(sg, f"A{k}", idx + 1, border=BOX, align=Alignment(horizontal="center"))
    put(sg, f"B{k}", g[0] if g else None, F_IN, fill=FILL_IN, border=BOX)
    put(sg, f"C{k}", None, F_IN, fill=FILL_IN, border=BOX)
    put(sg, f"D{k}", PHASES[g[4]] if g else None, F_IN, fill=FILL_IN, border=BOX)
    qty = g[1] if g else None
    put(sg, f"E{k}", qty, F_LINK if isinstance(qty, str) else F_IN, fmt=NUM1, fill=FILL_IN, border=BOX)
    put(sg, f"F{k}", g[2] if g else None, F_IN, fill=FILL_IN, border=BOX)
    rate = g[3] if g else None
    put(sg, f"G{k}", rate, F_CALC if isinstance(rate, str) else F_IN, fmt=CUR, fill=FILL_IN, border=BOX)
    put(sg, f"H{k}", f"=E{k}*G{k}", fmt=CUR, border=BOX)
    alloc_inputs(k, g[5] if g else {})
    alloc_row(k)
GC_TYPES = ["Labour", "Other Job Costs", "Subs & Safety", "Admin & Overhead"]
put(sg, f"R{SUB_FIRST - 1}", "COGS Type", F_HDR, fill=FILL_HDR, border=BOX, align=Alignment(horizontal="center", wrap_text=True))
put(sg, f"R{GC_FIRST - 1}", "COGS Type", F_HDR, fill=FILL_HDR, border=BOX, align=Alignment(horizontal="center", wrap_text=True))
for k in range(SUB_FIRST, SUB_LAST + 1):
    put(sg, f"R{k}", "Subs & Safety", border=BOX)
for k in range(GC_FIRST, GC_LAST + 1):
    nm_ = sg[f"B{k}"].value or ""
    if nm_.startswith(("Supervision", "Office Project Manager", "Bonus")):
        typ = "Labour"
    elif nm_.startswith("Health & Safety"):
        typ = "Subs & Safety"
    elif nm_:
        typ = "Other Job Costs"
    else:
        typ = None
    put(sg, f"R{k}", typ, F_IN, fill=FILL_IN, border=BOX)
dv_t = DataValidation(type="list", formula1='"' + ",".join(GC_TYPES) + '"', allow_blank=True)
sg.add_data_validation(dv_t)
dv_t.add(f"R{GC_FIRST}:R{GC_LAST}")
sg.column_dimensions["R"].width = 17
note(sg, f"R{GC_FIRST - 1}", "Which company cost category this line belongs to. Used for the back-check on BID SUMMARY.")
put(sg, f"C{GC_FIRST + 4}", "Dedicated office PM, per month", F_NOTE)
note(sg, f"G{GC_FIRST + 4}", "$7,500/month from original PRJ INFO D53. Confirmed: the overhead $/hr covers salaried "
                             "staff company wide; this is the PM assigned to this job, so it is not a double count. "
                             "Auto-splits by direct cost across divisions.")
note(sg, f"B{GC_FIRST}", "Confirmed: field supervision here is separate from the Foreman & PM share of the crew mix.")
note(sg, f"B{GC_FIRST + 5}", "Knowify charges the Health & Safety Coordinator's time to the job (237 hrs on 405 Marie Place).")
put(sg, f"B{GC_LAST + 1}", "GENERAL CONDITIONS TOTAL", F_BOLD)
for col in "HMNOP":
    put(sg, f"{col}{GC_LAST + 1}", f"=SUM({col}{GC_FIRST}:{col}{GC_LAST})", fmt=CUR, bold=True, fill=FILL_TOT, border=TOPLINE)
ok_cf(sg, f"Q{SUB_FIRST}:Q{GC_LAST}")
dv = DataValidation(type="list", formula1=PH_RANGE, allow_blank=True)
sg.add_data_validation(dv)
dv.add(f"D{SUB_FIRST}:D{SUB_LAST}")
dv.add(f"D{GC_FIRST}:D{GC_LAST}")
sg.freeze_panes = "C7"

# =====================================================================
# BID SUMMARY
# =====================================================================
bs = wb.create_sheet("BID SUMMARY")
setw(bs, {"A": 6, "B": 22, "C": 10, "D": 15, "E": 13, "F": 14, "G": 13, "H": 13, "I": 13, "J": 12, "K": 15,
          "L": 10, "M": 15, "N": 15, "O": 14, "P": 9, "Q": 9, "R": 11, "S": 10, "T": 9, "U": 9})
title(bs, "BID SUMMARY: full cost, then margin")
put(bs, "A2", '=IF(INPUTS!C4="","(enter project name on INPUTS)",INPUTS!C4)&"  |  "&INPUTS!C11&", "&INPUTS!C12', F_LINK)
header_row(bs, 5, ["Code", "Division", "Hours", "Labour $ (wage+burden)", "Overhead $", "Material $ (incl PST)",
                   "Consumables $", "Subtrades $", "Gen. Cond. $", "Contingency $", "TOTAL COST $",
                   "Self-Perform Margin", "Sell Price $", "CONTRACT $ (rounded)", "Gross Profit $", "Actual Margin",
                   "Markup on Cost", "Per Unit $", "Per Sq Ft $", "Hrs / Unit", "Subs Margin"])
for i, (code, nm) in enumerate(DIVS):
    r_ = 6 + i
    put(bs, f"A{r_}", f"=INPUTS!C{33 + i}", F_LINK, border=BOX, align=Alignment(horizontal="center"))
    put(bs, f"B{r_}", f"=INPUTS!B{33 + i}", F_LINK, border=BOX)
    put(bs, f"C{r_}", "=" + est_sumifs("W", f"$A{r_}"), fmt=NUM, border=BOX)
    put(bs, f"D{r_}", "=" + est_sumifs("X", f"$A{r_}"), fmt=CUR, border=BOX)
    put(bs, f"E{r_}", "=" + est_sumifs("Y", f"$A{r_}"), fmt=CUR, border=BOX)
    put(bs, f"F{r_}", "=" + est_sumifs("V", f"$A{r_}"), fmt=CUR, border=BOX)
    put(bs, f"G{r_}", f"=F{r_}*INPUTS!$C$25", fmt=CUR, border=BOX)
    put(bs, f"H{r_}", f"='SUBS & GC'!{SUB_COLS[i]}{SUB_LAST + 1}", F_LINK, fmt=CUR, border=BOX)
    put(bs, f"I{r_}", f"='SUBS & GC'!{SUB_COLS[i]}{GC_LAST + 1}", F_LINK, fmt=CUR, border=BOX)
    put(bs, f"J{r_}", f"=CONTINGENCY!$G${8 + i}", F_LINK, fmt=CUR, border=BOX)
    put(bs, f"K{r_}", f"=SUM(D{r_}:J{r_})", fmt=CUR, bold=True, fill=FILL_TOT, border=BOX)
    put(bs, f"L{r_}", f"=INPUTS!D{33 + i}", F_LINK, fmt=PCT, border=BOX)
    put(bs, f"U{r_}", f"=INPUTS!G{33 + i}", F_LINK, fmt=PCT, border=BOX)
    put(bs, f"M{r_}", f"=IFERROR((K{r_}-H{r_})/(1-L{r_}),0)+IFERROR(H{r_}/(1-U{r_}),0)", fmt=CUR, border=BOX)
    put(bs, f"N{r_}", f"=IF(M{r_}<=0,0,CEILING(M{r_},INPUTS!$C$28))", fmt=CUR, bold=True, fill=FILL_TOT, border=BOX)
    put(bs, f"O{r_}", f"=N{r_}-K{r_}", fmt=CUR, border=BOX)
    put(bs, f"P{r_}", f"=IFERROR(O{r_}/N{r_},0)", fmt=PCT, border=BOX)
    put(bs, f"Q{r_}", f"=IFERROR(O{r_}/K{r_},0)", fmt=PCT, border=BOX)
    put(bs, f"R{r_}", f"=IFERROR(N{r_}/INPUTS!$C$18,0)", fmt=CUR, border=BOX)
    put(bs, f"S{r_}", f"=IFERROR(N{r_}/INPUTS!$C$17,0)", fmt=CUR2, border=BOX)
    put(bs, f"T{r_}", f"=IFERROR(C{r_}/INPUTS!$C$18,0)", fmt=NUM1, border=BOX)
put(bs, "B10", "TOTAL MECHANICAL", F_BOLD)
for col in "CDEFGHIJKMNO":
    put(bs, f"{col}10", f"=SUM({col}6:{col}9)", fmt=NUM if col == "C" else CUR, bold=True, fill=FILL_TOT, border=TOPLINE)
put(bs, "P10", "=IFERROR(O10/N10,0)", fmt=PCT, bold=True, border=TOPLINE)
put(bs, "Q10", "=IFERROR(O10/K10,0)", fmt=PCT, bold=True, border=TOPLINE)
put(bs, "R10", "=IFERROR(N10/INPUTS!$C$18,0)", fmt=CUR, bold=True, border=TOPLINE)
put(bs, "S10", "=IFERROR(N10/INPUTS!$C$17,0)", fmt=CUR2, bold=True, border=TOPLINE)
put(bs, "T10", "=IFERROR(C10/INPUTS!$C$18,0)", fmt=NUM1, bold=True, border=TOPLINE)
put(bs, "M11", "GST (not in contract cost)", F_NOTE)
put(bs, "N11", "=N10*INPUTS!$C$24", fmt=CUR)
put(bs, "M12", "Contract + GST", F_NOTE)
put(bs, "N12", "=N10+N11", fmt=CUR, bold=True)
put(bs, "B11", "Sell price = (cost excluding subs) / (1 - self-perform margin) + subs / (1 - subs margin). Contingency sits inside cost, before the grand total.", F_NOTE)

FIX_BID = mat_equip(None)


def gc_type(t):
    return f"SUMIFS('SUBS & GC'!$H${GC_FIRST}:$H${GC_LAST},'SUBS & GC'!$R${GC_FIRST}:$R${GC_LAST},\"{t}\")"


put(bs, "B14", "BACK-CHECK vs COMPANY COST BREAKDOWN (% of all costs)", F_SUB)
header_row(bs, 15, ["Category", "", "This Bid $", "% of Bid Cost", "Knowify Job %", "Company Typical %",
                    "Bid vs Typical", "Flag"], start_col=2)
bs.merge_cells("B15:C15")
mix = [
    ("Materials (pipe, fittings, valves)", f"=F10+G10-({FIX_BID})", f"={REV}!G10", f"=INPUTS!C{CG_FIRST}"),
    ("Equipment (fixtures & equipment supply)", f"={FIX_BID}", f"={REV}!G11", f"=INPUTS!C{CG_FIRST + 1}"),
    ("Other Job Costs (+ contingency)", f"={gc_type('Other Job Costs')}+J10", f"={REV}!G12", f"=INPUTS!C{CG_FIRST + 2}"),
    ("Labour", f"=D10+{gc_type('Labour')}", f"={REV}!G13", f"=INPUTS!C{CG_FIRST + 3}"),
    ("Subs & Safety", f"=H10+{gc_type('Subs & Safety')}", f"={REV}!G14", f"=INPUTS!C{CG_FIRST + 4}"),
    ("Overhead (admin labour + general)", f"=E10+{gc_type('Admin & Overhead')}", f"={REV}!G15",
     f"=INPUTS!C{CG_FIRST + 5}+INPUTS!C{CG_FIRST + 6}"),
]
for i, (lab, f_bid, f_kn, f_typ) in enumerate(mix):
    r_ = 16 + i
    put(bs, f"B{r_}", lab, border=BOX)
    bs.merge_cells(f"B{r_}:C{r_}")
    put(bs, f"D{r_}", f_bid, fmt=CUR, border=BOX)
    put(bs, f"E{r_}", f"=IFERROR(D{r_}/$D$22,0)", fmt=PCT, border=BOX)
    put(bs, f"F{r_}", f_kn, F_LINK, fmt=PCT, border=BOX)
    put(bs, f"G{r_}", f_typ, F_LINK, fmt=PCT, border=BOX)
    put(bs, f"H{r_}", f"=E{r_}-G{r_}", fmt=PCT, border=BOX)
    put(bs, f"I{r_}", f'=IF($D$22=0,"",IF(ABS(H{r_})>0.03,"REVIEW","OK"))', bold=True, border=BOX,
        align=Alignment(horizontal="center"))
put(bs, "B22", "TOTAL COST", F_BOLD)
put(bs, "D22", "=SUM(D16:D21)", fmt=CUR, bold=True, border=TOPLINE)
put(bs, "E22", "=SUM(E16:E21)", fmt=PCT, bold=True, border=TOPLINE)
put(bs, "F22", "=SUM(F16:F21)", fmt=PCT, bold=True, border=TOPLINE)
put(bs, "G22", "=SUM(G16:G21)", fmt=PCT, bold=True, border=TOPLINE)
put(bs, "I22", '=IF(ABS(D22-K10)<1,"","CHECK")', bold=True)
put(bs, "B23", "Gross profit (% of contract)")
put(bs, "D23", "=O10", fmt=CUR)
put(bs, "E23", "=P10", fmt=PCT)
bs.conditional_formatting.add("I16:I22", CellIsRule(operator="equal", formula=['"REVIEW"'],
                              fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, color="9C5700", bold=True)))
ok_cf(bs, "I16:I22")
put(bs, "K16", "Flag = REVIEW when a category is more than 3 points off your company typical.", F_NOTE)
put(bs, "K17", "A one-off job can differ for good reasons; the flag asks you to check, not to change.", F_NOTE)
put(bs, "K18", "Knowify job % = projected job cost plus overhead on projected hours (KNOWIFY REVIEW).", F_NOTE)

put(bs, "B26", "CHECKS (all must read OK before the bid goes out)", F_SUB)


def valid_rows(sheet):
    def rng(c):
        return f"'{sheet}'!${c}${EST_FIRST}:${c}${EST_LAST}"
    return f"SUMPRODUCT(({rng('Z')}<>0)*(ISNA(MATCH({rng('B')},{DIV_RANGE},0))+ISNA(MATCH({rng('C')},{PH_RANGE},0))))"


checks = [
    ("P&H crew mix = 100% and labour ties to wage breakdown", f"='LABOUR RATES'!C{ph_t + 3}"),
    ("HVAC crew mix = 100% and labour ties to wage breakdown", f"='LABOUR RATES'!C{hv_t + 3}"),
    ("P+H labour $ = LABOUR RATES P&H cost by wage breakdown", f"=IF(ABS(D6+D7-'LABOUR RATES'!K{ph_t})<1,\"OK\",\"CHECK\")"),
    ("V+AC labour $ = LABOUR RATES HVAC cost by wage breakdown", f"=IF(ABS(D8+D9-'LABOUR RATES'!K{hv_t})<1,\"OK\",\"CHECK\")"),
    ("Every estimate row has a valid Div and Phase",
     f"=IF({valid_rows('PLUMBING EST')}+{valid_rows('HVAC EST')}=0,\"OK\",\"CHECK\")"),
    ("Subtrades fully allocated to divisions",
     f"=IF(ABS('SUBS & GC'!H{SUB_LAST + 1}-SUM('SUBS & GC'!M{SUB_LAST + 1}:P{SUB_LAST + 1}))<1,\"OK\",\"CHECK\")"),
    ("General conditions fully allocated to divisions",
     f"=IF(ABS('SUBS & GC'!H{GC_LAST + 1}-SUM('SUBS & GC'!M{GC_LAST + 1}:P{GC_LAST + 1}))<1,\"OK\",\"CHECK\")"),
    ("Level weights total 100%", f"=INPUTS!F{LV_TOT}"),
    ("BUDGET total cost = BID SUMMARY total cost (every sub/GC row has a phase)", "=IF(ABS(BUDGET!J11-K10)<1,\"OK\",\"CHECK\")"),
    ("SCHEDULE OF VALUES total = contract", "=IF(ABS('SCHEDULE OF VALUES'!C5-N10)<1,\"OK\",\"CHECK\")"),
    ("Every takeoff line is in the estimate (TAKEOFF CHECK)", f"=IF('{TK}'!I10=0,\"OK\",\"CHECK\")"),
    ("Every takeoff line with a quantity has a unit cost (TAKEOFF CHECK)", f"=IF('{TK}'!I12=0,\"OK\",\"CHECK\")"),
]
for i, (lab, f) in enumerate(checks):
    r_ = 27 + i
    put(bs, f"B{r_}", lab, border=BOX)
    bs.merge_cells(f"B{r_}:H{r_}")
    put(bs, f"I{r_}", f, bold=True, border=BOX, align=Alignment(horizontal="center"))
CHK_LAST = 27 + len(checks) - 1
put(bs, f"B{CHK_LAST + 1}", "OVERALL", F_BOLD)
put(bs, f"I{CHK_LAST + 1}", f'=IF(COUNTIF(I27:I{CHK_LAST},"OK")={len(checks)},"OK","CHECK")', bold=True,
    border=BOX, align=Alignment(horizontal="center"))
ok_cf(bs, f"I27:I{CHK_LAST + 1}")
put(bs, "K2", "Bid status:", F_BOLD, align=Alignment(horizontal="right"))
put(bs, "L2", f"=I{CHK_LAST + 1}", bold=True, align=Alignment(horizontal="center"))
ok_cf(bs, "L2")
bs.freeze_panes = "C6"

# =====================================================================
# BUDGET
# =====================================================================
bu = wb.create_sheet("BUDGET")
setw(bu, {"A": 3, "B": 42, "C": 11, "D": 15, "E": 14, "F": 15, "G": 14, "H": 14, "I": 14, "J": 16, "K": 13, "L": 16})
title(bu, "PROJECT BUDGET (cost, by division and phase)")
put(bu, "B2", '=IF(INPUTS!C4="","",INPUTS!C4)&"   Quote: "&INPUTS!C5', F_LINK)
put(bu, "B4", "PROJECT ROLL-UP", F_SUB)
header_row(bu, 5, ["Division", "Hours", "Labour $", "Overhead $", "Material $ (incl consumables)", "Subtrades $",
                   "Gen. Cond. $", "Contingency $", "TOTAL COST $", "Margin $", "CONTRACT $"], start_col=2)
for i in range(4):
    r_, s = 6 + i, 6 + i
    put(bu, f"B{r_}", f"='BID SUMMARY'!B{s}", F_LINK, border=BOX)
    for col, ref in {"C": f"C{s}", "D": f"D{s}", "E": f"E{s}", "F": f"F{s}+'BID SUMMARY'!G{s}", "G": f"H{s}",
                     "H": f"I{s}", "I": f"J{s}", "J": f"K{s}", "K": f"O{s}", "L": f"N{s}"}.items():
        put(bu, f"{col}{r_}", f"='BID SUMMARY'!{ref}", F_LINK, fmt=NUM if col == "C" else CUR, border=BOX,
            fill=FILL_TOT if col in "JL" else None)
put(bu, "B10", "TOTAL", F_BOLD)
for col in "CDEFGHIJKL":
    put(bu, f"{col}10", f"=SUM({col}6:{col}9)", fmt=NUM if col == "C" else CUR, bold=True, border=TOPLINE)
put(bu, "B11", "Total cost rebuilt from phase detail below (must equal J10)", F_NOTE)
put(bu, "B12", "Material on 03/04 rows shows under 05 Material Supply, or 06 Fixtures & Equipment Supply when the row is flagged Equip = Y.", F_NOTE)

BLOCK = 19
B_START = 14
block_tot_refs = []
BROW = {}
BCONS = {}
for d, (code, nm) in enumerate(DIVS):
    top = B_START + d * BLOCK
    put(bu, f"B{top}", f"=UPPER(INPUTS!B{33 + d})&\" ({code})\"", F_SUB, fill=FILL_SEC)
    for col in "CDEFGHIJK":
        bu[f"{col}{top}"].fill = FILL_SEC
    header_row(bu, top + 1, ["Phase", "Hours", "Labour $", "Overhead $", "Material $", "Subtrades $",
                             "Gen. Cond. $", "", "TOTAL BUDGET $", "% of Division"], start_col=2)
    r_c = top + 2 + len(PHASES)
    r_t = r_c + 2
    div = f'"{code}"'
    for p in range(len(PHASES)):
        r_ = top + 2 + p
        BROW[(d, p)] = r_
        ph = ph_ref(p)
        put(bu, f"B{r_}", f"={ph}", F_LINK, border=BOX)
        put(bu, f"C{r_}", "=" + est_sumifs("W", div, ph), fmt=NUM, border=BOX)
        put(bu, f"D{r_}", "=" + est_sumifs("X", div, ph), fmt=CUR, border=BOX)
        put(bu, f"E{r_}", "=" + est_sumifs("Y", div, ph), fmt=CUR, border=BOX)
        if p in LEVEL_PHASES:
            mat = "=0"
        elif p == MAT_SUPPLY:
            mat = "=" + mat_supply(div)
        elif p == FIX_EQ:
            mat = "=" + mat_equip(div)
        else:
            mat = "=" + est_sumifs("V", div, ph)
        put(bu, f"F{r_}", mat, fmt=CUR, border=BOX)
        dcol = SUB_COLS[d]
        put(bu, f"G{r_}", f"=SUMIFS('SUBS & GC'!${dcol}${SUB_FIRST}:${dcol}${SUB_LAST},'SUBS & GC'!$D${SUB_FIRST}:$D${SUB_LAST},{ph})",
            fmt=CUR, border=BOX)
        put(bu, f"H{r_}", f"=SUMIFS('SUBS & GC'!${dcol}${GC_FIRST}:${dcol}${GC_LAST},'SUBS & GC'!$D${GC_FIRST}:$D${GC_LAST},{ph})",
            fmt=CUR, border=BOX)
        put(bu, f"I{r_}", None, border=BOX)
        put(bu, f"J{r_}", f"=SUM(D{r_}:H{r_})", fmt=CUR, bold=True, border=BOX)
        put(bu, f"K{r_}", f"=IFERROR(J{r_}/$J${r_t},0)", fmt=PCT, border=BOX)
    BCONS[d] = r_c
    put(bu, f"B{r_c}", "Consumables allowance", border=BOX)
    put(bu, f"F{r_c}", f"='BID SUMMARY'!G{6 + d}", F_LINK, fmt=CUR, border=BOX)
    put(bu, f"J{r_c}", f"=F{r_c}", fmt=CUR, bold=True, border=BOX)
    put(bu, f"K{r_c}", f"=IFERROR(J{r_c}/$J${r_t},0)", fmt=PCT, border=BOX)
    put(bu, f"B{r_c + 1}", "Contingency", border=BOX)
    put(bu, f"J{r_c + 1}", f"='BID SUMMARY'!J{6 + d}", F_LINK, fmt=CUR, bold=True, border=BOX)
    put(bu, f"K{r_c + 1}", f"=IFERROR(J{r_c + 1}/$J${r_t},0)", fmt=PCT, border=BOX)
    put(bu, f"B{r_t}", "TOTAL COST BUDGET", F_BOLD)
    for col in "CDEFGH":
        put(bu, f"{col}{r_t}", f"=SUM({col}{top + 2}:{col}{r_c + 1})", fmt=NUM if col == "C" else CUR, bold=True, border=TOPLINE)
    put(bu, f"J{r_t}", f"=SUM(J{top + 2}:J{r_c + 1})", fmt=CUR, bold=True, fill=FILL_TOT, border=TOPLINE)
    put(bu, f"B{r_t + 1}", "Margin (fee)")
    put(bu, f"J{r_t + 1}", f"='BID SUMMARY'!O{6 + d}", F_LINK, fmt=CUR)
    put(bu, f"B{r_t + 2}", "CONTRACT VALUE", F_BOLD)
    put(bu, f"J{r_t + 2}", f"=J{r_t}+J{r_t + 1}", fmt=CUR, bold=True, fill=FILL_TOT, border=TOPLINE)
    block_tot_refs.append(f"J{r_t}")
put(bu, "J11", "=" + "+".join(block_tot_refs), fmt=CUR, bold=True)
put(bu, "K11", '=IF(ABS(J11-J10)<1,"OK","CHECK")', bold=True)
ok_cf(bu, "K11")
bu.freeze_panes = "C6"

# =====================================================================
# SCHEDULE OF VALUES (Knowify-style lines + Knowify cost budget per line)
# =====================================================================
sv = wb.create_sheet("SCHEDULE OF VALUES")
setw(sv, {"A": 11, "B": 52, "C": 15, "D": 9, "E": 15, "F": 9, "G": 15, "H": 13, "I": 2,
          "J": 14, "K": 14, "L": 14, "M": 14, "N": 10, "O": 2, "P": 10})
title(sv, "SCHEDULE OF VALUES")
put(sv, "A2", '=IF(INPUTS!C4="","",INPUTS!C4)&"   "&INPUTS!C11&", "&INPUTS!C12&"   GC: "&INPUTS!C8', F_LINK)
put(sv, "B4", "Contract value (from BID SUMMARY)", F_BOLD)
put(sv, "C4", "='BID SUMMARY'!N10", F_LINK, fmt=CUR, bold=True)
put(sv, "B5", "Total of schedule below", F_BOLD)
put(sv, "D5", '=IF(ABS(C5-C4)<1,"OK","CHECK")', bold=True)
ok_cf(sv, "D5")
put(sv, "B6", "Line names follow your Knowify setup. Columns J:N are the cost budget for the same line, ready to load into Knowify.", F_NOTE)
put(sv, "J7", "KNOWIFY COST BUDGET (no overhead, no margin)", F_BOLD)
header_row(sv, 8, ["Item #", "Job Phase / Contract Item", "Scheduled Value $", "% of Contract", "Completed to Date $",
                   "% Complete", "Balance to Finish $", "Holdback $", "", "Material $", "Labour $",
                   "Subtrades $", "Other $", "Hours", "", "Div factor"])
r_ = 9
div_header_rows = []
for d, (code, nm) in enumerate(DIVS):
    hdr = r_
    div_header_rows.append(hdr)
    put(sv, f"A{hdr}", code, F_BOLD, fill=FILL_SEC)
    put(sv, f"B{hdr}", f"=UPPER(INPUTS!B{33 + d})", F_BOLD, fill=FILL_SEC)
    phase_sum = "+".join(f"BUDGET!$J${BROW[(d, p)]}" for p in range(len(PHASES)))
    put(sv, f"P{hdr}", f"=IFERROR('BID SUMMARY'!N{6 + d}/({phase_sum}),0)", fmt="0.0000", fill=FILL_SEC)
    note(sv, f"P{hdr}", "Contract for this division / its total phase cost. Spreads margin, consumables and contingency evenly.")
    r_ += 1
    first_line = r_
    divname = f"INPUTS!$B${33 + d}"
    for p in range(len(PHASES)):
        br = BROW[(d, p)]
        if p in LEVEL_PHASES:
            word = "Rough-in" if p == 2 else "Finishing"
            for li in range(LV_N):
                lv = LV_FIRST + li
                w = f"INPUTS!$E${lv}"
                put(sv, f"A{r_}", f"{code}-{p + 1:02d}-{li + 1:02d}", border=BOX)
                put(sv, f"B{r_}", f'=IF(INPUTS!$B${lv}="","","{code} - "&INPUTS!$B${lv}&" "&{divname}&" {word} (Labour)")', border=BOX)
                put(sv, f"C{r_}", f"=(BUDGET!$J${br}-BUDGET!$G${br})*{w}*$P${hdr}", fmt=CUR, border=BOX)
                put(sv, f"K{r_}", f"=BUDGET!$D${br}*{w}", fmt=CUR, border=BOX)
                put(sv, f"M{r_}", f"=BUDGET!$H${br}*{w}", fmt=CUR, border=BOX)
                put(sv, f"N{r_}", f"=BUDGET!$C${br}*{w}", fmt=NUM, border=BOX)
                r_ += 1
        else:
            put(sv, f"A{r_}", f"{code}-{p + 1:02d}", border=BOX)
            put(sv, f"B{r_}", f'="{code} - "&MID({ph_ref(p)},4,80)', border=BOX)
            put(sv, f"C{r_}", f"=(BUDGET!$J${br}-BUDGET!$G${br})*$P${hdr}", fmt=CUR, border=BOX)
            mat = f"=BUDGET!$F${br}" + (f"+BUDGET!$F${BCONS[d]}" if p == MAT_SUPPLY else "")
            put(sv, f"J{r_}", mat, fmt=CUR, border=BOX)
            put(sv, f"K{r_}", f"=BUDGET!$D${br}", fmt=CUR, border=BOX)
            put(sv, f"M{r_}", f"=BUDGET!$H${br}", fmt=CUR, border=BOX)
            put(sv, f"N{r_}", f"=BUDGET!$C${br}", fmt=NUM, border=BOX)
            r_ += 1
    for k in range(SUB_FIRST, SUB_LAST + 1):
        put(sv, f"A{r_}", f"{code}-S{k - SUB_FIRST + 1:02d}", border=BOX)
        put(sv, f"B{r_}", f"=IF('SUBS & GC'!$B${k}=\"\",\"\",\"{code} - \"&'SUBS & GC'!$B${k})", border=BOX)
        put(sv, f"C{r_}", f"='SUBS & GC'!${SUB_COLS[d]}${k}*$P${hdr}", fmt=CUR, border=BOX)
        put(sv, f"L{r_}", f"='SUBS & GC'!${SUB_COLS[d]}${k}", fmt=CUR, border=BOX)
        r_ += 1
    last_line = r_ - 1
    for col in "CEGHJKLMN":
        put(sv, f"{col}{hdr}", f"=SUM({col}{first_line}:{col}{last_line})", fmt=NUM if col == "N" else CUR, bold=True, fill=FILL_SEC)
    put(sv, f"D{hdr}", f"=IFERROR(C{hdr}/$C$4,0)", fmt=PCT, bold=True, fill=FILL_SEC)
    for k in range(first_line, last_line + 1):
        put(sv, f"D{k}", f"=IFERROR(C{k}/$C$4,0)", fmt=PCT, border=BOX)
        put(sv, f"E{k}", None, F_IN, fmt=CUR, fill=FILL_IN, border=BOX)
        put(sv, f"F{k}", f"=IFERROR(E{k}/C{k},0)", fmt=PCT, border=BOX)
        put(sv, f"G{k}", f"=C{k}-E{k}", fmt=CUR, border=BOX)
        put(sv, f"H{k}", f"=E{k}*INPUTS!$C$29", fmt=CUR, border=BOX)
        for col in "JKLMN":
            if sv[f"{col}{k}"].value is None:
                put(sv, f"{col}{k}", None, border=BOX)
    grey_zero(sv, f"A{first_line}:N{last_line}", f"$C{first_line}")
    r_ += 1
put(sv, f"B{r_}", "TOTAL CONTRACT", F_BOLD)
for col in "CEGHJKLMN":
    put(sv, f"{col}{r_}", "=" + "+".join(f"{col}{h}" for h in div_header_rows), fmt=NUM if col == "N" else CUR,
        bold=True, fill=FILL_TOT, border=TOPLINE)
put(sv, f"D{r_}", f"=IFERROR(C{r_}/$C$4,0)", fmt=PCT, bold=True, border=TOPLINE)
put(sv, f"F{r_}", f"=IFERROR(E{r_}/C{r_},0)", fmt=PCT, bold=True, border=TOPLINE)
put(sv, "C5", f"=C{r_}", fmt=CUR, bold=True)
put(sv, f"B{r_ + 1}", "Check: Knowify budget + overhead + contingency + margin = contract", F_NOTE)
put(sv, f"J{r_ + 1}", f"=SUM(J{r_}:M{r_})+'BID SUMMARY'!E10+'BID SUMMARY'!J10+'BID SUMMARY'!O10", fmt=CUR)
put(sv, f"K{r_ + 1}", f'=IF(ABS(J{r_ + 1}-C4)<1,"OK","CHECK")', bold=True)
ok_cf(sv, f"K{r_ + 1}")
sv.freeze_panes = "C9"
sv.auto_filter.ref = f"A8:N{r_ - 1}"

# =====================================================================
# KNOWIFY raw paste tabs
# =====================================================================
kp = wb.create_sheet("KNOWIFY PHASES")
kt = wb.create_sheet("KNOWIFY TIME")
kp.column_dimensions["A"].width = 55
kt.column_dimensions["A"].width = 16
kt.column_dimensions["B"].width = 28
KN_JOB = KN_CONTRACT = KN_COMMITTED = None
if KNOWIFY_SRC:
    src = load_workbook(KNOWIFY_SRC, data_only=True)
    sp = src["Actual vs Budget - Phases"]
    for row in sp.iter_rows(min_row=1, max_row=sp.max_row, max_col=22):
        for c in row:
            if c.value is not None:
                kp.cell(row=c.row, column=c.column, value=c.value)
    st = src["Time"]
    n = 0
    for row in st.iter_rows(min_row=1, max_row=st.max_row, max_col=6):
        for c in row:
            v = c.value
            if v is None:
                continue
            if c.column == 1 and c.row >= KT_FIRST:
                n += 1
                v = f"Employee {n:02d}"
            cell = kt.cell(row=c.row, column=c.column, value=v)
            if isinstance(v, datetime.time):
                cell.number_format = "[h]:mm"
    ssum = src["Summary"]
    KN_JOB, KN_CONTRACT, KN_COMMITTED = ssum["A1"].value, ssum["B6"].value, ssum["B14"].value
for ws_ in (kp, kt):
    for row in ws_.iter_rows():
        for c in row:
            c.font = Font(name=FONT, size=10)

# =====================================================================
# KNOWIFY REVIEW
# =====================================================================
kr = wb.create_sheet("KNOWIFY REVIEW")
setw(kr, {"A": 3, "B": 50, "C": 16, "D": 14, "E": 30, "F": 15, "G": 14, "H": 14, "I": 14, "J": 14,
          "AA": 40, "AB": 6, "AC": 30, "AD": 5, "AT": 26, "AU": 10, "AV": 12, "AW": 20, "AZ": 20, "BA": 34,
          "BC": 28, "BD": 22})
title(kr, "KNOWIFY REVIEW: what the last job actually cost",
      "Reads KNOWIFY PHASES and KNOWIFY TIME. Blue cells: type from the export's Summary sheet.")
put(kr, "B4", "Job")
put(kr, "C4", KN_JOB, F_IN, fill=FILL_IN, border=BOX)
kr.merge_cells("C4:F4")
put(kr, "B5", "Contract (Knowify 'Total Amount')")
put(kr, "C5", KN_CONTRACT, F_IN, fmt=CUR, fill=FILL_IN, border=BOX)
put(kr, "B6", "Committed cost (Knowify 'Total Cost To Date (Committed)')")
put(kr, "C6", KN_COMMITTED, F_IN, fmt=CUR, fill=FILL_IN, border=BOX)

P = "'KNOWIFY PHASES'"
for r_ in range(KP_FIRST, KP_LAST + 1):
    a = f"{P}!A{r_}"
    put(kr, f"AA{r_}", f'=IF({a}="","",{a})')
    put(kr, f"AB{r_}", f'=IF(AA{r_}="","",IF(LEFT(AA{r_},4)="P - ","P",IF(LEFT(AA{r_},4)="H - ","H",'
                       f'IF(LEFT(AA{r_},5)="AC - ","AC",IF(OR(LEFT(AA{r_},4)="V - ",LEFT(AA{r_},7)="HVAC - "),"V","Other")))))')
    put(kr, f"AC{r_}", f'=IF(AA{r_}="","",IFERROR(LOOKUP(2,1/(($AZ$5:$AZ$60<>"")*ISNUMBER(SEARCH($AZ$5:$AZ$60,AA{r_}))),$BA$5:$BA$60),IF(AB{r_}="Other","Excluded","Unmapped")))')
    put(kr, f"AD{r_}", f'=IF(AA{r_}="","",IF(AND(AB{r_}<>"Other",OR(AND({P}!B{r_}="Closed",N({P}!D{r_})>=0.5),N({P}!D{r_})>=1)),1,0))')
    for col, srccol in zip(["AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN"],
                           ["L", "K", "I", "H", "O", "N", "R", "Q", "F", "E"]):
        put(kr, f"{col}{r_}", f"=N({P}!{srccol}{r_})")
    for col, (b_, a_) in zip(["AO", "AP", "AQ", "AR"], [("AE", "AF"), ("AG", "AH"), ("AI", "AJ"), ("AK", "AL")]):
        put(kr, f"{col}{r_}", f"=MAX({b_}{r_},{a_}{r_})")
header_row(kr, 3, ["Knowify Phase", "Div", "Mapped Phase", "Use", "Lab Bud", "Lab Act", "Mat Bud", "Mat Act",
                   "Sub Bud", "Sub Act", "Eq Bud", "Eq Act", "Tot Bud", "Tot Act", "Proj Lab", "Proj Mat",
                   "Proj Sub", "Proj Eq"], start_col=27)
put(kr, "AA2", "HELPER COLUMNS (one row per KNOWIFY PHASES row; do not edit)", F_BOLD)
T = "'KNOWIFY TIME'"
for r_ in range(KT_FIRST, KT_LAST + 1):
    d_ = f"{T}!D{r_}"
    put(kr, f"AT{r_}", f'=IF({T}!B{r_}="","",{T}!B{r_})')
    put(kr, f"AU{r_}", f'=IF(AT{r_}="",0,IF(ISNUMBER({d_}),{d_}*24,IFERROR(VALUE(LEFT({d_},FIND(":",{d_})-1))+VALUE(MID({d_},FIND(":",{d_})+1,2))/60,0)))')
    put(kr, f"AV{r_}", f"=N({T}!E{r_})")
    put(kr, f"AW{r_}", f'=IF(AT{r_}="","",IFERROR(INDEX($BD$5:$BD$60,MATCH(AT{r_},$BC$5:$BC$60,0)),"Unmapped"))')
header_row(kr, 2, ["Role", "Hours", "Cost", "Crew Class"], start_col=46)

put(kr, "AZ3", "PHASE KEYWORD MAP (last match wins)", F_BOLD)
header_row(kr, 4, ["Keyword in Knowify phase", "Maps to Phase"], start_col=52)
KW = [("Water Connection", 0), ("Shop Drawing", 0), ("Permit", 0), ("Mobilization", 0),
      ("Foundation", 1), ("Underground", 1), ("U/G", 1),
      ("Parkade", 2), ("Rough-in", 2), ("Roof", 2), ("Insulation", 2), ("Sprinkler", 2),
      ("Finishing", 3), ("Material Supply", 4),
      ("Fixtures", 5), ("Equipment Supply", 5), ("Mech Room", 5), ("Controls", 5),
      ("Testing", 6), ("Commissioning", 6), ("TAB ", 6),
      ("Completion", 7), ("Final Doc", 7),
      ("Office", 8), ("Safety", 8), ("Supervision", 8)]
for i in range(56):
    r_ = 5 + i
    kw = KW[i] if i < len(KW) else None
    put(kr, f"AZ{r_}", kw[0] if kw else None, F_IN, fill=FILL_IN, border=BOX)
    put(kr, f"BA{r_}", PHASES[kw[1]] if kw else None, F_IN, fill=FILL_IN, border=BOX)
dvk = DataValidation(type="list", formula1=PH_RANGE, allow_blank=True)
kr.add_data_validation(dvk)
dvk.add("BA5:BA60")
put(kr, "BC3", "ROLE MAP (Knowify role -> crew class)", F_BOLD)
header_row(kr, 4, ["Knowify Role (exact)", "Crew Class"], start_col=55)
EX = 5
ROLES = [("Apprentice Lvl 1", 0), ("Apprentice Lvl 2", 0), ("Apprentice Lvl 3", 1), ("Apprentice Lvl 4", 1),
         ("Apprentice Lvl 5", 1), ("Apprentice Lvl 6", 1), ("Apprentice Level 6", 1), ("Apprentice Lvl 7", 2),
         ("Apprentice Lvl 8", 2), ("Journeyman", 3), ("Journeyman Lvl 1", 3), ("Journeyman Lvl 2", 3),
         ("Foreman Lvl 1", 4), ("Foreman Lvl 2", 4), ("HVAC Field Supervisor", 4), ("Project Manager", EX),
         ("Health & Safety Coordinator", EX), ("Divisional Manager", EX), ("Admin", EX)]
ALLCLS = CLASSES + [CLASS_EXCL]
for i in range(56):
    r_ = 5 + i
    ro = ROLES[i] if i < len(ROLES) else None
    put(kr, f"BC{r_}", ro[0] if ro else None, F_IN, fill=FILL_IN, border=BOX)
    put(kr, f"BD{r_}", ALLCLS[ro[1]] if ro else None, F_IN, fill=FILL_IN, border=BOX)
dvc = DataValidation(type="list", formula1='"' + ",".join(ALLCLS) + '"', allow_blank=True)
kr.add_data_validation(dvc)
dvc.add("BD5:BD60")

# --- headline ---
put(kr, "B8", "HEADLINE", F_SUB)
header_row(kr, 9, ["Measure", "Amount", "Margin %"], start_col=2)
TOT_PROJ = f"SUM({rv('AO')})+SUM({rv('AP')})+SUM({rv('AQ')})+SUM({rv('AR')})"
for row, lab, f, m in [
    (10, "Knowify cost budget", f"=SUM({rv('AM')})", "=IFERROR(1-C10/$C$5,0)"),
    (11, "Cost to date (actual)", f"=SUM({rv('AN')})", None),
    (12, "Projected cost (larger of budget or actual, per phase)", f"={TOT_PROJ}", "=IFERROR(1-C12/$C$5,0)"),
    (13, "Committed cost (typed above)", "=C6", "=IFERROR(1-C13/$C$5,0)"),
    (14, "Labour hours to date (KNOWIFY TIME)", f"=SUM({REV}!$AU${KT_FIRST}:$AU${KT_LAST})", None),
    (15, "Company overhead on those hours (not in Knowify)", f"=C14*'LABOUR RATES'!$C${ph_t + 1}", None),
    (16, "Margin at projected cost, after overhead", None, "=IFERROR(1-(C12+C15)/$C$5,0)"),
    (17, "Margin at committed cost, after overhead", None, "=IFERROR(1-(C13+C15)/$C$5,0)"),
]:
    put(kr, f"B{row}", lab, F_BOLD if row >= 16 else F_CALC, border=BOX)
    put(kr, f"C{row}", f, fmt=NUM if row == 14 else CUR, border=BOX)
    put(kr, f"D{row}", m, fmt=PCT, border=BOX, bold=row >= 16, fill=FILL_TOT if row >= 16 else None)
put(kr, "B18", "Overhead counts hours so far only; remaining hours add more.", F_NOTE)

put(kr, "E9", "Company Category (projected)", F_HDR, fill=FILL_HDR, border=BOX)
put(kr, "F9", "Amount", F_HDR, fill=FILL_HDR, border=BOX)
put(kr, "G9", "% of All Costs", F_HDR, fill=FILL_HDR, border=BOX)
FIXKN = f"SUMIFS({rv('AP')},{rv('AC')},{ph_ref(FIX_EQ)})"
for row, lab, f in [
    (10, "Materials (excl. fixtures & equipment)", f"=SUM({rv('AP')})-{FIXKN}"),
    (11, "Equipment (fixtures & equipment supply)", f"={FIXKN}"),
    (12, "Other Job Costs (Knowify equipment/tools)", f"=SUM({rv('AR')})"),
    (13, "Labour (incl. office PM & safety time)", f"=SUM({rv('AO')})"),
    (14, "Subs", f"=SUM({rv('AQ')})"),
    (15, "Overhead on projected hours (not in Knowify)", f"=IFERROR(F13/{KN_BLENDED_ALL},0)*'LABOUR RATES'!$C${ph_t + 1}"),
]:
    put(kr, f"E{row}", lab, border=BOX)
    put(kr, f"F{row}", f, fmt=CUR, border=BOX)
    put(kr, f"G{row}", f"=IFERROR(F{row}/$F$16,0)", fmt=PCT, border=BOX)
put(kr, "E16", "TOTAL", F_BOLD)
put(kr, "F16", "=SUM(F10:F15)", fmt=CUR, bold=True, border=TOPLINE)
put(kr, "G16", "=SUM(G10:G15)", fmt=PCT, bold=True, border=TOPLINE)

put(kr, "B20", "Knowify phases all mapped", border=BOX)
put(kr, "C20", f'=IF(COUNTIF({rv("AC")},"Unmapped")=0,"OK","CHECK")', bold=True, border=BOX)
put(kr, "B21", "Knowify roles all mapped", border=BOX)
put(kr, "C21", f'=IF(COUNTIF({REV}!$AW${KT_FIRST}:$AW${KT_LAST},"Unmapped")=0,"OK","CHECK")', bold=True, border=BOX)
ok_cf(kr, "C20:C21")

put(kr, f"B{PERF_FIRST - 2}", "HOW EACH PHASE RAN vs BUDGET (closed phases that used at least half their budget + phases already at/over budget; change orders excluded)", F_SUB)
header_row(kr, PERF_FIRST - 1, ["Div / Phase", "# Phases", "Labour Budget", "Labour Actual", "Labour Act/Bud",
                                "Material Budget", "Material Actual", "Material Act/Bud", "Labour Hrs (est.)"], start_col=2)
r_ = PERF_FIRST
for code, nm in DIVS:
    for p in range(len(PHASES)):
        crit = f"{rv('AB')},\"{code}\",{rv('AC')},{ph_ref(p)},{rv('AD')},1"
        put(kr, f"B{r_}", f'="{code}  "&{ph_ref(p)}', border=BOX)
        put(kr, f"C{r_}", f"=COUNTIFS({crit})", fmt=NUM, border=BOX)
        put(kr, f"D{r_}", f"=SUMIFS({rv('AE')},{crit})", fmt=CUR, border=BOX)
        put(kr, f"E{r_}", f"=SUMIFS({rv('AF')},{crit})", fmt=CUR, border=BOX)
        put(kr, f"F{r_}", f'=IFERROR(E{r_}/D{r_},"")', fmt=FAC, border=BOX)
        put(kr, f"G{r_}", f"=SUMIFS({rv('AG')},{crit})", fmt=CUR, border=BOX)
        put(kr, f"H{r_}", f"=SUMIFS({rv('AH')},{crit})", fmt=CUR, border=BOX)
        put(kr, f"I{r_}", f'=IFERROR(H{r_}/G{r_},"")', fmt=FAC, border=BOX)
        put(kr, f"J{r_}", f"=IFERROR(E{r_}/{KN_BLENDED},0)", fmt=NUM, border=BOX)
        r_ += 1
for col in "FI":
    kr.conditional_formatting.add(f"{col}{PERF_FIRST}:{col}{PERF_LAST}", FormulaRule(
        formula=[f"AND(ISNUMBER({col}{PERF_FIRST}),{col}{PERF_FIRST}>1.1)"], fill=PatternFill("solid", fgColor="FFC7CE")))
grey_zero(kr, f"B{PERF_FIRST}:J{PERF_LAST}", f"$C{PERF_FIRST}")

put(kr, f"B{CREW_HDR}", "ACTUAL CREW (KNOWIFY TIME, all trades)", F_SUB)
header_row(kr, CREW_HDR + 1, ["Crew Class", "Hours", "Cost", "Actual $/hr", "Crew Mix", "Template Mix (P&H)"], start_col=2)
for i, cls in enumerate(CLASSES + [CLASS_EXCL, "Unmapped"]):
    rr = CREW_FIRST + i
    put(kr, f"B{rr}", cls, border=BOX)
    put(kr, f"C{rr}", f'=SUMIFS({REV}!$AU${KT_FIRST}:$AU${KT_LAST},{REV}!$AW${KT_FIRST}:$AW${KT_LAST},B{rr})', fmt=NUM, border=BOX)
    put(kr, f"D{rr}", f'=SUMIFS({REV}!$AV${KT_FIRST}:$AV${KT_LAST},{REV}!$AW${KT_FIRST}:$AW${KT_LAST},B{rr})', fmt=CUR, border=BOX)
    put(kr, f"E{rr}", f'=IFERROR(D{rr}/C{rr},0)', fmt=CUR2, border=BOX)
    if i < 5:
        put(kr, f"F{rr}", f'=IFERROR(C{rr}/$C${CREW_SUB},0)', fmt=PCT, border=BOX)
        put(kr, f"G{rr}", f"='LABOUR RATES'!D{6 + i}", F_LINK, fmt=PCT, border=BOX)
put(kr, f"B{CREW_SUB}", "FIELD CREW (used for rates and mix)", F_BOLD)
put(kr, f"B{CREW_ALL}", "ALL ROLES", F_BOLD)
for rr, rng in ((CREW_SUB, f"{CREW_FIRST}:{CREW_FIRST + 4}"), (CREW_ALL, f"{CREW_FIRST}:{CREW_FIRST + 6}")):
    a_, b_ = rng.split(":")
    put(kr, f"C{rr}", f"=SUM(C{a_}:C{b_})", fmt=NUM, bold=True, border=TOPLINE)
    put(kr, f"D{rr}", f"=SUM(D{a_}:D{b_})", fmt=CUR, bold=True, border=TOPLINE)
    put(kr, f"E{rr}", f"=IFERROR(D{rr}/C{rr},0)", fmt=CUR2, bold=True, fill=FILL_TOT, border=TOPLINE)
put(kr, f"F{CREW_SUB}", f"=SUM(F{CREW_FIRST}:F{CREW_FIRST + 4})", fmt=PCT, bold=True, border=TOPLINE)
put(kr, f"B{CREW_ALL + 1}", "Knowify cost includes burden (confirmed). Office PM, admin and safety time is priced on SUBS & GC, so it is left out of the crew mix.", F_NOTE)
kr.freeze_panes = "A4"

# =====================================================================
# CONTINGENCY
# =====================================================================
ct = wb.create_sheet("CONTINGENCY")
setw(ct, {"A": 3, "B": 44, "C": 18, "D": 14, "E": 15, "F": 15, "G": 17, "H": 14})
title(ct, "CONTINGENCY (added to cost before the grand total)",
      "General allowance on budgeted cost, plus any specific risks you want priced. Margin is applied on top.")
put(ct, "B4", "General allowance (% of budgeted cost)", F_BOLD)
put(ct, "C4", 0.02, F_IN, fmt=PCT, fill=FILL_KEY, border=BOX)
put(ct, "D4", "Standard 2% (estimator, Oct 2026). Override per division in column D.", F_NOTE)
header_row(ct, 7, ["Division", "Budgeted Cost (before contingency)", "Allowance %", "Allowance $",
                   "Specific Risks $", "TOTAL CONTINGENCY $", "% of Budgeted Cost"], start_col=2)
for i, (code, nm) in enumerate(DIVS):
    r_ = 8 + i
    put(ct, f"B{r_}", f"=INPUTS!B{33 + i}", F_LINK, border=BOX)
    put(ct, f"C{r_}", f"=SUM('BID SUMMARY'!D{6 + i}:I{6 + i})", F_LINK, fmt=CUR, border=BOX)
    put(ct, f"D{r_}", "=$C$4", F_IN, fmt=PCT, fill=FILL_IN, border=BOX)
    put(ct, f"E{r_}", f"=C{r_}*D{r_}", fmt=CUR, border=BOX)
    put(ct, f"F{r_}", f'=SUMIFS($E$16:$E$40,$D$16:$D$40,INPUTS!C{33 + i})', fmt=CUR, border=BOX)
    put(ct, f"G{r_}", f"=E{r_}+F{r_}", fmt=CUR, bold=True, fill=FILL_TOT, border=BOX)
    put(ct, f"H{r_}", f"=IFERROR(G{r_}/C{r_},0)", fmt=PCT, border=BOX)
put(ct, "B12", "TOTAL", F_BOLD)
for col in "CEFG":
    put(ct, f"{col}12", f"=SUM({col}8:{col}11)", fmt=CUR, bold=True, border=TOPLINE)
put(ct, "H12", "=IFERROR(G12/C12,0)", fmt=PCT, bold=True, border=TOPLINE)
put(ct, "B14", "SPECIFIC RISKS (optional)", F_SUB)
header_row(ct, 15, ["#", "Risk / Item", "", "Div", "Amount $", "Notes"], start_col=1)
ct.merge_cells("B15:C15")
for k in range(16, 41):
    put(ct, f"A{k}", k - 15, border=BOX, align=Alignment(horizontal="center"))
    put(ct, f"B{k}", None, F_IN, fill=FILL_IN, border=BOX)
    ct.merge_cells(f"B{k}:C{k}")
    put(ct, f"D{k}", None, F_IN, fill=FILL_IN, border=BOX)
    put(ct, f"E{k}", None, F_IN, fmt=CUR, fill=FILL_IN, border=BOX)
    put(ct, f"F{k}", None, F_IN, fill=FILL_IN, border=BOX)
ct.merge_cells("F15:H15")
for k in range(16, 41):
    ct.merge_cells(f"F{k}:H{k}")
dv_c = DataValidation(type="list", formula1=DIV_RANGE, allow_blank=True)
ct.add_data_validation(dv_c)
dv_c.add("D16:D40")
put(ct, "B42", "Examples: schedule compression, winter heat, unclear scope, long-lead price escalation.", F_NOTE)
ct.freeze_panes = "A8"

# =====================================================================
# RATE VARIANCE
# =====================================================================
va = wb.create_sheet("RATE VARIANCE")
setw(va, {"A": 3, "B": 30, "C": 13, "D": 13, "E": 13, "F": 11, "G": 13, "H": 13, "I": 13, "J": 13, "K": 22})
title(va, "RATE VARIANCE: template vs Knowify actuals vs last snapshot",
      "Shows how far actual labour cost has moved. Update the snapshot after you review a new Knowify paste.")
put(va, "B4", "Snapshot source", F_BOLD)
put(va, "C4", "0-MARC25102 - 405 Marie Place (Ashton), Knowify export 2026-10-06", F_IN, fill=FILL_IN, border=BOX)
va.merge_cells("C4:H4")
header_row(va, 6, ["Classification", "Template $/hr", "Knowify Now $/hr", "Variance $/hr", "Variance %",
                   "Snapshot $/hr", "Change Since Snapshot", "Knowify Mix Now", "Snapshot Mix", "Flag"], start_col=2)
SNAP = [(40.25, 0.3795), (47.10, 0.2599), (51.21, 0.0893), (61.55, 0.1250), (65.80, 0.1463)]
for i, cls in enumerate(CLASSES):
    r_ = 7 + i
    put(va, f"B{r_}", f"='LABOUR RATES'!B{6 + i}", F_LINK, border=BOX)
    put(va, f"C{r_}", f"='LABOUR RATES'!C{6 + i}", F_LINK, fmt=CUR2, border=BOX)
    put(va, f"D{r_}", f"='LABOUR RATES'!E{6 + i}", F_LINK, fmt=CUR2, border=BOX)
    put(va, f"E{r_}", f"=D{r_}-C{r_}", fmt=CUR2, border=BOX)
    put(va, f"F{r_}", f"=IFERROR(D{r_}/C{r_}-1,0)", fmt=PCT, border=BOX)
    put(va, f"G{r_}", SNAP[i][0], F_IN, fmt=CUR2, fill=FILL_IN, border=BOX)
    put(va, f"H{r_}", f"=IFERROR(D{r_}/G{r_}-1,0)", fmt=PCT, border=BOX)
    put(va, f"I{r_}", f"='LABOUR RATES'!F{6 + i}", F_LINK, fmt=PCT, border=BOX)
    put(va, f"J{r_}", SNAP[i][1], F_IN, fmt=PCT, fill=FILL_IN, border=BOX)
    put(va, f"K{r_}", f'=IF(ABS(H{r_})>0.05,"Moved >5% since snapshot","")', border=BOX)
put(va, "B12", "Weighted (crew mix)", F_BOLD)
put(va, "C12", "='LABOUR RATES'!C11", F_LINK, fmt=CUR2, bold=True, border=TOPLINE)
put(va, "D12", "='LABOUR RATES'!E11", F_LINK, fmt=CUR2, bold=True, border=TOPLINE)
put(va, "E12", "=D12-C12", fmt=CUR2, bold=True, border=TOPLINE)
put(va, "F12", "=IFERROR(D12/C12-1,0)", fmt=PCT, bold=True, border=TOPLINE)
put(va, "G12", "=SUMPRODUCT(G7:G11,J7:J11)", fmt=CUR2, bold=True, border=TOPLINE)
put(va, "H12", "=IFERROR(D12/G12-1,0)", fmt=PCT, bold=True, border=TOPLINE)
put(va, "I12", "=SUM(I7:I11)", fmt=PCT, bold=True, border=TOPLINE)
put(va, "J12", "=SUM(J7:J11)", fmt=PCT, bold=True, border=TOPLINE)
put(va, "K12", '=IF(ABS(H12)>0.05,"Moved >5% since snapshot","")', bold=True, border=TOPLINE)
va.conditional_formatting.add("K7:K12", FormulaRule(formula=['K7<>""'], fill=PatternFill("solid", fgColor="FFEB9C")))

put(va, "B15", "EFFECT ON THIS BID", F_SUB)
header_row(va, 16, ["Crew", "Rate Source", "Hours", "At Template $", "At Knowify $", "Difference $",
                    "At Snapshot $", "Rate Used $/hr"], start_col=2)
for i, (lbl, t) in enumerate((("P&H crew", ph_t), ("HVAC crew", hv_t))):
    r_ = 17 + i
    top = t - 7
    put(va, f"B{r_}", lbl, border=BOX)
    put(va, f"C{r_}", f"='LABOUR RATES'!G{top}", F_LINK, border=BOX)
    put(va, f"D{r_}", f"='LABOUR RATES'!C{t + 4}", F_LINK, fmt=NUM, border=BOX)
    put(va, f"E{r_}", f"=D{r_}*'LABOUR RATES'!C{t}", fmt=CUR, border=BOX)
    put(va, f"F{r_}", f"=D{r_}*'LABOUR RATES'!E{t}", fmt=CUR, border=BOX)
    put(va, f"G{r_}", f"=F{r_}-E{r_}", fmt=CUR, border=BOX)
    put(va, f"H{r_}", f"=D{r_}*$G$12", fmt=CUR, border=BOX)
    put(va, f"I{r_}", f"='LABOUR RATES'!I{t}", F_LINK, fmt=CUR2, border=BOX)
put(va, "B19", "TOTAL", F_BOLD)
for col in "DEFGH":
    put(va, f"{col}19", f"=SUM({col}17:{col}18)", fmt=NUM if col == "D" else CUR, bold=True, border=TOPLINE)

put(va, "B22", "OVERHEAD (reference only; bids use the manual $/hr on LABOUR RATES)", F_SUB)
header_row(va, 23, ["Crew", "Manual $/hr", "Implied $/hr", "Variance $/hr", "Hours", "Variance $ on Bid"], start_col=2)
for i, (lbl, t) in enumerate((("P&H crew", ph_t), ("HVAC crew", hv_t))):
    r_ = 24 + i
    put(va, f"B{r_}", lbl, border=BOX)
    put(va, f"C{r_}", f"='LABOUR RATES'!C{t + 1}", F_LINK, fmt=CUR2, border=BOX)
    put(va, f"D{r_}", f"={OH_RATIO}*'LABOUR RATES'!I{t}", fmt=CUR2, border=BOX)
    put(va, f"E{r_}", f"=C{r_}-D{r_}", fmt=CUR2, border=BOX)
    put(va, f"F{r_}", f"='LABOUR RATES'!C{t + 4}", F_LINK, fmt=NUM, border=BOX)
    put(va, f"G{r_}", f"=E{r_}*F{r_}", fmt=CUR, border=BOX)
put(va, "B27", "Implied $/hr = (Admin labour + General overhead) / Labour from the company cost breakdown on INPUTS, x the rate used.", F_NOTE)
put(va, "B28", "A negative variance means the manual rate recovers less overhead than the company books suggest.", F_NOTE)
put(va, "B30", "TO UPDATE THE SNAPSHOT: after pasting a new Knowify job and reviewing it, type the 'Knowify Now' values into the Snapshot columns.", F_NOTE)
va.freeze_panes = "A7"

# =====================================================================
# AUDIT NOTES
# =====================================================================
au = wb.create_sheet("AUDIT NOTES")
setw(au, {"A": 4, "B": 9, "C": 30, "D": 52, "E": 46, "F": 46})
title(au, "AUDIT NOTES: formula review of 'Template-Empty - WORK IN PROGRESS.xlsm'",
      "Values quoted are what the original file currently calculates (20,800 plumbing hours entered, no HVAC).")
header_row(au, 4, ["#", "Severity", "Where (original)", "What happens now", "Why it matters / $ impact", "How the new workbook handles it"])
FIND = [
    ("HIGH", "PRJ SUMMARY D12 via P&H SUMMARY L2 (rate = PRJ INFO D42)",
     "Bid labour = hours x $70.20, a simple average of the 5 rates (D42 = SUM(D36:D40)/5). Crew mix % is ignored.",
     "Crew-mix weighted rate is $64.75. On 20,800 hrs the bid carries $1,460,160 vs $1,346,696 from the crew mix (+$113,464).",
     "Labour $ = hours x SUMPRODUCT(rate, mix) on LABOUR RATES. One rate, used everywhere."),
    ("HIGH", "PRJ INFO H35 / H44 ($22/hr overhead)",
     "Overhead per hour is entered but never added to cost. The only formulas that touch H35 are PRJ SUMMARY C57:C58, by mistake.",
     "20,800 hrs x $22 = $457,600 of overhead not in the cost. Net of item 1, labour + OH is under-costed by $344,136 before markup.",
     "Overhead $ is its own column on every estimate line and its own column on BID SUMMARY."),
    ("HIGH", "P&H SUMMARY L11:L68 (rate = PRJ INFO H36)",
     "Every plumbing line item prices labour at $53/hr (Apprentice Lvl 1-2 only), while the summary row uses $70.20.",
     "Line items total $1,102,400; the bid total uses $1,460,160. BUDGET PLAN and the KNOWIFY export pull the $53 lines, so the field budget is $357,760 lower than what was bid. S.O.V column Y 'ADJST' plugs the gap ($411,440).",
     "Line labour and summary labour use the same rate, so BUDGET, BID SUMMARY and SOV tie with no plug."),
    ("HIGH", "HVAC SUMMARY H29 and H54 (= PRJ INFO H53)",
     "PRJ INFO H53 is an empty cell, so every HVAC line item shows $0 labour.",
     "Any HVAC budget built from line items would have zero labour dollars.",
     "HVAC EST uses the HVAC crew rate from LABOUR RATES, same formula as PLUMBING EST."),
    ("HIGH", "PRJ INFO J46:J49",
     "HVAC 'Cost by wage breakdown' multiplies I45 (Apprentice 1-2 hours) by each rate instead of its own row's hours.",
     "HVAC wage breakdown is wrong as soon as HVAC hours exist.",
     "Each row: Hours x Rate on its own row, plus a tie-out check."),
    ("MED", "PRJ INFO H45:H49",
     "HVAC rate column links to the plumbing rates D36:D40 (and is labelled 'P&H'), not the HVAC table D45:D49.",
     "Changing an HVAC rate in D45:D49 has no effect on HVAC labels; easy to misread.",
     "Separate HVAC crew table with its own rates and mix."),
    ("MED", "PRJ INFO I37",
     "Apprentice 3-6 hours use PRJ SUMMARY C11 (material-row hours only); the other four rows use C10 (all plumbing hours).",
     "Once equipment or fixture hours exist, the crew mix no longer totals 100% of hours.",
     "All rows use the same total-hours cell."),
    ("HIGH", "PRJ SUMMARY C57, C58",
     "Supervision months for Ventilation and A/C point to PRJ INFO H35 (the $22 overhead) instead of H43 (HVAC duration).",
     "Supervision would be charged for 22 months instead of 24.",
     "Supervision qty links to INPUTS HVAC duration."),
    ("MED", "PRJ SUMMARY row 59 (Supervision - Office)",
     "$7,500/month x 24 months = $180,000 (+15% = $207,000) is allocated 100% to Plumbing, using the HVAC duration.",
     "Plumbing looks $207,000 more expensive than it is; HVAC looks cheaper. Confirmed this is a real job cost (dedicated PM), separate from the per-hour overhead.",
     "Kept as 'Office Project Manager'. Auto-splits by direct cost across divisions (or set your own %)."),
    ("HIGH", "PRJ SUMMARY M55:T79 (\"50-50 Div's\" split)",
     "GC split formulas divide by C4+C5 with no IFERROR. With no HVAC entered they return a divide-by-zero (DIV/0) error.",
     "PRJ SUMMARY F6, G6, J6, W6 and J7 all show DIV/0 errors, so the grand total breaks on any plumbing-only job.",
     "Allocation uses % inputs or an IFERROR-guarded direct-cost share. No divide-by-zero."),
    ("MED", "S.O.V B1, AC1, M2:M6, rows 4-5; KNOWIFY F1",
     "Broken references (REF errors from deleted cells) in the S.O.V; the Knowify 'ALL GOOD' indicator therefore shows a REF error.",
     "The SOV check light can never go green, so it cannot be trusted.",
     "New SCHEDULE OF VALUES built from BUDGET with Knowify line names and a contract tie-out check."),
    ("MED", "PRJ SUMMARY V2 / W2",
     "Contract = ROUNDUP(total, -2) + V2, where V2 is a hardcoded $1,200 'ROLL-UP'.",
     "An unexplained $1,200 is added to the plumbing contract.",
     "Removed. Rounding is a visible input on INPUTS."),
    ("MED", "PRJ SUMMARY E11:E26 (15%)",
     "Markup on cost is applied, not margin. 15% markup = 13.0% margin.",
     "If the company target is 15% margin, every bid is about 2 points short.",
     "Margin entered per division; equivalent markup shown beside it."),
    ("LOW", "P&H SUMMARY M45; HVAC SUMMARY J29, J54",
     "Total-hours boxes leave out Finishing Misc hours (J37, J21, J46).",
     "The cost box on those tabs under-reports if finishing misc hours are used.",
     "Single hours column summed across every row."),
    ("LOW", "P&H SUMMARY row 26 and HVAC SUMMARY rows 13 / 37",
     "Condensate drains appear in both Plumbing and HVAC.",
     "Risk of pricing the same scope twice.",
     "Kept on both tabs, labelled 'if by Plumbing' / 'if by HVAC'. Use only one."),
    ("OK", "PRJ INFO E40 (Foreman & PM 14%) and PRJ SUMMARY rows 55-58",
     "Foreman & PM time is in the crew mix and supervision is also a GC line.",
     "Reviewed with estimator: these are separate costs, not a double count.",
     "Both kept: crew mix on LABOUR RATES, supervision on SUBS & GC."),
    ("LOW", "BUDGET PLAN J1",
     "Supervision/PM wage hardcoded at $50/hr, a fourth labour rate in the file.",
     "Budget supervision hours do not match the bid.",
     "One rate source: LABOUR RATES."),
    ("LOW", "PRJ SUMMARY X1 / Z1",
     "'PRJ DAYS' = hours / 8 is really man-days; 'MAN DAYS' = man-days / crew is really duration in days.",
     "Labels swapped; easy to misread schedule.",
     "Not carried over. Hours per unit shown on BID SUMMARY."),
]
for i, (sev, where, what, why, fix) in enumerate(FIND):
    r_ = 5 + i
    for j, v in enumerate([i + 1, sev, where, what, why, fix]):
        put(au, f"{get_column_letter(1 + j)}{r_}", v, border=BOX,
            align=Alignment(wrap_text=True, vertical="top", horizontal="center" if j < 2 else "left"))
    au[f"B{r_}"].font = Font(name=FONT, size=10, bold=True,
                             color={"HIGH": "C00000", "MED": "C65911", "LOW": "7F7F7F", "OK": "006100"}[sev])


def text_list(top, heading, items, marker=None):
    put(au, f"B{top}", heading, F_SUB)
    for i, q in enumerate(items):
        row = top + 1 + i
        if marker:
            put(au, f"B{row}", marker, Font(name=FONT, size=10, bold=True, color="006100"),
                align=Alignment(horizontal="center", vertical="top"))
        else:
            put(au, f"B{row}", i + 1, align=Alignment(horizontal="center", vertical="top"))
        put(au, f"C{row}", q, align=Alignment(wrap_text=True, vertical="top"))
        au.merge_cells(f"C{row}:F{row}")
        au.row_dimensions[row].height = 30
    return top + len(items) + 2


nxt = text_list(5 + len(FIND) + 2, "OPEN QUESTIONS FOR YOU", [
    "Office PM: the $32.71/hr comes from Admin Labour + General Overhead in your books. If the dedicated office PM's salary sits in "
    "Admin Labour, the $7,500/month GC line partly double counts with the overhead rate. Is the PM booked to job cost or to admin?",
    "438 West Pender: unit costs and hours per unit are blank. Paste or type your prices; TAKEOFF CHECK shows what is still unpriced.",
])
text_list(nxt, "ANSWERED (decisions built into this workbook)", [
    "Overhead is a manual entry per bid, now $32.71/hr (was $22): the rate implied by the company cost breakdown at Knowify labour rates. RATE VARIANCE tracks it.",
    "Subs margin standard 12%.",
    "Takeoff import built for the provider's Takeoff Summary format (438 West Pender sample). Estimate shows true cost first, then adjustment columns.",
    "A dedicated office project manager per job is a separate monthly cost, so the Office Project Manager GC line stays.",
    "Field supervision GC lines are not double counted with the Foreman & PM share of the crew mix.",
    "Template rates include nothing Knowify lacks; both are fully burdened. Bids now default to Knowify actual rates and mix; RATE VARIANCE tracks drift.",
    "Self-perform margin standard 15%, drop-down 8% to 20%. Subs margin drop-down 10% to 15%.",
    "Contingency: 2% of budgeted cost by default, on its own CONTINGENCY tab, added before the grand total.",
    "Company cost breakdown (MF/CM) entered on INPUTS and used as the back-check on BID SUMMARY.",
    "Knowify Job Summary export format confirmed. SOV and BUDGET use the same phase structure and line naming.",
], marker="OK")
au.freeze_panes = "A5"

# ---------- workbook-wide settings ----------
order = ["READ ME", "INPUTS", "LABOUR RATES", "TAKEOFF", "TAKEOFF CHECK", "PLUMBING EST", "HVAC EST", "SUBS & GC", "CONTINGENCY", "BID SUMMARY",
         "BUDGET", "SCHEDULE OF VALUES", "RATE VARIANCE", "KNOWIFY REVIEW", "KNOWIFY PHASES", "KNOWIFY TIME", "AUDIT NOTES"]
wb._sheets = [wb[n] for n in order]
tab_colors = {"READ ME": "7F7F7F", "INPUTS": "FFC000", "LABOUR RATES": "FFC000", "PLUMBING EST": "2F5597",
              "HVAC EST": "2F5597", "TAKEOFF": "A9D08E", "TAKEOFF CHECK": "70AD47", "SUBS & GC": "2F5597", "CONTINGENCY": "2F5597", "RATE VARIANCE": "70AD47", "BID SUMMARY": "C00000", "BUDGET": "548235",
              "SCHEDULE OF VALUES": "548235", "KNOWIFY REVIEW": "70AD47", "KNOWIFY PHASES": "A9D08E",
              "KNOWIFY TIME": "A9D08E", "AUDIT NOTES": "7F7F7F"}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = tab_colors[ws.title]
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if ws.title not in ("KNOWIFY PHASES", "KNOWIFY TIME"):
        ws.sheet_view.showGridLines = False
wb.active = wb.sheetnames.index("BID SUMMARY")
wb.save(OUT)
print("saved", OUT)
