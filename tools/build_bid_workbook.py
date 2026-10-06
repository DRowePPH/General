"""Builds 'New Excel bidding Spreadsheet.xlsx' for Paris Mechanical.

Rebuild of the 'Template-Empty - WORK IN PROGRESS.xlsm' bidding workbook:
full transparent cost (wage + burden, overhead, material incl. PST, subs, GCs)
with margin added on top, plus Budget and Schedule of Values tabs.

Usage: python tools/build_bid_workbook.py [output_path]
"""
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "New Excel bidding Spreadsheet.xlsx"

# ---------- styles ----------
FONT = "Arial"
BLUE = "0000FF"
GREEN = "008000"
NAVY = "1F3864"
F_IN = Font(name=FONT, size=10, color=BLUE)
F_CALC = Font(name=FONT, size=10, color="000000")
F_LINK = Font(name=FONT, size=10, color=GREEN)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=11, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
F_EX = Font(name=FONT, size=10, italic=True, color="7F7F7F")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FILL_EX = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOPLINE = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

CUR = '$#,##0;($#,##0);"-"'
CUR2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
NUM = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'

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


DIVS = [("P", "Plumbing"), ("H", "Hydronic"), ("V", "Ventilation"), ("AC", "Air Conditioning")]
PHASES = [
    "01 Mobilization & Submittals",
    "02 Underground",
    "03 Rough-in",
    "04 Finishing",
    "05 Equipment & Fixtures",
    "06 Testing, Balancing & Commissioning",
    "07 Closeout & Warranty",
    "08 General Conditions",
]
LEVELS = [("U/G", "Ground Work", None), ("P-02", "Parking 2", None), ("PRKG", "Parking 1", None),
          ("L-01", "Level 1", 24), ("L-02", "Level 2", 25), ("L-03", "Level 3", 25), ("L-04", "Level 4", 25),
          ("L-05", "Level 5", 24), ("L-06", "Level 6", None), ("L-07", "Level 7", None), ("L-08", "Level 8", None),
          ("L-09", "Level 9", None), ("L-10", "Level 10", None), ("ROOF", "Roof", None)]

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
    ("1. INPUTS", "Project info, durations, PST, consumables, contingency, target margin per division, level weights, your typical COGS %."),
    ("2. LABOUR RATES", "Crew mix per trade (same layout as the original PRJ INFO rates block). Weighted burdened wage + overhead $/hr = fully loaded rate."),
    ("3. PLUMBING EST / HVAC EST", "One row per item: Qty x Material $/unit and Qty x Hrs/unit. Same columns on both sheets so HVAC ties back exactly like Plumbing."),
    ("4. SUBS & GC", "Subtrade quotes and general conditions. Allocate by % to each division, or leave % blank to auto-split by direct cost."),
    ("5. CRM ACTUALS", "Paste item-level actuals and job-level COGS from Knowify. Estimates show your historical hrs/unit and $/unit beside each line."),
    ("6. BID SUMMARY", "Cost by type and division, then margin. Shows COGS % vs CRM history and vs your typical %. Holds every tie-out check."),
    ("7. BUDGET", "Cost budget per division and phase. This is what the field is held to (no margin in the phase lines)."),
    ("8. SCHEDULE OF VALUES", "Contract value split by division, phase, and level (rough-in and finishing by floor). Totals tie to the contract."),
    ("9. AUDIT NOTES", "Formula problems found in the original template, their dollar impact, and how this workbook handles them."),
]
for k, v in flow:
    put(ws, f"B{r}", k, F_BOLD)
    put(ws, f"C{r}", v, align=Alignment(wrap_text=True, vertical="top"))
    r += 1
r += 1
put(ws, f"B{r}", "COLOUR LEGEND", F_SUB); r += 1
legend = [
    ("Blue text, light yellow fill", "Input. Type here.", F_IN, FILL_IN),
    ("Bright yellow fill", "Key assumption. Review on every bid.", F_IN, FILL_KEY),
    ("Black text", "Formula. Do not type over.", F_CALC, None),
    ("Green text", "Link from another tab.", F_LINK, None),
    ("Grey italic row", "Example row showing the expected format. Not counted.", F_EX, FILL_EX),
]
for k, v, f, fl in legend:
    put(ws, f"B{r}", k, f, fill=fl)
    put(ws, f"C{r}", v)
    r += 1
r += 1
put(ws, f"B{r}", "MARGIN vs MARKUP", F_SUB); r += 1
for line in [
    "This workbook prices with MARGIN: Contract = Total Cost / (1 - Margin %).",
    "Example: cost $100,000 at 15% margin = $117,647 contract, $17,647 gross profit (17.6% markup).",
    "The original template used 15% MARKUP: $100,000 x 1.15 = $115,000, which is only a 13.0% margin.",
    "INPUTS shows the equivalent markup next to each margin so both views are visible.",
]:
    put(ws, f"C{r}", line); r += 1
r += 1
put(ws, f"B{r}", "STEP BY STEP", F_SUB); r += 1
for i, line in enumerate([
    "Fill INPUTS (project info, durations, margin). Check level units and weights.",
    "Confirm LABOUR RATES: burdened wage per classification, crew mix (must total 100%), overhead $/hr.",
    "Enter quantities on PLUMBING EST and HVAC EST. Use the Add columns for lump adjustments (negatives allowed).",
    "Compare your Hrs/Unit to the CRM Hrs/Unit column. Cells over +/-15% are highlighted.",
    "Enter sub quotes and GC items on SUBS & GC. Make sure every Check cell reads OK.",
    "Review BID SUMMARY: every check in the CHECKS block must read OK before the number goes out.",
    "BUDGET and SCHEDULE OF VALUES update automatically. Enter progress billing on the SOV each month.",
], 1):
    put(ws, f"C{r}", f"{i}. {line}"); r += 1

# =====================================================================
# INPUTS
# =====================================================================
inp = wb.create_sheet("INPUTS")
setw(inp, {"A": 3, "B": 40, "C": 24, "D": 16, "E": 16, "F": 14, "G": 40})
title(inp, "INPUTS", "Blue cells are inputs. Bright yellow = key assumption to review on every bid.")
put(inp, "B3", "PROJECT INFORMATION", F_SUB)
proj = [
    (4, "Project Name", None), (5, "Quote Number", None), (6, "Project Type", "Multi-Family"),
    (7, "Owner", None), (8, "General Contractor", None), (9, "Architect", None), (10, "Mech. Engineer", None),
    (11, "Site Address", "2650 East 41st Ave"), (12, "City", "Vancouver"), (13, "Proposal Date", None),
    (14, "Proposal Rev", None), (15, "Prepared By", "Paris Mechanical Ltd"), (16, "Drawings", "Issued For Tender"),
    (17, "Gross Floor Area (sq ft)", None),
]
for row, lab, val in proj:
    put(inp, f"B{row}", lab)
    put(inp, f"C{row}", val, F_IN, fill=FILL_IN, border=BOX)
inp["C13"].number_format = "yyyy-mm-dd"
inp["C17"].number_format = NUM
put(inp, "B18", "Number of Units")
put(inp, "C18", "=D55", fmt=NUM, border=BOX)
note(inp, "C18", "Sum of units from the LEVELS table below.")

put(inp, "B20", "KEY ASSUMPTIONS", F_SUB)
assum = [
    (21, "P&H project duration (months)", 24, NUM, "From original PRJ INFO H34."),
    (22, "HVAC project duration (months)", 24, NUM, "From original PRJ INFO H43."),
    (23, "PST on materials", 0.07, PCT, "BC PST 7%. From original PRJ INFO H52. Applied to material rows marked Y."),
    (24, "GST (shown on proposal only, not a cost)", 0.05, PCT, "From original PRJ INFO H51. GST is recoverable, so it is not in cost."),
    (25, "Consumables (% of material)", 0.02, PCT, "Original used 2% of material (P&H SUMMARY H46, HVAC SUMMARY D30)."),
    (26, "Contingency (% of direct cost)", 0.0, PCT, "Optional risk allowance. Kept as its own line so it stays visible."),
    (27, "Warranty reserve (% of equipment & fixture material)", 0.02, PCT, "Original: 2 years x 1% of equipment+fixtures (PRJ SUMMARY rows 75-78)."),
    (28, "Contract rounding (round up to nearest $)", 100, CUR, "Original rounded each division up to the nearest $100."),
    (29, "Holdback (BC Builders Lien Act)", 0.10, PCT, "Used on the SOV for progress billing."),
]
for row, lab, val, fmt, cmt in assum:
    put(inp, f"B{row}", lab)
    put(inp, f"C{row}", val, F_IN, fmt=fmt, fill=FILL_KEY, border=BOX)
    put(inp, f"G{row}", cmt, F_NOTE)

put(inp, "B31", "TARGET MARGIN BY DIVISION", F_SUB)
header_row(inp, 32, ["Division", "Code", "Target Margin %", "Equivalent Markup %", "Crew"], start_col=2)
for i, (code, name) in enumerate(DIVS):
    row = 33 + i
    put(inp, f"B{row}", name, border=BOX)
    put(inp, f"C{row}", code, border=BOX, align=Alignment(horizontal="center"))
    put(inp, f"D{row}", 0.15, F_IN, fmt=PCT, fill=FILL_KEY, border=BOX)
    put(inp, f"E{row}", f"=IFERROR(D{row}/(1-D{row}),0)", fmt=PCT, border=BOX)
    put(inp, f"F{row}", "P&H" if code in ("P", "H") else "HVAC", border=BOX, align=Alignment(horizontal="center"))
put(inp, "G33", "Margin = profit / contract. 15% is a placeholder: set your real target.", F_NOTE)

put(inp, "B38", "LEVELS (used to split rough-in and finishing by floor on the SOV)", F_SUB)
header_row(inp, 39, ["Tag", "Floor", "Units", "SOV Weight %"], start_col=2)
for i, (tag, floor, units) in enumerate(LEVELS):
    row = 40 + i
    put(inp, f"B{row}", tag, F_IN, fill=FILL_IN, border=BOX)
    put(inp, f"C{row}", floor, F_IN, fill=FILL_IN, border=BOX)
    put(inp, f"D{row}", units, F_IN, fmt=NUM, fill=FILL_IN, border=BOX)
    put(inp, f"E{row}", f"=IFERROR(D{row}/$D$55,0)", F_IN, fmt=PCT, fill=FILL_IN, border=BOX)
put(inp, "B55", "TOTAL", F_BOLD)
put(inp, "D55", "=SUM(D40:D53)", fmt=NUM, bold=True, border=TOPLINE)
put(inp, "E55", "=SUM(E40:E53)", fmt=PCT, bold=True, border=TOPLINE)
put(inp, "F55", '=IF(ABS(E55-1)<0.0001,"OK","CHECK")', bold=True)
ok_cf(inp, "F55")
put(inp, "G40", "Units per level from original PRJ INFO C6:C10.", F_NOTE)
put(inp, "G41", "Weight defaults to units share. Type over it to give parking,", F_NOTE)
put(inp, "G42", "U/G or roof a share. Weights must total 100%.", F_NOTE)

put(inp, "B57", "PHASES (drop-down list used on every tab)", F_SUB)
for i, ph in enumerate(PHASES):
    put(inp, f"B{58 + i}", ph, F_IN, fill=FILL_IN, border=BOX)
put(inp, "G58", "Rename freely, but keep 8 rows. 03 and 04 are split by level on the SOV.", F_NOTE)

put(inp, "B67", "YOUR TYPICAL COST OF GOODS (% of contract)", F_SUB)
header_row(inp, 68, ["Cost Type", "Typical % of Contract"], start_col=2)
COGS = ["Labour (wage + burden)", "Overhead", "Material (incl. PST + consumables)", "Subtrades",
        "General Conditions", "Contingency", "Gross Profit"]
for i, lab in enumerate(COGS):
    put(inp, f"B{69 + i}", lab, border=BOX)
    put(inp, f"C{69 + i}", None, F_IN, fmt=PCT, fill=FILL_IN, border=BOX)
put(inp, "B76", "TOTAL", F_BOLD)
put(inp, "C76", "=SUM(C69:C75)", fmt=PCT, bold=True, border=TOPLINE)
put(inp, "D76", '=IF(C76=0,"",IF(ABS(C76-1)<0.0001,"OK","CHECK"))', bold=True)
ok_cf(inp, "D76")
put(inp, "G69", "Enter your typical project COGS split here. BID SUMMARY compares this bid to it.", F_NOTE)
inp.freeze_panes = "A3"

# =====================================================================
# LABOUR RATES (mirrors original PRJ INFO rates block)
# =====================================================================
lr = wb.create_sheet("LABOUR RATES")
setw(lr, {"A": 3, "B": 30, "C": 16, "D": 13, "E": 16, "F": 13, "G": 18, "H": 4, "I": 60})
title(lr, "LABOUR RATES (crew mix)", "Same layout as the original PRJ INFO rates block. Both trades now tie back the same way.")
put(lr, "B2", "UPDATED 2026-06-29", F_BOLD, fill=FILL_KEY)


def crew_block(top, label, hours_formula):
    """Rows: top=section title, top+1 headers, top+2..top+6 classes, top+7 totals."""
    put(lr, f"B{top}", label, F_SUB)
    header_row(lr, top + 1, ["Classification", "Burdened Rate $/hr", "Crew Mix %", "Weighted $/hr",
                             "Hours", "Cost by Wage Breakdown"], start_col=2)
    classes = [("Apprentice Lvl 1-2", 53, 0.40), ("Apprentice Lvl 3-6", 62.5, 0.30),
               ("Apprentice Lvl 7-8", 66.5, 0.05), ("Journeyman Lvl 1-2", 73, 0.11), ("Foreman & PM", 96, 0.14)]
    first, last = top + 2, top + 6
    for i, (nm, rate, mix) in enumerate(classes):
        r_ = first + i
        put(lr, f"B{r_}", nm, F_IN, fill=FILL_IN, border=BOX)
        put(lr, f"C{r_}", rate, F_IN, fmt=CUR2, fill=FILL_KEY, border=BOX)
        put(lr, f"D{r_}", mix, F_IN, fmt=PCT, fill=FILL_KEY, border=BOX)
        put(lr, f"E{r_}", f"=C{r_}*D{r_}", fmt=CUR2, border=BOX)
        put(lr, f"F{r_}", f"=$C${last + 5}*D{r_}", fmt=NUM, border=BOX)
        put(lr, f"G{r_}", f"=F{r_}*C{r_}", fmt=CUR, border=BOX)
    t = last + 1
    put(lr, f"B{t}", "Weighted average burdened wage", F_BOLD)
    put(lr, f"D{t}", f"=SUM(D{first}:D{last})", fmt=PCT, bold=True, border=TOPLINE)
    put(lr, f"E{t}", f"=SUM(E{first}:E{last})", fmt=CUR2, bold=True, fill=FILL_TOT, border=TOPLINE)
    put(lr, f"F{t}", f"=SUM(F{first}:F{last})", fmt=NUM, bold=True, border=TOPLINE)
    put(lr, f"G{t}", f"=SUM(G{first}:G{last})", fmt=CUR, bold=True, border=TOPLINE)
    put(lr, f"B{t + 1}", "Overhead $/hr (shop, office, vehicles, insurance)")
    put(lr, f"C{t + 1}", 22, F_IN, fmt=CUR2, fill=FILL_KEY, border=BOX)
    put(lr, f"B{t + 2}", "FULLY LOADED RATE $/hr", F_BOLD)
    put(lr, f"C{t + 2}", f"=E{t}+C{t + 1}", fmt=CUR2, bold=True, fill=FILL_TOT, border=BOX)
    put(lr, f"B{t + 3}", "Crew mix = 100% and hours x rate ties to breakdown")
    put(lr, f"C{t + 3}", f'=IF(AND(ABS(D{t}-1)<0.0001,ABS(G{t}-F{t}*E{t})<1),"OK","CHECK")', bold=True)
    ok_cf(lr, f"C{t + 3}")
    put(lr, f"B{t + 4}", "Estimated hours (from estimate tabs)")
    put(lr, f"C{t + 4}", hours_formula, F_LINK, fmt=NUM, border=BOX)
    put(lr, f"I{first}", "Rates are fully burdened (confirmed): from original PRJ INFO D36:D40.", F_NOTE)
    put(lr, f"I{first + 1}", "Crew mix from original PRJ INFO E36:E40.", F_NOTE)
    put(lr, f"I{first + 2}", "Overhead $22/hr from original PRJ INFO H35/H44.", F_NOTE)
    put(lr, f"I{first + 3}", "Hours x mix x rate = the 'Cost by wage breakdown' column", F_NOTE)
    put(lr, f"I{first + 4}", "in your screenshot, so it reconciles to the estimate labour $.", F_NOTE)
    return t


EST_FIRST, EST_LAST = 7, 156


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


ph_t = crew_block(4, "PLUMBING & HYDRONIC CREW", "=" + est_sumifs("M", '"P"') + "+" + est_sumifs("M", '"H"'))
hv_t = crew_block(19, "HVAC CREW (ventilation & A/C)", "=" + est_sumifs("M", '"V"') + "+" + est_sumifs("M", '"AC"'))
# ph_t = 11 -> wage E11, OH C12, loaded C13, check C14, hours C15
# hv_t = 26 -> wage E26, OH C27, loaded C28, check C29, hours C30
PH_WAGE, PH_OH, PH_LOAD = f"'LABOUR RATES'!$E${ph_t}", f"'LABOUR RATES'!$C${ph_t + 1}", f"'LABOUR RATES'!$C${ph_t + 2}"
HV_WAGE, HV_OH, HV_LOAD = f"'LABOUR RATES'!$E${hv_t}", f"'LABOUR RATES'!$C${hv_t + 1}", f"'LABOUR RATES'!$C${hv_t + 2}"

cmp_r = hv_t + 7
put(lr, f"B{cmp_r}", "WHY THIS MATTERS: rate the original template used vs true loaded rate", F_SUB)
header_row(lr, cmp_r + 1, ["Rate Basis", "$/hr", "", "", "P&H Hours", "Labour + OH Cost"], start_col=2)
rows_cmp = [
    ("Original line items (Apprentice 1-2 only)", "=C6"),
    ("Original summary 'Budget with' (simple average)", "=AVERAGE(C6:C10)"),
    ("Crew-mix weighted burdened wage", f"=E{ph_t}"),
    ("Fully loaded (weighted wage + overhead)", f"=C{ph_t + 2}"),
]
for i, (lab, f) in enumerate(rows_cmp):
    r_ = cmp_r + 2 + i
    put(lr, f"B{r_}", lab, border=BOX)
    put(lr, f"C{r_}", f, fmt=CUR2, border=BOX)
    put(lr, f"F{r_}", f"=$C${ph_t + 4}", fmt=NUM, border=BOX)
    put(lr, f"G{r_}", f"=C{r_}*F{r_}", fmt=CUR, border=BOX)
put(lr, f"I{cmp_r + 2}", "At 20,800 hrs the original bid labour at $70.20 = $1,460,160.", F_NOTE)
put(lr, f"I{cmp_r + 3}", "True loaded cost at $86.75 = $1,804,296. Gap: $344,136 before margin.", F_NOTE)
lr.freeze_panes = "A3"

# =====================================================================
# ESTIMATE SHEETS
# =====================================================================
EST_HEAD = ["Item Code", "Div", "Phase", "Description", "Qty", "Unit", "Material $/Unit", "Hrs/Unit",
            "Add Material $", "Add Hrs", "PST? (Y/N)", "Material $ (incl PST)", "Hours", "Labour $ (wage+burden)",
            "Overhead $", "TOTAL COST $", "CRM Hrs/Unit", "CRM Mat $/Unit", "Hrs/Unit vs CRM"]

PL_ITEMS = [
    # code, div, phase idx, desc, unit, mat/unit, hrs/unit, example qty
    ("P-UG-SUMP", "P", 1, "Sumps & Interceptors", "ea", None, None, None),
    ("P-UG-STRM", "P", 1, "Storm Water & Drain Tile", "lf", None, None, None),
    ("P-UG-SAN", "P", 1, "Sanitary (underground)", "lf", None, None, None),
    ("P-UG-DW", "P", 1, "Domestic Water (underground)", "lf", None, None, None),
    ("P-RI-STRM", "P", 2, "DWV Storm Water", "lf", None, None, None),
    ("P-RI-SAN", "P", 2, "DWV Sanitary (Drainage/Vents)", "lf", None, None, None),
    ("P-RI-WR2", "P", 2, "2-Piece Bathroom (WC, LAV)", "ea", 550, 3, None),
    ("P-RI-WR3T", "P", 2, "3-Piece Bathroom (WC, LAV, TUB)", "ea", 860, 5, None),
    ("P-RI-WR3S", "P", 2, "3-Piece Bathroom (WC, LAV, SH)", "ea", 860, 5, None),
    ("P-RI-WR4T", "P", 2, "4-Piece Bathroom (WC, 2LAV, TUB)", "ea", 990, 6, None),
    ("P-RI-WR4S", "P", 2, "4-Piece Bathroom (WC, 2LAV, SH)", "ea", 990, 6, None),
    ("P-RI-WR4TS", "P", 2, "4-Piece Bathroom (WC, LAV, TUB & SH)", "ea", 1310, 6, None),
    ("P-RI-WR5", "P", 2, "5-Piece Bathroom (WC, 2LAV, TUB & SH)", "ea", 1460, 6, None),
    ("P-RI-LAUN", "P", 2, "Laundry - 4\" Wet Vent", "ea", 280, 2, None),
    ("P-RI-KS3", "P", 2, "Kitchen Sink - 3\" Vent", "ea", 240, 2, None),
    ("P-RI-KS4", "P", 2, "Kitchen Sink - 4\" Vent", "ea", 250, 2, None),
    ("P-RI-DWD", "P", 2, "Domestic Water Distribution Piping", "lf", None, None, None),
    ("P-RI-DWS", "P", 2, "Domestic Water Insuite Piping", "suite", 800, 12, 123),
    ("P-RI-IRR", "P", 2, "Irrigation", "LS", None, None, None),
    ("P-RI-COND", "P", 2, "Condensate Drains (if by Plumbing)", "LS", None, None, None),
    ("P-RI-GAS", "P", 2, "Gas & Boiler Vents", "LS", None, None, None),
    ("P-RI-DRN", "P", 2, "Drains & Hose Bibs (rough-in)", "ea", None, None, None),
    ("P-RI-WER", "P", 2, "Water Entry Room", "LS", None, None, None),
    ("P-RI-FLSH", "P", 2, "Roof Flashings", "ea", None, None, None),
    ("P-RI-CAN", "P", 2, "Canning / Cutting / Layout", "LS", None, None, None),
    ("P-RI-CLN", "P", 2, "Clean-up", "LS", None, None, None),
    ("P-FN-FIX", "P", 3, "Set Plumbing Fixtures", "ea", None, None, None),
    ("P-FN-AP", "P", 3, "Access Panels", "ea", None, None, None),
    ("P-FN-LBL", "P", 3, "Pipe Labelling", "LS", None, None, None),
    ("P-EQ-FIX", "P", 4, "Plumbing Fixtures Supply (WC, LAV, KS, SH, BT)", "LS", None, None, None),
    ("P-EQ-DRN", "P", 4, "Drains Supply (FD, AD, RD, HB, TD, TP)", "LS", None, None, None),
    ("P-EQ-EQP", "P", 4, "Plumbing Equipment (DHW, pumps, tanks)", "LS", None, None, None),
    ("P-TS-TEST", "P", 5, "Testing", "LS", None, None, None),
    ("P-CL-DOCS", "P", 6, "As-builts, O&M, Closeout", "LS", None, None, None),
    ("H-UG-PIPE", "H", 1, "Hydronic U/G Piping", "lf", None, None, None),
    ("H-RI-DIST", "H", 2, "Heating Distribution Piping", "lf", None, None, None),
    ("H-RI-MAN", "H", 2, "Manifolds", "ea", None, None, None),
    ("H-RI-SUITE", "H", 2, "Hydronic Insuite Piping", "suite", None, 8, None),
    ("H-RI-GLY", "H", 2, "Glycol Loop", "LS", None, None, None),
    ("H-RI-CAN", "H", 2, "Canning / Cutting / Layout", "LS", None, None, None),
    ("H-RI-CLN", "H", 2, "Clean-up", "LS", None, None, None),
    ("H-FN-AP", "H", 3, "Access Panels", "ea", None, None, None),
    ("H-FN-LBL", "H", 3, "Pipe Labelling", "LS", None, None, None),
    ("H-EQ-MECH", "H", 4, "Mechanical Room", "LS", None, None, None),
    ("H-EQ-EQP", "H", 4, "Hydronic Equipment (boilers, pumps, HX)", "LS", None, None, None),
    ("H-TS-TEST", "H", 5, "Testing", "LS", None, None, None),
]
HV_ITEMS = [
    ("V-RI-DUCT", "V", 2, "Distribution Ductwork", "lb", None, None, None),
    ("V-RI-SUITE", "V", 2, "Insuite Ductwork", "suite", None, None, None),
    ("V-RI-COND", "V", 2, "Condensate Drains (if by HVAC)", "LS", None, None, None),
    ("V-RI-INS", "V", 2, "Thermal Insulation (in-house)", "LS", None, None, None),
    ("V-RI-CAN", "V", 2, "Canning / Cutting / Layout", "LS", None, None, None),
    ("V-RI-CLN", "V", 2, "Clean-up", "LS", None, None, None),
    ("V-FN-GRL", "V", 3, "Grilles & Ventilation Terminations", "ea", None, None, None),
    ("V-FN-AP", "V", 3, "Access Panels", "ea", None, None, None),
    ("V-FN-LBL", "V", 3, "Duct Labelling", "LS", None, None, None),
    ("V-EQ-FAN", "V", 4, "Fans & ERVs", "ea", None, None, None),
    ("V-TS-TEST", "V", 5, "Testing", "LS", None, None, None),
    ("AC-RI-DUCT", "AC", 2, "Distribution Ductwork", "lb", None, None, None),
    ("AC-RI-SUITE", "AC", 2, "Insuite Ductwork", "suite", None, None, None),
    ("AC-RI-COND", "AC", 2, "Condensate Drains (if by HVAC)", "LS", None, None, None),
    ("AC-RI-LINE", "AC", 2, "Refrigeration Linesets", "lf", None, None, None),
    ("AC-RI-PTAC", "AC", 2, "PTAC Rough-in", "ea", None, None, None),
    ("AC-RI-LS12", "AC", 2, "AC-1.1/2 Lineset", "ea", None, None, None),
    ("AC-RI-CAN", "AC", 2, "Canning / Cutting / Layout", "LS", None, None, None),
    ("AC-RI-CLN", "AC", 2, "Clean-up", "LS", None, None, None),
    ("AC-FN-GRL", "AC", 3, "Diffusers & Registers", "ea", None, None, None),
    ("AC-FN-AP", "AC", 3, "Access Panels", "ea", None, None, None),
    ("AC-EQ-AHU", "AC", 4, "AHUs & Heat Pumps", "ea", None, None, None),
    ("AC-TS-TEST", "AC", 5, "Testing", "LS", None, None, None),
]

CRM_ITEM_FIRST, CRM_ITEM_LAST = 7, 2006
CRM_JOB_FIRST, CRM_JOB_LAST = 7, 106


def crm(col):
    return f"'CRM ACTUALS'!${col}${CRM_ITEM_FIRST}:${col}${CRM_ITEM_LAST}"


def build_est(name, items, subtitle):
    ws = wb.create_sheet(name)
    widths = {"A": 13, "B": 6, "C": 30, "D": 40, "E": 9, "F": 7, "G": 13, "H": 10, "I": 13, "J": 9, "K": 8,
              "L": 15, "M": 10, "N": 16, "O": 13, "P": 16, "Q": 11, "R": 12, "S": 11}
    setw(ws, widths)
    title(ws, name, subtitle)
    put(ws, "A3", "TOTALS", F_BOLD)
    for col in "LMNOP":
        put(ws, f"{col}3", f"=SUM({col}{EST_FIRST}:{col}{EST_LAST})", fmt=NUM if col == "M" else CUR, bold=True,
            fill=FILL_TOT, border=BOX)
    put(ws, "A4", "Labour $ = Hours x crew-mix weighted burdened wage. Overhead $ = Hours x overhead $/hr. "
                  "P/H rows use the P&H crew, V/AC rows use the HVAC crew (LABOUR RATES).", F_NOTE)
    header_row(ws, 6, EST_HEAD)
    # example/format row 5 is the header context; data starts row 7
    for i in range(EST_FIRST, EST_LAST + 1):
        idx = i - EST_FIRST
        item = items[idx] if idx < len(items) else None
        if item:
            code, div, ph, desc, unit, mat, hrs, qty = item
            vals = {"A": code, "B": div, "C": PHASES[ph], "D": desc, "E": qty, "F": unit, "G": mat, "H": hrs, "K": "Y"}
        else:
            vals = {"K": "Y"}
        for col in "ABCDEFGHIJK":
            c = put(ws, f"{col}{i}", vals.get(col), F_IN, fill=FILL_IN, border=BOX)
            if col in "EJ":
                c.number_format = NUM1
            if col in "GI":
                c.number_format = CUR2
            if col == "H":
                c.number_format = '0.00;(0.00);"-"'
        rate_w = f'IF(OR($B{i}="V",$B{i}="AC"),{HV_WAGE},{PH_WAGE})'
        rate_o = f'IF(OR($B{i}="V",$B{i}="AC"),{HV_OH},{PH_OH})'
        put(ws, f"L{i}", f'=($E{i}*$G{i}+$I{i})*(1+IF($K{i}="N",0,INPUTS!$C$23))', fmt=CUR, border=BOX)
        put(ws, f"M{i}", f"=$E{i}*$H{i}+$J{i}", fmt=NUM1, border=BOX)
        put(ws, f"N{i}", f"=$M{i}*{rate_w}", fmt=CUR, border=BOX)
        put(ws, f"O{i}", f"=$M{i}*{rate_o}", fmt=CUR, border=BOX)
        put(ws, f"P{i}", f"=$L{i}+$N{i}+$O{i}", fmt=CUR, bold=True, border=BOX)
        put(ws, f"Q{i}", f'=IF($A{i}="","",IFERROR(SUMIFS({crm("H")},{crm("D")},$A{i})/SUMIFS({crm("F")},{crm("D")},$A{i}),""))',
            fmt='0.00', border=BOX)
        put(ws, f"R{i}", f'=IF($A{i}="","",IFERROR(SUMIFS({crm("J")},{crm("D")},$A{i})/SUMIFS({crm("F")},{crm("D")},$A{i}),""))',
            fmt=CUR2, border=BOX)
        put(ws, f"S{i}", f'=IF(OR($Q{i}="",N($H{i})=0),"",$H{i}/$Q{i}-1)', fmt=PCT, border=BOX)
        if item and item[7]:
            for col in "EGH":
                ws[f"{col}{i}"].fill = FILL_KEY
            note(ws, f"E{i}", "EXAMPLE quantity: 123 units from PRJ INFO, $800 ea and 12 hrs ea from the original "
                              "'Insuite Piping - $800 ea @ 12 HRS ea' label. Replace with your takeoff.")
    rng = f"S{EST_FIRST}:S{EST_LAST}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(S{EST_FIRST}),ABS(S{EST_FIRST})>0.15)'],
                                  fill=PatternFill("solid", fgColor="FFC7CE")))
    dv_div = DataValidation(type="list", formula1="INPUTS!$C$33:$C$36", allow_blank=True)
    dv_ph = DataValidation(type="list", formula1="INPUTS!$B$58:$B$65", allow_blank=True)
    dv_yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    for dv, col in ((dv_div, "B"), (dv_ph, "C"), (dv_yn, "K")):
        ws.add_data_validation(dv)
        dv.add(f"{col}{EST_FIRST}:{col}{EST_LAST}")
    note(ws, "H6", "Labour hours per unit. Compare to CRM Hrs/Unit (actuals from Knowify). "
                   "Cells more than 15% off history turn red.")
    note(ws, "I6", "Lump material adjustment ($). Negative numbers allowed.")
    note(ws, "J6", "Lump hours adjustment. Negative numbers allowed.")
    note(ws, "A6", "Item Code links this line to CRM ACTUALS. Use the same code in your Knowify cost codes.")
    ws.freeze_panes = "E7"
    ws.auto_filter.ref = f"A6:S{EST_LAST}"
    return ws


build_est("PLUMBING EST", PL_ITEMS, "Plumbing (P) and Hydronic (H). Seed unit costs and hours from the original EQUIPMENT sheet.")
build_est("HVAC EST", HV_ITEMS, "Ventilation (V) and Air Conditioning (AC). Identical columns and formulas to PLUMBING EST.")

# =====================================================================
# SUBS & GC
# =====================================================================
sg = wb.create_sheet("SUBS & GC")
setw(sg, {"A": 5, "B": 30, "C": 22, "D": 30, "E": 8, "F": 9, "G": 13, "H": 15, "I": 8, "J": 8, "K": 8, "L": 8,
          "M": 13, "N": 13, "O": 13, "P": 13, "Q": 9})
title(sg, "SUBTRADES & GENERAL CONDITIONS", "Allocation %: enter a split per division, or leave all four blank to auto-split by direct cost.")
put(sg, "H3", "Direct cost by division", F_NOTE, align=Alignment(horizontal="right"))
put(sg, "H4", "Auto split share", F_NOTE, align=Alignment(horizontal="right"))
for i, (code, nm) in enumerate(DIVS):
    col = get_column_letter(9 + i)
    put(sg, f"{col}3", "=" + est_sumifs("P", f'"{code}"'), F_LINK, fmt=CUR)
    put(sg, f"{col}4", f"=IFERROR({col}3/SUM($I$3:$L$3),0)", fmt=PCT)
SG_HEAD = ["#", "Item", "Company / Basis", "Phase", "Qty", "Unit", "Rate $", "Amount $",
           "% P", "% H", "% V", "% AC", "Plumbing $", "Hydronic $", "Ventilation $", "A/C $", "Check"]

SUB_FIRST, SUB_LAST = 7, 31
GC_FIRST, GC_LAST = 36, 65


def alloc_row(r_):
    for i in range(4):
        pc = get_column_letter(9 + i)
        dc = get_column_letter(13 + i)
        put(sg, f"{dc}{r_}", f"=$H{r_}*IF(SUM($I{r_}:$L{r_})=0,{pc}$4,{pc}{r_})", fmt=CUR, border=BOX)
    put(sg, f"Q{r_}", f'=IF($H{r_}=0,"",IF(ABS(SUM(M{r_}:P{r_})-$H{r_})<0.01,"OK","CHECK"))', bold=True, border=BOX)


def alloc_inputs(r_, split):
    for i, code in enumerate(["P", "H", "V", "AC"]):
        pc = get_column_letter(9 + i)
        put(sg, f"{pc}{r_}", split.get(code), F_IN, fmt='0%;;""', fill=FILL_IN, border=BOX)


put(sg, "A5", "SUBTRADES (enter quotes as received, before margin)", F_SUB)
header_row(sg, 6, SG_HEAD)
SUBS = [
    ("Pipe Insulation & Heat Trace", "Adler", 2, {"P": 1}),
    ("Duct Insulation", None, 2, {"V": 0.5, "AC": 0.5}),
    ("P&H Controls", None, 4, {"P": 0.5, "H": 0.5}),
    ("HVAC Controls", "Olympic Controls", 4, {"V": 0.5, "AC": 0.5}),
    ("TAB (Test, Adjust, Balance)", "Western Mechanical", 5, {"V": 1}),
    ("Commissioning", None, 5, {"AC": 1}),
    ("Canning / Piping Plans", "Paris Mechanical", 0, {"P": 1}),
    ("Coring / Scanning", None, 2, {}),
    ("Fixtures Caulking", "Paris Mechanical", 3, {"P": 1}),
    ("Water Balancing", "Western Mechanical", 5, {"P": 1}),
    ("Sump Pumps", None, 4, {"P": 1}),
    ("Chemical Treatment", None, 5, {"H": 1}),
    ("P&H Seismic / Schedule", None, 2, {"P": 0.5, "H": 0.5}),
    ("HVAC Seismic / Schedule", None, 2, {"V": 0.5, "AC": 0.5}),
    ("As-built Drawings", "Paris Mechanical", 6, {}),
    ("O&M Manuals", "Paris Mechanical", 6, {}),
    ("Plumbing Bond", None, 0, {"P": 1}),
    ("HVAC Bond", None, 0, {"V": 0.5, "AC": 0.5}),
    ("Welding", None, 2, {"V": 1}),
    ("Gas Service / Sprinklers / Other", None, 2, {}),
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
GC = [
    ("Supervision - Plumbing", "=INPUTS!$C$21", "months", None, 7, {"P": 1}),
    ("Supervision - Hydronic", "=INPUTS!$C$21", "months", None, 7, {"H": 1}),
    ("Supervision - Ventilation", "=INPUTS!$C$22", "months", None, 7, {"V": 1}),
    ("Supervision - A/C", "=INPUTS!$C$22", "months", None, 7, {"AC": 1}),
    ("Office Project Manager (dedicated to this job)", "=MAX(INPUTS!$C$21,INPUTS!$C$22)", "months", 7500, 7, {}),
    ("Bonus - Plumbing", 1, "LS", None, 7, {"P": 1}),
    ("Bonus - Hydronic", 1, "LS", None, 7, {"H": 1}),
    ("Bonus - Ventilation", 1, "LS", None, 7, {"V": 1}),
    ("Bonus - A/C", 1, "LS", None, 7, {"AC": 1}),
    ("Bonus - Office", 1, "LS", None, 7, {}),
    ("Shop Drawings", 1, "LS", None, 0, {}),
    ("Plumbing & Gas Permit", 1, "LS", None, 0, {"P": 1}),
    ("Prints", 1, "LS", None, 0, {}),
    ("Rentals", 4, "months", None, 7, {}),
    ("Trailer", 12, "months", None, 7, {}),
    ("Legal Fees", "=MAX(INPUTS!$C$21,INPUTS!$C$22)", "months", None, 7, {}),
    ("Freight", "=MAX(INPUTS!$C$21,INPUTS!$C$22)", "months", None, 7, {}),
    ("Safety", "=MAX(INPUTS!$C$21,INPUTS!$C$22)", "months", None, 7, {}),
]
for j, (code, nm) in enumerate(DIVS):
    GC.append((f"Warranty Reserve - {nm}", 1, "LS",
               f'=ROUNDUP(INPUTS!$C$27*({est_sumifs("L", chr(34) + code + chr(34), "INPUTS!$B$62")}),-2)', 6, {code: 1}))
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
put(sg, f"C{GC_FIRST + 4}", "Dedicated office PM, per month", F_NOTE)
note(sg, f"G{GC_FIRST + 4}", "$7,500/month from original PRJ INFO D53. The original charged this 100% to Plumbing; "
                             "here it auto-splits by direct cost. Confirmed: the $22/hr overhead covers salaried staff company wide; "
                             "this line is the PM assigned to this job, so it is not a double count.")
note(sg, f"B{GC_FIRST}", "Confirmed: field supervision here is separate from the Foreman & PM share of the crew mix, "
                         "so it is not a double count.")
put(sg, f"B{GC_LAST + 1}", "GENERAL CONDITIONS TOTAL", F_BOLD)
for col in "HMNOP":
    put(sg, f"{col}{GC_LAST + 1}", f"=SUM({col}{GC_FIRST}:{col}{GC_LAST})", fmt=CUR, bold=True, fill=FILL_TOT, border=TOPLINE)
ok_cf(sg, f"Q{SUB_FIRST}:Q{GC_LAST}")
dv_ph2 = DataValidation(type="list", formula1="INPUTS!$B$58:$B$65", allow_blank=True)
sg.add_data_validation(dv_ph2)
dv_ph2.add(f"D{SUB_FIRST}:D{SUB_LAST}")
dv_ph2.add(f"D{GC_FIRST}:D{GC_LAST}")
sg.freeze_panes = "C7"

# =====================================================================
# CRM ACTUALS
# =====================================================================
cr = wb.create_sheet("CRM ACTUALS")
setw(cr, {"A": 10, "B": 26, "C": 7, "D": 13, "E": 6, "F": 11, "G": 7, "H": 12, "I": 14, "J": 14, "K": 22, "L": 3,
          "M": 10, "N": 26, "O": 7, "P": 14, "Q": 14, "R": 13, "S": 14, "T": 14, "U": 14, "V": 14, "W": 10})
title(cr, "CRM ACTUALS (Knowify)", "Paste exported actuals from row 7 down. Grey row 5 is a format example and is NOT counted.")
put(cr, "A3", "ITEM-LEVEL PRODUCTION (feeds CRM Hrs/Unit and CRM Mat $/Unit on the estimate tabs)", F_SUB)
header_row(cr, 6, ["Job #", "Job Name", "Year", "Item Code", "Div", "Qty Installed", "Unit", "Actual Hours",
                   "Actual Labour $", "Actual Material $", "Notes"])
ex = ["EXAMPLE", "Example job (not counted)", 2025, "P-RI-DWS", "P", 100, "suite", 1150, 74000, 82000, "Format example only"]
for i, v in enumerate(ex):
    c = put(cr, f"{get_column_letter(1 + i)}5", v, F_EX, fill=FILL_EX, border=BOX)
put(cr, "M3", "JOB-LEVEL COST OF GOODS HISTORY (feeds BID SUMMARY comparison)", F_SUB)
header_row(cr, 6, ["Job #", "Job Name", "Year", "Contract $", "Labour $", "Overhead $", "Material $",
                   "Subtrades $", "Gen. Cond. $", "Total Cost $", "GP %"], start_col=13)
exj = ["EXAMPLE", "Example job (not counted)", 2025, 1000000, 300000, 90000, 330000, 80000, 60000]
for i, v in enumerate(exj):
    put(cr, f"{get_column_letter(13 + i)}5", v, F_EX, fmt=CUR if i >= 3 else None, fill=FILL_EX, border=BOX)
put(cr, "V5", "=SUM(Q5:U5)", F_EX, fmt=CUR, fill=FILL_EX, border=BOX)
put(cr, "W5", "=IFERROR(1-V5/P5,0)", F_EX, fmt=PCT, fill=FILL_EX, border=BOX)
for k in range(CRM_ITEM_FIRST, CRM_ITEM_FIRST + 60):
    for col in "ABCDEFGHIJK":
        c = cr[f"{col}{k}"]
        c.font = F_IN
        c.border = BOX
        if col in "IJ":
            c.number_format = CUR
        if col in "FH":
            c.number_format = NUM1
for k in range(CRM_JOB_FIRST, CRM_JOB_LAST + 1):
    for col in "MNOPQRSTU":
        c = cr[f"{col}{k}"]
        c.font = F_IN
        c.border = BOX
        if col >= "P":
            c.number_format = CUR
    put(cr, f"V{k}", f'=IF(P{k}="","",SUM(Q{k}:U{k}))', fmt=CUR, border=BOX)
    put(cr, f"W{k}", f'=IF(N(P{k})=0,"",1-V{k}/P{k})', fmt=PCT, border=BOX)
note(cr, "D6", "Must match the Item Code on PLUMBING EST / HVAC EST. Set the same codes up as Knowify cost codes.")
cr.freeze_panes = "A7"

# =====================================================================
# BID SUMMARY
# =====================================================================
bs = wb.create_sheet("BID SUMMARY")
setw(bs, {"A": 6, "B": 20, "C": 10, "D": 14, "E": 13, "F": 14, "G": 12, "H": 13, "I": 13, "J": 12, "K": 15,
          "L": 9, "M": 15, "N": 15, "O": 14, "P": 9, "Q": 9, "R": 11, "S": 10, "T": 9})
title(bs, "BID SUMMARY: full cost, then margin")
put(bs, "A2", '=IF(INPUTS!C4="","(enter project name on INPUTS)",INPUTS!C4)&"  |  "&INPUTS!C11&", "&INPUTS!C12', F_LINK)
header_row(bs, 5, ["Code", "Division", "Hours", "Labour $ (wage+burden)", "Overhead $", "Material $ (incl PST)",
                   "Consumables $", "Subtrades $", "Gen. Cond. $", "Contingency $", "TOTAL COST $",
                   "Target Margin", "Sell Price $", "CONTRACT $ (rounded)", "Gross Profit $", "Actual Margin",
                   "Markup on Cost", "Per Unit $", "Per Sq Ft $", "Hrs / Unit"])
sub_cols = ["M", "N", "O", "P"]
for i, (code, nm) in enumerate(DIVS):
    r_ = 6 + i
    put(bs, f"A{r_}", f"=INPUTS!C{33 + i}", F_LINK, border=BOX, align=Alignment(horizontal="center"))
    put(bs, f"B{r_}", f"=INPUTS!B{33 + i}", F_LINK, border=BOX)
    put(bs, f"C{r_}", "=" + est_sumifs("M", f"$A{r_}"), fmt=NUM, border=BOX)
    put(bs, f"D{r_}", "=" + est_sumifs("N", f"$A{r_}"), fmt=CUR, border=BOX)
    put(bs, f"E{r_}", "=" + est_sumifs("O", f"$A{r_}"), fmt=CUR, border=BOX)
    put(bs, f"F{r_}", "=" + est_sumifs("L", f"$A{r_}"), fmt=CUR, border=BOX)
    put(bs, f"G{r_}", f"=F{r_}*INPUTS!$C$25", fmt=CUR, border=BOX)
    put(bs, f"H{r_}", f"='SUBS & GC'!{sub_cols[i]}{SUB_LAST + 1}", F_LINK, fmt=CUR, border=BOX)
    put(bs, f"I{r_}", f"='SUBS & GC'!{sub_cols[i]}{GC_LAST + 1}", F_LINK, fmt=CUR, border=BOX)
    put(bs, f"J{r_}", f"=SUM(D{r_}:I{r_})*INPUTS!$C$26", fmt=CUR, border=BOX)
    put(bs, f"K{r_}", f"=SUM(D{r_}:J{r_})", fmt=CUR, bold=True, fill=FILL_TOT, border=BOX)
    put(bs, f"L{r_}", f"=INPUTS!D{33 + i}", F_LINK, fmt=PCT, border=BOX)
    put(bs, f"M{r_}", f"=IFERROR(K{r_}/(1-L{r_}),0)", fmt=CUR, border=BOX)
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
put(bs, "N11", "=N10*INPUTS!$C$23", fmt=CUR)
put(bs, "M12", "Contract + GST", F_NOTE)
put(bs, "N12", "=N10+N11", fmt=CUR, bold=True)

put(bs, "B14", "COST OF GOODS MIX (whole job)", F_SUB)
header_row(bs, 15, ["Cost Type", "", "This Bid $", "% of Contract", "CRM History %", "Your Typical %", "Bid vs Typical"], start_col=2)
bs.merge_cells("B15:C15")
cj = lambda col: f"'CRM ACTUALS'!${col}${CRM_JOB_FIRST}:${col}${CRM_JOB_LAST}"
mix = [
    ("Labour (wage + burden)", "=D10", f"=IFERROR(SUM({cj('Q')})/SUM({cj('P')}),\"\")"),
    ("Overhead", "=E10", f"=IFERROR(SUM({cj('R')})/SUM({cj('P')}),\"\")"),
    ("Material (incl PST + consumables)", "=F10+G10", f"=IFERROR(SUM({cj('S')})/SUM({cj('P')}),\"\")"),
    ("Subtrades", "=H10", f"=IFERROR(SUM({cj('T')})/SUM({cj('P')}),\"\")"),
    ("General Conditions", "=I10", f"=IFERROR(SUM({cj('U')})/SUM({cj('P')}),\"\")"),
    ("Contingency", "=J10", '=""'),
    ("Gross Profit", "=O10", f"=IFERROR(1-SUM({cj('V')})/SUM({cj('P')}),\"\")"),
]
for i, (lab, f_bid, f_hist) in enumerate(mix):
    r_ = 16 + i
    put(bs, f"B{r_}", lab, border=BOX)
    bs.merge_cells(f"B{r_}:C{r_}")
    put(bs, f"D{r_}", f_bid, fmt=CUR, border=BOX)
    put(bs, f"E{r_}", f"=IFERROR(D{r_}/$N$10,0)", fmt=PCT, border=BOX)
    put(bs, f"F{r_}", f_hist, F_LINK, fmt=PCT, border=BOX)
    put(bs, f"G{r_}", f"=IF(INPUTS!C{69 + i}=\"\",\"\",INPUTS!C{69 + i})", F_LINK, fmt=PCT, border=BOX)
    put(bs, f"H{r_}", f'=IF(G{r_}="","",E{r_}-G{r_})', fmt=PCT, border=BOX)
put(bs, "B23", "TOTAL", F_BOLD)
put(bs, "D23", "=SUM(D16:D22)", fmt=CUR, bold=True, border=TOPLINE)
put(bs, "E23", "=SUM(E16:E22)", fmt=PCT, bold=True, border=TOPLINE)
put(bs, "I16", "Rounding adds a few dollars of profit, so 'This Bid $' total equals the contract.", F_NOTE)

put(bs, "B26", "CHECKS (all must read OK before the bid goes out)", F_SUB)
checks = [
    ("P&H crew mix = 100% and labour ties to wage breakdown", f"='LABOUR RATES'!C{ph_t + 3}"),
    ("HVAC crew mix = 100% and labour ties to wage breakdown", f"='LABOUR RATES'!C{hv_t + 3}"),
    ("P+H labour $ = LABOUR RATES P&H cost by wage breakdown", f"=IF(ABS(D6+D7-'LABOUR RATES'!G{ph_t})<1,\"OK\",\"CHECK\")"),
    ("V+AC labour $ = LABOUR RATES HVAC cost by wage breakdown", f"=IF(ABS(D8+D9-'LABOUR RATES'!G{hv_t})<1,\"OK\",\"CHECK\")"),
    ("Every estimate row has a valid Div and Phase",
     "=IF(SUMPRODUCT(('PLUMBING EST'!$P$7:$P$156<>0)*(ISNA(MATCH('PLUMBING EST'!$B$7:$B$156,INPUTS!$C$33:$C$36,0))+ISNA(MATCH('PLUMBING EST'!$C$7:$C$156,INPUTS!$B$58:$B$65,0))))"
     "+SUMPRODUCT(('HVAC EST'!$P$7:$P$156<>0)*(ISNA(MATCH('HVAC EST'!$B$7:$B$156,INPUTS!$C$33:$C$36,0))+ISNA(MATCH('HVAC EST'!$C$7:$C$156,INPUTS!$B$58:$B$65,0))))=0,\"OK\",\"CHECK\")"),
    ("Subtrades fully allocated to divisions",
     f"=IF(ABS('SUBS & GC'!H{SUB_LAST + 1}-SUM('SUBS & GC'!M{SUB_LAST + 1}:P{SUB_LAST + 1}))<1,\"OK\",\"CHECK\")"),
    ("General conditions fully allocated to divisions",
     f"=IF(ABS('SUBS & GC'!H{GC_LAST + 1}-SUM('SUBS & GC'!M{GC_LAST + 1}:P{GC_LAST + 1}))<1,\"OK\",\"CHECK\")"),
    ("Level weights total 100%", "=INPUTS!F55"),
    ("BUDGET total cost = BID SUMMARY total cost", "=IF(ABS(BUDGET!J11-K10)<1,\"OK\",\"CHECK\")"),
    ("SCHEDULE OF VALUES total = contract", "=IF(ABS('SCHEDULE OF VALUES'!C5-N10)<1,\"OK\",\"CHECK\")"),
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
setw(bu, {"A": 3, "B": 38, "C": 11, "D": 15, "E": 14, "F": 15, "G": 14, "H": 14, "I": 14, "J": 16, "K": 13, "L": 16})
title(bu, "PROJECT BUDGET (cost, by division and phase)")
put(bu, "B2", '=IF(INPUTS!C4="","",INPUTS!C4)&"   Quote: "&INPUTS!C5', F_LINK)
put(bu, "B4", "PROJECT ROLL-UP", F_SUB)
header_row(bu, 5, ["Division", "Hours", "Labour $", "Overhead $", "Material $ (incl consumables)", "Subtrades $",
                   "Gen. Cond. $", "Contingency $", "TOTAL COST $", "Margin $", "CONTRACT $"], start_col=2)
for i in range(4):
    r_ = 6 + i
    s = 6 + i
    put(bu, f"B{r_}", f"='BID SUMMARY'!B{s}", F_LINK, border=BOX)
    links = {"C": f"C{s}", "D": f"D{s}", "E": f"E{s}", "F": f"F{s}+'BID SUMMARY'!G{s}", "G": f"H{s}",
             "H": f"I{s}", "I": f"J{s}", "J": f"K{s}", "K": f"O{s}", "L": f"N{s}"}
    for col, ref in links.items():
        put(bu, f"{col}{r_}", f"='BID SUMMARY'!{ref}", F_LINK, fmt=NUM if col == "C" else CUR, border=BOX,
            fill=FILL_TOT if col in "JL" else None)
put(bu, "B10", "TOTAL", F_BOLD)
for col in "CDEFGHIJKL":
    put(bu, f"{col}10", f"=SUM({col}6:{col}9)", fmt=NUM if col == "C" else CUR, bold=True, border=TOPLINE)
# J11 = budget total built from the phase blocks (used for tie-out check)
put(bu, "B11", "Total cost rebuilt from phase detail below (must equal J10)", F_NOTE)

BLOCK = 17
B_START = 14
block_tot_refs = []
phase_cell = {}  # (div_idx, phase_idx) -> BUDGET!J cell
for d, (code, nm) in enumerate(DIVS):
    top = B_START + d * BLOCK
    put(bu, f"B{top}", f"{nm.upper()} ({code})", F_SUB, fill=FILL_SEC)
    for col in "CDEFGHIJ":
        bu[f"{col}{top}"].fill = FILL_SEC
    header_row(bu, top + 1, ["Phase", "Hours", "Labour $", "Overhead $", "Material $", "Subtrades $",
                             "Gen. Cond. $", "", "TOTAL BUDGET $", "% of Division"], start_col=2)
    for p in range(8):
        r_ = top + 2 + p
        ph = f"INPUTS!$B${58 + p}"
        put(bu, f"B{r_}", f"={ph}", F_LINK, border=BOX)
        put(bu, f"C{r_}", "=" + est_sumifs("M", f'"{code}"', ph), fmt=NUM, border=BOX)
        put(bu, f"D{r_}", "=" + est_sumifs("N", f'"{code}"', ph), fmt=CUR, border=BOX)
        put(bu, f"E{r_}", "=" + est_sumifs("O", f'"{code}"', ph), fmt=CUR, border=BOX)
        put(bu, f"F{r_}", "=" + est_sumifs("L", f'"{code}"', ph), fmt=CUR, border=BOX)
        dcol = sub_cols[d]
        put(bu, f"G{r_}", f"=SUMIFS('SUBS & GC'!${dcol}${SUB_FIRST}:${dcol}${SUB_LAST},'SUBS & GC'!$D${SUB_FIRST}:$D${SUB_LAST},{ph})",
            fmt=CUR, border=BOX)
        put(bu, f"H{r_}", f"=SUMIFS('SUBS & GC'!${dcol}${GC_FIRST}:${dcol}${GC_LAST},'SUBS & GC'!$D${GC_FIRST}:$D${GC_LAST},{ph})",
            fmt=CUR, border=BOX)
        put(bu, f"I{r_}", None, border=BOX)
        put(bu, f"J{r_}", f"=SUM(D{r_}:H{r_})", fmt=CUR, bold=True, border=BOX)
        put(bu, f"K{r_}", f"=IFERROR(J{r_}/$J${top + 12},0)", fmt=PCT, border=BOX)
        phase_cell[(d, p)] = f"BUDGET!$J${r_}"
    r_c = top + 10
    put(bu, f"B{r_c}", "Consumables allowance", border=BOX)
    put(bu, f"F{r_c}", f"='BID SUMMARY'!G{6 + d}", F_LINK, fmt=CUR, border=BOX)
    put(bu, f"J{r_c}", f"=F{r_c}", fmt=CUR, bold=True, border=BOX)
    put(bu, f"K{r_c}", f"=IFERROR(J{r_c}/$J${top + 12},0)", fmt=PCT, border=BOX)
    put(bu, f"B{r_c + 1}", "Contingency", border=BOX)
    put(bu, f"J{r_c + 1}", f"='BID SUMMARY'!J{6 + d}", F_LINK, fmt=CUR, bold=True, border=BOX)
    put(bu, f"K{r_c + 1}", f"=IFERROR(J{r_c + 1}/$J${top + 12},0)", fmt=PCT, border=BOX)
    r_t = top + 12
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
# SCHEDULE OF VALUES
# =====================================================================
sv = wb.create_sheet("SCHEDULE OF VALUES")
setw(sv, {"A": 11, "B": 46, "C": 16, "D": 11, "E": 16, "F": 11, "G": 16, "H": 14, "I": 3, "J": 12})
title(sv, "SCHEDULE OF VALUES")
put(sv, "A2", '=IF(INPUTS!C4="","",INPUTS!C4)&"   "&INPUTS!C11&", "&INPUTS!C12&"   GC: "&INPUTS!C8', F_LINK)
put(sv, "B4", "Contract value (from BID SUMMARY)", F_BOLD)
put(sv, "C4", "='BID SUMMARY'!N10", F_LINK, fmt=CUR, bold=True)
put(sv, "B5", "Total of schedule below", F_BOLD)
put(sv, "D5", '=IF(ABS(C5-C4)<1,"OK","CHECK")', bold=True)
ok_cf(sv, "D5")
put(sv, "B6", "Margin, consumables and contingency are spread across each division's lines in proportion to cost.", F_NOTE)
header_row(sv, 8, ["Item #", "Description", "Scheduled Value $", "% of Contract", "Completed to Date $",
                   "% Complete", "Balance to Finish $", "Holdback $", "", "Div factor"])
r_ = 9
div_header_rows = []
for d, (code, nm) in enumerate(DIVS):
    hdr = r_
    div_header_rows.append(hdr)
    put(sv, f"A{hdr}", code, F_BOLD, fill=FILL_SEC)
    put(sv, f"B{hdr}", f"{nm.upper()}", F_BOLD, fill=FILL_SEC)
    phase_sum = "+".join(phase_cell[(d, p)] for p in range(8))
    put(sv, f"J{hdr}", f"=IFERROR('BID SUMMARY'!N{6 + d}/({phase_sum}),0)", fmt="0.0000", fill=FILL_SEC)
    note(sv, f"J{hdr}", "Contract for this division / sum of its phase costs. Spreads margin and allowances evenly.")
    r_ += 1
    first_line = r_
    n = 0
    for p in range(8):
        if p in (2, 3):
            for li, (tag, floor, _) in enumerate(LEVELS):
                n += 1
                put(sv, f"A{r_}", f"{code}-{p + 1:02d}-{li + 1:02d}", border=BOX)
                word = "Rough-in" if p == 2 else "Finishing"
                put(sv, f"B{r_}", f'="{code} - "&INPUTS!$B${40 + li}&" {word}"', border=BOX)
                put(sv, f"C{r_}", f"={phase_cell[(d, p)]}*INPUTS!$E${40 + li}*$J${hdr}", fmt=CUR, border=BOX)
                r_ += 1
        else:
            n += 1
            put(sv, f"A{r_}", f"{code}-{p + 1:02d}", border=BOX)
            put(sv, f"B{r_}", f'="{code} - "&MID(INPUTS!$B${58 + p},4,60)', border=BOX)
            put(sv, f"C{r_}", f"={phase_cell[(d, p)]}*$J${hdr}", fmt=CUR, border=BOX)
            r_ += 1
    last_line = r_ - 1
    put(sv, f"C{hdr}", f"=SUM(C{first_line}:C{last_line})", fmt=CUR, bold=True, fill=FILL_SEC)
    put(sv, f"D{hdr}", f"=IFERROR(C{hdr}/$C$4,0)", fmt=PCT, bold=True, fill=FILL_SEC)
    put(sv, f"E{hdr}", f"=SUM(E{first_line}:E{last_line})", fmt=CUR, bold=True, fill=FILL_SEC)
    put(sv, f"G{hdr}", f"=SUM(G{first_line}:G{last_line})", fmt=CUR, bold=True, fill=FILL_SEC)
    put(sv, f"H{hdr}", f"=SUM(H{first_line}:H{last_line})", fmt=CUR, bold=True, fill=FILL_SEC)
    for k in range(first_line, last_line + 1):
        put(sv, f"D{k}", f"=IFERROR(C{k}/$C$4,0)", fmt=PCT, border=BOX)
        put(sv, f"E{k}", None, F_IN, fmt=CUR, fill=FILL_IN, border=BOX)
        put(sv, f"F{k}", f"=IFERROR(E{k}/C{k},0)", fmt=PCT, border=BOX)
        put(sv, f"G{k}", f"=C{k}-E{k}", fmt=CUR, border=BOX)
        put(sv, f"H{k}", f"=E{k}*INPUTS!$C$29", fmt=CUR, border=BOX)
    r_ += 1
put(sv, f"B{r_}", "TOTAL CONTRACT", F_BOLD)
for col in "CEGH":
    put(sv, f"{col}{r_}", "=" + "+".join(f"{col}{h}" for h in div_header_rows), fmt=CUR, bold=True, fill=FILL_TOT, border=TOPLINE)
put(sv, f"D{r_}", f"=IFERROR(C{r_}/$C$4,0)", fmt=PCT, bold=True, border=TOPLINE)
put(sv, f"F{r_}", f"=IFERROR(E{r_}/C{r_},0)", fmt=PCT, bold=True, border=TOPLINE)
put(sv, "C5", f"=C{r_}", fmt=CUR, bold=True)
sv.freeze_panes = "C9"
sv.auto_filter.ref = f"A8:H{r_ - 1}"

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
     "Bid labour = hours x $70.20, which is a simple average of the 5 rates (D42 = SUM(D36:D40)/5). Crew mix % is ignored.",
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
     "HVAC wage breakdown is wrong as soon as HVAC hours exist (it would show 5 x the apprentice hours).",
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
     "Supervision would be charged for 22 months instead of 24 (or whatever the overhead rate is).",
     "Supervision qty links to INPUTS HVAC duration."),
    ("MED", "PRJ SUMMARY row 59 (Supervision - Office)",
     "$7,500/month x 24 months = $180,000 (+15% = $207,000) is allocated 100% to Plumbing, using the HVAC duration.",
     "Plumbing looks $207,000 more expensive than it is; HVAC looks cheaper. Confirmed this is a real job cost (dedicated PM), separate from the $22/hr overhead.",
     "Kept as 'Office Project Manager'. Auto-splits by direct cost across divisions (or set your own %)."),
    ("HIGH", "PRJ SUMMARY M55:T79 (\"50-50 Div's\" split)",
     "GC split formulas divide by C4+C5 with no IFERROR. With no HVAC entered they return a divide-by-zero (DIV/0) error.",
     "PRJ SUMMARY F6, G6, J6, W6 and J7 all show DIV/0 errors, so the grand total breaks on any plumbing-only job.",
     "Allocation uses % inputs or an IFERROR-guarded direct-cost share. No divide-by-zero."),
    ("MED", "S.O.V B1, AC1, M2:M6, rows 4-5; KNOWIFY F1",
     "Broken references (REF errors from deleted cells) in the S.O.V; the Knowify 'ALL GOOD' indicator therefore shows a REF error.",
     "The SOV check light can never go green, so it cannot be trusted.",
     "New SCHEDULE OF VALUES built from BUDGET with a contract tie-out check."),
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
     "The cost box on those tabs under-reports if finishing misc hours are used (summary row 2 is fine).",
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
    vals = [i + 1, sev, where, what, why, fix]
    for j, v in enumerate(vals):
        c = put(au, f"{get_column_letter(1 + j)}{r_}", v, border=BOX,
                align=Alignment(wrap_text=True, vertical="top", horizontal="center" if j < 2 else "left"))
    au[f"B{r_}"].font = Font(name=FONT, size=10, bold=True,
                             color={"HIGH": "C00000", "MED": "C65911", "LOW": "7F7F7F", "OK": "006100"}[sev])
q0 = 5 + len(FIND) + 2
put(au, f"B{q0}", "OPEN QUESTIONS FOR YOU", F_SUB)
QS = [
    "Target margin per division (placeholder is 15%). Do subs carry a lower margin than self-perform work?",
    "Your typical COGS % (labour, material, subs, GC, profit) to fill INPUTS rows 69-75.",
    "What does a Knowify job cost export look like (column names)? Item Codes on the estimate can match your Knowify cost codes.",
    "Should the SOV split rough-in and finishing by floor (current), or by phase only?",
    "Do the price-list and takeoff tabs (P&H MATERIAL, HVAC MATERIAL, EQUIPMENT) need to come across, or will takeoff totals be typed in as Qty x unit cost?",
]
for i, q in enumerate(QS):
    put(au, f"B{q0 + 1 + i}", i + 1, align=Alignment(horizontal="center", vertical="top"))
    put(au, f"C{q0 + 1 + i}", q, align=Alignment(wrap_text=True, vertical="top"))
    au.merge_cells(f"C{q0 + 1 + i}:F{q0 + 1 + i}")
    au.row_dimensions[q0 + 1 + i].height = 28
a0 = q0 + len(QS) + 2
put(au, f"B{a0}", "ANSWERED (decisions built into this workbook)", F_SUB)
ANS = [
    "$22/hr overhead covers all salaried staff company wide. A dedicated office project manager per job is a separate monthly cost, so the Office Project Manager GC line stays.",
    "Field supervision GC lines are not double counted with the Foreman & PM share of the crew mix. Both stay.",
    "Labour rates on LABOUR RATES are fully burdened (CPP, EI, WCB, vacation, benefits). No extra burden % is added.",
]
for i, a in enumerate(ANS):
    put(au, f"B{a0 + 1 + i}", "OK", Font(name=FONT, size=10, bold=True, color="006100"), align=Alignment(horizontal="center", vertical="top"))
    put(au, f"C{a0 + 1 + i}", a, align=Alignment(wrap_text=True, vertical="top"))
    au.merge_cells(f"C{a0 + 1 + i}:F{a0 + 1 + i}")
    au.row_dimensions[a0 + 1 + i].height = 28
au.freeze_panes = "A5"

# ---------- workbook-wide settings ----------
order = ["READ ME", "INPUTS", "LABOUR RATES", "PLUMBING EST", "HVAC EST", "SUBS & GC", "CRM ACTUALS",
         "BID SUMMARY", "BUDGET", "SCHEDULE OF VALUES", "AUDIT NOTES"]
wb._sheets = [wb[n] for n in order]
tab_colors = {"READ ME": "7F7F7F", "INPUTS": "FFC000", "LABOUR RATES": "FFC000", "PLUMBING EST": "2F5597",
              "HVAC EST": "2F5597", "SUBS & GC": "2F5597", "CRM ACTUALS": "70AD47", "BID SUMMARY": "C00000",
              "BUDGET": "548235", "SCHEDULE OF VALUES": "548235", "AUDIT NOTES": "7F7F7F"}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = tab_colors[ws.title]
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False
wb.active = wb.sheetnames.index("BID SUMMARY")
wb.save(OUT)
print("saved", OUT)
