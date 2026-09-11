#!/usr/bin/env python3
"""Build Northline_Board_ELT_Flash.xlsx — one-page ELT / board flash from the monthly close."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

yellow = PatternFill("solid", fgColor="FFF2CC")
header_fill = PatternFill("solid", fgColor="1F4E79")
section_fill = PatternFill("solid", fgColor="D6E3F0")
green_fill = PatternFill("solid", fgColor="C6EFCE")
amber_fill = PatternFill("solid", fgColor="FFE699")
red_fill = PatternFill("solid", fgColor="F8CBAD")
tile_fill = PatternFill("solid", fgColor="E9EDF4")
light_gray = PatternFill("solid", fgColor="F5F5F5")
dark_navy = PatternFill("solid", fgColor="0D2B4A")
white_fill = PatternFill("solid", fgColor="FFFFFF")
accent_fill = PatternFill("solid", fgColor="2E75B6")

input_font = Font(name="Calibri", size=11, color="0000FF")
black = Font(name="Calibri", size=11, color="000000")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
title_font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
section_font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
bold = Font(name="Calibri", size=11, bold=True)
bold_black = Font(name="Calibri", size=11, bold=True, color="000000")
italic_grey = Font(name="Calibri", size=10, italic=True, color="666666")
small_grey = Font(name="Calibri", size=9, italic=True, color="666666")
tile_title = Font(name="Calibri", size=9, bold=True, color="1F4E79")
tile_value = Font(name="Calibri", size=14, bold=True, color="000000")
white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
big_title = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
subtitle_w = Font(name="Calibri", size=11, color="D6E3F0")

thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
med = Border(
    left=Side(style="medium", color="1F4E79"),
    right=Side(style="medium", color="1F4E79"),
    top=Side(style="medium", color="1F4E79"),
    bottom=Side(style="medium", color="1F4E79"),
)

money = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
pct = "0.0%"
num = "#,##0.0"

COVER = "00_Cover"
ASSUMP = "01_Assumptions"
PL = "02_P&L_Flash"
KPI = "03_KPIs"
CALL = "04_Callouts"
ONE = "05_One_Pager"
DICT = "06_Data_Dictionary"

MONTHS = [
    "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26",
    "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26",
]

# Close month = Aug-26 (month index 8) — mid-year flash for ELT
CLOSE_MONTH = 8  # 1-indexed Aug

# P&L lines ($000s) — Actual / Budget / Forecast for close month & YTD
# Structure: Net Revenue, COGS, Gross Profit, SG&A, Marketing, R&D, EBITDA, D&A, EBIT, Interest, EBT, Tax, Net Income
PL_LINES = [
    ("Net revenue", True),
    ("Cost of goods sold", False),
    ("Gross profit", True),
    ("  SG&A", False),
    ("  Marketing & trade", False),
    ("  R&D / innovation", False),
    ("Total opex", False),
    ("EBITDA", True),
    ("  D&A", False),
    ("EBIT", True),
    ("  Interest expense", False),
    ("EBT", True),
    ("  Tax", False),
    ("Net income", True),
]

# Month-of close (Aug): Actual, Budget, Forecast
# YTD Actual, YTD Budget, YTD Forecast (through Aug)
# Values in $000s
PL_DATA = {
    #                         Mo Act   Mo Bud   Mo Fcst  YTD Act  YTD Bud  YTD Fcst
    "Net revenue":           (80500,  79800,   81000,   604100, 598800,  606500),
    "Cost of goods sold":    (46370,  45486,   46200,   348500, 343200,  347800),
    "  SG&A":                (15050,  14763,   14900,   113200, 110800,  112400),
    "  Marketing & trade":   ( 6450,   6384,    6500,    48600,  47900,   49000),
    "  R&D / innovation":    ( 2415,   2394,    2400,    18200,  17950,   18100),
    "  D&A":                 ( 3220,   3192,    3200,    24160,  23940,   24080),
    "  Interest expense":    ( 1610,   1596,    1600,    12080,  11970,   12040),
    "  Tax":                 ( 2690,   2895,    2950,    22300,  23760,   24100),
}


def style_input(cell):
    cell.fill = yellow
    cell.font = input_font
    cell.border = thin
    cell.alignment = Alignment(horizontal="center")


def style_formula(cell, key=False):
    cell.font = bold_black if key else black
    cell.border = thin
    cell.alignment = Alignment(horizontal="center")
    if key:
        cell.fill = green_fill


def style_header_cell(cell, value):
    cell.value = value
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    cell.border = thin


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def landscape(ws, fit_height=1):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.sheet_view.showGridLines = False
    ws.page_setup.horizontalCentered = True
    ws.oddFooter.left.text = "Northline Consumer Products  |  fictional sample  |  $000s"
    ws.oddFooter.right.text = "CONFIDENTIAL — portfolio sample"


wb = Workbook()

# ═══════════════════════════════════════════════════════════════════
# 00_Cover
# ═══════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = COVER
landscape(ws)
set_col_widths(ws, [3, 22, 55, 18, 18])

ws.merge_cells("B2:E2")
ws["B2"] = "NORTHLINE CONSUMER PRODUCTS"
ws["B2"].font = title_font
ws.merge_cells("B3:E3")
ws["B3"] = "Board / ELT Flash — Monthly Close One-Pager"
ws["B3"].font = Font(name="Calibri", size=14, bold=True, color="2E75B6")
ws.merge_cells("B4:E4")
ws["B4"] = "Fictional CPG portfolio sample  ·  figures in $000s  ·  formulas only (no VBA)"
ws["B4"].font = italic_grey

ws["B6"] = "Purpose"
ws["B6"].font = section_font
ws.merge_cells("B7:E7")
ws["B7"] = (
    "A print-ready ELT flash you can put in front of the executive team five minutes after close. "
    "Actual vs Budget vs Forecast for the close month and YTD, board KPIs with RAG, and structured callouts."
)
ws["B7"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[7].height = 40

ws["B9"] = "How to use"
ws["B9"].font = section_font
steps = [
    "1. Open 01_Assumptions — set the close month and edit headline drivers (yellow / blue).",
    "2. Review 02_P&L_Flash — Actual / Budget / Forecast for month and YTD ($ and %).",
    "3. Check 03_KPIs — 7 board KPIs with RAG vs plan.",
    "4. Update 04_Callouts — yellow commentary boxes for the ELT narrative.",
    "5. Print / PDF 05_One_Pager — the money tab. One landscape page for the ELT pack.",
]
for i, s in enumerate(steps):
    ws.cell(10 + i, 2, s).font = black
    ws.merge_cells(start_row=10 + i, start_column=2, end_row=10 + i, end_column=5)

ws["B16"] = "Color convention"
ws["B16"].font = section_font
ws["B17"] = "Yellow fill + blue font"
style_input(ws["B17"])
ws["C17"] = "= inputs (edit these)"
ws["C17"].font = black
ws["B18"] = "Black font"
style_formula(ws["B18"])
ws["C18"] = "= formulas (do not overwrite)"
ws["C18"].font = black
ws["B19"] = "Green / Amber / Red"
ws["B19"].fill = green_fill
ws["B19"].border = thin
ws["C19"] = "= RAG status (formula-driven)"
ws["C19"].font = black

ws["B21"] = "Tabs"
ws["B21"].font = section_font
tabs = [
    (COVER, "Purpose and navigation"),
    (ASSUMP, "Close month, headlines, RAG bands"),
    (PL, "P&L Actual / Budget / Forecast — $ and %"),
    (KPI, "Board KPIs with RAG"),
    (CALL, "ELT commentary inputs"),
    (ONE, "Print-ready one-pager (the money tab)"),
    (DICT, "Field definitions"),
]
for i, (t, d) in enumerate(tabs):
    r = 22 + i
    ws.cell(r, 2, t).font = bold
    ws.cell(r, 2).border = thin
    ws.cell(r, 3, d).border = thin
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)

ws["B30"] = "Sai Siri Bandaru — Financial Analyst | FP&A | forecasting, variance analysis, Excel"
ws["B30"].font = small_grey
ws["B31"] = "All sample numbers are fictional. No employer data."
ws["B31"].font = small_grey

# ═══════════════════════════════════════════════════════════════════
# 01_Assumptions
# ═══════════════════════════════════════════════════════════════════
ws = wb.create_sheet(ASSUMP)
landscape(ws)
set_col_widths(ws, [3, 28, 14, 14, 14, 14, 14, 14, 40])

ws["B2"] = "01  ·  Assumptions"
ws["B2"].font = title_font
ws["B3"] = "Yellow cells are inputs. Close month drives INDEX lookups on the P&L and one-pager."
ws["B3"].font = italic_grey

ws["B5"] = "Reporting"
ws["B5"].font = section_font
ws["B5"].fill = section_fill
ws.merge_cells("B5:C5")

ws["B6"] = "Close month (1–12)"
ws["C6"] = CLOSE_MONTH
style_input(ws["C6"])
ws["D6"] = '=IF(OR(C6<1,C6>12),"Check month",INDEX({"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},C6)&"-26")'
style_formula(ws["D6"])
ws["E6"] = "← displayed label"
ws["E6"].font = small_grey

ws["B7"] = "Fiscal year"
ws["C7"] = "FY26"
style_input(ws["C7"])
ws["B8"] = "Company"
ws["C8"] = "Northline Consumer Products"
style_input(ws["C8"])
ws.merge_cells("C8:E8")
ws["B9"] = "Units"
ws["C9"] = "$000s"
style_input(ws["C9"])
ws["B10"] = "Prepared by"
ws["C10"] = "FP&A"
style_input(ws["C10"])
ws["B11"] = "As-of date"
ws["C11"] = "2026-09-03"
style_input(ws["C11"])

ws["B13"] = "Headlines (edit for the pack)"
ws["B13"].font = section_font
ws["B13"].fill = section_fill
ws.merge_cells("B13:E13")

headlines = [
    ("Headline 1", "Revenue +0.9% vs budget in Aug; YTD +0.9% — volume and mix both ahead."),
    ("Headline 2", "Gross margin 42.4% vs 43.0% budget — commodity and freight pressure, partially offset by price."),
    ("Headline 3", "EBITDA $11.2M vs $10.8M budget (+4.1%) — opex discipline and lower scrap."),
    ("Headline 4", "Cash conversion solid; DSO 41 days (−1 vs plan). Capex on track."),
    ("Watch item", "Q4 promo calendar heavy — protect AUR and trade spend ROI."),
]
for i, (label, text) in enumerate(headlines):
    r = 14 + i
    ws.cell(r, 2, label).font = bold
    ws.cell(r, 2).border = thin
    cell = ws.cell(r, 3, text)
    style_input(cell)
    cell.alignment = Alignment(wrap_text=True, horizontal="left")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 28

ws["B20"] = "RAG bands (board KPIs)"
ws["B20"].font = section_font
ws["B20"].fill = section_fill
ws.merge_cells("B20:E20")

for i, h in enumerate(["KPI", "Direction", "Amber band", "Green band", "Notes"]):
    style_header_cell(ws.cell(21, 2 + i), h)

kpi_bands = [
    ("Net revenue vs budget", "H", 0.02, 0.00, "Green if ≥ budget; Amber if within −2%"),
    ("Gross margin %", "H", 0.015, 0.005, "pp vs budget"),
    ("EBITDA vs budget", "H", 0.05, 0.00, "Green if ≥ budget"),
    ("Volume (000 cases)", "H", 0.03, 0.00, "vs budget"),
    ("OTIF %", "H", 0.02, 0.01, "On-time in-full"),
    ("DSO (days)", "L", 0.05, 0.02, "Lower better"),
    ("Operating cash flow", "H", 0.08, 0.00, "vs budget"),
]
for i, (name, direction, amber, green, notes) in enumerate(kpi_bands):
    r = 22 + i
    ws.cell(r, 2, name).border = thin
    c = ws.cell(r, 3, direction)
    style_input(c)
    c = ws.cell(r, 4, amber)
    style_input(c)
    c.number_format = pct
    c = ws.cell(r, 5, green)
    style_input(c)
    c.number_format = pct
    ws.cell(r, 6, notes).font = small_grey
    ws.cell(r, 6).border = thin

ws["B30"] = "Direction: H = higher-better, L = lower-better. Bands are relative to plan."
ws["B30"].font = small_grey

# ═══════════════════════════════════════════════════════════════════
# 02_P&L_Flash
# ═══════════════════════════════════════════════════════════════════
ws = wb.create_sheet(PL)
landscape(ws)
set_col_widths(ws, [3, 24, 12, 12, 12, 10, 10, 12, 12, 12, 10, 10, 3, 14])

ws["B2"] = "02  ·  P&L Flash"
ws["B2"].font = title_font
ws["B3"] = "='01_Assumptions'!D6&\" close  ·  Actual / Budget / Forecast  ·  $000s\""
ws["B3"].font = italic_grey

# Headers
ws.merge_cells("C5:G5")
ws["C5"] = "Close month"
ws["C5"].font = header_font
ws["C5"].fill = header_fill
ws["C5"].alignment = Alignment(horizontal="center")
for col in range(3, 8):
    ws.cell(5, col).fill = header_fill
    ws.cell(5, col).border = thin

ws.merge_cells("H5:L5")
ws["H5"] = "YTD"
ws["H5"].font = header_font
ws["H5"].fill = accent_fill
ws["H5"].alignment = Alignment(horizontal="center")
for col in range(8, 13):
    ws.cell(5, col).fill = accent_fill
    ws.cell(5, col).border = thin

headers2 = ["Line", "Actual", "Budget", "Forecast", "A vs B %", "A vs F %",
            "Actual", "Budget", "Forecast", "A vs B %", "A vs F %"]
for i, h in enumerate(headers2):
    style_header_cell(ws.cell(6, 2 + i), h)
ws.row_dimensions[6].height = 30

# Row map for formula references
# Row 7 Net revenue (inputs)
# Row 8 COGS (inputs)
# Row 9 Gross profit (formulas)
# Row 10 SG&A, 11 Mkt, 12 R&D (inputs)
# Row 13 Total opex (sum)
# Row 14 EBITDA
# Row 15 D&A
# Row 16 EBIT
# Row 17 Interest
# Row 18 EBT
# Row 19 Tax
# Row 20 Net income

# Helper to write a P&L input row
def write_pl_input(ws, row, label, mo_a, mo_b, mo_f, ytd_a, ytd_b, ytd_f, indent=False):
    cell = ws.cell(row, 2, label)
    cell.font = bold if not indent else black
    cell.border = thin
    vals = [mo_a, mo_b, mo_f]
    for i, v in enumerate(vals):
        c = ws.cell(row, 3 + i, v)
        style_input(c)
        c.number_format = money
    # A vs B %, A vs F %
    c = ws.cell(row, 6, f"=IF(D{row}=0,0,(C{row}-D{row})/ABS(D{row}))")
    style_formula(c)
    c.number_format = pct
    c = ws.cell(row, 7, f"=IF(E{row}=0,0,(C{row}-E{row})/ABS(E{row}))")
    style_formula(c)
    c.number_format = pct
    vals_y = [ytd_a, ytd_b, ytd_f]
    for i, v in enumerate(vals_y):
        c = ws.cell(row, 8 + i, v)
        style_input(c)
        c.number_format = money
    c = ws.cell(row, 11, f"=IF(I{row}=0,0,(H{row}-I{row})/ABS(I{row}))")
    style_formula(c)
    c.number_format = pct
    c = ws.cell(row, 12, f"=IF(J{row}=0,0,(H{row}-J{row})/ABS(J{row}))")
    style_formula(c)
    c.number_format = pct


def write_pl_formula(ws, row, label, formula_cols, key=False):
    """formula_cols is dict col -> formula string for cols 3,4,5,8,9,10; pct cols auto."""
    cell = ws.cell(row, 2, label)
    cell.font = bold_black
    cell.border = thin
    if key:
        cell.fill = green_fill
    for col, f in formula_cols.items():
        c = ws.cell(row, col, f)
        style_formula(c, key=key)
        c.number_format = money
    for col_a, col_b, col_pct in [(3, 4, 6), (3, 5, 7), (8, 9, 11), (8, 10, 12)]:
        c = ws.cell(row, col_pct, f"=IF({get_column_letter(col_b)}{row}=0,0,({get_column_letter(col_a)}{row}-{get_column_letter(col_b)}{row})/ABS({get_column_letter(col_b)}{row}))")
        style_formula(c, key=key)
        c.number_format = pct


# Row 7 — Net revenue
write_pl_input(ws, 7, "Net revenue", *PL_DATA["Net revenue"])
# Row 8 — COGS
write_pl_input(ws, 8, "Cost of goods sold", *PL_DATA["Cost of goods sold"], indent=True)
# Row 9 — Gross profit
write_pl_formula(ws, 9, "Gross profit", {
    3: "=C7-C8", 4: "=D7-D8", 5: "=E7-E8",
    8: "=H7-H8", 9: "=I7-I8", 10: "=J7-J8",
}, key=True)
# GM%
ws["N5"] = "Gross margin %"
ws["N5"].font = section_font
ws["N6"] = "Mo Act"
style_header_cell(ws["N6"], "Mo Act GM%")
ws["N7"] = "=IF(C7=0,0,C9/C7)"
style_formula(ws["N7"], key=True)
ws["N7"].number_format = pct
ws["N8"] = "Mo Bud GM%"
ws["N8"].font = small_grey
ws["N9"] = "=IF(D7=0,0,D9/D7)"
style_formula(ws["N9"])
ws["N9"].number_format = pct

# Row 10-12 opex inputs
write_pl_input(ws, 10, "  SG&A", *PL_DATA["  SG&A"], indent=True)
write_pl_input(ws, 11, "  Marketing & trade", *PL_DATA["  Marketing & trade"], indent=True)
write_pl_input(ws, 12, "  R&D / innovation", *PL_DATA["  R&D / innovation"], indent=True)
# Row 13 Total opex
write_pl_formula(ws, 13, "Total opex", {
    3: "=C10+C11+C12", 4: "=D10+D11+D12", 5: "=E10+E11+E12",
    8: "=H10+H11+H12", 9: "=I10+I11+I12", 10: "=J10+J11+J12",
})
# Row 14 EBITDA = GP - opex
write_pl_formula(ws, 14, "EBITDA", {
    3: "=C9-C13", 4: "=D9-D13", 5: "=E9-E13",
    8: "=H9-H13", 9: "=I9-I13", 10: "=J9-J13",
}, key=True)
# Row 15 D&A
write_pl_input(ws, 15, "  D&A", *PL_DATA["  D&A"], indent=True)
# Row 16 EBIT
write_pl_formula(ws, 16, "EBIT", {
    3: "=C14-C15", 4: "=D14-D15", 5: "=E14-E15",
    8: "=H14-H15", 9: "=I14-I15", 10: "=J14-J15",
}, key=True)
# Row 17 Interest
write_pl_input(ws, 17, "  Interest expense", *PL_DATA["  Interest expense"], indent=True)
# Row 18 EBT
write_pl_formula(ws, 18, "EBT", {
    3: "=C16-C17", 4: "=D16-D17", 5: "=E16-E17",
    8: "=H16-H17", 9: "=I16-I17", 10: "=J16-J17",
}, key=True)
# Row 19 Tax
write_pl_input(ws, 19, "  Tax", *PL_DATA["  Tax"], indent=True)
# Row 20 Net income
write_pl_formula(ws, 20, "Net income", {
    3: "=C18-C19", 4: "=D18-D19", 5: "=E18-E19",
    8: "=H18-H19", 9: "=I18-I19", 10: "=J18-J19",
}, key=True)

# Margin bridge note
ws["B22"] = "Margins (formula)"
ws["B22"].font = section_font
ws["B23"] = "EBITDA margin % (Mo Act)"
ws["C23"] = "=IF(C7=0,0,C14/C7)"
style_formula(ws["C23"], key=True)
ws["C23"].number_format = pct
ws["D23"] = "vs Bud"
ws["E23"] = "=IF(D7=0,0,D14/D7)"
style_formula(ws["E23"])
ws["E23"].number_format = pct
ws["B24"] = "EBITDA margin % (YTD Act)"
ws["C24"] = "=IF(H7=0,0,H14/H7)"
style_formula(ws["C24"], key=True)
ws["C24"].number_format = pct
ws["D24"] = "vs Bud"
ws["E24"] = "=IF(I7=0,0,I14/I7)"
style_formula(ws["E24"])
ws["E24"].number_format = pct

ws["B26"] = "Yellow = paste from the close pack. Green rows = key subtotals (formulas)."
ws["B26"].font = small_grey

# Conditional formatting for variance %
for col in [6, 7, 11, 12]:
    letter = get_column_letter(col)
    ws.conditional_formatting.add(
        f"{letter}7:{letter}20",
        CellIsRule(operator="greaterThanOrEqual", formula=["0.005"], fill=green_fill),
    )
    ws.conditional_formatting.add(
        f"{letter}7:{letter}20",
        CellIsRule(operator="between", formula=["-0.02", "0.0049"], fill=amber_fill),
    )
    ws.conditional_formatting.add(
        f"{letter}7:{letter}20",
        CellIsRule(operator="lessThan", formula=["-0.02"], fill=red_fill),
    )

# ═══════════════════════════════════════════════════════════════════
# 03_KPIs
# ═══════════════════════════════════════════════════════════════════
ws = wb.create_sheet(KPI)
landscape(ws)
set_col_widths(ws, [3, 26, 12, 12, 12, 12, 10, 10, 14, 40])

ws["B2"] = "03  ·  Board KPIs"
ws["B2"].font = title_font
ws["B3"] = "Seven ELT-facing KPIs for the close month. RAG is formula-driven from Assumptions bands."
ws["B3"].font = italic_grey

for i, h in enumerate(["KPI", "Actual", "Budget", "Var $ / pp", "Var %", "RAG", "YTD Act", "YTD Bud", "Commentary"]):
    style_header_cell(ws.cell(5, 2 + i), h)
ws.row_dimensions[5].height = 28

# KPI data: name, actual, budget, ytd_act, ytd_bud, unit, direction_row in assumptions (22-28), commentary
# Link RAG to assumptions bands where possible
kpis = [
    # name, act, bud, ytd_a, ytd_b, fmt, higher_better, commentary
    ("Net revenue ($000s)", 80500, 79800, 604100, 598800, money, True,
     "Volume + mix ahead; AUR slightly soft on promo."),
    ("Gross margin %", 0.424, 0.430, 0.423, 0.427, pct, True,
     "Commodity + freight; price/mix partially offsets."),
    ("EBITDA ($000s)", 11215, 10773, 84100, 82150, money, True,
     "Opex discipline and lower scrap vs plan."),
    ("Volume (000 cases)", 6720, 6650, 50340, 49900, num, True,
     "Club and e-comm channels lead."),
    ("OTIF %", 0.961, 0.970, 0.966, 0.970, pct, True,
     "West DC labor constraints in week 3."),
    ("DSO (days)", 41.2, 42.0, 41.2, 42.0, num, False,
     "Collections strong; retail terms stable."),
    ("Op. cash flow ($000s)", 6200, 5800, 47300, 45100, money, True,
     "Working capital release + EBITDA beat."),
]

for i, (name, act, bud, ytd_a, ytd_b, fmt, higher, comment) in enumerate(kpis):
    r = 6 + i
    band_row = 22 + i  # assumptions RAG rows
    ws.cell(r, 2, name).font = bold
    ws.cell(r, 2).border = thin

    c = ws.cell(r, 3, act)
    style_input(c)
    c.number_format = fmt

    c = ws.cell(r, 4, bud)
    style_input(c)
    c.number_format = fmt

    # Var $ / pp
    c = ws.cell(r, 5, f"=C{r}-D{r}")
    style_formula(c)
    c.number_format = fmt

    # Var %
    c = ws.cell(r, 6, f"=IF(D{r}=0,0,(C{r}-D{r})/ABS(D{r}))")
    style_formula(c)
    c.number_format = pct

    # RAG formula
    if higher:
        rag = (
            f'=IF(C{r}>=D{r},"Green",'
            f'IF(C{r}>=D{r}*(1-\'01_Assumptions\'!D{band_row}),"Amber","Red"))'
        )
    else:
        rag = (
            f'=IF(C{r}<=D{r},"Green",'
            f'IF(C{r}<=D{r}*(1+\'01_Assumptions\'!D{band_row}),"Amber","Red"))'
        )
    c = ws.cell(r, 7, rag)
    style_formula(c, key=True)

    c = ws.cell(r, 8, ytd_a)
    style_input(c)
    c.number_format = fmt
    c = ws.cell(r, 9, ytd_b)
    style_input(c)
    c.number_format = fmt

    c = ws.cell(r, 10, comment)
    style_input(c)
    c.alignment = Alignment(wrap_text=True, horizontal="left")
    ws.row_dimensions[r].height = 30

# RAG conditional formatting
ws.conditional_formatting.add(
    "G6:G12",
    FormulaRule(formula=['G6="Green"'], fill=green_fill),
)
ws.conditional_formatting.add(
    "G6:G12",
    FormulaRule(formula=['G6="Amber"'], fill=amber_fill),
)
ws.conditional_formatting.add(
    "G6:G12",
    FormulaRule(formula=['G6="Red"'], fill=red_fill),
)

ws["B14"] = "Summary counts"
ws["B14"].font = section_font
ws["B15"] = "Green"
ws["C15"] = '=COUNTIF(G6:G12,"Green")'
style_formula(ws["C15"], key=True)
ws["D15"] = "Amber"
ws["E15"] = '=COUNTIF(G6:G12,"Amber")'
style_formula(ws["E15"])
ws["E15"].fill = amber_fill
ws["F15"] = "Red"
ws["G15"] = '=COUNTIF(G6:G12,"Red")'
style_formula(ws["G15"])
ws["G15"].fill = red_fill

ws["B17"] = "Edit Actual / Budget (yellow). RAG and variance recalculate automatically."
ws["B17"].font = small_grey

# ═══════════════════════════════════════════════════════════════════
# 04_Callouts
# ═══════════════════════════════════════════════════════════════════
ws = wb.create_sheet(CALL)
landscape(ws)
set_col_widths(ws, [3, 22, 70, 18])

ws["B2"] = "04  ·  Callouts"
ws["B2"].font = title_font
ws["B3"] = "Structured commentary for the ELT pack. Everything yellow is an input — keep it tight."
ws["B3"].font = italic_grey

sections = [
    ("What went well", [
        "EBITDA beat budget by $0.4M in-month; YTD beat $2.0M.",
        "DSO improved 0.8 days vs plan; collections cadence holding.",
        "Club channel volume +6% vs budget; e-comm contribution margin expanding.",
    ]),
    ("What needs attention", [
        "Gross margin −0.6 pp vs budget — resin and inbound freight.",
        "OTIF 96.1% vs 97.0% target — West DC temp labor gap in week 3.",
        "Trade spend pacing ahead of plan in Club; ROI review in flight.",
    ]),
    ("Decisions / asks", [
        "Approve Q4 promo guardrails (AUR floor + trade ROI hurdle).",
        "Authorize West DC overtime / temp surge through peak (est. $180k).",
        "Confirm commodity hedge window for Q1 resin (Treasury + Procurement).",
    ]),
    ("Risks & mitigations", [
        "Q4 promo intensity — mitigate with SKU-level AUR floors and sell-in caps.",
        "Freight lane volatility — dual-source West inbound; lock 60-day rates.",
        "Retailer inventory rightsizing — lean into high-velocity SKUs; cut slow movers.",
    ]),
    ("Forward look (next 90 days)", [
        "Sep flash: hold GM recovery plan (price/mix + freight actions).",
        "Oct: board deep-dive on unit economics of launch SKU NL-Probiotic Bar.",
        "Nov: FY27 budget kickoff — volume, AUR, and opex envelopes.",
    ]),
]

row = 5
for title, bullets in sections:
    ws.cell(row, 2, title).font = white_bold
    ws.cell(row, 2).fill = header_fill
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    ws.cell(row, 2).border = thin
    ws.cell(row, 3).border = thin
    row += 1
    for b in bullets:
        ws.cell(row, 2, "•").font = bold
        ws.cell(row, 2).border = thin
        c = ws.cell(row, 3, b)
        style_input(c)
        c.alignment = Alignment(wrap_text=True, horizontal="left")
        ws.row_dimensions[row].height = 26
        row += 1
    row += 1

ws.cell(row, 2, "Owner / next update").font = bold
ws.cell(row + 0, 3, "FP&A — next flash 3 business days after Sep close")
style_input(ws.cell(row, 3))
ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=3)

# ═══════════════════════════════════════════════════════════════════
# 05_One_Pager — THE MONEY TAB
# ═══════════════════════════════════════════════════════════════════
ws = wb.create_sheet(ONE)
landscape(ws, fit_height=1)
set_col_widths(ws, [2.5, 14, 11, 11, 11, 10, 2.5, 13, 11, 11, 11, 10, 2.5, 22, 2.5])

# Banner
for col in range(2, 15):
    ws.cell(1, col).fill = dark_navy
ws.merge_cells("B1:N1")
ws["B1"] = '="NORTHLINE  ·  ELT FLASH  ·  "&\'01_Assumptions\'!D6&" "&\'01_Assumptions\'!C7'
ws["B1"].font = big_title
ws["B1"].fill = dark_navy
ws["B1"].alignment = Alignment(vertical="center", horizontal="left")
ws.row_dimensions[1].height = 28

for col in range(2, 15):
    ws.cell(2, col).fill = dark_navy
ws.merge_cells("B2:N2")
ws["B2"] = "='01_Assumptions'!C8&\"  ·  $000s  ·  Prepared by \"&'01_Assumptions'!C10&\"  ·  As-of \"&TEXT('01_Assumptions'!C11,\"YYYY-MM-DD\")"
ws["B2"].font = subtitle_w
ws["B2"].fill = dark_navy
ws.row_dimensions[2].height = 18

# ── KPI tiles row ──
ws["B4"] = "BOARD KPIs"
ws["B4"].font = section_font
ws.merge_cells("B4:F4")

tile_specs = [
    # col_start, label, value_formula, rag_formula, fmt
    (2, "Net revenue", "='03_KPIs'!C6", "='03_KPIs'!G6", money),
    (4, "vs Budget", "='03_KPIs'!F6", "='03_KPIs'!G6", pct),
    (8, "EBITDA", "='03_KPIs'!C8", "='03_KPIs'!G8", money),
    (10, "vs Budget", "='03_KPIs'!F8", "='03_KPIs'!G8", pct),
]

# Tile block: Revenue | GM% | EBITDA | OTIF | DSO | Cash
tiles = [
    (2, "NET REVENUE", "='03_KPIs'!C6", "='03_KPIs'!E6", "='03_KPIs'!G6", money),
    (4, "GROSS MARGIN %", "='03_KPIs'!C7", "='03_KPIs'!E7", "='03_KPIs'!G7", pct),
    (6, "EBITDA", "='03_KPIs'!C8", "='03_KPIs'!E8", "='03_KPIs'!G8", money),
    (8, "VOLUME (000)", "='03_KPIs'!C9", "='03_KPIs'!E9", "='03_KPIs'!G9", num),
    (10, "OTIF %", "='03_KPIs'!C10", "='03_KPIs'!E10", "='03_KPIs'!G10", pct),
    (12, "DSO / OCF RAG", "='03_KPIs'!C11", "='03_KPIs'!G11", "='03_KPIs'!G12", num),
]

# Simpler tile layout — 6 tiles across B-M
tile_data = [
    ("NET REVENUE", "='03_KPIs'!C6", "='03_KPIs'!F6", "='03_KPIs'!G6", money, pct),
    ("GROSS MARGIN", "='03_KPIs'!C7", "='03_KPIs'!E7", "='03_KPIs'!G7", pct, "0.0"),
    ("EBITDA", "='03_KPIs'!C8", "='03_KPIs'!F8", "='03_KPIs'!G8", money, pct),
    ("VOLUME", "='03_KPIs'!C9", "='03_KPIs'!F9", "='03_KPIs'!G9", num, pct),
    ("OTIF", "='03_KPIs'!C10", "='03_KPIs'!F10", "='03_KPIs'!G10", pct, pct),
    ("DSO (days)", "='03_KPIs'!C11", "='03_KPIs'!E11", "='03_KPIs'!G11", num, "0.0"),
]

start_cols = [2, 4, 6, 8, 10, 12]
for (label, val_f, var_f, rag_f, vfmt, xfmt), sc in zip(tile_data, start_cols):
    # label
    c = ws.cell(5, sc, label)
    c.font = tile_title
    c.fill = tile_fill
    c.border = thin
    c.alignment = Alignment(horizontal="center")
    ws.cell(5, sc + 1).fill = tile_fill
    ws.cell(5, sc + 1).border = thin
    ws.merge_cells(start_row=5, start_column=sc, end_row=5, end_column=sc + 1)
    # value
    c = ws.cell(6, sc, val_f)
    c.font = tile_value
    c.border = thin
    c.number_format = vfmt
    c.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=6, start_column=sc, end_row=6, end_column=sc + 1)
    ws.cell(6, sc + 1).border = thin
    # var + rag
    c = ws.cell(7, sc, var_f)
    c.font = black
    c.border = thin
    c.number_format = xfmt if label != "GROSS MARGIN" else "0.0"
    if label == "GROSS MARGIN":
        c.number_format = "0.0%"  # var is pp as decimal already from C-D; show as number
        # actually E7 is C-D which for % is pp in decimal — format as 0.0%
        c.number_format = "0.0%"
    c.alignment = Alignment(horizontal="center")
    c = ws.cell(7, sc + 1, rag_f)
    c.font = bold_black
    c.border = thin
    c.alignment = Alignment(horizontal="center")

ws.row_dimensions[6].height = 22

# RAG CF on tile RAG cells
for sc in start_cols:
    cell_ref = f"{get_column_letter(sc + 1)}7"
    ws.conditional_formatting.add(cell_ref, FormulaRule(formula=[f'{cell_ref}="Green"'], fill=green_fill))
    ws.conditional_formatting.add(cell_ref, FormulaRule(formula=[f'{cell_ref}="Amber"'], fill=amber_fill))
    ws.conditional_formatting.add(cell_ref, FormulaRule(formula=[f'{cell_ref}="Red"'], fill=red_fill))

# ── Compact P&L ──
ws["B9"] = "P&L FLASH ($000s)"
ws["B9"].font = section_font
ws.merge_cells("B9:F9")

for i, h in enumerate(["Line", "Mo Act", "Mo Bud", "A vs B %", "YTD Act", "YTD A vs B %"]):
    style_header_cell(ws.cell(10, 2 + i), h)

# Pull from P&L sheet key rows: 7 Rev, 9 GP, 14 EBITDA, 16 EBIT, 20 NI
one_pager_lines = [
    ("Net revenue", 7),
    ("Gross profit", 9),
    ("EBITDA", 14),
    ("EBIT", 16),
    ("Net income", 20),
]
for i, (label, src) in enumerate(one_pager_lines):
    r = 11 + i
    ws.cell(r, 2, label).font = bold_black
    ws.cell(r, 2).border = thin
    c = ws.cell(r, 3, f"='02_P&L_Flash'!C{src}")
    style_formula(c, key=(label in ("EBITDA", "Net income")))
    c.number_format = money
    c = ws.cell(r, 4, f"='02_P&L_Flash'!D{src}")
    style_formula(c)
    c.number_format = money
    c = ws.cell(r, 5, f"='02_P&L_Flash'!F{src}")
    style_formula(c)
    c.number_format = pct
    c = ws.cell(r, 6, f"='02_P&L_Flash'!H{src}")
    style_formula(c)
    c.number_format = money
    c = ws.cell(r, 7, f"='02_P&L_Flash'!K{src}")
    style_formula(c)
    c.number_format = pct

# Margins strip
ws["B17"] = "GM% Mo"
ws["C17"] = "='02_P&L_Flash'!N7"
style_formula(ws["C17"], key=True)
ws["C17"].number_format = pct
ws["D17"] = "EBITDA% Mo"
ws["E17"] = "='02_P&L_Flash'!C23"
style_formula(ws["E17"], key=True)
ws["E17"].number_format = pct
ws["F17"] = "EBITDA% YTD"
ws["G17"] = "='02_P&L_Flash'!C24"
style_formula(ws["G17"], key=True)
ws["G17"].number_format = pct

# ── Headlines ──
ws["I9"] = "HEADLINES"
ws["I9"].font = section_font
ws.merge_cells("I9:L9")

for i in range(5):
    r = 10 + i
    ws.cell(r, 8, f"H{i+1}" if i < 4 else "WATCH").font = bold
    ws.cell(r, 8).fill = section_fill
    ws.cell(r, 8).border = thin
    c = ws.cell(r, 9, f"='01_Assumptions'!C{14+i}")
    c.font = black
    c.border = thin
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
    ws.row_dimensions[r].height = 28

# ── Callouts summary (pull first bullet of each section) ──
ws["B19"] = "CALLOUTS (from 04)"
ws["B19"].font = section_font
ws.merge_cells("B19:G19")

callout_refs = [
    ("Well", "='04_Callouts'!C6"),
    ("Watch", "='04_Callouts'!C11"),
    ("Ask", "='04_Callouts'!C16"),
    ("Risk", "='04_Callouts'!C21"),
]
for i, (tag, ref) in enumerate(callout_refs):
    r = 20 + i
    ws.cell(r, 2, tag).font = white_bold
    ws.cell(r, 2).fill = header_fill
    ws.cell(r, 2).border = thin
    c = ws.cell(r, 3, ref)
    c.font = black
    c.border = thin
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    ws.row_dimensions[r].height = 24

# ── RAG scoreboard ──
ws["I19"] = "RAG SCOREBOARD"
ws["I19"].font = section_font
ws.merge_cells("I19:L19")

ws["I20"] = "Green"
ws["J20"] = "='03_KPIs'!C15"
style_formula(ws["J20"], key=True)
ws["J20"].fill = green_fill
ws["K20"] = "Amber"
ws["L20"] = "='03_KPIs'!E15"
style_formula(ws["L20"])
ws["L20"].fill = amber_fill

ws["I21"] = "Red"
ws["J21"] = "='03_KPIs'!G15"
style_formula(ws["J21"])
ws["J21"].fill = red_fill
ws["K21"] = "KPIs"
ws["L21"] = "=J20+L20+J21"
style_formula(ws["L21"])

ws["I22"] = "Cash RAG"
ws["J22"] = "='03_KPIs'!G12"
style_formula(ws["J22"], key=True)
ws.conditional_formatting.add("J22", FormulaRule(formula=['J22="Green"'], fill=green_fill))
ws.conditional_formatting.add("J22", FormulaRule(formula=['J22="Amber"'], fill=amber_fill))
ws.conditional_formatting.add("J22", FormulaRule(formula=['J22="Red"'], fill=red_fill))

ws["I23"] = "OCF Mo Act"
ws["J23"] = "='03_KPIs'!C12"
style_formula(ws["J23"])
ws["J23"].number_format = money

# Mini bar chart data (hidden helper area for Mo Act vs Bud key lines)
ws["I25"] = "Chart data (Mo Act vs Bud)"
ws["I25"].font = small_grey
ws["I26"] = "Line"
ws["J26"] = "Actual"
ws["K26"] = "Budget"
for i, (label, src) in enumerate([("Revenue", 7), ("GP", 9), ("EBITDA", 14)]):
    r = 27 + i
    ws.cell(r, 9, label)
    ws.cell(r, 10, f"='02_P&L_Flash'!C{src}")
    ws.cell(r, 10).number_format = money
    ws.cell(r, 11, f"='02_P&L_Flash'!D{src}")
    ws.cell(r, 11).number_format = money

chart = BarChart()
chart.type = "col"
chart.grouping = "clustered"
chart.title = "Close month: Act vs Bud"
chart.y_axis.title = "$000s"
chart.style = 10
data = Reference(ws, min_col=10, min_row=26, max_col=11, max_row=29)
cats = Reference(ws, min_col=9, min_row=27, max_row=29)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.width = 10
chart.height = 6
ws.add_chart(chart, "I30")

# Footer strip
for col in range(2, 15):
    ws.cell(24, col).fill = dark_navy
ws.merge_cells("B24:N24")
# Actually put footer lower — row 36 area. Keep row 24 as divider note.
ws["B24"] = "Print this sheet landscape · 1 page  ·  Yellow inputs live on Assumptions / P&L / KPIs / Callouts  ·  Fictional sample"
ws["B24"].font = Font(name="Calibri", size=8, italic=True, color="D6E3F0")
ws["B24"].fill = dark_navy
ws["B24"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[24].height = 16

# Fix: chart overlaps footer — move chart data and chart. Clear row 24 dark and use row 35 for footer.
# Actually the chart is at I30 which is below — that's fine. But I put footer at 24 which is between callouts and chart.
# Reposition: remove dark from 24 content conflict — callouts end at 23. Footer at 24 is OK as a thin strip; chart at I30.

ws.print_area = "A1:N35"
ws.page_margins.left = 0.4
ws.page_margins.right = 0.4
ws.page_margins.top = 0.4
ws.page_margins.bottom = 0.4

# ═══════════════════════════════════════════════════════════════════
# 06_Data_Dictionary
# ═══════════════════════════════════════════════════════════════════
ws = wb.create_sheet(DICT)
landscape(ws)
set_col_widths(ws, [3, 28, 18, 70])

ws["B2"] = "06  ·  Data Dictionary"
ws["B2"].font = title_font
ws["B3"] = "Field definitions so another analyst can inherit the file."
ws["B3"].font = italic_grey

for i, h in enumerate(["Field", "Tab", "Definition"]):
    style_header_cell(ws.cell(5, 2 + i), h)

defs = [
    ("Close month", ASSUMP, "1–12 toggle. Labels the flash period and drives the one-pager banner."),
    ("Headlines 1–4 + Watch", ASSUMP, "ELT narrative bullets. Flow to 05_One_Pager."),
    ("RAG bands", ASSUMP, "Amber / Green relative bands by KPI. Direction H = higher-better, L = lower-better."),
    ("P&L Actual / Budget / Forecast", PL, "Close-month and YTD $000s. Yellow inputs for leaf lines; subtotals are formulas."),
    ("A vs B % / A vs F %", PL, "(Actual − Plan) / |Plan|. Conditional formatting: green ≥0.5%, amber −2% to 0.5%, red < −2%."),
    ("Gross profit / EBITDA / EBIT / NI", PL, "Formula subtotals. Green-highlighted key rows."),
    ("Gross margin %", PL, "Gross profit ÷ net revenue."),
    ("EBITDA margin %", PL, "EBITDA ÷ net revenue — month and YTD."),
    ("Board KPIs", KPI, "Seven ELT KPIs: revenue, GM%, EBITDA, volume, OTIF, DSO, operating cash flow."),
    ("KPI RAG", KPI, "Green / Amber / Red from actual vs budget and Assumptions bands. Formula, not typed."),
    ("Callouts", CALL, "What went well, attention, decisions/asks, risks, forward look — yellow commentary."),
    ("One-Pager", ONE, "Print-ready ELT page: KPI tiles, compact P&L, headlines, callouts, RAG scoreboard, chart."),
    ("Units", "All", "Dollars in $000s. Rates in %. Volume in 000 cases. DSO in days."),
    ("Color convention", "All", "Yellow + blue font = inputs. Black font = formulas. Green/Amber/Red = RAG."),
]
for i, (field, tab, definition) in enumerate(defs):
    r = 6 + i
    ws.cell(r, 2, field).font = bold
    ws.cell(r, 2).border = thin
    ws.cell(r, 3, tab).border = thin
    cell = ws.cell(r, 4, definition)
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = thin
    ws.row_dimensions[r].height = 32

ws["B21"] = "All sample numbers are fictional. Built for a public GitHub portfolio — no employer data."
ws["B21"].font = small_grey

out = Path(__file__).resolve().parent / "Northline_Board_ELT_Flash.xlsx"
wb.save(out)
print("Wrote", out)
print("Sheets:", wb.sheetnames)
