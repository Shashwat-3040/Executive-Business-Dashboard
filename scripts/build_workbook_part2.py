"""
Stage 2 — adds formula-driven KPI sheets on top of the raw tables built in
build_workbook.py. Every KPI is a live SUMIFS/COUNTIFS/AVERAGEIFS/INDEX-MATCH
formula referencing the raw sheets — nothing is a hardcoded Python value.
"""
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PATH = "/home/claude/executive-dashboard/excel/Executive_Business_Performance_Dashboard.xlsx"
EXPORTS = "/home/claude/executive-dashboard/data/exports"

FONT_NAME = "Arial"
NAVY = "1F3864"
WHITE = "FFFFFF"
header_font = Font(name=FONT_NAME, bold=True, color=WHITE, size=10)
header_fill = PatternFill("solid", fgColor=NAVY)
label_font = Font(name=FONT_NAME, bold=True, size=10)
normal_font = Font(name=FONT_NAME, size=10)
pct_fmt = "0.0%"
money_fmt = "$#,##0"
ratio_fmt = "0.00"

def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

def add_table(ws, name, ref, style="TableStyleMedium9"):
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(tbl)

def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

wb = load_workbook(PATH)

# Raw data ranges (known row counts from stage 1)
RS_LAST = 110700   # Raw_Sales data through row 110700 (header row 1)
DC_LAST = 6001     # Dim_Customer
OD_LAST = 721      # OpEx_Data
IS_LAST = 601      # Inventory_Summary

RS = "Raw_Sales"
DC = "Dim_Customer"
OD = "OpEx_Data"
IS = "Inventory_Summary"

# Raw_Sales columns: A Year, B Month, C MonthName, D Region, E Store, F StoreType,
# G Department, H Category, I Product, J Quantity, K UnitPrice, L UnitCost,
# M DiscountPct, N GrossRevenue, O NetRevenue, P TotalCost, Q GrossProfit
def rs_col(letter):
    return f"'{RS}'!${letter}$2:${letter}${RS_LAST}"

# Dim_Customer columns: A CustomerID, B CustomerName, C Region, D MembershipType,
# E JoinDate, F ChurnDate, G IsActive
def dc_col(letter):
    return f"'{DC}'!${letter}$2:${letter}${DC_LAST}"

# OpEx_Data columns: A Region, B Year, C Month, D Category, E Amount
def od_col(letter):
    return f"'{OD}'!${letter}$2:${letter}${OD_LAST}"

# Inventory_Summary columns: A Region, B Department, C Year, D Month, E MonthlyCOGS, F AvgInventoryValue
def is_col(letter):
    return f"'{IS}'!${letter}$2:${letter}${IS_LAST}"

regions = ["Central", "East", "North", "South", "West"]
departments = ["Electronics", "Apparel", "Home & Living", "Grocery", "Beauty & Personal Care"]
years = [2024, 2025]

# ============================================================================
# SHEET: Sales_Summary  (Region x Year x Month, SUMIFS from Raw_Sales)
# ============================================================================
ws = wb.create_sheet("Sales_Summary")
headers = ["Year", "Month", "Region", "NetRevenue", "GrossProfit", "Units", "Transactions"]
for j, h in enumerate(headers, start=1):
    ws.cell(row=1, column=j, value=h)
style_header_row(ws, 1, len(headers))

row = 2
month_rows = []
for y in years:
    for m in range(1, 13):
        for reg in regions:
            ws.cell(row=row, column=1, value=y)
            ws.cell(row=row, column=2, value=m)
            ws.cell(row=row, column=3, value=reg)
            ws.cell(row=row, column=4,
                     value=f"=SUMIFS({rs_col('O')},{rs_col('A')},A{row},{rs_col('B')},B{row},{rs_col('D')},C{row})")
            ws.cell(row=row, column=5,
                     value=f"=SUMIFS({rs_col('Q')},{rs_col('A')},A{row},{rs_col('B')},B{row},{rs_col('D')},C{row})")
            ws.cell(row=row, column=6,
                     value=f"=SUMIFS({rs_col('J')},{rs_col('A')},A{row},{rs_col('B')},B{row},{rs_col('D')},C{row})")
            ws.cell(row=row, column=7,
                     value=f"=COUNTIFS({rs_col('A')},A{row},{rs_col('B')},B{row},{rs_col('D')},C{row})")
            for c in (4, 5):
                ws.cell(row=row, column=c).number_format = money_fmt
            row += 1
last_row_ss = row - 1
add_table(ws, "SalesSummary", f"A1:G{last_row_ss}")
ws.freeze_panes = "A2"
autosize(ws, [7, 7, 14, 14, 14, 10, 13])
print("Sales_Summary rows:", last_row_ss - 1)

# ============================================================================
# SHEET: EBITDA  (Region x Year)
# ============================================================================
ws = wb.create_sheet("EBITDA")
headers = ["Region", "Year", "NetRevenue", "COGS", "GrossProfit", "GrossMarginPct",
           "CashOpEx", "D_and_A", "EBITDA", "EBITDAMarginPct"]
for j, h in enumerate(headers, start=1):
    ws.cell(row=1, column=j, value=h)
style_header_row(ws, 1, len(headers))

row = 2
for reg in regions:
    for y in years:
        ws.cell(row=row, column=1, value=reg)
        ws.cell(row=row, column=2, value=y)
        ws.cell(row=row, column=3, value=f"=SUMIFS({rs_col('O')},{rs_col('D')},A{row},{rs_col('A')},B{row})")
        ws.cell(row=row, column=4, value=f"=SUMIFS({rs_col('P')},{rs_col('D')},A{row},{rs_col('A')},B{row})")
        ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
        ws.cell(row=row, column=6, value=f"=E{row}/C{row}")
        ws.cell(row=row, column=7,
                 value=f'=SUMIFS({od_col("E")},{od_col("A")},A{row},{od_col("B")},B{row},{od_col("D")},"<>Depreciation & Amortization")')
        ws.cell(row=row, column=8,
                 value=f'=SUMIFS({od_col("E")},{od_col("A")},A{row},{od_col("B")},B{row},{od_col("D")},"Depreciation & Amortization")')
        ws.cell(row=row, column=9, value=f"=C{row}-D{row}-G{row}")
        ws.cell(row=row, column=10, value=f"=I{row}/C{row}")
        for c in (3, 4, 5, 7, 8, 9):
            ws.cell(row=row, column=c).number_format = money_fmt
        for c in (6, 10):
            ws.cell(row=row, column=c).number_format = pct_fmt
        row += 1
last_row_eb = row - 1

# company total row
ws.cell(row=row, column=1, value="TOTAL / COMPANY").font = label_font
ws.cell(row=row, column=2, value="")
for c, letter in [(3, "C"), (4, "D"), (5, "E"), (7, "G"), (8, "H"), (9, "I")]:
    ws.cell(row=row, column=c, value=f"=SUM({letter}2:{letter}{last_row_eb})")
    ws.cell(row=row, column=c).number_format = money_fmt
ws.cell(row=row, column=6, value=f"=E{row}/C{row}").number_format = pct_fmt
ws.cell(row=row, column=10, value=f"=I{row}/C{row}").number_format = pct_fmt
last_row_eb_total = row
add_table(ws, "EbitdaTbl", f"A1:J{last_row_eb}")
ws.freeze_panes = "A2"
autosize(ws, [10, 6, 14, 14, 14, 14, 12, 12, 14, 15])
print("EBITDA rows:", last_row_eb - 1, "total row:", last_row_eb_total)

# ============================================================================
# SHEET: Margin_Leakage  (Region gross margin vs. company average)
# ============================================================================
ws = wb.create_sheet("Margin_Leakage")
headers = ["Region", "NetRevenue", "COGS", "GrossMarginPct", "CompanyAvgMarginPct",
           "MarginGap_pp", "PctLeakageVsAvg"]
for j, h in enumerate(headers, start=1):
    ws.cell(row=1, column=j, value=h)
style_header_row(ws, 1, len(headers))
row = 2
first_ml_row = row
for reg in regions:
    ws.cell(row=row, column=1, value=reg)
    ws.cell(row=row, column=2, value=f"=SUMIFS({rs_col('O')},{rs_col('D')},A{row})")
    ws.cell(row=row, column=3, value=f"=SUMIFS({rs_col('P')},{rs_col('D')},A{row})")
    ws.cell(row=row, column=4, value=f"=(B{row}-C{row})/B{row}")
    ws.cell(row=row, column=2).number_format = money_fmt
    ws.cell(row=row, column=3).number_format = money_fmt
    ws.cell(row=row, column=4).number_format = pct_fmt
    row += 1
last_ml_row = row - 1
avg_range = f"D{first_ml_row}:D{last_ml_row}"
for r in range(first_ml_row, last_ml_row + 1):
    ws.cell(row=r, column=5, value=f"=AVERAGE(${avg_range})")
    ws.cell(row=r, column=6, value=f"=D{r}-E{r}")
    ws.cell(row=r, column=7, value=f"=(E{r}-D{r})/E{r}")
    ws.cell(row=r, column=5).number_format = pct_fmt
    ws.cell(row=r, column=6).number_format = pct_fmt
    ws.cell(row=r, column=7).number_format = pct_fmt
add_table(ws, "MarginLeakage", f"A1:G{last_ml_row}")
autosize(ws, [10, 14, 14, 15, 18, 14, 16])
print("Margin_Leakage rows:", last_ml_row - 1)

# ============================================================================
# SHEET: Churn  (Region x MembershipType)
# ============================================================================
ws = wb.create_sheet("Churn")
headers = ["Region", "MembershipType", "TotalCustomers", "ChurnedCustomers", "ChurnRatePct"]
for j, h in enumerate(headers, start=1):
    ws.cell(row=1, column=j, value=h)
style_header_row(ws, 1, len(headers))
row = 2
memberships = ["Standard", "Silver", "Gold"]
for reg in regions:
    for mem in memberships:
        ws.cell(row=row, column=1, value=reg)
        ws.cell(row=row, column=2, value=mem)
        ws.cell(row=row, column=3, value=f"=COUNTIFS({dc_col('C')},A{row},{dc_col('D')},B{row})")
        ws.cell(row=row, column=4, value=f'=COUNTIFS({dc_col("C")},A{row},{dc_col("D")},B{row},{dc_col("F")},"<>")')
        ws.cell(row=row, column=5, value=f"=D{row}/C{row}")
        ws.cell(row=row, column=5).number_format = pct_fmt
        row += 1
last_row_ch = row - 1
# region-level rollup rows
region_rollup_start = row
for reg in regions:
    ws.cell(row=row, column=1, value=reg)
    ws.cell(row=row, column=2, value="ALL TIERS")
    ws.cell(row=row, column=3, value=f"=COUNTIFS({dc_col('C')},A{row})")
    ws.cell(row=row, column=4, value=f'=COUNTIFS({dc_col("C")},A{row},{dc_col("F")},"<>")')
    ws.cell(row=row, column=5, value=f"=D{row}/C{row}")
    ws.cell(row=row, column=5).number_format = pct_fmt
    row += 1
last_row_ch_total = row - 1
# overall
ws.cell(row=row, column=1, value="COMPANY").font = label_font
ws.cell(row=row, column=2, value="ALL")
ws.cell(row=row, column=3, value=f"=COUNTA({dc_col('A')})")
ws.cell(row=row, column=4, value=f'=COUNTIFS({dc_col("F")},"<>")')
ws.cell(row=row, column=5, value=f"=D{row}/C{row}")
ws.cell(row=row, column=5).number_format = pct_fmt
add_table(ws, "ChurnTbl", f"A1:E{last_row_ch}")
autosize(ws, [10, 16, 15, 17, 13])
print("Churn rows:", last_row_ch - 1, "region rollups thru", last_row_ch_total)

# ============================================================================
# SHEET: Inventory_Turnover  (Region x Department, annualized from 24mo data)
# ============================================================================
ws = wb.create_sheet("Inventory_Turnover")
headers = ["Region", "Department", "TwoYearCOGS", "AvgInventoryValue", "AnnualTurnoverRatio"]
for j, h in enumerate(headers, start=1):
    ws.cell(row=1, column=j, value=h)
style_header_row(ws, 1, len(headers))
row = 2
for reg in regions:
    for dept in departments:
        ws.cell(row=row, column=1, value=reg)
        ws.cell(row=row, column=2, value=dept)
        ws.cell(row=row, column=3, value=f"=SUMIFS({is_col('E')},{is_col('A')},A{row},{is_col('B')},B{row})")
        ws.cell(row=row, column=4, value=f"=AVERAGEIFS({is_col('F')},{is_col('A')},A{row},{is_col('B')},B{row})")
        ws.cell(row=row, column=5, value=f"=(C{row}/2)/D{row}")
        ws.cell(row=row, column=3).number_format = money_fmt
        ws.cell(row=row, column=4).number_format = money_fmt
        ws.cell(row=row, column=5).number_format = ratio_fmt
        row += 1
last_row_it = row - 1
# department-level rollup (all regions)
dept_rollup_start = row
for dept in departments:
    ws.cell(row=row, column=1, value="ALL REGIONS")
    ws.cell(row=row, column=2, value=dept)
    ws.cell(row=row, column=3, value=f"=SUMIFS({is_col('E')},{is_col('B')},B{row})")
    ws.cell(row=row, column=4, value=f"=AVERAGEIFS({is_col('F')},{is_col('B')},B{row})")
    ws.cell(row=row, column=5, value=f"=(C{row}/2)/D{row}")
    ws.cell(row=row, column=3).number_format = money_fmt
    ws.cell(row=row, column=4).number_format = money_fmt
    ws.cell(row=row, column=5).number_format = ratio_fmt
    row += 1
dept_rollup_end = row - 1
add_table(ws, "InventoryTurnover", f"A1:E{last_row_it}")
autosize(ws, [14, 22, 16, 18, 18])
print("Inventory_Turnover rows:", last_row_it - 1, "dept rollups thru", dept_rollup_end)

wb.save(PATH)
print("Stage 2 saved.")
print("ROW_MAP", dict(
    RS_LAST=RS_LAST, DC_LAST=DC_LAST, OD_LAST=OD_LAST, IS_LAST=IS_LAST,
    last_row_ss=last_row_ss, last_row_eb=last_row_eb, last_row_eb_total=last_row_eb_total,
    first_ml_row=first_ml_row, last_ml_row=last_ml_row,
    last_row_ch=last_row_ch, region_rollup_start=region_rollup_start, last_row_ch_total=last_row_ch_total,
    last_row_it=last_row_it, dept_rollup_start=dept_rollup_start, dept_rollup_end=dept_rollup_end,
))
