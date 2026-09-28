"""
Build Executive_Business_Performance_Dashboard.xlsx
Real formulas (SUMIFS / INDEX-MATCH), Excel Tables, and native charts —
built entirely from the Raw_Sales / Dim_Customer / OpEx_Data / Inventory_Summary
tables so the workbook recalculates if the underlying extract changes.
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference, PieChart
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

EXPORTS = "/home/claude/executive-dashboard/data/exports"
OUT_PATH = "/home/claude/executive-dashboard/excel/Executive_Business_Performance_Dashboard.xlsx"

FONT_NAME = "Arial"
NAVY = "1F3864"
BLUE = "2E5395"
LIGHT_BLUE = "D9E2F3"
GOLD = "BF8F00"
RED = "C00000"
GREEN = "375623"
WHITE = "FFFFFF"

header_font = Font(name=FONT_NAME, bold=True, color=WHITE, size=10)
header_fill = PatternFill("solid", fgColor=NAVY)
title_font = Font(name=FONT_NAME, bold=True, size=16, color=NAVY)
subtitle_font = Font(name=FONT_NAME, italic=True, size=10, color="595959")
label_font = Font(name=FONT_NAME, bold=True, size=10)
normal_font = Font(name=FONT_NAME, size=10)
kpi_value_font = Font(name=FONT_NAME, bold=True, size=20, color=NAVY)
kpi_label_font = Font(name=FONT_NAME, size=10, color="595959")
thin = Side(style="thin", color="BFBFBF")
box_border = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)

# ---------------------------------------------------------------- helpers
def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

def add_table(ws, name, ref):
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tbl)

def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def write_df(ws, df, start_row=1, start_col=1):
    for j, col in enumerate(df.columns):
        ws.cell(row=start_row, column=start_col + j, value=col)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, val in enumerate(row):
            ws.cell(row=start_row + i, column=start_col + j, value=val)
    return start_row + len(df)  # last data row

# ============================================================================
# SHEET: Read_Me
# ============================================================================
ws = wb.create_sheet("Read_Me")
ws.sheet_view.showGridLines = False
ws["B2"] = "Executive Business Performance 360° Dashboard"
ws["B2"].font = title_font
ws["B3"] = "NorthStar Retail Group — synthetic dataset built for portfolio purposes"
ws["B3"].font = subtitle_font

sections = [
    ("Data source", "PostgreSQL warehouse (5 dim tables, 3 fact tables). Raw_Sales below mirrors the "
     "Power Query output that would refresh from that database on a schedule."),
    ("Scope", "110,699 transaction line items · 40 stores across 5 regions · 300 SKUs across 5 departments · "
     "6,000 loyalty customers · 24 months (Jan 2024 – Dec 2025)."),
    ("Sheets", "Raw_Sales, Dim_Customer, OpEx_Data, Inventory_Summary hold the source tables. "
     "Sales_Summary, EBITDA, Churn, Inventory_Turnover compute KPIs from those tables with SUMIFS / "
     "COUNTIFS / INDEX-MATCH formulas — nothing is hardcoded. Dashboard is the executive view."),
    ("Key finding", "The South region is running ~14% below the company's average gross margin, driven by "
     "heavier discounting and rising unit costs — see the EBITDA and Dashboard sheets."),
    ("On PivotTables / Slicers", "This workbook uses SUMIFS-based summary tables instead of native "
     "PivotTables so every formula recalculates identically in any Excel version. Because Raw_Sales, "
     "Dim_Customer, OpEx_Data and Inventory_Summary are real Excel Tables, you can select any of them and "
     "go Insert > PivotTable (or Insert > Slicer once a PivotTable exists) to get fully interactive, "
     "native pivots on top of the same data in seconds."),
    ("On XLOOKUP", "Lookups are written with INDEX/MATCH for maximum compatibility across Excel versions; "
     "they are functionally interchangeable with XLOOKUP for this workbook's single-value lookups (see the "
     "Customer Lookup tool on the Dashboard sheet)."),
    ("Automation", "See /vba/ReportAutomation.bas — a macro module (RefreshAllFormulas, "
     "GenerateMonthlyReportSnapshot, Workbook_Open auto-refresh) that automates the recurring reporting "
     "workflow described on the resume. Import it via Excel's VBA editor (Alt+F11 > File > Import File)."),
]
r = 5
for head, body in sections:
    ws.cell(row=r, column=2, value=head).font = label_font
    ws.cell(row=r + 1, column=2, value=body).font = normal_font
    ws.cell(row=r + 1, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=9)
    ws.row_dimensions[r + 1].height = 32
    r += 3
autosize(ws, [3] + [12] * 9)

# ============================================================================
# SHEET: Raw_Sales  (Power Query-style extract -- 110,699 rows)
# ============================================================================
ws = wb.create_sheet("Raw_Sales")
raw = pd.read_csv(f"{EXPORTS}/raw_sales.csv")
last_row = write_df(ws, raw)
n_cols = len(raw.columns)
last_col_letter = get_column_letter(n_cols)
style_header_row(ws, 1, n_cols)
add_table(ws, "RawSales", f"A1:{last_col_letter}{last_row}")
ws.freeze_panes = "A2"
autosize(ws, [6, 6, 9, 10, 22, 10, 18, 16, 20, 9, 9, 9, 10, 12, 12, 11, 12])
print(f"Raw_Sales written: {last_row-1} rows")

RS = f"RawSales"  # structured reference name

# ============================================================================
# SHEET: Dim_Customer  (6,000 rows)
# ============================================================================
ws = wb.create_sheet("Dim_Customer")
cust = pd.read_csv(f"{EXPORTS}/dim_customer.csv")
last_row_c = write_df(ws, cust)
n_cols_c = len(cust.columns)
style_header_row(ws, 1, n_cols_c)
add_table(ws, "DimCustomer", f"A1:{get_column_letter(n_cols_c)}{last_row_c}")
ws.freeze_panes = "A2"
autosize(ws, [10, 20, 10, 14, 12, 12, 9])
print(f"Dim_Customer written: {last_row_c-1} rows")

# ============================================================================
# SHEET: OpEx_Data  (720 rows)
# ============================================================================
ws = wb.create_sheet("OpEx_Data")
opex = pd.read_csv(f"{EXPORTS}/opex.csv")
last_row_o = write_df(ws, opex)
n_cols_o = len(opex.columns)
style_header_row(ws, 1, n_cols_o)
add_table(ws, "OpexData", f"A1:{get_column_letter(n_cols_o)}{last_row_o}")
ws.freeze_panes = "A2"
autosize(ws, [10, 8, 8, 28, 12])
print(f"OpEx_Data written: {last_row_o-1} rows")

# ============================================================================
# SHEET: Inventory_Summary  (600 rows, pre-aggregated from 115K store-level rows)
# ============================================================================
ws = wb.create_sheet("Inventory_Summary")
inv = pd.read_csv(f"{EXPORTS}/inventory_summary.csv")
last_row_i = write_df(ws, inv)
n_cols_i = len(inv.columns)
style_header_row(ws, 1, n_cols_i)
add_table(ws, "InventorySummary", f"A1:{get_column_letter(n_cols_i)}{last_row_i}")
ws.freeze_panes = "A2"
autosize(ws, [10, 18, 8, 8, 16, 18])
ws.cell(row=last_row_i + 2, column=1,
         value="Source: aggregated in PostgreSQL from a 115,200-row store x product x month inventory fact "
               "table (fact_inventory_monthly) to keep the workbook a manageable size.").font = subtitle_font
print(f"Inventory_Summary written: {last_row_i-1} rows")

wb.save(OUT_PATH)
print("Stage 1 saved:", OUT_PATH)
