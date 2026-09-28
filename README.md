# Executive Business Performance 360° Dashboard

**NorthStar Retail Group — a synthetic multi-region retail dataset built to demonstrate
an end-to-end analytics workflow: PostgreSQL → Power Query-style extraction → Excel
executive reporting.**

> Portfolio project. All data is synthetically generated (see [`scripts/generate_data.py`](scripts/generate_data.py)) — no real company, customer, or transaction data is used.

## The story

NorthStar is a 40-store retail chain spanning 5 regions (Central, East, North, South,
West), selling across 5 departments (Electronics, Apparel, Home & Living, Grocery,
Beauty & Personal Care) to a 6,000-member loyalty base. Two years of transactional
history (FY2024–FY2025) were generated and loaded into a PostgreSQL warehouse, then
extracted into an executive Excel dashboard that tracks sales, profitability,
customer, and operational KPIs.

**Headline finding:** the South region is running **13.9% below the company's
average gross margin** — driven by heavier discounting and rising unit costs — which
shows up consistently across the SQL analysis and the Excel dashboard.

| KPI | Result |
|---|---|
| Total net revenue (FY24–FY25) | $21.1M |
| YoY revenue growth | +10.7% |
| Company EBITDA margin | 13.9% |
| Overall customer churn rate | 20.5% |
| Avg. inventory turnover | 3.8x / year |
| **South region margin leakage vs. company avg** | **13.9%** |

## Architecture

```
PostgreSQL (star schema)          Excel workbook
┌─────────────────────┐           ┌──────────────────────────────┐
│ dim_region           │           │ Raw_Sales      (110,699 rows)│
│ dim_store             │  extract  │ Dim_Customer     (6,000 rows)│
│ dim_department        │ ────────► │ OpEx_Data          (720 rows)│
│ dim_product           │  (Power   │ Inventory_Summary  (600 rows)│
│ dim_customer           │   Query   │       │                       │
│ dim_date               │   /SQL)   │       ▼                       │
│ fact_sales (110,699)   │           │ Sales_Summary  (SUMIFS)       │
│ fact_inventory_monthly │           │ EBITDA         (SUMIFS)       │
│   (115,200)            │           │ Margin_Leakage (AVERAGE)      │
│ fact_opex_monthly (720)│           │ Churn          (COUNTIFS)     │
└─────────────────────┘           │ Inventory_Turnover (AVERAGEIFS)│
                                    │ Chart_Data     (INDEX/MATCH)  │
                                    │       │                       │
                                    │       ▼                       │
                                    │ Dashboard (KPI cards, charts, │
                                    │  customer lookup tool)        │
                                    └──────────────────────────────┘
```

Every KPI sheet in the workbook is formula-driven (`SUMIFS`, `COUNTIFS`,
`AVERAGEIFS`, `INDEX`/`MATCH`) against the raw tables — nothing is a hardcoded
value, so the dashboard recalculates correctly if the underlying extract changes.
**6,811 formulas, verified with zero calculation errors.**

## Repo structure

```
├── data/
│   ├── generate_data.py output (dim_*.csv, fact_*.csv)
│   └── exports/                  # flattened extracts consumed by the workbook
├── sql/
│   ├── 01_schema.sql             # star schema DDL
│   ├── 02_load_data.sql          # \copy data load
│   └── 03_kpi_queries.sql        # EBITDA, YoY growth, churn, turnover, margin leakage
├── scripts/
│   ├── generate_data.py          # synthetic data generator (Faker + numpy/pandas)
│   ├── pull_for_excel.py         # pulls aggregates from Postgres for the workbook
│   ├── build_workbook.py         # Stage 1: raw data sheets
│   ├── build_workbook_part2.py   # Stage 2: KPI formula sheets
│   ├── build_workbook_part3.py   # Stage 3: Chart_Data helper sheet
│   └── build_workbook_part4.py   # Stage 4: Dashboard (KPI cards, charts, lookup tool)
├── excel/
│   └── Executive_Business_Performance_Dashboard_1.xlsx
├── vba/
│   └── ReportAutomation.bas      # macro module (see below)
└── docs/
```

## Reproducing it

```bash
# 1. Generate the synthetic dataset
python3 scripts/generate_data.py

# 2. Load into PostgreSQL
createdb executive_dashboard
psql -d executive_dashboard -f sql/01_schema.sql
psql -d executive_dashboard -f sql/02_load_data.sql

# 3. Run the KPI queries directly (optional, for the SQL-only view)
psql -d executive_dashboard -f sql/03_kpi_queries.sql

# 4. Pull extracts + build the Excel workbook
python3 scripts/pull_for_excel.py
python3 scripts/build_workbook.py
python3 scripts/build_workbook_part2.py
python3 scripts/build_workbook_part3.py
python3 scripts/build_workbook_part4.py
```

## On PivotTables, XLOOKUP, and Slicers

The workbook uses SUMIFS/COUNTIFS-based summary tables instead of native
PivotTables so every formula recalculates identically across Excel versions and
platforms. Because `Raw_Sales`, `Dim_Customer`, `OpEx_Data`, and
`Inventory_Summary` are real Excel Tables, you can select any of them and go
**Insert > PivotTable** (then **Insert > Slicer** once a PivotTable exists) to get
fully interactive native pivots and slicers on top of the same data in seconds.
Lookups are written with `INDEX`/`MATCH` for maximum compatibility; they are
functionally interchangeable with `XLOOKUP` for the single-value lookups used
here (see the Customer Lookup tool on the Dashboard sheet).

## VBA automation (`vba/ReportAutomation.bas`)

A macro module that automates the recurring reporting workflow:

- **`Workbook_Open`** — auto-refreshes all data connections and formulas whenever the file is opened.
- **`RefreshAllFormulas`** — forces a full recalculation after a Power Query refresh.
- **`GenerateMonthlyReportSnapshot`** — copies the current KPI headlines into a new, timestamped sheet as static values, for board packs or month-end archives.
- **`ExportDashboardToPDF`** — exports the Dashboard sheet to a dated PDF.

To use it: open the workbook in Excel, `Alt+F11` → **File > Import File** →
select `ReportAutomation.bas`, then save as `.xlsm`.

## Tech stack

PostgreSQL 16 · Python (pandas, numpy, Faker, psycopg2) for data generation ·
Advanced Excel (SUMIFS, COUNTIFS, AVERAGEIFS, INDEX/MATCH, Excel Tables, native
charts) · VBA
