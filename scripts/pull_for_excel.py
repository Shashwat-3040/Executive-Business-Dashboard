import psycopg2
import pandas as pd

conn = psycopg2.connect(host="localhost", dbname="executive_dashboard", user="postgres", password="postgres")

def q(sql):
    return pd.read_sql(sql, conn)

OUT = "/home/claude/executive-dashboard/data/exports"

# 1. Raw sales flat table (the "Power Query output" -> feeds all SUMIFS formulas)
raw_sales = q("""
    SELECT d.year AS "Year", d.month AS "Month", d.month_name AS "MonthName",
           r.region_name AS "Region", s.store_name AS "Store", s.store_type AS "StoreType",
           dep.department_name AS "Department", p.category AS "Category", p.product_name AS "Product",
           f.quantity AS "Quantity", f.unit_price AS "UnitPrice", f.unit_cost AS "UnitCost",
           f.discount_pct AS "DiscountPct", f.gross_revenue AS "GrossRevenue",
           f.net_revenue AS "NetRevenue", f.total_cost AS "TotalCost",
           ROUND(f.net_revenue - f.total_cost, 2) AS "GrossProfit"
    FROM fact_sales f
    JOIN dim_date d ON d.date_id = f.date_id
    JOIN dim_region r ON r.region_id = f.region_id
    JOIN dim_store s ON s.store_id = f.store_id
    JOIN dim_product p ON p.product_id = f.product_id
    JOIN dim_department dep ON dep.department_id = f.department_id
    ORDER BY d.year, d.month, r.region_name
""")
raw_sales.to_csv(f"{OUT}/raw_sales.csv", index=False)
print("raw_sales", raw_sales.shape)

# 2. Customer dimension (for lookup tool + churn analysis)
dim_customer = q("""
    SELECT c.customer_id AS "CustomerID", c.customer_name AS "CustomerName",
           r.region_name AS "Region", c.membership_type AS "MembershipType",
           c.join_date AS "JoinDate", c.churn_date AS "ChurnDate",
           c.is_active AS "IsActive"
    FROM dim_customer c JOIN dim_region r ON r.region_id = c.region_id
    ORDER BY c.customer_id
""")
dim_customer.to_csv(f"{OUT}/dim_customer.csv", index=False)
print("dim_customer", dim_customer.shape)

# 3. OpEx data (region x month x category)
opex = q("""
    SELECT r.region_name AS "Region", (o.month_id/100) AS "Year", (o.month_id%100) AS "Month",
           o.category AS "Category", o.amount AS "Amount"
    FROM fact_opex_monthly o JOIN dim_region r ON r.region_id = o.region_id
    ORDER BY r.region_name, o.month_id
""")
opex.to_csv(f"{OUT}/opex.csv", index=False)
print("opex", opex.shape)

# 4. Inventory summary (region x department x month) -- pre-aggregated from the
#    115K-row store-level inventory fact so the workbook stays a manageable size.
inv_summary = q("""
    WITH inv_value AS (
        SELECT s.region_id, p.department_id, i.month_id,
               -- SUM (not AVG) gives the total dollar inventory held across all
               -- store-product rows in this region/department/month
               SUM((i.beginning_units + i.ending_units)/2.0 * p.unit_cost) AS total_inv_value,
               SUM(i.units_sold * p.unit_cost) AS monthly_cogs
        FROM fact_inventory_monthly i
        JOIN dim_store s ON s.store_id = i.store_id
        JOIN dim_product p ON p.product_id = i.product_id
        GROUP BY s.region_id, p.department_id, i.month_id
    )
    SELECT r.region_name AS "Region", dep.department_name AS "Department",
           (iv.month_id/100) AS "Year", (iv.month_id%100) AS "Month",
           ROUND(iv.monthly_cogs,2) AS "MonthlyCOGS", ROUND(iv.total_inv_value,2) AS "TotalInventoryValue"
    FROM inv_value iv
    JOIN dim_region r ON r.region_id = iv.region_id
    JOIN dim_department dep ON dep.department_id = iv.department_id
    ORDER BY r.region_name, dep.department_name, iv.month_id
""")
inv_summary.to_csv(f"{OUT}/inventory_summary.csv", index=False)
print("inventory_summary", inv_summary.shape)

conn.close()
print("done")
