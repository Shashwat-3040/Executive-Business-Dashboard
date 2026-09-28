-- ============================================================================
-- Executive KPI Queries
-- Powers: EBITDA margin, YoY revenue growth, churn rate, inventory turnover,
--         and the regional margin-leakage investigation.
-- ============================================================================

-- ---------------------------------------------------------------------------
-- 1. EBITDA & EBITDA MARGIN by region and year
--    EBITDA = Net Revenue - COGS - OpEx (excl. D&A, which is added back)
-- ---------------------------------------------------------------------------
WITH revenue_cogs AS (
    SELECT
        r.region_name,
        d.year,
        SUM(f.net_revenue) AS net_revenue,
        SUM(f.total_cost)  AS cogs
    FROM fact_sales f
    JOIN dim_region r ON r.region_id = f.region_id
    JOIN dim_date d   ON d.date_id  = f.date_id
    GROUP BY r.region_name, d.year
),
opex_agg AS (
    SELECT
        r.region_name,
        (o.month_id / 100) AS year,
        SUM(o.amount) FILTER (WHERE o.category <> 'Depreciation & Amortization') AS cash_opex,
        SUM(o.amount) FILTER (WHERE o.category = 'Depreciation & Amortization')  AS d_and_a
    FROM fact_opex_monthly o
    JOIN dim_region r ON r.region_id = o.region_id
    GROUP BY r.region_name, (o.month_id / 100)
)
SELECT
    rc.region_name,
    rc.year,
    rc.net_revenue,
    rc.cogs,
    oa.cash_opex,
    oa.d_and_a,
    ROUND(rc.net_revenue - rc.cogs - oa.cash_opex, 2) AS ebitda,
    ROUND(100.0 * (rc.net_revenue - rc.cogs - oa.cash_opex) / NULLIF(rc.net_revenue, 0), 2) AS ebitda_margin_pct
FROM revenue_cogs rc
JOIN opex_agg oa ON oa.region_name = rc.region_name AND oa.year = rc.year
ORDER BY rc.year, rc.region_name;


-- ---------------------------------------------------------------------------
-- 2. YoY REVENUE GROWTH (company-wide and by region)
-- ---------------------------------------------------------------------------
WITH yearly_rev AS (
    SELECT r.region_name, d.year, SUM(f.net_revenue) AS net_revenue
    FROM fact_sales f
    JOIN dim_region r ON r.region_id = f.region_id
    JOIN dim_date d   ON d.date_id  = f.date_id
    GROUP BY r.region_name, d.year
)
SELECT
    region_name,
    year,
    net_revenue,
    LAG(net_revenue) OVER (PARTITION BY region_name ORDER BY year) AS prior_year_revenue,
    ROUND(100.0 * (net_revenue - LAG(net_revenue) OVER (PARTITION BY region_name ORDER BY year))
          / NULLIF(LAG(net_revenue) OVER (PARTITION BY region_name ORDER BY year), 0), 2) AS yoy_growth_pct
FROM yearly_rev
ORDER BY region_name, year;

-- company-wide roll-up
WITH company_yearly AS (
    SELECT d.year, SUM(f.net_revenue) AS net_revenue
    FROM fact_sales f JOIN dim_date d ON d.date_id = f.date_id
    GROUP BY d.year
)
SELECT year, net_revenue,
       ROUND(100.0 * (net_revenue - LAG(net_revenue) OVER (ORDER BY year))
             / NULLIF(LAG(net_revenue) OVER (ORDER BY year), 0), 2) AS yoy_growth_pct
FROM company_yearly ORDER BY year;


-- ---------------------------------------------------------------------------
-- 3. CHURN RATE by region and membership tier
--    Churn rate = customers who churned in period / active customers at period start
-- ---------------------------------------------------------------------------
SELECT
    r.region_name,
    c.membership_type,
    COUNT(*) AS total_customers,
    COUNT(*) FILTER (WHERE c.churn_date IS NOT NULL) AS churned_customers,
    ROUND(100.0 * COUNT(*) FILTER (WHERE c.churn_date IS NOT NULL) / COUNT(*), 2) AS churn_rate_pct
FROM dim_customer c
JOIN dim_region r ON r.region_id = c.region_id
GROUP BY r.region_name, c.membership_type
ORDER BY churn_rate_pct DESC;

-- overall churn rate
SELECT
    ROUND(100.0 * COUNT(*) FILTER (WHERE churn_date IS NOT NULL) / COUNT(*), 2) AS overall_churn_rate_pct
FROM dim_customer;


-- ---------------------------------------------------------------------------
-- 4. INVENTORY TURNOVER by region / department
--    Turnover = COGS (annualized) / Average Inventory Value
-- ---------------------------------------------------------------------------
WITH inv_value AS (
    SELECT
        s.region_id,
        p.department_id,
        i.month_id,
        -- SUM (not AVG) gives the total dollar inventory held across all
        -- store-product rows in this region/department/month
        SUM((i.beginning_units + i.ending_units) / 2.0 * p.unit_cost) AS total_inv_value,
        SUM(i.units_sold * p.unit_cost) AS monthly_cogs
    FROM fact_inventory_monthly i
    JOIN dim_store s ON s.store_id = i.store_id
    JOIN dim_product p ON p.product_id = i.product_id
    GROUP BY s.region_id, p.department_id, i.month_id
)
SELECT
    r.region_name,
    dep.department_name,
    ROUND(SUM(iv.monthly_cogs), 2) AS two_year_cogs,
    ROUND(AVG(iv.total_inv_value), 2) AS avg_inventory_value,
    -- data spans 24 months, so /2 annualizes the ratio
    ROUND((SUM(iv.monthly_cogs) / 2.0) / NULLIF(AVG(iv.total_inv_value), 0), 2) AS annual_inventory_turnover_ratio
FROM inv_value iv
JOIN dim_region r ON r.region_id = iv.region_id
JOIN dim_department dep ON dep.department_id = iv.department_id
GROUP BY r.region_name, dep.department_name
ORDER BY inventory_turnover_ratio ASC;


-- ---------------------------------------------------------------------------
-- 5. REGIONAL MARGIN LEAKAGE — the headline finding
--    Compares each region's gross margin % against the company average
-- ---------------------------------------------------------------------------
WITH region_margin AS (
    SELECT
        r.region_name,
        SUM(f.net_revenue) AS net_revenue,
        SUM(f.total_cost)  AS total_cost,
        ROUND(100.0 * (SUM(f.net_revenue) - SUM(f.total_cost)) / NULLIF(SUM(f.net_revenue), 0), 2) AS gross_margin_pct
    FROM fact_sales f
    JOIN dim_region r ON r.region_id = f.region_id
    GROUP BY r.region_name
),
company_avg AS (
    SELECT AVG(gross_margin_pct) AS avg_margin_pct FROM region_margin
)
SELECT
    rm.region_name,
    rm.net_revenue,
    rm.total_cost,
    rm.gross_margin_pct,
    ca.avg_margin_pct,
    ROUND(rm.gross_margin_pct - ca.avg_margin_pct, 2) AS margin_gap_vs_avg_pp,
    ROUND(100.0 * (ca.avg_margin_pct - rm.gross_margin_pct) / NULLIF(ca.avg_margin_pct, 0), 2) AS pct_leakage_vs_avg
FROM region_margin rm CROSS JOIN company_avg ca
ORDER BY gross_margin_pct ASC;


-- ---------------------------------------------------------------------------
-- 6. MASTER EXPORT — one flat query per month/region, used to feed
--    the Excel Power Query refresh (Data tab -> PivotTables)
-- ---------------------------------------------------------------------------
SELECT
    d.year, d.month, d.month_name, r.region_name, s.store_name, s.store_type,
    dep.department_name, p.category, p.product_name,
    f.quantity, f.unit_price, f.unit_cost, f.discount_pct,
    f.gross_revenue, f.net_revenue, f.total_cost,
    ROUND(f.net_revenue - f.total_cost, 2) AS gross_profit
FROM fact_sales f
JOIN dim_date d ON d.date_id = f.date_id
JOIN dim_region r ON r.region_id = f.region_id
JOIN dim_store s ON s.store_id = f.store_id
JOIN dim_product p ON p.product_id = f.product_id
JOIN dim_department dep ON dep.department_id = f.department_id;
