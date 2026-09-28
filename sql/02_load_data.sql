-- ============================================================================
-- Load synthetic CSVs into the warehouse (run after 01_schema.sql)
-- Paths assume the /data folder sits alongside /sql at the project root.
-- ============================================================================

\copy dim_region FROM 'data/dim_region.csv' WITH (FORMAT csv, HEADER true);
\copy dim_department FROM 'data/dim_department.csv' WITH (FORMAT csv, HEADER true);
\copy dim_store FROM 'data/dim_store.csv' WITH (FORMAT csv, HEADER true);
\copy dim_product FROM 'data/dim_product.csv' WITH (FORMAT csv, HEADER true);
\copy dim_customer FROM 'data/dim_customer.csv' WITH (FORMAT csv, HEADER true);
\copy dim_date(date, date_id, year, month, quarter, month_name, month_id) FROM 'data/dim_date.csv' WITH (FORMAT csv, HEADER true);
\copy fact_sales FROM 'data/fact_sales.csv' WITH (FORMAT csv, HEADER true);
\copy fact_inventory_monthly FROM 'data/fact_inventory_monthly.csv' WITH (FORMAT csv, HEADER true);
\copy fact_opex_monthly FROM 'data/fact_opex_monthly.csv' WITH (FORMAT csv, HEADER true);

-- sanity checks
SELECT 'dim_region' AS tbl, COUNT(*) FROM dim_region
UNION ALL SELECT 'dim_store', COUNT(*) FROM dim_store
UNION ALL SELECT 'dim_product', COUNT(*) FROM dim_product
UNION ALL SELECT 'dim_customer', COUNT(*) FROM dim_customer
UNION ALL SELECT 'dim_date', COUNT(*) FROM dim_date
UNION ALL SELECT 'fact_sales', COUNT(*) FROM fact_sales
UNION ALL SELECT 'fact_inventory_monthly', COUNT(*) FROM fact_inventory_monthly
UNION ALL SELECT 'fact_opex_monthly', COUNT(*) FROM fact_opex_monthly;
