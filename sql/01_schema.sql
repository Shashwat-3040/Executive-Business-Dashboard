-- ============================================================================
-- Executive Business Performance 360 Dashboard
-- NorthStar Retail Group -- Data Warehouse Schema (star schema)
-- ============================================================================

DROP TABLE IF EXISTS fact_opex_monthly CASCADE;
DROP TABLE IF EXISTS fact_inventory_monthly CASCADE;
DROP TABLE IF EXISTS fact_sales CASCADE;
DROP TABLE IF EXISTS dim_customer CASCADE;
DROP TABLE IF EXISTS dim_product CASCADE;
DROP TABLE IF EXISTS dim_store CASCADE;
DROP TABLE IF EXISTS dim_department CASCADE;
DROP TABLE IF EXISTS dim_region CASCADE;
DROP TABLE IF EXISTS dim_date CASCADE;

CREATE TABLE dim_region (
    region_id     INTEGER PRIMARY KEY,
    region_name   TEXT NOT NULL
);

CREATE TABLE dim_department (
    department_id     INTEGER PRIMARY KEY,
    department_name   TEXT NOT NULL
);

CREATE TABLE dim_store (
    store_id     INTEGER PRIMARY KEY,
    store_name   TEXT NOT NULL,
    region_id    INTEGER REFERENCES dim_region(region_id),
    city         TEXT,
    store_type   TEXT,
    open_date    DATE
);

CREATE TABLE dim_product (
    product_id     INTEGER PRIMARY KEY,
    product_name   TEXT NOT NULL,
    department_id  INTEGER REFERENCES dim_department(department_id),
    category       TEXT,
    unit_cost      NUMERIC(10,2),
    unit_price     NUMERIC(10,2)
);

CREATE TABLE dim_customer (
    customer_id       INTEGER PRIMARY KEY,
    customer_name     TEXT,
    region_id         INTEGER REFERENCES dim_region(region_id),
    membership_type   TEXT,
    join_date         DATE,
    churn_date        DATE,
    is_active         BOOLEAN
);

CREATE TABLE dim_date (
    date_id      INTEGER PRIMARY KEY,
    date         DATE NOT NULL,
    year         INTEGER,
    month        INTEGER,
    quarter      INTEGER,
    month_name   TEXT,
    month_id     INTEGER
);

CREATE TABLE fact_sales (
    transaction_id   BIGINT PRIMARY KEY,
    date_id          INTEGER REFERENCES dim_date(date_id),
    store_id         INTEGER REFERENCES dim_store(store_id),
    region_id        INTEGER REFERENCES dim_region(region_id),
    product_id       INTEGER REFERENCES dim_product(product_id),
    department_id    INTEGER REFERENCES dim_department(department_id),
    customer_id      INTEGER REFERENCES dim_customer(customer_id),
    quantity         INTEGER,
    unit_price       NUMERIC(10,2),
    unit_cost        NUMERIC(10,2),
    discount_pct     NUMERIC(5,4),
    gross_revenue    NUMERIC(12,2),
    net_revenue      NUMERIC(12,2),
    total_cost       NUMERIC(12,2)
);

CREATE TABLE fact_inventory_monthly (
    inventory_id      BIGINT PRIMARY KEY,
    store_id          INTEGER REFERENCES dim_store(store_id),
    product_id        INTEGER REFERENCES dim_product(product_id),
    month_id          INTEGER,
    beginning_units   INTEGER,
    units_received    INTEGER,
    units_sold        INTEGER,
    ending_units      INTEGER
);

CREATE TABLE fact_opex_monthly (
    opex_id      BIGINT PRIMARY KEY,
    region_id    INTEGER REFERENCES dim_region(region_id),
    month_id     INTEGER,
    category     TEXT,
    amount       NUMERIC(12,2)
);

CREATE INDEX idx_sales_date ON fact_sales(date_id);
CREATE INDEX idx_sales_store ON fact_sales(store_id);
CREATE INDEX idx_sales_region ON fact_sales(region_id);
CREATE INDEX idx_sales_product ON fact_sales(product_id);
CREATE INDEX idx_sales_customer ON fact_sales(customer_id);
CREATE INDEX idx_inventory_month ON fact_inventory_monthly(month_id);
CREATE INDEX idx_opex_month ON fact_opex_monthly(month_id);
