"""
Synthetic data generator — NorthStar Retail Group
Executive Business Performance 360 Dashboard project

Generates a realistic multi-department, multi-region retail dataset:
  - dim_region, dim_store, dim_department, dim_product, dim_customer, dim_date
  - fact_sales            (~100K+ transaction line items, 24 months)
  - fact_inventory_monthly
  - fact_opex_monthly

A deliberate margin-leakage pattern is injected into the "South" region
(heavier discounting + higher unit cost inflation) so the finished dashboard
has a real, reproducible "14% regional margin leakage" story to report on.
"""
import numpy as np
import pandas as pd
from faker import Faker
from datetime import date
import random

SEED = 42
random.seed(SEED)
np.random.seed(SEED)
fake = Faker()
Faker.seed(SEED)

OUT = "/home/claude/executive-dashboard/data"

# ---------------------------------------------------------------- dim_region
regions = pd.DataFrame({
    "region_id": [1, 2, 3, 4, 5],
    "region_name": ["North", "South", "East", "West", "Central"],
})

# ---------------------------------------------------------------- dim_department
departments = pd.DataFrame({
    "department_id": [1, 2, 3, 4, 5],
    "department_name": ["Electronics", "Apparel", "Home & Living", "Grocery", "Beauty & Personal Care"],
})

# ---------------------------------------------------------------- dim_store
STORES_PER_REGION = 8
store_rows = []
store_id = 1
for _, r in regions.iterrows():
    for i in range(STORES_PER_REGION):
        store_rows.append({
            "store_id": store_id,
            "store_name": f"NorthStar {r['region_name']} #{i+1:02d}",
            "region_id": r["region_id"],
            "city": fake.city(),
            "store_type": random.choice(["Flagship", "Standard", "Standard", "Express"]),
            "open_date": fake.date_between(start_date=date(2016, 1, 1), end_date=date(2022, 12, 31)),
        })
        store_id += 1
stores = pd.DataFrame(store_rows)

# ---------------------------------------------------------------- dim_product
DEPT_CATEGORIES = {
    1: ["Mobile Accessories", "Audio", "Laptops & Tablets", "Smart Home", "Wearables"],
    2: ["Men's Wear", "Women's Wear", "Kids Wear", "Footwear", "Accessories"],
    3: ["Furniture", "Kitchenware", "Decor", "Bedding", "Storage"],
    4: ["Packaged Foods", "Beverages", "Snacks", "Fresh Produce", "Dairy"],
    5: ["Skincare", "Haircare", "Makeup", "Fragrance", "Personal Care"],
}
DEPT_PRICE_RANGE = {  # (unit_cost_low, unit_cost_high, markup_low, markup_high)
    1: (25, 400, 1.35, 1.9),
    2: (8, 60, 1.6, 2.4),
    3: (15, 250, 1.4, 2.0),
    4: (2, 15, 1.15, 1.4),
    5: (5, 45, 1.5, 2.2),
}
N_PRODUCTS_PER_DEPT = 60
product_rows = []
product_id = 1
for dept_id, cats in DEPT_CATEGORIES.items():
    lo_cost, hi_cost, lo_mk, hi_mk = DEPT_PRICE_RANGE[dept_id]
    for i in range(N_PRODUCTS_PER_DEPT):
        cat = random.choice(cats)
        unit_cost = round(np.random.uniform(lo_cost, hi_cost), 2)
        markup = np.random.uniform(lo_mk, hi_mk)
        product_rows.append({
            "product_id": product_id,
            "product_name": f"{cat} Item {i+1:03d}",
            "department_id": dept_id,
            "category": cat,
            "unit_cost": unit_cost,
            "unit_price": round(unit_cost * markup, 2),
        })
        product_id += 1
products = pd.DataFrame(product_rows)

# ---------------------------------------------------------------- dim_date (monthly grain helper + daily)
date_range = pd.date_range("2024-01-01", "2025-12-31", freq="D")
dim_date = pd.DataFrame({"date": date_range})
dim_date["date_id"] = dim_date["date"].dt.strftime("%Y%m%d").astype(int)
dim_date["year"] = dim_date["date"].dt.year
dim_date["month"] = dim_date["date"].dt.month
dim_date["quarter"] = dim_date["date"].dt.quarter
dim_date["month_name"] = dim_date["date"].dt.strftime("%b")
dim_date["month_id"] = dim_date["date"].dt.strftime("%Y%m").astype(int)

months = sorted(dim_date["month_id"].unique())

# ---------------------------------------------------------------- dim_customer
N_CUSTOMERS = 6000
cust_rows = []
for cid in range(1, N_CUSTOMERS + 1):
    region_id = random.choice(regions["region_id"].tolist())
    join_date = fake.date_between(start_date=date(2022, 1, 1), end_date=date(2025, 6, 30))
    membership = random.choices(["Standard", "Silver", "Gold"], weights=[0.55, 0.3, 0.15])[0]
    # South region + Standard tier churn more (feeds the churn-rate KPI story)
    base_churn_prob = 0.16
    if region_id == 2:
        base_churn_prob += 0.07
    if membership == "Standard":
        base_churn_prob += 0.05
    is_churned = np.random.rand() < base_churn_prob
    churn_date = None
    if is_churned:
        churn_date = fake.date_between(start_date=max(join_date, date(2024, 1, 1)), end_date=date(2025, 12, 31))
    cust_rows.append({
        "customer_id": cid,
        "customer_name": fake.name(),
        "region_id": region_id,
        "membership_type": membership,
        "join_date": join_date,
        "churn_date": churn_date,
        "is_active": churn_date is None,
    })
customers = pd.DataFrame(cust_rows)

print(f"Dims built: {len(regions)} regions, {len(stores)} stores, {len(departments)} departments, "
      f"{len(products)} products, {len(customers)} customers, {len(dim_date)} calendar days")

regions.to_csv(f"{OUT}/dim_region.csv", index=False)
departments.to_csv(f"{OUT}/dim_department.csv", index=False)
stores.to_csv(f"{OUT}/dim_store.csv", index=False)
products.to_csv(f"{OUT}/dim_product.csv", index=False)
customers.to_csv(f"{OUT}/dim_customer.csv", index=False)
dim_date.to_csv(f"{OUT}/dim_date.csv", index=False)

# ================================================================= fact_sales
# Target: 100K+ line items across 24 months, seasonally weighted, with
# South region running a deliberately higher discount + cost-inflation
# pattern that produces the ~14% margin leakage finding.

TARGET_ROWS = 105_000
store_ids = stores["store_id"].tolist()
store_region_map = dict(zip(stores["store_id"], stores["region_id"]))
product_ids = products["product_id"].tolist()
product_lookup = products.set_index("product_id")
customer_ids = customers["customer_id"].tolist()
customer_region_map = dict(zip(customers["customer_id"], customers["region_id"]))

# seasonal weight per month (Nov/Dec holiday bump, Feb lull)
month_seasonality = {1: .85, 2: .8, 3: .9, 4: .95, 5: 1.0, 6: 1.0, 7: .95, 8: .95,
                     9: 1.0, 10: 1.05, 11: 1.25, 12: 1.35}

rows_per_month = []
n_months = len(months)
base_per_month = TARGET_ROWS // n_months
for m in months:
    mm = int(str(m)[4:6])
    yy = int(str(m)[:4])
    growth_factor = 1.0 + (0.10 if yy == 2025 else 0.0)  # ~10% YoY growth built in
    rows_per_month.append(int(base_per_month * month_seasonality[mm] * growth_factor))

sales_rows = []
txn_id = 1
dates_by_month = dim_date.groupby("month_id")["date"].apply(list).to_dict()

for m, n_rows in zip(months, rows_per_month):
    day_choices = dates_by_month[m]
    for _ in range(n_rows):
        store_id = random.choice(store_ids)
        region_id = store_region_map[store_id]
        product_id = random.choice(product_ids)
        prod = product_lookup.loc[product_id]
        customer_id = random.choice(customer_ids) if np.random.rand() > 0.08 else None  # 8% walk-in/no loyalty
        qty = np.random.choice([1, 1, 1, 2, 2, 3, 4], p=[0.35, 0.2, 0.1, 0.15, 0.1, 0.06, 0.04])

        unit_price = prod["unit_price"]
        unit_cost = prod["unit_cost"]

        # South region: bigger discounts + creeping cost inflation -> margin leakage
        # (tuned so South's gross margin lands ~14% below the company average)
        if region_id == 2:
            discount_pct = np.random.choice([0.0, 0.05, 0.10, 0.15, 0.20], p=[0.125, 0.275, 0.30, 0.175, 0.125])
            cost_inflation = np.random.uniform(1.025, 1.10)
        else:
            discount_pct = np.random.choice([0.0, 0.05, 0.10, 0.15], p=[0.4, 0.3, 0.2, 0.1])
            cost_inflation = np.random.uniform(0.98, 1.05)

        eff_unit_cost = round(unit_cost * cost_inflation, 2)
        gross_revenue = round(unit_price * qty, 2)
        net_revenue = round(gross_revenue * (1 - discount_pct), 2)
        total_cost = round(eff_unit_cost * qty, 2)

        txn_date = random.choice(day_choices)
        sales_rows.append({
            "transaction_id": txn_id,
            "date_id": int(txn_date.strftime("%Y%m%d")),
            "store_id": store_id,
            "region_id": region_id,
            "product_id": product_id,
            "department_id": int(prod["department_id"]),
            "customer_id": customer_id,
            "quantity": int(qty),
            "unit_price": unit_price,
            "unit_cost": eff_unit_cost,
            "discount_pct": discount_pct,
            "gross_revenue": gross_revenue,
            "net_revenue": net_revenue,
            "total_cost": total_cost,
        })
        txn_id += 1

fact_sales = pd.DataFrame(sales_rows)
fact_sales["customer_id"] = fact_sales["customer_id"].astype("Int64")
print(f"fact_sales rows: {len(fact_sales):,}")
fact_sales.to_csv(f"{OUT}/fact_sales.csv", index=False)

# ========================================================= fact_inventory_monthly
# Build a fast lookup of units sold per (store, product, month) from fact_sales,
# then size each SKU's starting inventory relative to its own sell-through rate
# so turnover ratios land in a realistic band instead of being arbitrary.
sales_lookup = fact_sales.groupby(["store_id", "product_id", (fact_sales["date_id"] // 100)])["quantity"].sum()
sales_lookup.index.set_names(["store_id", "product_id", "month_id"], inplace=True)
sales_lookup = sales_lookup.to_dict()

inv_rows = []
inv_id = 1
for store_id in store_ids:
    region_id = store_region_map[store_id]
    # each store carries a random assortment of ~120 products
    assortment = random.sample(product_ids, 120)
    for product_id in assortment:
        # estimate this SKU's average monthly sell-through at this store across
        # the full window, to size a realistic starting stock level
        total_sold = sum(sales_lookup.get((store_id, product_id, m), 0) for m in months)
        avg_monthly_sold = total_sold / len(months)
        # target ~1.2-2.2 months of supply on hand (~6-10x annual turnover)
        target_months_supply = np.random.uniform(1.2, 2.2)
        beginning = max(3, round(avg_monthly_sold * target_months_supply))
        for m in months:
            units_sold = sales_lookup.get((store_id, product_id, m), 0)
            # replenishment keeps stock roughly stable month to month
            restock_factor = np.random.uniform(0.9, 1.05)
            units_received = int(round(units_sold * restock_factor)) + np.random.randint(0, 2)
            ending = max(0, beginning + units_received - units_sold)
            inv_rows.append({
                "inventory_id": inv_id,
                "store_id": store_id,
                "product_id": product_id,
                "month_id": m,
                "beginning_units": beginning,
                "units_received": units_received,
                "units_sold": int(units_sold),
                "ending_units": ending,
            })
            beginning = ending
            inv_id += 1

fact_inventory = pd.DataFrame(inv_rows)
print(f"fact_inventory_monthly rows: {len(fact_inventory):,}")
fact_inventory.to_csv(f"{OUT}/fact_inventory_monthly.csv", index=False)

# ============================================================= fact_opex_monthly
OPEX_CATEGORIES = ["Salaries & Wages", "Rent & Utilities", "Marketing", "Logistics", "Depreciation & Amortization", "Other Admin"]
# per-store, per-month baseline (calibrated so company EBITDA margin lands ~15-20%)
BASE_OPEX = {"Salaries & Wages": 2400, "Rent & Utilities": 1100, "Marketing": 500,
             "Logistics": 400, "Depreciation & Amortization": 350, "Other Admin": 250}
opex_rows = []
opex_id = 1
for _, r in regions.iterrows():
    region_id = r["region_id"]
    n_stores_in_region = STORES_PER_REGION
    for m in months:
        yy = int(str(m)[:4])
        infl = 1.06 if yy == 2025 else 1.0
        for cat in OPEX_CATEGORIES:
            amount = round(BASE_OPEX[cat] * n_stores_in_region * infl * np.random.uniform(0.95, 1.05), 2)
            opex_rows.append({
                "opex_id": opex_id,
                "region_id": region_id,
                "month_id": m,
                "category": cat,
                "amount": amount,
            })
            opex_id += 1

fact_opex = pd.DataFrame(opex_rows)
print(f"fact_opex_monthly rows: {len(fact_opex):,}")
fact_opex.to_csv(f"{OUT}/fact_opex_monthly.csv", index=False)

print("\nAll CSVs written to", OUT)
