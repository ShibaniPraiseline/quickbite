
import pandas as pd
from pathlib import Path

# --------------------------------------------------
# 1. FILE PATHS
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = BASE_DIR / "cleaned_data" / "Quick_Bite_Combined_Data.xlsx"
OUTPUT_DIR = BASE_DIR / "analysis"
OUTPUT_FILE = OUTPUT_DIR / "Quick_Bite_Analysis.xlsx"

# Create analysis folder if it does not exist
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------
# 2. LOAD COMBINED DATA
# --------------------------------------------------

print("Loading combined dataset...")

df = pd.read_excel(INPUT_FILE)

print("Dataset loaded successfully!")
print("Original shape:", df.shape)


# --------------------------------------------------
# 3. DATA PREPARATION
# --------------------------------------------------

# Convert numeric columns
df["order_amount"] = pd.to_numeric(
    df["order_amount"], errors="coerce"
)

df["delivery_minutes"] = pd.to_numeric(
    df["delivery_minutes"], errors="coerce"
)

df["rating"] = pd.to_numeric(
    df["rating"], errors="coerce"
)

# Standardize order status
df["order_status"] = (
    df["order_status"]
    .astype("string")
    .str.strip()
    .str.lower()
)

# Standardize text columns used in grouping
for column in ["area_name", "city", "cuisine", "vehicle_type", "order_day"]:
    if column in df.columns:
        df[column] = df[column].replace(r"^\s*$", pd.NA, regex=True)

# Consider only delivered orders for delivery performance
delivered = df[
    df["order_status"] == "delivered"
].copy()

# Calculate late delivery:
# More than 45 minutes = Late
delivered["is_late"] = (
    delivered["delivery_minutes"] > 45
)


# --------------------------------------------------
# 4. AREA PERFORMANCE
# --------------------------------------------------

print("\nGenerating Area Performance...")

# Business performance:
# Use ALL orders to calculate total orders and revenue.
area_business = (
    df.groupby(
        ["area_name", "city"],
        dropna=False
    )
    .agg(
        total_orders=("order_id", "nunique"),
        total_revenue=("order_amount", "sum")
    )
    .reset_index()
)

# Delivery performance:
# Use only delivered orders.
area_delivery = (
    delivered.groupby(
        ["area_name", "city"],
        dropna=False
    )
    .agg(
        delivered_orders=("order_id", "nunique"),
        avg_delivery_time=("delivery_minutes", "mean"),
        avg_rating=("rating", "mean"),
        late_deliveries=("is_late", "sum"),
        orders_with_delivery_time=("delivery_minutes", "count")
    )
    .reset_index()
)

# Merge business and delivery performance
area_performance = area_business.merge(
    area_delivery,
    on=["area_name", "city"],
    how="left"
)

# Fill missing count values
count_columns = [
    "delivered_orders",
    "late_deliveries",
    "orders_with_delivery_time"
]

area_performance[count_columns] = (
    area_performance[count_columns]
    .fillna(0)
)

# Calculate late delivery percentage
area_performance["late_percentage"] = (
    area_performance["late_deliveries"]
    .div(
        area_performance["orders_with_delivery_time"]
        .replace(0, pd.NA)
    )
    * 100
)

# Rank areas based on revenue
area_performance = area_performance.sort_values(
    by="total_revenue",
    ascending=False
).reset_index(drop=True)

area_performance.insert(
    0,
    "revenue_rank",
    area_performance["total_revenue"]
    .rank(method="dense", ascending=False)
    .astype("Int64")
)

# Round numerical values
area_performance["total_revenue"] = (
    area_performance["total_revenue"].round(2)
)

area_performance["avg_delivery_time"] = (
    area_performance["avg_delivery_time"].round(2)
)

area_performance["avg_rating"] = (
    area_performance["avg_rating"].round(2)
)

area_performance["late_percentage"] = (
    area_performance["late_percentage"].round(2)
)


# --------------------------------------------------
# 5. TOP CUISINE BY DAY
# --------------------------------------------------

print("Generating Top Cuisine by Day...")

# Count orders for each cuisine on each day
cuisine_counts = (
    df.dropna(subset=["order_day", "cuisine"])
    .groupby(["order_day", "cuisine"])["order_id"]
    .nunique()
    .reset_index(name="order_count")
)

# Calculate total orders for each day
daily_totals = (
    df.dropna(subset=["order_day"])
    .groupby("order_day")["order_id"]
    .nunique()
    .reset_index(name="total_day_orders")
)

# Merge cuisine counts with daily totals
cuisine_analysis = cuisine_counts.merge(
    daily_totals,
    on="order_day",
    how="left"
)

# Calculate percentage share
cuisine_analysis["order_share_percentage"] = (
    cuisine_analysis["order_count"]
    / cuisine_analysis["total_day_orders"]
    * 100
)

# Find the highest order count for each day
cuisine_analysis["max_orders"] = (
    cuisine_analysis.groupby("order_day")["order_count"]
    .transform("max")
)

# Keep only the top cuisine(s) for each day.
# If multiple cuisines tie, all tied cuisines are retained.
top_cuisine = cuisine_analysis[
    cuisine_analysis["order_count"]
    == cuisine_analysis["max_orders"]
].copy()

# Sort by day
top_cuisine = top_cuisine.sort_values(
    by=["order_day", "cuisine"]
).reset_index(drop=True)

# Round percentage
top_cuisine["order_share_percentage"] = (
    top_cuisine["order_share_percentage"].round(2)
)

# Remove helper column
top_cuisine = top_cuisine.drop(
    columns=["max_orders"]
)


# --------------------------------------------------
# 6. VEHICLE PERFORMANCE
# --------------------------------------------------

print("Generating Vehicle Performance...")

# Only delivered orders are considered
vehicle_performance = (
    delivered.groupby("vehicle_type", dropna=False)
    .agg(
        total_delivered_orders=("order_id", "nunique"),
        avg_delivery_time=("delivery_minutes", "mean"),
        median_delivery_time=("delivery_minutes", "median"),
        avg_rating=("rating", "mean"),
        late_deliveries=("is_late", "sum"),
        orders_with_delivery_time=("delivery_minutes", "count")
    )
    .reset_index()
)

# Replace missing vehicle names with Unknown
vehicle_performance["vehicle_type"] = (
    vehicle_performance["vehicle_type"]
    .fillna("Unknown")
)

# Calculate late delivery percentage
vehicle_performance["late_percentage"] = (
    vehicle_performance["late_deliveries"]
    .div(
        vehicle_performance["orders_with_delivery_time"]
        .replace(0, pd.NA)
    )
    * 100
)

# Round numerical values
numeric_columns = [
    "avg_delivery_time",
    "median_delivery_time",
    "avg_rating",
    "late_percentage"
]

vehicle_performance[numeric_columns] = (
    vehicle_performance[numeric_columns].round(2)
)

# Sort by average delivery time
vehicle_performance = vehicle_performance.sort_values(
    by="avg_delivery_time",
    ascending=True,
    na_position="last"
).reset_index(drop=True)


# --------------------------------------------------
# 7. SAVE ANALYSIS OUTPUT
# --------------------------------------------------

print("\nSaving analysis workbook...")

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    area_performance.to_excel(
        writer,
        sheet_name="Area Performance",
        index=False
    )

    top_cuisine.to_excel(
        writer,
        sheet_name="Top Cuisine by Day",
        index=False
    )

    vehicle_performance.to_excel(
        writer,
        sheet_name="Vehicle Performance",
        index=False
    )


# --------------------------------------------------
# 8. FINAL SUMMARY
# --------------------------------------------------

print("\nAnalysis completed successfully!")

print("Combined dataset rows:", len(df))
print("Delivered orders:", len(delivered))

print("\nArea Performance:")
print(area_performance.head())

print("\nTop Cuisine by Day:")
print(top_cuisine.head())

print("\nVehicle Performance:")
print(vehicle_performance)

print("\nOutput saved at:")
print(OUTPUT_FILE)
