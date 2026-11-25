
import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = BASE_DIR / "data" / "QuickBite_Assessment.xlsx"

OUTPUT_FOLDER = BASE_DIR / "cleaned_data"

OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)


areas = pd.read_excel(INPUT_FILE, sheet_name="areas")

restaurants = pd.read_excel(INPUT_FILE, sheet_name="restaurants")

delivery_partners = pd.read_excel(
    INPUT_FILE, sheet_name="delivery_partners"
)

customers = pd.read_excel(INPUT_FILE, sheet_name="customers")

orders = pd.read_excel(INPUT_FILE, sheet_name="orders")

feedback = pd.read_excel(INPUT_FILE, sheet_name="feedback")


print("All six sheets loaded successfully!")

print("\nOriginal Data Shape:")
print("Areas:", areas.shape)
print("Restaurants:", restaurants.shape)
print("Delivery Partners:", delivery_partners.shape)
print("Customers:", customers.shape)
print("Orders:", orders.shape)
print("Feedback:", feedback.shape)

areas = areas.drop_duplicates()

restaurants = restaurants.drop_duplicates()

delivery_partners = delivery_partners.drop_duplicates()

customers = customers.drop_duplicates()

orders = orders.drop_duplicates()

feedback = feedback.drop_duplicates()


areas = areas.drop_duplicates(subset="area_id", keep="first")

restaurants = restaurants.drop_duplicates(
    subset="restaurant_id", keep="first"
)

delivery_partners = delivery_partners.drop_duplicates(
    subset="partner_id", keep="first"
)

customers = customers.drop_duplicates(
    subset="customer_id", keep="first"
)

orders = orders.drop_duplicates(
    subset="order_id", keep="first"
)

feedback = feedback.drop_duplicates(
    subset="feedback_id", keep="first"
)


print("\nDuplicate records removed!")


areas = areas.dropna(
    subset=["area_id", "area_name", "city", "zone"]
)


restaurants = restaurants.dropna(
    subset=[
        "restaurant_id",
        "restaurant_name",
        "area_id",
        "cuisine"
    ]
)

delivery_partners = delivery_partners.dropna(
    subset=[
        "partner_id",
        "partner_name",
        "phone",
        "vehicle_type"
    ]
)


# Customers

customers = customers.dropna(
    subset=[
        "customer_id",
        "customer_name",
        "phone"
    ]
)


# Orders

orders = orders.dropna(
    subset=[
        "order_id",
        "customer_id",
        "restaurant_id",
        "order_datetime",
        "order_amount",
        "payment_mode",
        "order_status"
    ]
)


# Feedback

feedback = feedback.dropna(
    subset=[
        "feedback_id",
        "order_id",
        "feedback_date"
    ]
)


print("Mandatory missing values handled!")


# 5. REMOVE EXTRA SPACES FROM TEXT COLUMNS

all_tables = [
    areas,
    restaurants,
    delivery_partners,
    customers,
    orders,
    feedback
]

for df in all_tables:

    text_columns = df.select_dtypes(include=["object"]).columns

    for col in text_columns:

        df[col] = df[col].apply(
            lambda x: x.strip() if isinstance(x, str) else x
        )


# 6. STANDARDIZE CITY NAMES

city_mapping = {
    "blr": "Bengaluru",
    "bengaluru": "Bengaluru",
    "chennai": "Chennai",
    "madras": "Chennai",
    "hyd": "Hyderabad",
    "hyderabad": "Hyderabad",
    "mumbai": "Mumbai"
}

areas["city"] = (
    areas["city"]
    .str.lower()
    .map(city_mapping)
    .fillna(areas["city"])
)


# 7. STANDARDIZE CUISINE NAMES

cuisine_mapping = {
    "biriyani": "Biryani",
    "biryani": "Biryani",
    "n. indian": "North Indian",
    "north-indian": "North Indian",
    "north indian": "North Indian",
    "south-indian": "South Indian",
    "south indian": "South Indian",
    "dessert": "Desserts",
    "desserts": "Desserts",
    "italian - pizza": "Pizza",
    "pizza": "Pizza"
}

restaurants["cuisine"] = (
    restaurants["cuisine"]
    .str.lower()
    .map(cuisine_mapping)
    .fillna(restaurants["cuisine"].str.title())
)


# 8. STANDARDIZE VEHICLE TYPES

vehicle_mapping = {
    "bike": "Motorbike",
    "bi ke": "Motorbike",
    "motor bike": "Motorbike",
    "motorbike": "Motorbike",
    "bicycle": "Bicycle",
    "cycle": "Bicycle",
    "activa": "Scooter",
    "scooty": "Scooter",
    "scooter": "Scooter",
    "ev": "EV Scooter",
    "e-scooter": "EV Scooter",
    "ev scooter": "EV Scooter"
}

delivery_partners["vehicle_type"] = (
    delivery_partners["vehicle_type"]
    .str.lower()
    .map(vehicle_mapping)
    .fillna(delivery_partners["vehicle_type"].str.title())
)


# 9. STANDARDIZE PAYMENT MODES

payment_mapping = {
    "cash": "Cash",
    "cash on delivery": "Cash on Delivery",
    "cod": "Cash on Delivery",
    "upi": "UPI",
    "gpay": "UPI",
    "phonepe": "UPI",
    "card": "Card",
    "credit card": "Card",
    "debit card": "Card",
    "wallet": "Wallet",
    "paytm wallet": "Wallet",
    "amazon pay": "Amazon Pay"
}

orders["payment_mode"] = (
    orders["payment_mode"]
    .str.lower()
    .map(payment_mapping)
    .fillna(orders["payment_mode"].str.title())
)


# 10. STANDARDIZE ORDER STATUS

status_mapping = {
    "delivered": "Delivered",
    "completed": "Delivered",
    "cancelled": "Cancelled",
    "canceled": "Cancelled"
}

orders["order_status"] = (
    orders["order_status"]
    .str.lower()
    .map(status_mapping)
    .fillna(orders["order_status"].str.title())
)


# 11. STANDARDIZE COUPON CODES

orders["coupon_code"] = (
    orders["coupon_code"]
    .astype("string")
    .str.strip()
    .str.upper()
)


# 12. CONVERT DATE COLUMNS

orders["order_datetime"] = pd.to_datetime(
    orders["order_datetime"],
    errors="coerce",
    dayfirst=True
)

customers["signup_date"] = pd.to_datetime(
    customers["signup_date"],
    errors="coerce",
    dayfirst=True
)

restaurants["onboard_date"] = pd.to_datetime(
    restaurants["onboard_date"],
    errors="coerce",
    dayfirst=True
)

delivery_partners["join_date"] = pd.to_datetime(
    delivery_partners["join_date"],
    errors="coerce",
    dayfirst=True
)

feedback["feedback_date"] = pd.to_datetime(
    feedback["feedback_date"],
    errors="coerce",
    dayfirst=True
)


# 13. CONVERT NUMERIC COLUMNS

orders["order_amount"] = pd.to_numeric(
    orders["order_amount"],
    errors="coerce"
)

orders["delivery_minutes"] = pd.to_numeric(
    orders["delivery_minutes"],
    errors="coerce"
)

restaurants["rating"] = pd.to_numeric(
    restaurants["rating"],
    errors="coerce"
)

feedback["rating"] = pd.to_numeric(
    feedback["rating"],
    errors="coerce"
)


# Remove orders with invalid mandatory numeric values

orders = orders.dropna(
    subset=["order_datetime", "order_amount"]
)

orders = orders[orders["order_amount"] >= 0]

orders = orders[
    orders["delivery_minutes"].isna()
    | (orders["delivery_minutes"] >= 0)
]


# 14. SAVE CLEANED DATA INTO ONE EXCEL FILE

cleaned_file = (
    OUTPUT_FOLDER / "Quick_Bite_Data_Cleaning.xlsx"
)

with pd.ExcelWriter(
    cleaned_file,
    engine="openpyxl"
) as writer:

    areas.to_excel(
        writer,
        sheet_name="areas",
        index=False
    )

    restaurants.to_excel(
        writer,
        sheet_name="restaurants",
        index=False
    )

    delivery_partners.to_excel(
        writer,
        sheet_name="delivery_partners",
        index=False
    )

    customers.to_excel(
        writer,
        sheet_name="customers",
        index=False
    )

    orders.to_excel(
        writer,
        sheet_name="orders",
        index=False
    )

    feedback.to_excel(
        writer,
        sheet_name="feedback",
        index=False
    )


print("\nTask 1 completed!")
print("Cleaned Excel file saved:", cleaned_file)


# ==========================================
# TASK 2 - DATA INTEGRATION
# ==========================================


# 15. PREPARE FEEDBACK DATA

# One order can have multiple feedback records.
# Aggregate them before merging to avoid duplicate orders.

feedback_summary = feedback.groupby("order_id").agg(
    rating=("rating", "mean"),
    comment=(
        "comment",
        lambda x: "; ".join(
            x.dropna().astype(str).unique()
        )
    )
).reset_index()


# 16. MERGE ORDERS WITH CUSTOMERS

combined = orders.merge(
    customers[
        ["customer_id", "customer_name"]
    ],
    on="customer_id",
    how="inner",
    validate="many_to_one"
)


# 17. MERGE RESTAURANT DETAILS

combined = combined.merge(
    restaurants[
        [
            "restaurant_id",
            "restaurant_name",
            "area_id",
            "cuisine"
        ]
    ],
    on="restaurant_id",
    how="inner",
    validate="many_to_one"
)


# 18. MERGE AREA DETAILS

combined = combined.merge(
    areas[
        ["area_id", "area_name", "city"]
    ],
    on="area_id",
    how="inner",
    validate="many_to_one"
)


# 19. MERGE DELIVERY PARTNER DETAILS

# Partner information is optional.
# Keep orders even when partner_id is missing.

combined = combined.merge(
    delivery_partners[
        [
            "partner_id",
            "partner_name",
            "vehicle_type"
        ]
    ],
    on="partner_id",
    how="left",
    validate="many_to_one"
)


# 20. MERGE FEEDBACK DETAILS

combined = combined.merge(
    feedback_summary,
    on="order_id",
    how="left",
    validate="many_to_one"
)


# 21. CREATE ORDER DATE, TIME AND DAY

combined["order_date"] = (
    combined["order_datetime"].dt.date
)

combined["order_time"] = (
    combined["order_datetime"].dt.strftime("%H:%M")
)

combined["order_day"] = (
    combined["order_datetime"].dt.day_name()
)


# 22. CALCULATE LATE DELIVERY

# A delivery is late only when it takes more than 45 minutes.

combined["is_late"] = pd.NA

valid_delivery = combined["delivery_minutes"].notna()

combined.loc[
    valid_delivery, "is_late"
] = combined.loc[
    valid_delivery, "delivery_minutes"
].apply(
    lambda x: "Yes" if x > 45 else "No"
)


# 23. SELECT REQUIRED 19 COLUMNS

combined = combined[
    [
        "order_id",
        "order_date",
        "order_time",
        "order_day",
        "customer_name",
        "restaurant_name",
        "cuisine",
        "area_name",
        "city",
        "partner_name",
        "vehicle_type",
        "order_amount",
        "payment_mode",
        "coupon_code",
        "order_status",
        "delivery_minutes",
        "is_late",
        "rating",
        "comment"
    ]
]


# 24. CHECK FINAL DATA

print("\nFinal Combined Dataset:")

print("Rows:", len(combined))

print("Columns:", len(combined.columns))

print("Duplicate Order IDs:", combined["order_id"].duplicated().sum())


# Make sure one row exists for each order

assert combined["order_id"].is_unique, (
    "Duplicate orders found in the combined dataset!"
)

assert len(combined.columns) == 19, (
    "Expected 19 columns in the combined dataset!"
)


# 25. SAVE COMBINED DATASET

combined_file = (
    OUTPUT_FOLDER / "Quick_Bite_Combined_Data.xlsx"
)

combined.to_excel(
    combined_file,
    sheet_name="Combined Data",
    index=False
)


print("\nTask 2 completed!")

print("Combined Excel file saved:", combined_file)

print("\nAll cleaning and integration operations completed!")
