
import pandas as pd
import numpy as np
from pathlib import Path

# --------------------------------------------------
# 1. FILE PATHS
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = BASE_DIR / "data" / "QuickBite_Assessment.xlsx"

CLEANED_DIR = BASE_DIR / "cleaned_data"
CLEANED_DIR.mkdir(parents=True, exist_ok=True)

CLEANED_FILE = CLEANED_DIR / "Quick_Bite_Data_Cleaning.xlsx"
COMBINED_FILE = CLEANED_DIR / "Quick_Bite_Combined_Data.xlsx"


# --------------------------------------------------
# 2. READ ALL EXCEL SHEETS
# --------------------------------------------------

print("Reading original Excel workbook...")

sheets = pd.read_excel(INPUT_FILE, sheet_name=None)

for name, df in sheets.items():
    # Remove accidental spaces from column names
    df.columns = df.columns.str.strip()

    # Convert empty strings into missing values
    df = df.replace(r"^\s*$", np.nan, regex=True)

    # Remove leading and trailing spaces from text
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].apply(
            lambda x: x.strip() if isinstance(x, str) else x
        )

    sheets[name] = df

print("All six sheets loaded successfully.")


# --------------------------------------------------
# 3. STANDARDIZE DATA
# --------------------------------------------------

# Standardize city names
city_map = {
    "BLR": "Bengaluru",
    "BENGALURU": "Bengaluru",
    "HYD": "Hyderabad",
    "HYDERABAD": "Hyderabad",
    "MADRAS": "Chennai",
    "CHENNAI": "Chennai",
    "MUMBAI": "Mumbai"
}

areas = sheets["areas"]
areas["city"] = (
    areas["city"].astype("string").str.strip().str.upper()
    .map(city_map)
)

# Standardize cuisines
cuisine_map = {
    "BIRIYANI": "Biryani",
    "BIRYANI": "Biryani",
    "N. INDIAN": "North Indian",
    "NORTH-INDIAN": "North Indian",
    "NORTH INDIAN": "North Indian",
    "SOUTH-INDIAN": "South Indian",
    "SOUTH INDIAN": "South Indian",
    "DESSERT": "Desserts"
}

restaurants = sheets["restaurants"]

restaurants["cuisine"] = (
    restaurants["cuisine"].astype("string").str.strip().str.upper()
    .map(cuisine_map)
    .fillna(restaurants["cuisine"])
)

# Standardize vehicle types
vehicle_map = {
    "BIKE": "Motorbike",
    "MOTOR BIKE": "Motorbike",
    "MOTORBIKE": "Motorbike",
    "SCOOTY": "Scooter",
    "ACTIVA": "Scooter",
    "SCOOTER": "Scooter",
    "BICYCLE": "Bicycle",
    "CYCLE": "Bicycle",
    "EV": "EV Scooter",
    "E-SCOOTER": "EV Scooter",
    "EV SCOOTER": "EV Scooter"
}

partners = sheets["delivery_partners"]

partners["vehicle_type"] = (
    partners["vehicle_type"].astype("string").str.strip().str.upper()
    .map(vehicle_map)
)

# Standardize payment modes
payment_map = {
    "CASH": "Cash",
    "COD": "Cash on Delivery",
    "CASH ON DELIVERY": "Cash on Delivery",
    "UPI": "UPI",
    "GPAY": "GPay",
    "PHONEPE": "PhonePe",
    "CARD": "Card",
    "CREDIT CARD": "Credit Card",
    "DEBIT CARD": "Debit Card",
    "WALLET": "Wallet",
    "AMAZON PAY": "Amazon Pay",
    "PAYTM WALLET": "Paytm Wallet"
}

orders = sheets["orders"]

orders["payment_mode"] = (
    orders["payment_mode"].astype("string").str.strip().str.upper()
    .map(payment_map)
)

# Standardize order statuses
status_map = {
    "DELIVERED": "Delivered",
    "COMPLETED": "Delivered",
    "CANCELLED": "Cancelled",
    "CANCELED": "Cancelled"
}

orders["order_status"] = (
    orders["order_status"].astype("string").str.strip().str.upper()
    .map(status_map)
)

# Standardize coupon codes
orders["coupon_code"] = orders["coupon_code"].str.upper()


# --------------------------------------------------
# 4. STANDARDIZE DATE AND NUMERIC COLUMNS
# --------------------------------------------------

date_columns = {
    "restaurants": ["onboard_date"],
    "delivery_partners": ["join_date"],
    "customers": ["signup_date"],
    "orders": ["order_datetime"],
    "feedback": ["feedback_date"]
}

for sheet_name, columns in date_columns.items():
    for col in columns:
        sheets[sheet_name][col] = pd.to_datetime(
            sheets[sheet_name][col],
            format="mixed",
            dayfirst=True,
            errors="coerce"
        )

# Convert numeric columns
restaurants["rating"] = pd.to_numeric(
    restaurants["rating"], errors="coerce"
)

orders["order_amount"] = pd.to_numeric(
    orders["order_amount"], errors="coerce"
)

orders["delivery_minutes"] = pd.to_numeric(
    orders["delivery_minutes"], errors="coerce"
)

feedback = sheets["feedback"]

feedback["rating"] = pd.to_numeric(
    feedback["rating"], errors="coerce"
)


# --------------------------------------------------
# 5. REMOVE DUPLICATES
# --------------------------------------------------

key_columns = {
    "areas": "area_id",
    "restaurants": "restaurant_id",
    "delivery_partners": "partner_id",
    "customers": "customer_id",
    "orders": "order_id",
    "feedback": "feedback_id"
}

for name, key in key_columns.items():

    before = len(sheets[name])

    # Keep the first occurrence of each unique ID
    sheets[name] = sheets[name].drop_duplicates(
        subset=key,
        keep="first"
    )

    removed = before - len(sheets[name])

    print(f"{name}: Removed {removed} duplicate records.")


# --------------------------------------------------
# 6. REMOVE ROWS WITH MISSING MANDATORY DATA
# --------------------------------------------------

mandatory_columns = {
    "areas": ["area_id", "area_name", "city", "zone"],

    "restaurants": [
        "restaurant_id", "restaurant_name",
        "area_id", "cuisine"
    ],

    "delivery_partners": [
        "partner_id", "partner_name", "vehicle_type"
    ],

    "customers": [
        "customer_id", "customer_name"
    ],

    "orders": [
        "order_id", "customer_id", "restaurant_id",
        "order_datetime", "order_amount",
        "payment_mode", "order_status"
    ],

    "feedback": [
        "feedback_id", "order_id", "rating"
    ]
}

for name, columns in mandatory_columns.items():

    before = len(sheets[name])

    sheets[name] = sheets[name].dropna(
        subset=columns
    )

    removed = before - len(sheets[name])

    print(f"{name}: Removed {removed} rows with missing mandatory data.")


# Delivered orders must have a delivery time
orders = sheets["orders"]

before = len(orders)

orders = orders[
    (orders["order_status"] != "Delivered") |
    (orders["delivery_minutes"].notna())
].copy()

print(
    "Orders: Removed",
    before - len(orders),
    "delivered orders with missing delivery time."
)

sheets["orders"] = orders


# --------------------------------------------------
# 7. SAVE ALL CLEANED TABLES
# --------------------------------------------------

with pd.ExcelWriter(CLEANED_FILE, engine="openpyxl") as writer:

    for name, df in sheets.items():
        df.to_excel(
            writer,
            sheet_name=name,
            index=False
        )

print("\nCleaned tables saved successfully!")


# --------------------------------------------------
# 8. COMBINE ALL SIX TABLES
# --------------------------------------------------

orders = sheets["orders"]
customers = sheets["customers"]
restaurants = sheets["restaurants"]
areas = sheets["areas"]
partners = sheets["delivery_partners"]
feedback = sheets["feedback"]

# Join orders with customers
combined = orders.merge(
    customers[["customer_id", "customer_name"]],
    on="customer_id",
    how="left",
    validate="many_to_one"
)

# Join restaurants
combined = combined.merge(
    restaurants[
        ["restaurant_id", "restaurant_name", "cuisine", "area_id"]
    ],
    on="restaurant_id",
    how="left",
    validate="many_to_one"
)

# Join areas
combined = combined.merge(
    areas[["area_id", "area_name", "city"]],
    on="area_id",
    how="left",
    validate="many_to_one"
)

# Join delivery partners
combined = combined.merge(
    partners[["partner_id", "partner_name", "vehicle_type"]],
    on="partner_id",
    how="left",
    validate="many_to_one"
)

# Join feedback
combined = combined.merge(
    feedback[["order_id", "rating", "comment"]],
    on="order_id",
    how="left",
    validate="one_to_one"
)
print("\nFOREIGN KEY VALIDATION")

print(
    "Orders without matching customer:",
    combined["customer_name"].isna().sum()
)

print(
    "Orders without matching restaurant:",
    combined["restaurant_name"].isna().sum()
)

print(
    "Orders without matching delivery partner:",
    combined["partner_name"].isna().sum()
)

print(
    "Orders without feedback:",
    combined["rating"].isna().sum()
)


# --------------------------------------------------
# 9. CREATE REQUIRED COLUMNS
# --------------------------------------------------

combined["order_date"] = combined["order_datetime"].dt.date

combined["order_time"] = combined["order_datetime"].dt.strftime(
    "%H:%M:%S"
)

combined["order_day"] = combined["order_datetime"].dt.day_name()

combined["is_late"] = (
    (combined["order_status"] == "Delivered") &
    (combined["delivery_minutes"] > 45)
)

# Select the exact required column order
final_columns = [
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

combined = combined[final_columns]

# Save combined workbook
combined.to_excel(
    COMBINED_FILE,
    index=False,
    engine="openpyxl"
)

print("\nCombined dataset saved successfully!")

print("\nFinal combined dataset shape:", combined.shape)
print("\nFINAL DATA VALIDATION")

print("Total rows:", len(combined))
print("Total columns:", len(combined.columns))

print("Unique order IDs:", combined["order_id"].nunique())

print("Duplicate order IDs:",
      combined["order_id"].duplicated().sum())

print("Missing values:")
print(combined.isnull().sum())

print("\nCleaning and combination completed!")
