
import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment


# --------------------------------------------------
# 1. FILE PATHS
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = BASE_DIR / "cleaned_data" / "Quick_Bite_Combined_Data.xlsx"
OUTPUT_DIR = BASE_DIR / "analysis"
OUTPUT_FILE = OUTPUT_DIR / "Quick_Bite_Analysis.xlsx"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------
# 2. READ THE COMBINED DATASET
# --------------------------------------------------

print("Reading combined dataset...")

df = pd.read_excel(INPUT_FILE)

print("Dataset loaded successfully!")
print("Original dataset shape:", df.shape)


# Check that the required columns exist
required_columns = [
    "order_id",
    "order_date",
    "order_day",
    "area_name",
    "city",
    "cuisine",
    "vehicle_type",
    "order_amount",
    "order_status",
    "delivery_minutes",
    "rating"
]

missing_columns = [col for col in required_columns if col not in df.columns]

if missing_columns:
    raise ValueError(
        f"Required columns are missing from the combined dataset: {missing_columns}"
    )


# --------------------------------------------------
# 3. FILTER DELIVERED ORDERS
# --------------------------------------------------

df["order_status"] = df["order_status"].astype("string").str.strip()

delivered_df = df[
    df["order_status"].str.casefold() == "delivered"
].copy()

print("\nDelivered orders available for analysis:", len(delivered_df))


# Convert numeric columns
delivered_df["order_amount"] = pd.to_numeric(
    delivered_df["order_amount"], errors="coerce"
)

delivered_df["delivery_minutes"] = pd.to_numeric(
    delivered_df["delivery_minutes"], errors="coerce"
)

delivered_df["rating"] = pd.to_numeric(
    delivered_df["rating"], errors="coerce"
)


# Late delivery means delivery time is greater than 45 minutes
delivered_df["is_late"] = (
    delivered_df["delivery_minutes"] > 45
).astype(int)


# --------------------------------------------------
# 4. IDENTIFY VALID CATEGORY VALUES
# --------------------------------------------------

def valid_category(series):
    """
    Return True for actual category values.
    Exclude missing values and the placeholder
    'Unknown / Unassigned'.
    """
    cleaned = series.astype("string").str.strip()

    return (
        cleaned.notna()
        & cleaned.ne("")
        & cleaned.str.casefold().ne("unknown / unassigned")
    )


# --------------------------------------------------
# 5. AREA PERFORMANCE
# --------------------------------------------------

print("\nCalculating Area Performance...")

# Only use orders with a known area and city
area_df = delivered_df[
    valid_category(delivered_df["area_name"])
    & valid_category(delivered_df["city"])
].copy()

area_performance = (
    area_df.groupby(["area_name", "city"], as_index=False)
    .agg(
        total_orders=("order_id", "nunique"),
        total_revenue=("order_amount", "sum"),
        avg_order_value=("order_amount", "mean"),
        avg_delivery_minutes=("delivery_minutes", "mean"),
        late_deliveries=("is_late", "sum"),
        avg_rating=("rating", "mean")
    )
)

# Calculate late delivery percentage
area_performance["late_delivery_pct"] = (
    area_performance["late_deliveries"]
    / area_performance["total_orders"]
    * 100
)

# Sort areas by revenue, highest first
area_performance = area_performance.sort_values(
    by=["total_revenue", "area_name"],
    ascending=[False, True]
).reset_index(drop=True)

# Assign sequential revenue ranks
area_performance.insert(
    0,
    "revenue_rank",
    range(1, len(area_performance) + 1)
)

# Round decimal values
area_decimal_columns = [
    "total_revenue",
    "avg_order_value",
    "avg_delivery_minutes",
    "avg_rating",
    "late_delivery_pct"
]

area_performance[area_decimal_columns] = (
    area_performance[area_decimal_columns].round(2)
)

print("Area Performance completed!")
print("Known areas included:", len(area_performance))
print(
    "Delivered orders excluded because area/city was unknown:",
    len(delivered_df) - len(area_df)
)


# --------------------------------------------------
# 6. TOP CUISINE BY DAY
# --------------------------------------------------

print("\nCalculating Top Cuisine by Day...")

# Prepare weekday values
cuisine_df = delivered_df.copy()

cuisine_df["order_day"] = (
    cuisine_df["order_day"]
    .astype("string")
    .str.strip()
    .str.title()
)

# Standard weekday order
weekday_order = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
]

# Count all delivered orders for each day.
# This denominator includes orders whose cuisine is unknown.
daily_totals = (
    cuisine_df.groupby("order_day")["order_id"]
    .nunique()
    .rename("total_day_orders")
)

# Only known cuisines are considered for the top-cuisine ranking
known_cuisine_df = cuisine_df[
    valid_category(cuisine_df["cuisine"])
    & cuisine_df["order_day"].isin(weekday_order)
].copy()

# Count orders for each cuisine on each day
cuisine_counts = (
    known_cuisine_df.groupby(
        ["order_day", "cuisine"], as_index=False
    )
    .agg(cuisine_orders=("order_id", "nunique"))
)

# Add the total delivered orders for each day
cuisine_counts["total_day_orders"] = (
    cuisine_counts["order_day"].map(daily_totals)
)

# Find the maximum cuisine order count for each day
daily_maximum = (
    cuisine_counts.groupby("order_day")["cuisine_orders"]
    .transform("max")
)

# Keep every cuisine tied for the highest count
top_cuisine = cuisine_counts[
    cuisine_counts["cuisine_orders"] == daily_maximum
].copy()

# Calculate the percentage share of all delivered orders that day
top_cuisine["order_share_pct"] = (
    top_cuisine["cuisine_orders"]
    / top_cuisine["total_day_orders"]
    * 100
)

# Sort days in calendar order
top_cuisine["order_day"] = pd.Categorical(
    top_cuisine["order_day"],
    categories=weekday_order,
    ordered=True
)

top_cuisine = top_cuisine.sort_values(
    ["order_day", "cuisine"]
).reset_index(drop=True)

# Convert categorical weekday back to ordinary text
top_cuisine["order_day"] = (
    top_cuisine["order_day"].astype(str)
)

top_cuisine["order_share_pct"] = (
    top_cuisine["order_share_pct"].round(2)
)

top_cuisine = top_cuisine[
    [
        "order_day",
        "cuisine",
        "cuisine_orders",
        "total_day_orders",
        "order_share_pct"
    ]
]

print("Top Cuisine by Day completed!")
print("Top cuisine result rows:", len(top_cuisine))


# --------------------------------------------------
# 7. VEHICLE PERFORMANCE
# --------------------------------------------------

print("\nCalculating Vehicle Performance...")

# Only include orders with a known vehicle type
vehicle_df = delivered_df[
    valid_category(delivered_df["vehicle_type"])
].copy()

vehicle_performance = (
    vehicle_df.groupby("vehicle_type", as_index=False)
    .agg(
        total_deliveries=("order_id", "nunique"),
        avg_delivery_minutes=("delivery_minutes", "mean"),
        late_deliveries=("is_late", "sum"),
        avg_rating=("rating", "mean")
    )
)

# Calculate late delivery percentage
vehicle_performance["late_delivery_pct"] = (
    vehicle_performance["late_deliveries"]
    / vehicle_performance["total_deliveries"]
    * 100
)

# Best-performing vehicle types appear first:
# lowest late delivery percentage, then lowest average delivery time
vehicle_performance = vehicle_performance.sort_values(
    by=["late_delivery_pct", "avg_delivery_minutes"],
    ascending=[True, True]
).reset_index(drop=True)

# Round decimal values
vehicle_decimal_columns = [
    "avg_delivery_minutes",
    "avg_rating",
    "late_delivery_pct"
]

vehicle_performance[vehicle_decimal_columns] = (
    vehicle_performance[vehicle_decimal_columns].round(2)
)

print("Vehicle Performance completed!")
print("Known vehicle types included:", len(vehicle_performance))
print(
    "Delivered orders excluded because vehicle type was unknown:",
    len(delivered_df) - len(vehicle_df)
)


# --------------------------------------------------
# 8. SAVE THE THREE ANALYSIS SHEETS
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
# 9. FORMAT THE EXCEL WORKBOOK
# --------------------------------------------------

workbook = load_workbook(OUTPUT_FILE)

header_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

header_font = Font(
    bold=True,
    color="FFFFFF"
)

for worksheet in workbook.worksheets:

    # Format header row
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # Freeze the header row
    worksheet.freeze_panes = "A2"

    # Enable filters
    worksheet.auto_filter.ref = worksheet.dimensions

    # Adjust column widths
    for column_cells in worksheet.columns:

        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            if cell.value is not None:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        worksheet.column_dimensions[column_letter].width = (
            min(max_length + 3, 30)
        )


# Format Area Performance values
ws = workbook["Area Performance"]

area_headers = {
    cell.value: cell.column
    for cell in ws[1]
}

for row in range(2, ws.max_row + 1):

    for column_name in ["total_revenue", "avg_order_value"]:
        col = area_headers[column_name]
        ws.cell(row, col).number_format = '₹#,##0.00'

    for column_name in [
        "avg_delivery_minutes",
        "avg_rating",
        "late_delivery_pct"
    ]:
        col = area_headers[column_name]
        ws.cell(row, col).number_format = "0.00"


# Format Top Cuisine by Day values
ws = workbook["Top Cuisine by Day"]

cuisine_headers = {
    cell.value: cell.column
    for cell in ws[1]
}

for row in range(2, ws.max_row + 1):
    col = cuisine_headers["order_share_pct"]
    ws.cell(row, col).number_format = "0.00"


# Format Vehicle Performance values
ws = workbook["Vehicle Performance"]

vehicle_headers = {
    cell.value: cell.column
    for cell in ws[1]
}

for row in range(2, ws.max_row + 1):

    for column_name in [
        "avg_delivery_minutes",
        "avg_rating",
        "late_delivery_pct"
    ]:
        col = vehicle_headers[column_name]
        ws.cell(row, col).number_format = "0.00"


# Save formatting changes
workbook.save(OUTPUT_FILE)


# --------------------------------------------------
# 10. FINAL SUMMARY
# --------------------------------------------------

total_delivered = delivered_df["order_id"].nunique()

total_late = int(delivered_df["is_late"].sum())

if total_delivered > 0:
    overall_late_pct = (
        total_late / total_delivered
    ) * 100
else:
    overall_late_pct = 0


print("\n========== FINAL ANALYSIS SUMMARY ==========")

print("Delivered orders analysed:", total_delivered)
print("Late deliveries:", total_late)

print(
    "Overall late delivery percentage:",
    round(overall_late_pct, 2),
    "%"
)

print(
    "Area Performance rows:",
    len(area_performance)
)

print(
    "Top Cuisine by Day rows:",
    len(top_cuisine)
)

print(
    "Vehicle Performance rows:",
    len(vehicle_performance)
)

print("\nOutput file:", OUTPUT_FILE)

print("\nAll three analyses completed successfully!")
