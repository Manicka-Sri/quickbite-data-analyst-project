# QuickBite Data Analyst Assessment

## Project Overview

This project was completed as part of the Data Analyst Assessment for FortisureIT. It focuses on cleaning, combining, and analysing food delivery data from a fictional food delivery application called QuickBite using Python.

The objective is to transform raw datasets into meaningful business insights through data cleaning, data integration, and exploratory data analysis.

## Technologies Used

- Python 3
- Pandas
- NumPy
- OpenPyXL
- Microsoft Excel
- Visual Studio Code
- Git and GitHub

## Project Structure

```text
quickbite-assessment/
│
├── data/
│   └── QuickBite_Assessment.xlsx
│
├── scripts/
│   ├── cleaning.py
│   └── analysis.py
│
├── cleaned_data/
│   ├── Quick_Bite_Data_Cleaning.xlsx
│   └── Quick_Bite_Combined_Data.xlsx
│
├── analysis/
│   └── Quick_Bite_Analysis.xlsx
│
└── README.md
```

## Dataset Description

The original Excel dataset contains six tables:

1. Orders
2. Customers
3. Restaurants
4. Areas
5. Delivery Partners
6. Feedback

These datasets contain information about customer orders, restaurant details, delivery partners, locations, and customer feedback.

## Task 1: Data Cleaning

The raw datasets were cleaned using Python and Pandas.

The following operations were performed:

- Removed duplicate records.
- Handled missing values in mandatory fields.
- Standardised data for consistency.
- Removed delivered orders with missing delivery-time information.
- Validated relationships between tables using their respective IDs.
- Preserved records with missing related information where appropriate.

**Output:** `cleaned_data/Quick_Bite_Data_Cleaning.xlsx`

This workbook contains all six cleaned datasets in separate Excel sheets.

## Task 2: Data Merging

The six cleaned datasets were combined using appropriate joins based on their common identifiers.

Left joins were used to retain order records even when some related customer, restaurant, delivery partner, or feedback information was unavailable.

The final combined dataset contains:

- 1,234 records
- 19 columns
- One record per order

**Output:** `cleaned_data/Quick_Bite_Combined_Data.xlsx`

## Task 3: Data Analysis

The analysis was performed using only the combined dataset.

Only delivered orders were considered for the analysis.

A delivery was classified as late when its delivery time exceeded 45 minutes.

### 1. Area Performance Analysis

Analysed delivery and business performance across different areas.

Metrics calculated:

- Total orders
- Total revenue
- Average order value
- Average delivery time
- Number of late deliveries
- Average customer rating
- Late delivery percentage
- Revenue ranking

### 2. Top Cuisine by Day

Identified the most popular cuisine for each day of the week based on the number of delivered orders.

Ties were preserved when multiple cuisines had the same highest order count.

### 3. Vehicle Performance Analysis

Compared delivery performance across vehicle types.

Metrics calculated:

- Total orders
- Average delivery time
- Number of late deliveries
- Average customer rating
- Late delivery percentage

**Output:** `analysis/Quick_Bite_Analysis.xlsx`

The workbook contains three separate sheets:

- Area Performance
- Top Cuisine by Day
- Vehicle Performance

## Key Analysis Results

| Metric | Result |
|---|---:|
| Total combined records | 1,234 |
| Delivered orders analysed | 1,138 |
| Late deliveries | 234 |
| Overall late delivery percentage | 20.56% |
| Areas included in analysis | 18 |
| Top cuisine result rows | 8 |
| Vehicle types analysed | 4 |

## How to Run the Project

### Step 1: Install Python libraries

```bash
pip install pandas numpy openpyxl
```

### Step 2: Run the data cleaning script

```bash
python scripts/cleaning.py
```

This generates the cleaned datasets and the combined dataset.

### Step 3: Run the analysis script

```bash
python scripts/analysis.py
```

This generates the final analysis workbook.

**Note:** Run the cleaning script before the analysis script because the analysis script reads the combined Excel workbook generated during the cleaning and merging process.

## Conclusion

This project demonstrates the use of Python for practical data analyst tasks, including data cleaning, data preprocessing, data merging, and business analysis.

The final analysis provides insights into area-wise performance, popular cuisines, and delivery vehicle efficiency, helping identify opportunities to improve delivery operations and customer experience.

---

**Project:** QuickBite Data Analyst Assessment  
**Repository:** [GitHub - QuickBite Data Analyst Project](https://github.com/Manicka-Sri/quickbite-data-analyst-project)