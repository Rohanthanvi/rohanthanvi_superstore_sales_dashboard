# =============================================================================
# Super Store Sales Data Cleaning Script
# =============================================================================
# Student  : Rohan Thanvi
# Roll No  : 23053625
# Program  : Power BI Capstone Project
# Dataset  : Super Store Sales Dataset (Kaggle)
# =============================================================================

import pandas as pd
import os

print("=" * 60)
print("   Super Store Sales — Data Cleaning Script")
print("   Rohan Thanvi | Roll No: 23053625")
print("=" * 60)
print()

# File path — update if needed
DATASET_PATH = "SuperStore_Sales_Dataset.xls"

# =============================================================================
# STEP 1 — LOAD DATASET
# =============================================================================
print(">>> STEP 1: Loading dataset...")

try:
    # Try reading as CSV (file is actually CSV despite .xls extension)
    df = pd.read_csv(DATASET_PATH, encoding="utf-8-sig")
    print(f"    ✔ Loaded as CSV successfully.")
except Exception:
    try:
        df = pd.read_excel(DATASET_PATH, engine="xlrd")
        print(f"    ✔ Loaded as Excel successfully.")
    except Exception as e:
        print(f"    ✘ Failed to load: {e}")
        exit()

print(f"    Shape       : {df.shape[0]} rows × {df.shape[1]} columns")
print(f"    Columns     : {list(df.columns)}")
print()

# =============================================================================
# STEP 2 — INITIAL INSPECTION
# =============================================================================
print(">>> STEP 2: Initial inspection...")
print(f"    Total rows          : {len(df)}")
print(f"    Total columns       : {len(df.columns)}")
print(f"    Total null values   : {df.isnull().sum().sum()}")
print(f"    Duplicate rows      : {df.duplicated().sum()}")
print()
print("    Null values per column:")
nulls = df.isnull().sum()
for col, n in nulls[nulls > 0].items():
    print(f"      {col:<25} : {n} nulls")
print()

# =============================================================================
# STEP 3 — REMOVE DUPLICATES
# =============================================================================
print(">>> STEP 3: Removing duplicate rows...")
before = len(df)
df.drop_duplicates(inplace=True)
print(f"    Duplicates removed  : {before - len(df)}")
print(f"    Rows remaining      : {len(df)}")
print()

# =============================================================================
# STEP 4 — STRIP & CLEAN COLUMN NAMES
# =============================================================================
print(">>> STEP 4: Cleaning column names...")
df.columns = df.columns.str.strip()
# Remove any unnamed/index columns
df = df.loc[:, ~df.columns.str.contains('^Unnamed|^ind')]
# Remove Row ID column if present (not needed for analysis)
if 'Row ID' in df.columns:
    df.drop(columns=['Row ID'], inplace=True)
    print("    Dropped: Row ID (not needed for analysis)")
print(f"    Final columns: {list(df.columns)}")
print()

# =============================================================================
# STEP 5 — FIX DATA TYPES
# =============================================================================
print(">>> STEP 5: Fixing data types...")

# Fix date columns
for date_col in ['Order Date', 'Ship Date']:
    if date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
        print(f"    {date_col} → datetime64")

# Fix numeric columns
for num_col in ['Sales', 'Quantity', 'Profit']:
    if num_col in df.columns:
        df[num_col] = pd.to_numeric(df[num_col], errors='coerce')
        print(f"    {num_col} → float64")

print()

# =============================================================================
# STEP 6 — HANDLE MISSING VALUES
# =============================================================================
print(">>> STEP 6: Handling missing/invalid values...")

# Drop rows with missing critical fields
critical = ['Order ID', 'Order Date', 'Sales', 'Profit', 'Category']
before = len(df)
df.dropna(subset=[c for c in critical if c in df.columns], inplace=True)
print(f"    Rows dropped (missing critical fields): {before - len(df)}")

# Fix Returns column — contains #N/A strings, replace with 0
if 'Returns' in df.columns:
    df['Returns'] = df['Returns'].replace('#N/A', 0)
    df['Returns'] = pd.to_numeric(df['Returns'], errors='coerce').fillna(0)
    print("    'Returns' column: #N/A replaced with 0")

# Fill missing text fields with 'Unknown'
for text_col in ['Ship Mode', 'Segment', 'Region', 'Payment Mode']:
    if text_col in df.columns:
        missing = df[text_col].isnull().sum()
        if missing > 0:
            df[text_col].fillna('Unknown', inplace=True)
            print(f"    {text_col}: {missing} nulls filled with 'Unknown'")

print(f"    Rows remaining      : {len(df)}")
print()

# =============================================================================
# STEP 7 — REMOVE INVALID RECORDS
# =============================================================================
print(">>> STEP 7: Removing invalid records...")
before = len(df)

# Remove rows where Sales is zero or negative
if 'Sales' in df.columns:
    df = df[df['Sales'] > 0]

# Remove rows where Quantity is zero or negative
if 'Quantity' in df.columns:
    df = df[df['Quantity'] > 0]

print(f"    Invalid records removed : {before - len(df)}")
print(f"    Rows remaining          : {len(df)}")
print()

# =============================================================================
# STEP 8 — STANDARDIZE TEXT COLUMNS
# =============================================================================
print(">>> STEP 8: Standardizing text columns...")
text_cols = ['Ship Mode', 'Segment', 'Country', 'City', 'State',
             'Region', 'Category', 'Sub-Category', 'Payment Mode']
for col in text_cols:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.title()
print("    ✔ All text columns stripped and title-cased.")
print()

# =============================================================================
# STEP 9 — ADD USEFUL DERIVED COLUMNS
# =============================================================================
print(">>> STEP 9: Adding derived columns...")

if 'Order Date' in df.columns:
    df['Order Month'] = df['Order Date'].dt.month_name()
    df['Order Year']  = df['Order Date'].dt.year
    df['Order Quarter'] = df['Order Date'].dt.quarter.map(
        {1:'Q1', 2:'Q2', 3:'Q3', 4:'Q4'})
    print("    Added: Order Month, Order Year, Order Quarter")

if 'Sales' in df.columns and 'Profit' in df.columns:
    df['Profit Margin (%)'] = (
        (df['Profit'] / df['Sales']) * 100
    ).round(2)
    print("    Added: Profit Margin (%)")

if 'Order Date' in df.columns and 'Ship Date' in df.columns:
    df['Days to Ship'] = (df['Ship Date'] - df['Order Date']).dt.days
    print("    Added: Days to Ship")

print()

# =============================================================================
# STEP 10 — SUMMARY STATISTICS
# =============================================================================
print(">>> STEP 10: Summary statistics after cleaning...")
print()

total_sales    = df['Sales'].sum() if 'Sales' in df.columns else 0
total_profit   = df['Profit'].sum() if 'Profit' in df.columns else 0
total_qty      = df['Quantity'].sum() if 'Quantity' in df.columns else 0
total_orders   = df['Order ID'].nunique() if 'Order ID' in df.columns else 0
profit_margin  = (total_profit / total_sales * 100).round(2) if total_sales > 0 else 0

print(f"    Total Sales           : $ {total_sales:,.2f}")
print(f"    Total Profit          : $ {total_profit:,.2f}")
print(f"    Total Quantity Sold   : {total_qty:,.0f} units")
print(f"    Total Unique Orders   : {total_orders:,}")
print(f"    Overall Profit Margin : {profit_margin}%")
print()

if 'Category' in df.columns and 'Sales' in df.columns:
    print("    Top Categories by Sales:")
    cat = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
    for c, s in cat.items():
        print(f"      {c:<25} $ {s:,.2f}")
    print()

if 'Sub-Category' in df.columns and 'Profit' in df.columns:
    print("    Top 5 Sub-Categories by Profit:")
    sub = df.groupby('Sub-Category')['Profit'].sum().sort_values(ascending=False).head(5)
    for s, p in sub.items():
        print(f"      {s:<25} $ {p:,.2f}")
    print()

if 'Region' in df.columns and 'Sales' in df.columns:
    print("    Sales by Region:")
    reg = df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
    for r, s in reg.items():
        print(f"      {r:<25} $ {s:,.2f}")
    print()

if 'Segment' in df.columns and 'Sales' in df.columns:
    print("    Sales by Segment:")
    seg = df.groupby('Segment')['Sales'].sum().sort_values(ascending=False)
    for s, v in seg.items():
        print(f"      {s:<25} $ {v:,.2f}")
    print()

# =============================================================================
# STEP 11 — EXPORT CLEANED FILE
# =============================================================================
print(">>> STEP 11: Exporting cleaned dataset...")
output_path = "SuperStore_Sales_Cleaned.csv"
df.to_csv(output_path, index=False)
print(f"    ✔ {output_path} saved successfully!")
print(f"    Final shape: {df.shape[0]} rows × {df.shape[1]} columns")
print()
print("=" * 60)
print("   Data cleaning complete! File is ready for Power BI.")
print("=" * 60)
