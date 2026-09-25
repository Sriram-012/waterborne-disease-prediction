"""
Step 1: Data Loading & Cleaning
Dataset: Kaggle water_potability.csv
Source: https://www.kaggle.com/datasets/adityakadiwal/water-potability
Download: Go to the link above → Download → place water_potability.csv in this folder
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd
import numpy as np

# ──────────────────────────────────────────────
# 1A. Load the dataset
# ──────────────────────────────────────────────
df = pd.read_csv('water_potability.csv')

print("Shape:", df.shape)
print("\nColumn types:")
print(df.dtypes)
print("\nFirst 5 rows:")
print(df.head())
# Expected Output:
# Shape: (3276, 10)
# Columns: ph(float), Hardness(float), Solids(float), Chloramines(float),
#          Sulfate(float), Conductivity(float), Organic_carbon(float),
#          Trihalomethanes(float), Turbidity(float), Potability(int)

# ──────────────────────────────────────────────
# 1B. Check missing values
# ──────────────────────────────────────────────
print("\n── Missing Values ──")
print(df.isnull().sum())
print(f"\nTotal missing: {df.isnull().sum().sum()}")
# Expected Output:
# ph                 491
# Hardness             0
# Solids               0
# Chloramines          0
# Sulfate            781
# Conductivity         0
# Organic_carbon       0
# Trihalomethanes    162
# Turbidity            0
# Potability           0
# Total missing: 1434

# ──────────────────────────────────────────────
# 1C. Check duplicates
# ──────────────────────────────────────────────
dupes = df.duplicated().sum()
print(f"\nDuplicate rows: {dupes}")
# Expected Output: Duplicate rows: 0

# ──────────────────────────────────────────────
# 1D. Median imputation for missing values
# ──────────────────────────────────────────────
# WHY median? Because pH, Sulfate, Trihalomethanes can have outliers — median is robust to them.
for col in ['ph', 'Sulfate', 'Trihalomethanes']:
    median_val = df[col].median()
    df[col] = df[col].fillna(median_val)
    print(f"Filled {col} NaN with median = {median_val:.2f}")
# Expected Output:
# Filled ph NaN with median = 7.04
# Filled Sulfate NaN with median = 333.07
# Filled Trihalomethanes NaN with median = 66.62

print(f"\nMissing after cleaning: {df.isnull().sum().sum()}")
# Expected Output: Missing after cleaning: 0

print("\n── Class Distribution ──")
print(df['Potability'].value_counts())
print(df['Potability'].value_counts(normalize=True))
# Expected Output:
# 0    1998  (~61%)
# 1    1278  (~39%)
# Note: Imbalanced — more non-potable. This affects accuracy expectations.

# ──────────────────────────────────────────────
# 1E. Basic statistics
# ──────────────────────────────────────────────
print("\n── Descriptive Statistics ──")
print(df.describe().round(2))

# Save cleaned data
df.to_csv('water_potability_cleaned.csv', index=False)
print("\n✓ Cleaned data saved to water_potability_cleaned.csv")
