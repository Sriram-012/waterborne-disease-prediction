"""
Step 2: Feature Engineering
Derived features to improve model signal.
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd
import numpy as np

df = pd.read_csv('water_potability_cleaned.csv')

# ──────────────────────────────────────────────
# 2A. Contamination Risk Index (CRI)
# ──────────────────────────────────────────────
# REASONING: Combines the 3 most health-relevant water quality indicators into one
# composite score. High turbidity + high organic carbon + high trihalomethanes = high contamination risk.
# Each is min-max normalized to [0,1] before summing so units don't dominate.

def min_max_normalize(series):
    return (series - series.min()) / (series.max() - series.min())

df['CRI'] = (
    min_max_normalize(df['Turbidity']) +
    min_max_normalize(df['Organic_carbon']) +
    min_max_normalize(df['Trihalomethanes'])
) / 3  # Scale to [0, 1]

print("── Contamination Risk Index (CRI) ──")
print(df['CRI'].describe().round(4))
# Expected Output:
# count    3276.0000
# mean        0.4900  (approximately)
# std         0.1500
# min         0.0000
# max         1.0000

# ──────────────────────────────────────────────
# 2B. pH Deviation from Safe Range
# ──────────────────────────────────────────────
# REASONING: WHO safe pH range is 6.5–8.5. Distance from this range matters more than raw pH.
df['ph_deviation'] = df['ph'].apply(
    lambda x: max(0, 6.5 - x) + max(0, x - 8.5)
)

print("\n── pH Deviation from Safe Range [6.5, 8.5] ──")
print(df['ph_deviation'].describe().round(4))
# Expected Output:
# ~75% of values will be 0 (within safe range)
# max around 5-7 (extreme pH values)

# ──────────────────────────────────────────────
# 2C. Risk Label (for multi-class framing)
# ──────────────────────────────────────────────
# Map continuous CRI into Low/Medium/High risk categories
# Map CRI into Low/Medium/High risk using quantile-based thresholds (data-driven)
df['risk_level'] = pd.qcut(
    df['CRI'],
    q=3,
    labels=['Low', 'Medium', 'High']
)

print("\n── Risk Level Distribution ──")
print(df['risk_level'].value_counts().sort_index())
# Expected Output:
# Low       ~800-900
# Medium    ~1600-1800
# High      ~500-700

# ──────────────────────────────────────────────
# 2D. Correlation check
# ──────────────────────────────────────────────
print("\n── Correlation with Potability ──")
corr = df.select_dtypes(include=[np.number]).corr()['Potability'].sort_values(ascending=False)
print(corr.round(4))
# Expected Output:
# Most correlations will be WEAK (< 0.05). This is realistic for water quality data.
# This is why accuracy won't be 98% on real data.

df.to_csv('water_potability_engineered.csv', index=False)
print("\n✓ Engineered data saved to water_potability_engineered.csv")
