"""
Step 4: Feature Importance Analysis
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier

# ──────────────────────────────────────────────
# 4A. Retrain RF and extract importances
# ──────────────────────────────────────────────
df = pd.read_csv('water_potability_engineered.csv')

feature_cols = ['ph', 'Hardness', 'Solids', 'Chloramines', 'Sulfate',
                'Conductivity', 'Organic_carbon', 'Trihalomethanes',
                'Turbidity', 'CRI', 'ph_deviation']

X = df[feature_cols]
y = df['Potability']

X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42, stratify=y
)

rf = RandomForestClassifier(
    n_estimators=200, max_depth=10,
    min_samples_split=5, random_state=42, class_weight='balanced'
)
rf.fit(X_train, y_train)

# ──────────────────────────────────────────────
# 4B. Feature importance (Gini impurity-based)
# ──────────────────────────────────────────────
importances = pd.Series(rf.feature_importances_, index=feature_cols)
importances = importances.sort_values(ascending=True)

print("── Feature Importances (Random Forest) ──")
for feat, imp in importances.items():
    bar = '█' * int(imp * 100)
    print(f"{feat:20s} {imp:.4f} {bar}")
# Expected Output (approximate order, importance ~0.07-0.12 each):
# ph_deviation          ~0.03-0.05  (lowest — most values are 0)
# CRI                   ~0.07-0.09
# Chloramines           ~0.08-0.10
# Turbidity             ~0.08-0.10
# Organic_carbon        ~0.08-0.10
# Trihalomethanes       ~0.08-0.10
# ph                    ~0.09-0.11
# Conductivity          ~0.09-0.11
# Sulfate               ~0.10-0.12
# Hardness              ~0.10-0.12
# Solids                ~0.12-0.15  (often highest)
#
# HOW TO INTERPRET:
# - All features have SIMILAR importance (~0.08-0.12)
# - No single feature dominates → confirms weak individual signals
# - This is EXPECTED for water quality data
# - If one feature had >0.40 importance, suspect data leakage

# ──────────────────────────────────────────────
# 4C. Plot and save
# ──────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 6))
importances.plot(kind='barh', ax=ax, color='steelblue', edgecolor='black')
ax.set_title('Feature Importance — Random Forest (Gini)', fontsize=14)
ax.set_xlabel('Importance Score')
ax.axvline(x=1/len(feature_cols), color='red', linestyle='--',
           label=f'Uniform baseline ({1/len(feature_cols):.3f})')
ax.legend()
plt.tight_layout()
plt.savefig('feature_importance.png', dpi=150)
print("\n✓ Plot saved to feature_importance.png")

# ──────────────────────────────────────────────
# 4D. Correlation heatmap
# ──────────────────────────────────────────────
fig2, ax2 = plt.subplots(figsize=(12, 8))
corr_matrix = df[feature_cols + ['Potability']].corr()
sns.heatmap(corr_matrix, annot=True, fmt='.2f', cmap='RdBu_r',
            center=0, ax=ax2, square=True)
ax2.set_title('Feature Correlation Heatmap', fontsize=14)
plt.tight_layout()
plt.savefig('correlation_heatmap.png', dpi=150)
print("✓ Heatmap saved to correlation_heatmap.png")
