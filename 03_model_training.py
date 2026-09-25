"""
Step 3: Model Training — Random Forest & Logistic Regression
REALISTIC ACCURACY EXPECTATION: 62-68% on this dataset.
Why? Because features have very weak correlation with potability in real data.
Your synthetic 98% was because the fake data had artificial patterns.
"""
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import (
    accuracy_score, classification_report, confusion_matrix
)
import warnings
warnings.filterwarnings('ignore')

# ──────────────────────────────────────────────
# 3A. Load & prepare data
# ──────────────────────────────────────────────
df = pd.read_csv('water_potability_engineered.csv')

# Features: all numeric columns except target & derived risk_level
feature_cols = ['ph', 'Hardness', 'Solids', 'Chloramines', 'Sulfate',
                'Conductivity', 'Organic_carbon', 'Trihalomethanes',
                'Turbidity', 'CRI', 'ph_deviation']

X = df[feature_cols]
y = df['Potability']  # Binary: 0 = Not Potable, 1 = Potable

print(f"Features: {X.shape[1]}, Samples: {X.shape[0]}")
print(f"Class balance: {y.value_counts().to_dict()}")
# Expected Output:
# Features: 11, Samples: 3276
# Class balance: {0: 1998, 1: 1278}

# ──────────────────────────────────────────────
# 3B. Train/Test split (80/20, stratified)
# ──────────────────────────────────────────────
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42, stratify=y
)
print(f"\nTrain: {X_train.shape[0]}, Test: {X_test.shape[0]}")
# Expected Output: Train: 2620, Test: 656

# Scale features (needed for Logistic Regression, harmless for RF)
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

# ──────────────────────────────────────────────
# 3C. Model 1: Random Forest
# ──────────────────────────────────────────────
print("\n" + "="*50)
print("MODEL 1: RANDOM FOREST")
print("="*50)

rf = RandomForestClassifier(
    n_estimators=200,
    max_depth=10,
    min_samples_split=5,
    random_state=42,
    class_weight='balanced'  # Handle class imbalance
)
rf.fit(X_train, y_train)  # RF doesn't need scaling

y_pred_rf = rf.predict(X_test)
acc_rf = accuracy_score(y_test, y_pred_rf)

print(f"\nAccuracy: {acc_rf:.4f}")
print(f"\nClassification Report:")
print(classification_report(y_test, y_pred_rf, target_names=['Not Potable', 'Potable']))
print("Confusion Matrix:")
print(confusion_matrix(y_test, y_pred_rf))
# Expected Output:
# Accuracy: 0.64 - 0.68  (THIS IS REALISTIC for this dataset)
# Precision/Recall for class 1 (Potable) will be lower (~0.50-0.58)
# because the class is underrepresented and features are weakly correlated.

# ──────────────────────────────────────────────
# 3D. Model 2: Logistic Regression
# ──────────────────────────────────────────────
print("\n" + "="*50)
print("MODEL 2: LOGISTIC REGRESSION")
print("="*50)

lr = LogisticRegression(
    max_iter=1000,
    random_state=42,
    class_weight='balanced'
)
lr.fit(X_train_scaled, y_train)

y_pred_lr = lr.predict(X_test_scaled)
acc_lr = accuracy_score(y_test, y_pred_lr)

print(f"\nAccuracy: {acc_lr:.4f}")
print(f"\nClassification Report:")
print(classification_report(y_test, y_pred_lr, target_names=['Not Potable', 'Potable']))
print("Confusion Matrix:")
print(confusion_matrix(y_test, y_pred_lr))
# Expected Output:
# Accuracy: 0.58 - 0.64
# LR will likely perform WORSE than RF here because relationships aren't linear.

# ──────────────────────────────────────────────
# 3E. Comparison summary
# ──────────────────────────────────────────────
print("\n" + "="*50)
print("MODEL COMPARISON")
print("="*50)
print(f"Random Forest Accuracy:      {acc_rf:.4f}")
print(f"Logistic Regression Accuracy: {acc_lr:.4f}")
print(f"Better model: {'Random Forest' if acc_rf > acc_lr else 'Logistic Regression'}")
print("\n⚠ If accuracy > 75%, double-check for data leakage.")
print("⚠ If accuracy > 90%, something is WRONG (likely target leakage or overfitting).")
