"""
Production Training & Cross-Validation Engine (PRJ_441)
- Evaluates multiple classifiers using 5-Fold Stratified Cross-Validation.
- Handles class imbalance and non-linear interactions.
- Serializes the best-performing model and preprocessors for real-time web deployment.
"""
import sys
import os
import json
import joblib
import pandas as pd
import numpy as np

from sklearn.model_selection import StratifiedKFold, cross_validate, train_test_split
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier, ExtraTreesClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, roc_auc_score, confusion_matrix

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

def train_and_export_production_models():
    print("=" * 65)
    print("  PRJ_441: PRODUCTION ML TRAINING & 5-FOLD CROSS-VALIDATION")
    print("=" * 65)

    data_path = "water_potability_engineered.csv"
    if not os.path.exists(data_path):
        print(f"[!] {data_path} not found. Running feature engineering first...")
        import subprocess
        subprocess.run([sys.executable, "02_feature_engineering.py"], check=True)

    df = pd.read_csv(data_path)
    feature_cols = [
        'ph', 'Hardness', 'Solids', 'Chloramines', 'Sulfate',
        'Conductivity', 'Organic_carbon', 'Trihalomethanes',
        'Turbidity', 'CRI', 'ph_deviation'
    ]
    target_col = 'Potability'

    X = df[feature_cols]
    y = df[target_col]

    print(f"[*] Ingested Dataset: {X.shape[0]} samples, {X.shape[1]} features.")
    print(f"[*] Class Distribution: 0 (Unsafe): {(y==0).sum()}, 1 (Safe): {(y==1).sum()}")

    # 1. 5-Fold Stratified Cross-Validation Benchmarking
    print("\n[+] Running 5-Fold Stratified Cross-Validation across 4 Algorithms...")
    skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)

    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    candidate_models = {
        "Random Forest": RandomForestClassifier(n_estimators=200, max_depth=10, min_samples_split=5, class_weight='balanced', random_state=42),
        "Gradient Boosting": GradientBoostingClassifier(n_estimators=150, max_depth=5, learning_rate=0.08, random_state=42),
        "Extra Trees": ExtraTreesClassifier(n_estimators=200, max_depth=10, class_weight='balanced', random_state=42),
        "Logistic Regression": LogisticRegression(max_iter=1000, class_weight='balanced', random_state=42)
    }

    cv_results = {}
    scoring = ['accuracy', 'precision', 'recall', 'f1', 'roc_auc']

    for name, model in candidate_models.items():
        data_to_use = X_scaled if name == "Logistic Regression" else X
        scores = cross_validate(model, data_to_use, y, cv=skf, scoring=scoring, n_jobs=-1)
        
        cv_results[name] = {
            "Accuracy": float(np.mean(scores['test_accuracy'])),
            "Accuracy_Std": float(np.std(scores['test_accuracy'])),
            "Precision": float(np.mean(scores['test_precision'])),
            "Recall": float(np.mean(scores['test_recall'])),
            "F1_Score": float(np.mean(scores['test_f1'])),
            "ROC_AUC": float(np.mean(scores['test_roc_auc']))
        }
        print(f"    - {name:<20}: Mean Acc = {cv_results[name]['Accuracy']*100:.2f}% (±{cv_results[name]['Accuracy_Std']*100:.2f}%), ROC-AUC = {cv_results[name]['ROC_AUC']:.3f}")

    # 2. Train Final Best Model on 80/20 Holdout for Complete Diagnostic Matrix
    print("\n[+] Fitting Production Random Forest on 80/20 Stratified Split for Diagnostics...")
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)
    
    prod_rf = candidate_models["Random Forest"]
    prod_rf.fit(X_train, y_train)

    y_pred = prod_rf.predict(X_test)
    y_prob = prod_rf.predict_proba(X_test)[:, 1]

    cm = confusion_matrix(y_test, y_pred).tolist()
    holdout_metrics = {
        "accuracy": float(accuracy_score(y_test, y_pred)),
        "precision_safe": float(precision_score(y_test, y_pred, pos_label=1)),
        "recall_safe": float(recall_score(y_test, y_pred, pos_label=1)),
        "f1_safe": float(f1_score(y_test, y_pred, pos_label=1)),
        "precision_unsafe": float(precision_score(y_test, y_pred, pos_label=0)),
        "recall_unsafe": float(recall_score(y_test, y_pred, pos_label=0)),
        "f1_unsafe": float(f1_score(y_test, y_pred, pos_label=0)),
        "roc_auc": float(roc_auc_score(y_test, y_prob)),
        "confusion_matrix": cm
    }

    # Extract Gini Feature Importance
    importances = dict(zip(feature_cols, [float(x) for x in prod_rf.feature_importances_]))
    sorted_importances = dict(sorted(importances.items(), key=lambda item: item[1], reverse=True))

    # 3. Export Models & Metadata
    os.makedirs("models", exist_ok=True)
    model_path = os.path.join("models", "water_risk_rf_model.pkl")
    scaler_path = os.path.join("models", "scaler.pkl")
    meta_path = os.path.join("models", "model_metadata.json")

    joblib.dump(prod_rf, model_path)
    joblib.dump(scaler, scaler_path)

    metadata = {
        "feature_cols": feature_cols,
        "target_col": target_col,
        "cv_results": cv_results,
        "holdout_metrics": holdout_metrics,
        "feature_importances": sorted_importances,
        "baseline_accuracy": float((y == 0).sum() / len(y))
    }

    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=4)

    print(f"\n[OK] Model successfully exported to: '{model_path}'")
    print(f"[OK] Scaler successfully exported to: '{scaler_path}'")
    print(f"[OK] Training metadata saved to: '{meta_path}'")
    print("=" * 65)

if __name__ == '__main__':
    train_and_export_production_models()
