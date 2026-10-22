"""
Master Pipeline Runner for Waterborne Disease Outbreak Risk Prediction (PRJ_441)
Runs Steps 1 through 4 sequentially.
"""
import sys
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

steps = [
    ("Step 1: Data Loading & Cleaning", "01_data_loading_and_cleaning.py"),
    ("Step 2: Feature Engineering", "02_feature_engineering.py"),
    ("Step 3: Baseline Model Training (RF vs LR)", "03_model_training.py"),
    ("Step 4: Feature Importance & Diagnostic Plots", "04_feature_importance.py"),
    ("Step 5: GIS Disease Outbreak Surveillance Engine", "gis_outbreak_mapper.py"),
    ("Step 6: Production 5-Fold CV & Model Serialization", "train_production_model.py")
]

print("=" * 60)
print("  WATERBORNE DISEASE OUTBREAK RISK PREDICTION (PRJ_441)")
print("  ML + GIS End-to-End Pipeline Execution (100% Milestone)")
print("=" * 60)

for title, script in steps:
    print("\n" + "#" * 60)
    print(f"  EXECUTING: {title} ({script})")
    print("#" * 60 + "\n")
    
    result = subprocess.run([sys.executable, script], check=False)
    if result.returncode != 0:
        print(f"\n[ERROR] {script} failed with exit code {result.returncode}!")
        sys.exit(result.returncode)

import webbrowser
import os

print("\n" + "=" * 60)
print("  [SUCCESS] All pipeline stages executed successfully!")
print("  Generated outputs:")
print("   - water_potability_cleaned.csv")
print("   - water_potability_engineered.csv")
print("   - feature_importance.png")
print("   - correlation_heatmap.png")
print("   - outbreak_risk_map.html (Interactive GIS Surveillance Map)")
print("   - models/water_risk_rf_model.pkl (Serialized Production Model)")
print("   - models/model_metadata.json (5-Fold CV Benchmarks)")
print("   - IEEE_Project_Report_PRJ_441.md (Complete IEEE Report)")
print("=" * 60)
print("\n[!] TO LAUNCH THE INTERACTIVE REAL-TIME WEB DASHBOARD:")
print("    streamlit run app.py\n")

# Automatically open the GIS map in the browser
map_path = os.path.abspath("outbreak_risk_map.html")
if os.path.exists(map_path):
    print(f"\n[GIS] Launching interactive surveillance map in browser: {map_path}")
    webbrowser.open(f"file://{map_path}")
