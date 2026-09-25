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
    ("Step 3: Model Training (RF vs LR)", "03_model_training.py"),
    ("Step 4: Feature Importance & Plots", "04_feature_importance.py")
]

print("=" * 60)
print("  WATERBORNE DISEASE OUTBREAK RISK PREDICTION (PRJ_441)")
print("  ML Pipeline Execution - 0 to 60% Milestone")
print("=" * 60)

for title, script in steps:
    print("\n" + "#" * 60)
    print(f"  EXECUTING: {title} ({script})")
    print("#" * 60 + "\n")
    
    result = subprocess.run([sys.executable, script], check=False)
    if result.returncode != 0:
        print(f"\n[ERROR] {script} failed with exit code {result.returncode}!")
        sys.exit(result.returncode)

print("\n" + "=" * 60)
print("  [SUCCESS] All 4 pipeline stages executed successfully!")
print("  Generated outputs:")
print("   - water_potability_cleaned.csv")
print("   - water_potability_engineered.csv")
print("   - feature_importance.png")
print("   - correlation_heatmap.png")
print("=" * 60)
