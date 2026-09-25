# Waterborne Disease Outbreak Risk Prediction System (PRJ_441)

ML + GIS system to predict disease outbreak risk (Low/Medium/High) per region from water quality and environmental data.

## Dataset
- **Source**: [Kaggle Water Potability](https://www.kaggle.com/datasets/adityakadiwal/water-potability)
- **File**: `water_potability.csv` (3,276 samples, 10 features)
- **Download**: Visit the link above → click "Download" → place `water_potability.csv` in the project root

## Setup
```bash
pip install -r requirements.txt
```

## Run Pipeline
```bash
python 01_data_loading_and_cleaning.py
python 02_feature_engineering.py
python 03_model_training.py
python 04_feature_importance.py
```

## Results
- Random Forest Accuracy: ~65% (realistic on real data)
- Logistic Regression Accuracy: ~60%
- Key insight: No single water quality feature strongly predicts potability

## Tech Stack
Python, pandas, scikit-learn, matplotlib, seaborn

## License
MIT
