"""
Object-Oriented ML Pipeline for Waterborne Disease Outbreak Risk Prediction (PRJ_441)
Demonstrates Encapsulation, Single Responsibility, and Modular Class Architecture.
"""
import sys
import os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


class DataLoader:
    """Encapsulates data ingestion and initial validation."""
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.data = None

    def load_data(self) -> pd.DataFrame:
        if not os.path.exists(self.filepath):
            raise FileNotFoundError(f"Dataset not found at {self.filepath}")
        self.data = pd.read_csv(self.filepath)
        return self.data.copy()

    def get_summary(self) -> dict:
        if self.data is None:
            raise ValueError("Data has not been loaded yet.")
        return {
            "shape": self.data.shape,
            "columns": list(self.data.columns),
            "missing_count": self.data.isnull().sum().to_dict(),
            "duplicates": int(self.data.duplicated().sum())
        }


class DataCleaner:
    """Handles missing value imputation and data integrity."""
    def __init__(self, impute_cols: list = None):
        self.impute_cols = impute_cols or ['ph', 'Sulfate', 'Trihalomethanes']
        self.medians = {}

    def fit_transform(self, df: pd.DataFrame) -> pd.DataFrame:
        cleaned_df = df.copy()
        for col in self.impute_cols:
            if col in cleaned_df.columns:
                median_val = cleaned_df[col].median()
                self.medians[col] = median_val
                cleaned_df[col] = cleaned_df[col].fillna(median_val)
        return cleaned_df

    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        cleaned_df = df.copy()
        for col, median_val in self.medians.items():
            if col in cleaned_df.columns:
                cleaned_df[col] = cleaned_df[col].fillna(median_val)
        return cleaned_df


class FeatureEngineer:
    """Calculates domain-specific health and risk metrics."""
    def __init__(self, safe_ph_min: float = 6.5, safe_ph_max: float = 8.5):
        self.safe_ph_min = safe_ph_min
        self.safe_ph_max = safe_ph_max

    @staticmethod
    def _min_max(series: pd.Series) -> pd.Series:
        rng = series.max() - series.min()
        return (series - series.min()) / rng if rng != 0 else series

    def compute_cri(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        t_norm = self._min_max(df['Turbidity'])
        c_norm = self._min_max(df['Organic_carbon'])
        m_norm = self._min_max(df['Trihalomethanes'])
        df['CRI'] = (t_norm + c_norm + m_norm) / 3.0
        return df

    def compute_ph_deviation(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df['ph_deviation'] = df['ph'].apply(
            lambda x: max(0.0, self.safe_ph_min - x) + max(0.0, x - self.safe_ph_max)
        )
        return df

    def bin_risk_levels(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df['risk_level'] = pd.qcut(
            df['CRI'],
            q=3,
            labels=['Low', 'Medium', 'High']
        )
        return df

    def engineer_all(self, df: pd.DataFrame) -> pd.DataFrame:
        df = self.compute_cri(df)
        df = self.compute_ph_deviation(df)
        df = self.bin_risk_levels(df)
        return df


class ModelTrainer:
    """Encapsulates training, scaling, evaluation, and feature attribution."""
    def __init__(self, feature_cols: list, random_state: int = 42):
        self.feature_cols = feature_cols
        self.random_state = random_state
        self.scaler = StandardScaler()
        self.rf_model = None
        self.lr_model = None

    def prepare_split(self, df: pd.DataFrame, target_col: str = 'Potability', test_size: float = 0.2):
        X = df[self.feature_cols]
        y = df[target_col]
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=test_size, random_state=self.random_state, stratify=y
        )
        X_train_scaled = self.scaler.fit_transform(X_train)
        X_test_scaled = self.scaler.transform(X_test)
        return X_train, X_test, X_train_scaled, X_test_scaled, y_train, y_test

    def train_random_forest(self, X_train, y_train):
        self.rf_model = RandomForestClassifier(
            n_estimators=200,
            max_depth=10,
            min_samples_split=5,
            random_state=self.random_state,
            class_weight='balanced'
        )
        self.rf_model.fit(X_train, y_train)
        return self.rf_model

    def train_logistic_regression(self, X_train_scaled, y_train):
        self.lr_model = LogisticRegression(
            max_iter=1000,
            random_state=self.random_state,
            class_weight='balanced'
        )
        self.lr_model.fit(X_train_scaled, y_train)
        return self.lr_model

    def evaluate(self, model, X_test, y_test) -> dict:
        preds = model.predict(X_test)
        acc = accuracy_score(y_test, preds)
        cm = confusion_matrix(y_test, preds)
        report = classification_report(y_test, preds, output_dict=True)
        return {"accuracy": acc, "confusion_matrix": cm, "report": report}

    def get_feature_importances(self) -> pd.Series:
        if self.rf_model is None:
            raise ValueError("Random Forest model has not been trained.")
        return pd.Series(self.rf_model.feature_importances_, index=self.feature_cols).sort_values(ascending=True)


class Visualizer:
    """Handles generation of visual diagnostic charts."""
    def __init__(self, output_dir: str = "."):
        self.output_dir = output_dir

    def plot_feature_importance(self, importances: pd.Series, filename: str = "feature_importance.png"):
        filepath = os.path.join(self.output_dir, filename)
        fig, ax = plt.subplots(figsize=(10, 6))
        importances.plot(kind='barh', ax=ax, color='steelblue', edgecolor='black')
        baseline = 1.0 / len(importances)
        ax.axvline(x=baseline, color='red', linestyle='--', label=f'Uniform baseline ({baseline:.3f})')
        ax.set_title('Feature Importance (Random Forest Gini)', fontsize=14)
        ax.set_xlabel('Importance Score')
        ax.legend()
        plt.tight_layout()
        plt.savefig(filepath, dpi=150)
        plt.close()
        return filepath

    def plot_correlation_heatmap(self, df: pd.DataFrame, cols: list, filename: str = "correlation_heatmap.png"):
        filepath = os.path.join(self.output_dir, filename)
        fig, ax = plt.subplots(figsize=(12, 8))
        corr_matrix = df[cols].corr()
        sns.heatmap(corr_matrix, annot=True, fmt='.2f', cmap='RdBu_r', center=0, ax=ax, square=True)
        ax.set_title('Feature Correlation Heatmap', fontsize=14)
        plt.tight_layout()
        plt.savefig(filepath, dpi=150)
        plt.close()
        return filepath


class WaterborneDiseasePipeline:
    """Master orchestrator class composing all pipeline stages."""
    def __init__(self, data_path: str = "water_potability.csv"):
        self.data_path = data_path
        self.loader = DataLoader(filepath=self.data_path)
        self.cleaner = DataCleaner()
        self.engineer = FeatureEngineer()
        self.feature_cols = [
            'ph', 'Hardness', 'Solids', 'Chloramines', 'Sulfate',
            'Conductivity', 'Organic_carbon', 'Trihalomethanes',
            'Turbidity', 'CRI', 'ph_deviation'
        ]
        self.trainer = ModelTrainer(feature_cols=self.feature_cols)
        self.visualizer = Visualizer()

    def run(self):
        print("=" * 60)
        print("  Running Object-Oriented Waterborne Disease ML Pipeline")
        print("=" * 60)

        # 1. Load
        raw_df = self.loader.load_data()
        summary = self.loader.get_summary()
        print(f"[DataLoader] Loaded shape: {summary['shape']}, Duplicates: {summary['duplicates']}")

        # 2. Clean
        cleaned_df = self.cleaner.fit_transform(raw_df)
        print(f"[DataCleaner] Missing values after median imputation: {cleaned_df.isnull().sum().sum()}")

        # 3. Engineer
        engineered_df = self.engineer.engineer_all(cleaned_df)
        print(f"[FeatureEngineer] Added CRI, ph_deviation, and risk_level features.")

        # 4. Train & Evaluate
        X_train, X_test, X_train_s, X_test_s, y_train, y_test = self.trainer.prepare_split(engineered_df)
        
        self.trainer.train_random_forest(X_train, y_train)
        rf_metrics = self.trainer.evaluate(self.trainer.rf_model, X_test, y_test)
        print(f"[ModelTrainer] Random Forest Accuracy: {rf_metrics['accuracy']:.4f}")

        self.trainer.train_logistic_regression(X_train_s, y_train)
        lr_metrics = self.trainer.evaluate(self.trainer.lr_model, X_test_s, y_test)
        print(f"[ModelTrainer] Logistic Regression Accuracy: {lr_metrics['accuracy']:.4f}")

        # 5. Visualize
        importances = self.trainer.get_feature_importances()
        fi_path = self.visualizer.plot_feature_importance(importances)
        heatmap_path = self.visualizer.plot_correlation_heatmap(engineered_df, self.feature_cols + ['Potability'])
        print(f"[Visualizer] Generated charts: '{fi_path}', '{heatmap_path}'")

        print("=" * 60)
        print("  [SUCCESS] OOP Pipeline Execution Completed Successfully!")
        print("=" * 60)


if __name__ == '__main__':
    pipeline = WaterborneDiseasePipeline(data_path="water_potability.csv")
    pipeline.run()
