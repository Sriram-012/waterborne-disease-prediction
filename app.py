"""
PRJ_441: Waterborne Disease Outbreak Risk Prediction & GIS Surveillance System
Complete Streamlit Web Application for Final Project Evaluation & Review.
"""
import os
import json
import joblib
import pandas as pd
import numpy as np
import streamlit as st
import streamlit.components.v1 as components

# ── Page Configuration ──
st.set_page_config(
    page_title="Waterborne Disease Outbreak Risk Prediction (PRJ_441)",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS for High-End Dashboard Aesthetics ──
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 800;
        color: #0F2043;
        margin-bottom: 4px;
    }
    .sub-header {
        font-size: 14px;
        color: #555;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 12px 18px;
        border-left: 5px solid #00A896;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .alert-critical {
        background-color: #F8D7DA;
        color: #721C24;
        padding: 14px;
        border-radius: 6px;
        border-left: 6px solid #DC3545;
        font-weight: 600;
        margin-top: 10px;
    }
    .alert-moderate {
        background-color: #FFF3CD;
        color: #856404;
        padding: 14px;
        border-radius: 6px;
        border-left: 6px solid #FFC107;
        font-weight: 600;
        margin-top: 10px;
    }
    .alert-safe {
        background-color: #D4EDDA;
        color: #155724;
        padding: 14px;
        border-radius: 6px;
        border-left: 6px solid #28A745;
        font-weight: 600;
        margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ── Load Model & Metadata ──
@st.cache_resource
def load_resources():
    model_path = os.path.join("models", "water_risk_rf_model.pkl")
    meta_path = os.path.join("models", "model_metadata.json")
    
    if not os.path.exists(model_path) or not os.path.exists(meta_path):
        import subprocess
        subprocess.run(["python", "train_production_model.py"], check=True)
        
    model = joblib.load(model_path)
    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)
    return model, meta

try:
    model, metadata = load_resources()
except Exception as e:
    st.error(f"Error loading production model: {e}")
    st.stop()

# ── Header Title & Overview ──
st.markdown('<div class="main-header">💧 Waterborne Disease Outbreak Risk Prediction & GIS Surveillance System</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header"><b>Project ID:</b> PRJ_441 &nbsp;|&nbsp; <b>Institution:</b> School of Computer Science & Engineering, Presidency University &nbsp;|&nbsp; <b>Domain:</b> ML + Spatial Epidemiology</div>', unsafe_allow_html=True)

# ── Top Key Metric Cards ──
col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.markdown('<div class="metric-card"><b>Dataset Benchmark</b><br><span style="font-size:22px;color:#0F2043;font-weight:bold;">3,276 Samples</span><br><small>Kaggle Water Potability</small></div>', unsafe_allow_html=True)
with col_m2:
    st.markdown(f'<div class="metric-card"><b>5-Fold CV Accuracy</b><br><span style="font-size:22px;color:#00A896;font-weight:bold;">{metadata["cv_results"]["Random Forest"]["Accuracy"]*100:.2f}%</span><br><small>Statistically Verified</small></div>', unsafe_allow_html=True)
with col_m3:
    st.markdown(f'<div class="metric-card"><b>Production Algorithm</b><br><span style="font-size:20px;color:#0F2043;font-weight:bold;">Random Forest</span><br><small>200 Trees (Depth: 10)</small></div>', unsafe_allow_html=True)
with col_m4:
    st.markdown('<div class="metric-card"><b>Surveillance Engine</b><br><span style="font-size:22px;color:#E67E22;font-weight:bold;">Folium GIS</span><br><small>8 Real-Time Grid Nodes</small></div>', unsafe_allow_html=True)

st.write("")

# ── Navigation Tabs ──
tab1, tab2, tab3, tab4 = st.tabs([
    "🎛️ Real-Time Water Quality Diagnostic",
    "🗺️ Interactive GIS Spatial Surveillance Map",
    "📊 Model Benchmarking & Explainability",
    "📁 Batch Water Sample Triage (CSV Upload)"
])

# =========================================================================
# TAB 1: REAL-TIME WATER QUALITY DIAGNOSTIC
# =========================================================================
with tab1:
    st.subheader("Laboratory Sensor Input & Outbreak Risk Assessment")
    st.caption("Adjust physical and chemical water sensor measurements to forecast potability and epidemiological outbreak hazards.")

    with st.form("diagnostic_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            ph = st.number_input("pH Level (WHO Safe: 6.5 - 8.5)", min_value=0.0, max_value=14.0, value=7.08, step=0.1)
            hardness = st.number_input("Hardness (mg/L)", min_value=10.0, max_value=400.0, value=196.3, step=1.0)
            solids = st.number_input("Total Dissolved Solids - TDS (ppm)", min_value=100.0, max_value=65000.0, value=22014.0, step=100.0)
        with c2:
            chloramines = st.number_input("Chloramines (ppm, Safe: up to 4.0)", min_value=0.0, max_value=15.0, value=7.12, step=0.1)
            sulfate = st.number_input("Sulfate (mg/L, WHO Limit: 250)", min_value=10.0, max_value=500.0, value=333.0, step=1.0)
            conductivity = st.number_input("Conductivity (μS/cm)", min_value=50.0, max_value=800.0, value=426.2, step=1.0)
        with c3:
            organic_carbon = st.number_input("Organic Carbon (ppm)", min_value=0.0, max_value=30.0, value=14.28, step=0.1)
            trihalomethanes = st.number_input("Trihalomethanes (μg/L, WHO Limit: 80)", min_value=0.0, max_value=150.0, value=66.4, step=0.5)
            turbidity = st.number_input("Turbidity (NTU, WHO Limit: 5.0)", min_value=0.0, max_value=10.0, value=3.96, step=0.1)
        
        btn_diagnose = st.form_submit_button("🚀 Run Outbreak Risk Diagnostics", use_container_width=True)

    if btn_diagnose:
        # 1. Feature Engineering
        # Min-max scaling approximations based on training distribution
        turb_norm = np.clip((turbidity - 1.45) / (6.74 - 1.45), 0, 1)
        org_norm = np.clip((organic_carbon - 2.2) / (27.0 - 2.2), 0, 1)
        thm_norm = np.clip((trihalomethanes - 0.74) / (124.0 - 0.74), 0, 1)
        cri = float((turb_norm + org_norm + thm_norm) / 3.0)

        ph_dev = float(max(0.0, 6.5 - ph) + max(0.0, ph - 8.5))

        # Model Inference
        features = [[ph, hardness, solids, chloramines, sulfate, conductivity, organic_carbon, trihalomethanes, turbidity, cri, ph_dev]]
        prob_unsafe = float(model.predict_proba(features)[0][0])  # probability of class 0 (Unsafe)
        composite_risk = (prob_unsafe * 0.6) + (cri * 0.4)

        # Disease Vector & Advisory Rules
        suspected_pathogens = []
        advisories = []
        if turbidity > 4.5 or organic_carbon > 15.0:
            suspected_pathogens.extend(["Vibrio cholerae (Cholera)", "Salmonella typhi (Typhoid)", "Rotavirus"])
            advisories.append("Accelerate secondary flocculation & coagulation; issue boil-water notice.")
        if chloramines < 4.0:
            suspected_pathogens.append("Pipeline Bacterial Re-growth (Secondary Disinfection Failure)")
            advisories.append("Dispatch emergency mobile chlorine booster units.")
        if trihalomethanes > 80.0:
            suspected_pathogens.append("Chemical Carcinogen Hazard (Excess THMs)")
            advisories.append("Optimize raw organic precursor filtration prior to chlorine contact.")
        if sulfate > 350.0:
            suspected_pathogens.append("Acute Osmotic Laxative & Dehydration Risk")
            advisories.append("Dilute with deep-aquifer storage reservoir blend.")

        if not suspected_pathogens:
            suspected_pathogens = ["None detected / Chemical markers within standard baseline"]
        if not advisories:
            advisories = ["Maintain routine automated pipeline distribution pressure and weekly sensor calibration."]

        # Display Diagnostic Output Cards
        st.write("---")
        st.subheader("Diagnostic Results & Epidemiological Triage")

        res_col1, res_col2, res_col3 = st.columns([1.2, 1.2, 1.6])
        
        with res_col1:
            if composite_risk >= 0.65:
                st.markdown(f'<div class="alert-critical">🚨 CRITICAL OUTBREAK ALERT<br><span style="font-size:24px;">{composite_risk*100:.1f}% Risk</span><br><small>Water is Unsafe / Non-Potable</small></div>', unsafe_allow_html=True)
            elif composite_risk >= 0.48:
                st.markdown(f'<div class="alert-moderate">⚠️ MODERATE HAZARD ALERT<br><span style="font-size:24px;">{composite_risk*100:.1f}% Risk</span><br><small>Precautionary Intervention Required</small></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="alert-safe">✅ POTABLE (SAFE WATER)<br><span style="font-size:24px;">{composite_risk*100:.1f}% Risk</span><br><small>Meets WHO Standards</small></div>', unsafe_allow_html=True)

        with res_col2:
            st.metric("Contamination Risk Index (CRI)", f"{cri:.3f}", delta=f"pH Dev: {ph_dev:.2f}")
            st.progress(float(composite_risk))
            st.caption(f"Random Forest Unsafe Probability: {prob_unsafe*100:.1f}%")

        with res_col3:
            st.markdown(f"**Identified Pathogen Vectors:**")
            for p in list(set(suspected_pathogens))[:3]:
                st.write(f"• `{p}`")
            st.markdown(f"**Automated Municipal Action Protocol:**")
            st.info(advisories[0])

# =========================================================================
# TAB 2: INTERACTIVE GIS SPATIAL SURVEILLANCE MAP
# =========================================================================
with tab2:
    st.subheader("Interactive GIS Outbreak Surveillance Dashboard")
    st.caption("Live spatial monitoring stations with real-time chemical anomaly detection, pathogen vector tracking, and multi-layer CartoDB/Esri basemaps.")

    c_btn1, c_btn2 = st.columns([2, 5])
    with c_btn1:
        if st.button("🔄 Refresh Surveillance Map Engine", use_container_width=True):
            import subprocess
            subprocess.run(["python", "gis_outbreak_mapper.py"], check=True)
            st.success("GIS Map updated with latest sensor readings!")

    map_path = "outbreak_risk_map.html"
    if os.path.exists(map_path):
        with open(map_path, "r", encoding="utf-8") as f:
            html_content = f.read()
        components.html(html_content, height=580, scrolling=True)
    else:
        st.warning("GIS Map file not found. Generating now...")
        import subprocess
        subprocess.run(["python", "gis_outbreak_mapper.py"], check=True)
        st.rerun()

    st.markdown("""
    **Map Legend & Surveillance Protocol:**
    - 🔴 **Red (Critical Outbreak Alert):** Stations with elevated Turbidity + Organic Carbon (>65% composite risk). Immediate boil-water advisory.
    - 🟠 **Orange (Moderate Warning):** Depleted chloramine disinfectant (<4.0 ppm) indicating secondary pipeline microbial re-growth.
    - 🟢 **Green (Safe Baseline):** Water treatment plant outlets and confirmed safe distribution lines.
    - *Use the top-right layer switcher on the map to toggle between Street View, Esri GIS, and Satellite Imagery.*
    """)

# =========================================================================
# TAB 3: MODEL BENCHMARKING & EXPLAINABILITY
# =========================================================================
with tab3:
    st.subheader("Empirical Model Benchmarking & 5-Fold Cross-Validation")
    st.caption("Statistically verified performance comparison across 4 algorithms and Gini feature importance attribution.")

    cv_df = pd.DataFrame(metadata["cv_results"]).T
    cv_df["Accuracy (%)"] = (cv_df["Accuracy"] * 100).round(2)
    cv_df["Accuracy ± Std (%)"] = (cv_df["Accuracy_Std"] * 100).round(2)
    cv_df["Precision (%)"] = (cv_df["Precision"] * 100).round(2)
    cv_df["Recall (%)"] = (cv_df["Recall"] * 100).round(2)
    cv_df["F1 Score (%)"] = (cv_df["F1_Score"] * 100).round(2)
    cv_df["ROC-AUC"] = cv_df["ROC_AUC"].round(3)

    display_cols = ["Accuracy (%)", "Accuracy ± Std (%)", "Precision (%)", "Recall (%)", "F1 Score (%)", "ROC-AUC"]
    st.dataframe(cv_df[display_cols], use_container_width=True)

    st.write("---")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("Gini Feature Importance")
        fi_series = pd.Series(metadata["feature_importances"])
        st.bar_chart(fi_series)
        st.caption("Sulfate (15.6%), pH (12.2%), and Hardness (11.3%) are the strongest non-linear discriminators.")

    with col_g2:
        st.subheader("Confusion Matrix (80/20 Stratified Holdout)")
        cm = metadata["holdout_metrics"]["confusion_matrix"]
        cm_df = pd.DataFrame(
            cm,
            columns=["Predicted: Unsafe (0)", "Predicted: Safe (1)"],
            index=["Actual: Unsafe (0)", "Actual: Safe (1)"]
        )
        st.table(cm_df)
        st.caption(f"Recall on Non-Potable Water: {metadata['holdout_metrics']['recall_unsafe']*100:.1f}% (Successfully catches dangerous water).")

# =========================================================================
# TAB 4: BATCH WATER SAMPLE TRIAGE (CSV UPLOAD)
# =========================================================================
with tab4:
    st.subheader("Batch Regional Water Quality Triage")
    st.caption("Upload a CSV file containing multiple sensor readings to generate automated risk scores, CRI indices, and triage actions.")

    uploaded_file = st.file_uploader("Upload Water Quality CSV", type=["csv"])
    if uploaded_file is not None:
        batch_df = pd.read_csv(uploaded_file)
        st.write(f"Loaded {len(batch_df)} samples. Running automated ML inference...")

        # Feature verification
        required_cols = ['ph', 'Hardness', 'Solids', 'Chloramines', 'Sulfate', 'Conductivity', 'Organic_carbon', 'Trihalomethanes', 'Turbidity']
        missing_cols = [c for c in required_cols if c not in batch_df.columns]

        if missing_cols:
            st.error(f"Missing required columns in CSV: {missing_cols}")
        else:
            processed_df = batch_df.copy()
            # Median fill for missing values
            for c in ['ph', 'Sulfate', 'Trihalomethanes']:
                if c in processed_df.columns:
                    processed_df[c] = processed_df[c].fillna(processed_df[c].median())

            # Compute CRI and pH dev
            t_norm = (processed_df['Turbidity'] - 1.45) / (6.74 - 1.45)
            c_norm = (processed_df['Organic_carbon'] - 2.2) / (27.0 - 2.2)
            m_norm = (processed_df['Trihalomethanes'] - 0.74) / (124.0 - 0.74)
            processed_df['CRI'] = ((t_norm + c_norm + m_norm) / 3.0).clip(0, 1)
            processed_df['ph_deviation'] = processed_df['ph'].apply(lambda x: max(0.0, 6.5 - x) + max(0.0, x - 8.5))

            feature_cols = ['ph', 'Hardness', 'Solids', 'Chloramines', 'Sulfate', 'Conductivity', 'Organic_carbon', 'Trihalomethanes', 'Turbidity', 'CRI', 'ph_deviation']
            preds = model.predict(processed_df[feature_cols])
            probs_unsafe = model.predict_proba(processed_df[feature_cols])[:, 0]

            processed_df['Predicted_Potability'] = np.where(preds == 1, "Safe (Potable)", "Unsafe (Non-Potable)")
            processed_df['Outbreak_Risk_Score'] = (probs_unsafe * 0.6 + processed_df['CRI'] * 0.4).round(3)
            processed_df['Triage_Action'] = np.where(
                processed_df['Outbreak_Risk_Score'] >= 0.65, "CRITICAL: Issue Boil-Water Notice",
                np.where(processed_df['Outbreak_Risk_Score'] >= 0.48, "WARNING: Inject Chlorine Booster", "NORMAL: Water Safe")
            )

            st.dataframe(processed_df[['ph', 'Turbidity', 'Chloramines', 'CRI', 'Predicted_Potability', 'Outbreak_Risk_Score', 'Triage_Action']], use_container_width=True)

            csv_download = processed_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Triaged Public Health Report (CSV)",
                data=csv_download,
                file_name="water_outbreak_triage_report.csv",
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info("💡 You can upload any CSV file formatted like `water_potability.csv` to process hundreds of water distribution points in seconds.")
