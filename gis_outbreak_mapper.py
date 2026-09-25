"""
Real-World GIS Disease Outbreak Risk Mapping Engine (PRJ_441)
Translates Water Quality ML Predictions into Spatial Outbreak Risk Intelligence
and Actionable Public Health Advisories.
"""
import sys
import os
import pandas as pd
import numpy as np
import folium
from folium import plugins
from sklearn.ensemble import RandomForestClassifier

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


def get_pathogen_and_advisory(row, rf_risk_prob):
    """
    Translates chemical/physical anomalies and ML risk probability
    into epidemiological disease risks and municipal advisories.
    """
    issues = []
    pathogens = []
    advisories = []

    # 1. Biological/Microbial Vector Condition: High Turbidity & Organic Carbon
    if row['Turbidity'] > 4.5 or row['Organic_carbon'] > 16.0:
        issues.append("Elevated microbial habitat (high turbidity / organic load)")
        pathogens.extend(["Vibrio cholerae (Cholera)", "Salmonella typhi (Typhoid)", "Rotavirus"])
        advisories.append("Issue immediate boil-water advisory; accelerate secondary coagulation.")

    # 2. Inadequate Disinfection: Low Chloramines
    if row['Chloramines'] < 4.0:
        issues.append("Sub-therapeutic disinfectant residual (< 4.0 ppm)")
        pathogens.append("Bacterial re-growth in secondary distribution pipelines")
        advisories.append("Dispatch emergency chlorine booster injection units.")

    # 3. Chemical Toxicity: Elevated Trihalomethanes & Sulfate
    if row['Trihalomethanes'] > 80.0:
        issues.append("Trihalomethanes exceed WHO limit (80 ug/L)")
        pathogens.append("Chemical toxicity hazard / chronic carcinogenic risk")
        advisories.append("Optimize precursor organic removal before chlorination.")

    if row['Sulfate'] > 350.0:
        issues.append("Excess sulfate concentration (> 350 mg/L)")
        pathogens.append("Acute gastrointestinal irritation / osmotic diarrhea")
        advisories.append("Dilute with alternate deep-aquifer reservoir supply.")

    # Extreme pH
    if row['ph'] < 6.5 or row['ph'] > 8.5:
        issues.append(f"pH out of safe physiological envelope ({row['ph']:.2f})")
        advisories.append("Adjust chemical lime / acid dosers for corrosion control.")

    # Calculate final Outbreak Risk Level
    composite_risk = (rf_risk_prob * 0.6) + (row['CRI'] * 0.4)

    if composite_risk >= 0.65:
        tier = "CRITICAL OUTBREAK ALERT"
        color = "red"
        icon_color = "red"
    elif composite_risk >= 0.48:
        tier = "MODERATE RISK - MONITOR"
        color = "orange"
        icon_color = "orange"
    else:
        tier = "LOW RISK - POTABLE"
        color = "green"
        icon_color = "green"

    if not pathogens:
        pathogens = ["None detected / within baseline safety limits"]
    if not advisories:
        advisories = ["Maintain routine sensor calibration and regular distribution pressure."]

    return {
        "composite_risk": composite_risk,
        "risk_tier": tier,
        "marker_color": color,
        "primary_threats": list(set(pathogens))[:3],
        "advisory": advisories[0],
        "key_issues": issues if issues else ["Parameters within acceptable baseline standard"]
    }


def generate_gis_surveillance_map(output_file: str = "outbreak_risk_map.html"):
    print("[GIS Engine] Initializing Spatial Water Quality Surveillance Engine...")

    # Load cleaned engineered data and train model
    data_path = "water_potability_engineered.csv"
    if not os.path.exists(data_path):
        from pipeline_oop import WaterborneDiseasePipeline
        pipeline = WaterborneDiseasePipeline("water_potability.csv")
        pipeline.run()

    df = pd.read_csv(data_path)
    feature_cols = [
        'ph', 'Hardness', 'Solids', 'Chloramines', 'Sulfate',
        'Conductivity', 'Organic_carbon', 'Trihalomethanes',
        'Turbidity', 'CRI', 'ph_deviation'
    ]

    # Fit RF model
    rf = RandomForestClassifier(n_estimators=150, max_depth=10, random_state=42, class_weight='balanced')
    rf.fit(df[feature_cols], df['Potability'])

    # Sample representative monitoring stations (Regional Public Health surveillance points)
    # Geocoded across urban, peri-urban, and water treatment network nodes
    stations = [
        {"id": "STN-101", "name": "Central Water Treatment Plant - Outlet", "lat": 12.9716, "lon": 77.5946, "idx": 15},
        {"id": "STN-102", "name": "Sector 4 High-Density Settlement Pipeline", "lat": 12.9900, "lon": 77.5700, "idx": 1},
        {"id": "STN-103", "name": "East Lake Reservoir Intake", "lat": 12.9550, "lon": 77.6400, "idx": 24},
        {"id": "STN-104", "name": "North River Basin Catchment Well", "lat": 13.0300, "lon": 77.5900, "idx": 45},
        {"id": "STN-105", "name": "Industrial Corridor Drainage Boundary", "lat": 12.9200, "lon": 77.6100, "idx": 88},
        {"id": "STN-106", "name": "South Agricultural Runoff Canal", "lat": 12.9100, "lon": 77.5500, "idx": 102},
        {"id": "STN-107", "name": "Metropolitan Hospital District Feeder", "lat": 12.9800, "lon": 77.6200, "idx": 150},
        {"id": "STN-108", "name": "Old City Sub-Surface Borewell Station", "lat": 12.9650, "lon": 77.5350, "idx": 210},
    ]

    # Initialize Folium GIS map centered at regional surveillance center
    m = folium.Map(location=[12.9716, 77.5946], zoom_start=12, tiles="OpenStreetMap")

    # Add heatmap layer for outbreak intensity
    heat_data = []

    print("[GIS Engine] Analyzing station readings & synthesizing outbreak advisories...")
    for stn in stations:
        row = df.iloc[stn["idx"]]
        features = df.iloc[[stn["idx"]]][feature_cols]
        
        # RF prediction probabilities (probability of non-potable / contamination = class 0)
        prob_non_potable = rf.predict_proba(features)[0][0]
        
        assessment = get_pathogen_and_advisory(row, prob_non_potable)
        heat_data.append([stn["lat"], stn["lon"], assessment["composite_risk"]])

        popup_html = f"""
        <div style="font-family: Arial, sans-serif; width: 310px; font-size: 13px; line-height: 1.4;">
            <div style="background-color: {'#d9534f' if assessment['marker_color']=='red' else '#f0ad4e' if assessment['marker_color']=='orange' else '#5cb85c'}; 
                        color: white; padding: 8px; border-radius: 4px; font-weight: bold; font-size: 14px;">
                {assessment['risk_tier']}
            </div>
            <p style="margin: 8px 0 4px 0;"><strong>Station:</strong> {stn['name']} (<code>{stn['id']}</code>)</p>
            <hr style="border: 0; border-top: 1px solid #ddd; margin: 6px 0;">
            <table style="width: 100%; font-size: 12px; margin-bottom: 6px;">
                <tr><td><strong>Outbreak Prob:</strong></td><td style="color:{assessment['marker_color']}; font-weight: bold;">{assessment['composite_risk']*100:.1f}%</td></tr>
                <tr><td><strong>Turbidity:</strong></td><td>{row['Turbidity']:.2f} NTU</td></tr>
                <tr><td><strong>pH:</strong></td><td>{row['ph']:.2f}</td></tr>
                <tr><td><strong>Chloramines:</strong></td><td>{row['Chloramines']:.2f} ppm</td></tr>
                <tr><td><strong>Sulfate:</strong></td><td>{row['Sulfate']:.1f} mg/L</td></tr>
            </table>
            <div style="background-color: #f8f9fa; padding: 6px; border-left: 3px solid #0275d8; margin-bottom: 6px;">
                <strong>Suspected Pathogen Hazards:</strong><br>
                <span style="color: #333; font-size: 11px;">{", ".join(assessment['primary_threats'])}</span>
            </div>
            <div style="background-color: #fff3cd; padding: 6px; border-left: 3px solid #ffc107; font-size: 11px;">
                <strong>Municipal Action Protocol:</strong><br>
                {assessment['advisory']}
            </div>
        </div>
        """

        # Add Marker with styled Popup
        folium.CircleMarker(
            location=[stn["lat"], stn["lon"]],
            radius=12 if assessment["marker_color"] == "red" else 9,
            popup=folium.Popup(popup_html, max_width=350),
            color=assessment["marker_color"],
            fill=True,
            fill_color=assessment["marker_color"],
            fill_opacity=0.85,
            tooltip=f"{stn['id']}: {stn['name']} [{assessment['risk_tier']}]"
        ).add_to(m)

    # Add Outbreak Intensity Heatmap Layer
    plugins.HeatMap(heat_data, radius=25, blur=18, max_zoom=13).add_to(m)

    # Add GIS Legend
    legend_html = '''
     <div style="position: fixed; 
     bottom: 30px; left: 30px; width: 230px; height: 135px; 
     background-color: white; border:2px solid grey; z-index:9999; font-size:12px;
     padding: 10px; border-radius: 6px; opacity: 0.95; font-family: Arial;">
     <b>Epidemiological Outbreak Risk</b><br>
     <i style="background:red; width:12px; height:12px; float:left; margin-right:8px; border-radius:50%;"></i> Critical (Outbreak Imminent)<br>
     <i style="background:orange; width:12px; height:12px; float:left; margin-right:8px; border-radius:50%;"></i> Moderate (Action Required)<br>
     <i style="background:green; width:12px; height:12px; float:left; margin-right:8px; border-radius:50%;"></i> Low Risk (Water Safe)<br>
     <hr style="margin: 5px 0;">
     <small>PRJ_441 GIS Spatial Surveillance</small>
     </div>
     '''
    m.get_root().html.add_child(folium.Element(legend_html))

    m.save(output_file)
    print(f"[GIS Engine] [OK] Outbreak risk surveillance map successfully generated: '{output_file}'")
    return output_file


if __name__ == '__main__':
    generate_gis_surveillance_map()
