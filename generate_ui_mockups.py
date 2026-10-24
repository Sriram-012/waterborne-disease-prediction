"""
generate_ui_mockups.py
Generates clean, high-resolution UI screenshots representing the Streamlit dashboard and GIS surveillance map.
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

def make_ui_dashboard():
    fig, ax = plt.subplots(figsize=(10, 6.5), dpi=200)
    ax.axis('off')

    # Background window
    bg = patches.Rectangle((0, 0), 1, 1, facecolor='#F8F9FA', edgecolor='#CBD5E1', linewidth=1.5)
    ax.add_patch(bg)

    # Top Navigation bar
    top_bar = patches.Rectangle((0, 0.90), 1, 0.10, facecolor='#1F4E79')
    ax.add_patch(top_bar)
    ax.text(0.04, 0.95, "💧 Waterborne Disease Outbreak Risk Prediction & Spatial Surveillance (PRJ_441)",
            fontsize=11, fontweight='bold', color='white', va='center')
    ax.text(0.85, 0.95, "Presidency University | v2.4",
            fontsize=9, color='#D9E1F2', va='center')

    # Tabs
    tab_bar = patches.Rectangle((0, 0.84), 1, 0.06, facecolor='#E2E8F0')
    ax.add_patch(tab_bar)
    tabs = ["Tab 1: Real-Time Risk Calculator", "Tab 2: Spatial GIS Surveillance Map", "Tab 3: Model Performance Analytics", "Tab 4: Batch CSV Triage"]
    for idx, t in enumerate(tabs):
        tab_bg = patches.Rectangle((0.02 + idx*0.24, 0.84), 0.23, 0.06,
                                   facecolor='#FFFFFF' if idx == 0 else '#E2E8F0',
                                   edgecolor='#CBD5E1')
        ax.add_patch(tab_bg)
        ax.text(0.135 + idx*0.24, 0.87, t, ha='center', va='center',
                fontsize=8, fontweight='bold' if idx == 0 else 'normal',
                color='#1F4E79' if idx == 0 else '#475569')

    # Left Column: Sliders Box
    left_box = patches.FancyBboxPatch((0.03, 0.08), 0.44, 0.72, boxstyle="round,pad=0.01",
                                      facecolor='white', edgecolor='#CBD5E1', linewidth=1)
    ax.add_patch(left_box)
    ax.text(0.05, 0.76, "Water Quality Sensor Inputs", fontsize=10.5, fontweight='bold', color='#1E293B')

    sliders = [
        ("pH Level (WHO Safe Range: 6.5 - 8.5)", "7.15", 0.70),
        ("Sulfate Concentration (mg/L)", "338.5", 0.63),
        ("Hardness (CaCO3 mg/L)", "198.2", 0.56),
        ("Chloramines (Disinfectant ppm)", "7.12", 0.49),
        ("Total Dissolved Solids (ppm)", "21,500", 0.42),
        ("Turbidity (NTU)", "4.15", 0.35),
        ("Organic Carbon (Microbial Nutrient ppm)", "14.28", 0.28),
        ("Trihalomethanes (Disinfection Byproduct µg/L)", "66.4", 0.21),
        ("Contamination Risk Index (CRI - Derived)", "0.512", 0.14)
    ]

    for label, val, y_pos in sliders:
        ax.text(0.05, y_pos + 0.02, label, fontsize=7.5, color='#334155')
        # Slider track
        ax.plot([0.05, 0.36], [y_pos, y_pos], color='#CBD5E1', linewidth=4, solid_capstyle='round')
        ax.plot([0.05, 0.22], [y_pos, y_pos], color='#2E75B6', linewidth=4, solid_capstyle='round')
        ax.plot(0.22, y_pos, marker='o', color='#1F4E79', markersize=7)
        ax.text(0.42, y_pos, val, ha='right', va='center', fontsize=8, fontweight='bold', color='#1F4E79')

    # Right Column: Prediction Output Card
    right_box = patches.FancyBboxPatch((0.50, 0.45), 0.47, 0.35, boxstyle="round,pad=0.01",
                                       facecolor='#F0F9FF', edgecolor='#0284C7', linewidth=1.5)
    ax.add_patch(right_box)
    ax.text(0.53, 0.76, "Random Forest ML Risk Classification", fontsize=10.5, fontweight='bold', color='#0369A1')

    # Risk Badge
    badge = patches.FancyBboxPatch((0.53, 0.63), 0.41, 0.08, boxstyle="round,pad=0.01",
                                   facecolor='#DCFCE7', edgecolor='#16A34A', linewidth=1.5)
    ax.add_patch(badge)
    ax.text(0.735, 0.67, "SAFE / LOW OUTBREAK RISK (Class 1)", ha='center', va='center',
            fontsize=10, fontweight='bold', color='#15803D')

    ax.text(0.53, 0.58, "• Outbreak Probability: 28.4% (Below 48% Threshold)", fontsize=8.5, color='#334155')
    ax.text(0.53, 0.53, "• Potability Confidence: 71.6%", fontsize=8.5, color='#334155')
    ax.text(0.53, 0.48, "• Recommended Action: Normal Baseline Chlorination Monitoring", fontsize=8.5, color='#334155')

    # Right Bottom: Spatial Surveillance Preview Card
    right_box2 = patches.FancyBboxPatch((0.50, 0.08), 0.47, 0.33, boxstyle="round,pad=0.01",
                                        facecolor='white', edgecolor='#CBD5E1', linewidth=1)
    ax.add_patch(right_box2)
    ax.text(0.53, 0.37, "Municipal Ward Surveillance Status", fontsize=10, fontweight='bold', color='#1E293B')
    ax.text(0.53, 0.32, "Active Monitoring Nodes: 12 Bangalore Wards", fontsize=8, color='#64748B')

    wards = [
        ("Ward 112 (Yelahanka Lake)", "0.78 (CRITICAL)", "#DC2626"),
        ("Ward 85 (Hebbal Intake)", "0.54 (WARNING)", "#D97706"),
        ("Ward 42 (Malleswaram)", "0.28 (SAFE)", "#16A34A"),
        ("Ward 174 (HSR Layout)", "0.32 (SAFE)", "#16A34A")
    ]
    for idx, (w_name, status, col) in enumerate(wards):
        y_w = 0.26 - idx * 0.045
        ax.text(0.53, y_w, f"• {w_name}:", fontsize=8, color='#334155')
        ax.text(0.92, y_w, status, ha='right', fontsize=8, fontweight='bold', color=col)

    plt.tight_layout()
    out = 'report_assets/ui_dashboard.png'
    plt.savefig(out, bbox_inches='tight')
    plt.close()
    print(f"Created {out}")

def make_gis_map_view():
    fig, ax = plt.subplots(figsize=(10, 6.5), dpi=200)
    ax.axis('off')

    # Simulated GIS Map Background
    bg = patches.Rectangle((0, 0), 1, 1, facecolor='#E5F3FF', edgecolor='#94A3B8', linewidth=1.5)
    ax.add_patch(bg)

    # Top Banner
    ax.text(0.04, 0.95, "GIS Outbreak Risk Surveillance Map (Folium / Leaflet Engine)",
            fontsize=12, fontweight='bold', color='#0F172A')
    ax.text(0.04, 0.91, "Bengaluru Urban Municipal Wards | Base Layers: CartoDB Positron & Esri Satellite",
            fontsize=8.5, color='#475569')

    # Water bodies
    lake1 = patches.Ellipse((0.35, 0.65), 0.18, 0.10, angle=30, facecolor='#93C5FD', edgecolor='#3B82F6')
    lake2 = patches.Ellipse((0.68, 0.38), 0.22, 0.12, angle=-20, facecolor='#93C5FD', edgecolor='#3B82F6')
    ax.add_patch(lake1)
    ax.add_patch(lake2)
    ax.text(0.35, 0.65, "Hebbal Lake", fontsize=7.5, color='#1E40AF', ha='center', va='center')
    ax.text(0.68, 0.38, "Bellandur Basin", fontsize=7.5, color='#1E40AF', ha='center', va='center')

    # Road network simulation
    ax.plot([0.1, 0.9], [0.5, 0.5], color='#CBD5E1', linewidth=2)
    ax.plot([0.5, 0.5], [0.1, 0.9], color='#CBD5E1', linewidth=2)
    ax.plot([0.2, 0.8], [0.2, 0.8], color='#E2E8F0', linewidth=1.5)

    # Risk Nodes
    nodes = [
        (0.32, 0.68, "Ward 112 (Yelahanka)", "Risk: 0.78 (Critical)", "#EF4444", "CRITICAL"),
        (0.72, 0.42, "Ward 150 (Bellandur)", "Risk: 0.71 (Critical)", "#EF4444", "CRITICAL"),
        (0.55, 0.52, "Ward 85 (Hebbal)", "Risk: 0.54 (Warning)", "#F59E0B", "WARNING"),
        (0.40, 0.35, "Ward 94 (Shivajinagar)", "Risk: 0.50 (Warning)", "#F59E0B", "WARNING"),
        (0.25, 0.30, "Ward 42 (Malleswaram)", "Risk: 0.28 (Safe)", "#10B981", "SAFE"),
        (0.60, 0.22, "Ward 174 (HSR Layout)", "Risk: 0.31 (Safe)", "#10B981", "SAFE"),
        (0.78, 0.70, "Ward 65 (K.R. Puram)", "Risk: 0.25 (Safe)", "#10B981", "SAFE")
    ]

    for x, y, name, r_text, col, status in nodes:
        # Outer pulse circle
        ax.plot(x, y, marker='o', markersize=14, color=col, alpha=0.3)
        # Inner solid pin
        ax.plot(x, y, marker='o', markersize=8, color=col)
        # Label popup
        popup = patches.FancyBboxPatch((x + 0.02, y - 0.02), 0.18, 0.05,
                                       boxstyle="round,pad=0.005",
                                       facecolor='white', edgecolor=col, linewidth=1)
        ax.add_patch(popup)
        ax.text(x + 0.03, y + 0.012, name, fontsize=6.5, fontweight='bold', color='#1E293B')
        ax.text(x + 0.03, y - 0.012, r_text, fontsize=6, color=col, fontweight='bold')

    # Map Legend
    leg_box = patches.FancyBboxPatch((0.05, 0.06), 0.32, 0.20, boxstyle="round,pad=0.01",
                                     facecolor='white', edgecolor='#94A3B8', linewidth=1)
    ax.add_patch(leg_box)
    ax.text(0.07, 0.22, "Surveillance Risk Legend", fontsize=8.5, fontweight='bold', color='#0F172A')
    
    leg_items = [
        ("Critical Outbreak Alert (Risk >= 0.65)", "#EF4444"),
        ("Moderate Hazard Warning (0.48 - 0.65)", "#F59E0B"),
        ("Low Baseline Risk (Risk < 0.48)", "#10B981")
    ]
    for idx, (label, col) in enumerate(leg_items):
        y_leg = 0.17 - idx * 0.045
        ax.plot(0.08, y_leg, marker='o', markersize=7, color=col)
        ax.text(0.10, y_leg, label, fontsize=7, color='#334155', va='center')

    plt.tight_layout()
    out = 'report_assets/gis_map_view.png'
    plt.savefig(out, bbox_inches='tight')
    plt.close()
    print(f"Created {out}")

if __name__ == '__main__':
    make_ui_dashboard()
    make_gis_map_view()
    print("UI mockups generated successfully.")
