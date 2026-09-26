"""
Generates the official Presidency University Review-2 PowerPoint Presentation (.pptx)
for PRJ_441: Waterborne Disease Outbreak Risk Prediction System (ML + GIS).
Creates a 100% editable PPTX file preserving the exact template structure and theme.
"""
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Color Palette (Presidency University Theme: Deep Navy & Teal) ──
COLOR_PRIMARY_NAVY = RGBColor(15, 32, 67)     # #0F2043 (Deep Academic Navy)
COLOR_ACCENT_TEAL  = RGBColor(0, 168, 150)    # #00A896 (Presidency Teal Accent)
COLOR_DARK_TEXT    = RGBColor(40, 44, 52)     # #282C34 (Dark Slate Text)
COLOR_MUTED_TEXT   = RGBColor(100, 110, 125)  # #646E7D (Subtitle / Muted Text)
COLOR_BG_CARD      = RGBColor(245, 248, 250)  # #F5F8FA (Light Soft Card BG)
COLOR_WHITE        = RGBColor(255, 255, 255)  # #FFFFFF
COLOR_BORDER       = RGBColor(218, 225, 233)  # #DAE1E9 (Card Border)
COLOR_ALERT_RED    = RGBColor(220, 53, 69)    # #DC3545
COLOR_ALERT_ORANGE = RGBColor(240, 140, 30)   # #F08C1E
COLOR_ALERT_GREEN  = RGBColor(40, 167, 69)    # #28A745


def add_header(slide, title_text, category_text="CSS7102 - Mini Project Review-2"):
    """Adds standard Presidency University branded header to slides."""
    # Top Accent Stripe (Teal)
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = COLOR_ACCENT_TEAL
    stripe.line.fill.background()

    # Category Text (Small caps over title)
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.25), Inches(11.7), Inches(0.35))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = f"PRESIDENCY UNIVERSITY  •  SCHOOL OF COMPUTER SCIENCE AND ENGINEERING  •  {category_text.upper()}"
    p_cat.font.size = Pt(9.5)
    p_cat.font.bold = True
    p_cat.font.color.rgb = COLOR_ACCENT_TEAL

    # Main Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.7), Inches(0.75))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.size = Pt(22)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_PRIMARY_NAVY

    # Subtle divider line
    div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.3), Inches(11.733), Inches(0.02))
    div.fill.solid()
    div.fill.fore_color.rgb = COLOR_BORDER
    div.line.fill.background()


def add_card(slide, left, top, width, height, bg_color=COLOR_BG_CARD, border_color=COLOR_BORDER):
    """Creates a clean rounded container box for structured information."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
    else:
        card.line.fill.background()
    return card


def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (PRESIDENCY UNIVERSITY TEMPLATE)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)

    # Top Brand Header Banner
    header_rect = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.2))
    header_rect.fill.solid()
    header_rect.fill.fore_color.rgb = COLOR_PRIMARY_NAVY
    header_rect.line.fill.background()

    # Teal Banner Accent Line
    banner_stripe = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(13.333), Inches(0.08))
    banner_stripe.fill.solid()
    banner_stripe.fill.fore_color.rgb = COLOR_ACCENT_TEAL
    banner_stripe.line.fill.background()

    # University Header Text
    u_box = slide1.shapes.add_textbox(Inches(0.8), Inches(0.2), Inches(11.7), Inches(0.8))
    tf_u = u_box.text_frame
    p_u1 = tf_u.paragraphs[0]
    p_u1.text = "PRESIDENCY UNIVERSITY"
    p_u1.font.size = Pt(18)
    p_u1.font.bold = True
    p_u1.font.color.rgb = COLOR_WHITE
    p_u2 = tf_u.add_paragraph()
    p_u2.text = "SCHOOL OF COMPUTER SCIENCE AND ENGINEERING  |  DEPARTMENT OF CSE"
    p_u2.font.size = Pt(11)
    p_u2.font.color.rgb = COLOR_ACCENT_TEAL

    # Review Subheading Pill
    pill = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.5), Inches(0.45))
    pill.fill.solid()
    pill.fill.fore_color.rgb = COLOR_ACCENT_TEAL
    pill.line.fill.background()
    tf_pill = pill.text_frame
    p_pill = tf_pill.paragraphs[0]
    p_pill.text = "CSS7102 - MINI PROJECT REVIEW-2"
    p_pill.font.size = Pt(11)
    p_pill.font.bold = True
    p_pill.font.color.rgb = COLOR_WHITE
    p_pill.alignment = PP_ALIGN.CENTER

    # Project Title & ID
    title_box = slide1.shapes.add_textbox(Inches(0.8), Inches(2.05), Inches(11.7), Inches(1.4))
    tf_main = title_box.text_frame
    tf_main.word_wrap = True
    p_m1 = tf_main.paragraphs[0]
    p_m1.text = "Waterborne Disease Outbreak Risk Prediction System (ML + GIS)"
    p_m1.font.size = Pt(26)
    p_m1.font.bold = True
    p_m1.font.color.rgb = COLOR_PRIMARY_NAVY
    p_m2 = tf_main.add_paragraph()
    p_m2.text = "Project ID: PRJ_441  |  Domain: Machine Learning & Spatial Epidemiology Surveillance"
    p_m2.font.size = Pt(13)
    p_m2.font.color.rgb = COLOR_ACCENT_TEAL

    # Left Column Card: Supervision & Program Details
    add_card(slide1, Inches(0.8), Inches(3.6), Inches(5.6), Inches(3.4))
    sup_box = slide1.shapes.add_textbox(Inches(1.0), Inches(3.75), Inches(5.2), Inches(3.1))
    tf_sup = sup_box.text_frame
    tf_sup.word_wrap = True

    p = tf_sup.paragraphs[0]
    p.text = "FACULTY & PROGRAM DETAILS"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY

    details = [
        ("Under the Supervision of:", "Faculty Mentor / Project Guide"),
        ("Designation:", "Assistant / Associate Professor, Dept. of CSE"),
        ("Name of the Program:", "B.Tech in Computer Science and Engineering"),
        ("Name of HoD (Team Leader's):", "Head of Department, CSE"),
        ("Program Project Coordinator:", "Program Coordinator, CSE"),
        ("School Project Coordinator:", "Mr. Muthuraju V"),
        ("Panel No:", "Panel - [Your Panel Number]")
    ]
    for label, val in details:
        p = tf_sup.add_paragraph()
        p.text = f"• {label} {val}"
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_DARK_TEXT

    # Right Column: Student Group Members Table
    add_card(slide1, Inches(6.8), Inches(3.6), Inches(5.7), Inches(3.4))
    tbl_title = slide1.shapes.add_textbox(Inches(7.0), Inches(3.75), Inches(5.3), Inches(0.4))
    p_tbl = tbl_title.text_frame.paragraphs[0]
    p_tbl.text = "GROUP MEMBERS DETAILS (3 MEMBERS)"
    p_tbl.font.size = Pt(12)
    p_tbl.font.bold = True
    p_tbl.font.color.rgb = COLOR_PRIMARY_NAVY

    # Table Shape
    table_shape = slide1.shapes.add_table(4, 3, Inches(7.0), Inches(4.3), Inches(5.3), Inches(2.4))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(0.8)
    tbl.columns[1].width = Inches(2.7)
    tbl.columns[2].width = Inches(1.8)

    headers = ["Sl.No", "Student Name", "Roll Number"]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_PRIMARY_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    students = [
        ("1", "Sriram (Team Leader)", "[Your Roll No]"),
        ("2", "[Student Name 2]", "[Roll Number 2]"),
        ("3", "[Student Name 3]", "[Roll Number 3]")
    ]
    for row_idx, (sno, name, roll) in enumerate(students, start=1):
        tbl.cell(row_idx, 0).text = sno
        tbl.cell(row_idx, 1).text = name
        tbl.cell(row_idx, 2).text = roll
        for c in range(3):
            cell = tbl.cell(row_idx, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_DARK_TEXT
            if c == 0 or c == 2:
                p.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 2: INDEX / AGENDA
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Presentation Index & Review Structure")

    agenda_items = [
        ("01", "Problem Statement", "The 48-hour biological lab lag, water contamination crisis, and target stakeholders."),
        ("02", "Objectives", "Measurable goals: dataset cleaning, CRI feature engineering, ML models, and GIS mapping."),
        ("03", "Literature / Research Understanding", "Existing water quality studies, baseline limitations, and identified research gap."),
        ("04", "Methodology & Architecture", "5-tier end-to-end pipeline, Object-Oriented software design (`pipeline_oop.py`)."),
        ("05", "Implementation Progress (Review-2)", "Achieved 60% milestone: data cleaning, Random Forest vs. LR, and diagnostic plots."),
        ("06", "GIS Outbreak Surveillance & Live Demo", "Spatial risk mapping (`outbreak_risk_map.html`), pathogen vectors, and municipal alerts."),
        ("07", "Individual Contribution & GitHub Link", "Exact role breakdown across 3 members, verified public GitHub repository link."),
        ("08", "Conclusion & References", "Summary of findings, remaining 40% roadmap, and IEEE/APA citations.")
    ]

    for i, (num, title, desc) in enumerate(agenda_items):
        col = i % 2
        row = i // 2
        left = Inches(0.8 + col * 5.95)
        top = Inches(1.55 + row * 1.35)

        add_card(slide2, left, top, Inches(5.75), Inches(1.2))

        # Number Badge
        badge = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.15), top + Inches(0.2), Inches(0.8), Inches(0.8))
        badge.fill.solid()
        badge.fill.fore_color.rgb = COLOR_PRIMARY_NAVY
        badge.line.fill.background()
        p_b = badge.text_frame.paragraphs[0]
        p_b.text = num
        p_b.font.size = Pt(14)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_WHITE
        p_b.alignment = PP_ALIGN.CENTER

        # Title & Subtext
        t_box = slide2.shapes.add_textbox(left + Inches(1.05), top + Inches(0.15), Inches(4.5), Inches(0.9))
        tf = t_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_PRIMARY_NAVY

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = COLOR_DARK_TEXT

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Problem Statement: The 48-Hour Biological Lab Delay")

    cards_p3 = [
        ("THE CORE REAL-WORLD PROBLEM", 
         "• Traditional microbiological detection for waterborne pathogens (e.g., Vibrio cholerae causing Cholera, Salmonella typhi causing Typhoid) requires culturing bacteria on agar plates in a biology laboratory.\n"
         "• This culture process requires 24 to 72 hours to yield definitive results.\n"
         "• During this 48-hour testing lag, contaminated municipal water continues to flow through residential pipelines, infecting communities before official warnings can be issued.",
         COLOR_ALERT_RED),
        ("THE EXISTING GAP & LIMITATION",
         "• Existing municipal systems rely on manual, delayed laboratory testing schedules rather than automated early-warning detection.\n"
         "• Standard single-parameter testing (e.g., checking pH alone) fails because water safety is governed by complex, multi-factor chemical interactions.\n"
         "• Raw lab reports lack spatial GIS integration, leaving disaster teams without geographical visualization of active contamination clusters.",
         COLOR_ALERT_ORANGE),
        ("WHO IS AFFECTED & WHY SOLVE IT?",
         "• Affected Stakeholders: Urban residential sectors, high-density settlements, flood-affected regions, and municipal health boards.\n"
         "• Human Impact: Over 500,000 diarrheal deaths occur globally every year due to unsafe drinking water (WHO).\n"
         "• Proposed Solution: Rapid chemical sensors (pH, turbidity, chloramines, sulfate) deliver readings in seconds. Our system uses Machine Learning as an instant proxy to predict outbreak risks before lab cultures arrive.",
         COLOR_ALERT_GREEN)
    ]

    for i, (title, body, accent) in enumerate(cards_p3):
        top = Inches(1.55 + i * 1.85)
        add_card(slide3, Inches(0.8), top, Inches(11.733), Inches(1.7))

        # Color Accent Left Bar
        accent_bar = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(0.15), Inches(1.7))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = accent
        accent_bar.line.fill.background()

        t_box = slide3.shapes.add_textbox(Inches(1.1), top + Inches(0.1), Inches(11.2), Inches(1.5))
        tf = t_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_PRIMARY_NAVY

        for line in body.split("\n"):
            p = tf.add_paragraph()
            p.text = line
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_DARK_TEXT

    # =========================================================================
    # SLIDE 4: OBJECTIVES
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Project Objectives: Measurable Technical Milestones")

    objectives = [
        ("Objective 1: Real-World Data Ingestion & Missing Value Imputation",
         "Ingest 3,276 real-world laboratory water test samples from the Kaggle benchmark dataset and implement Median Imputation on 1,434 missing values across pH, Sulfate, and Trihalomethanes, preserving 100% of samples without data loss."),
        ("Objective 2: Domain-Driven Feature Engineering",
         "Formulate and compute the Contamination Risk Index (CRI) combining normalized Turbidity, Organic Carbon, and Trihalomethanes, and compute absolute pH deviation outside the WHO drinking water safety envelope [6.5, 8.5]."),
        ("Objective 3: Non-Linear Machine Learning Benchmarking",
         "Train, evaluate, and benchmark a 200-tree Random Forest Classifier against a Logistic Regression baseline on an 80/20 stratified train-test split with balanced class weighting to handle the 61/39 potability class imbalance."),
        ("Objective 4: Interpretability & Diagnostic Attribution",
         "Quantify feature importance using Gini Impurity reduction to identify primary chemical drivers (Sulfate: 15.6%, pH: 12.2%, Hardness: 11.3%) and generate correlation heatmaps verifying feature independence."),
        ("Objective 5: GIS Spatial Surveillance & Municipal Advisory Automation",
         "Build an interactive Folium GIS surveillance dashboard (`outbreak_risk_map.html`) mapping monitoring stations into Critical, Moderate, and Safe tiers with automated disease vector alerts and municipal action protocols.")
    ]

    for i, (title, desc) in enumerate(objectives):
        top = Inches(1.55 + i * 1.1)
        add_card(slide4, Inches(0.8), top, Inches(11.733), Inches(0.98))

        # Objective Circle Badge
        badge = slide4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.0), top + Inches(0.18), Inches(0.6), Inches(0.6))
        badge.fill.solid()
        badge.fill.fore_color.rgb = COLOR_ACCENT_TEAL
        badge.line.fill.background()
        p_b = badge.text_frame.paragraphs[0]
        p_b.text = f"O{i+1}"
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_WHITE
        p_b.alignment = PP_ALIGN.CENTER

        t_box = slide4.shapes.add_textbox(Inches(1.75), top + Inches(0.08), Inches(10.6), Inches(0.85))
        tf = t_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_PRIMARY_NAVY

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = COLOR_DARK_TEXT

    # =========================================================================
    # SLIDE 5: LITERATURE / RESEARCH UNDERSTANDING
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Literature Survey & Research Understanding")

    # Table of Existing Work
    tbl_shape5 = slide5.shapes.add_table(4, 4, Inches(0.8), Inches(1.55), Inches(11.733), Inches(3.2))
    tbl5 = tbl_shape5.table
    tbl5.columns[0].width = Inches(2.2)
    tbl5.columns[1].width = Inches(3.4)
    tbl5.columns[2].width = Inches(3.2)
    tbl5.columns[3].width = Inches(2.933)

    cols5 = ["Study / Reference", "Approach & Technology Used", "Key Findings & Achievements", "Identified Limitations"]
    for c_idx, h in enumerate(cols5):
        cell = tbl5.cell(0, c_idx)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_PRIMARY_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    rows5 = [
        ("WHO Water Quality Guidelines (4th Edition, 2022)",
         "Established health-based thresholds for pH (6.5-8.5), Turbidity (<5 NTU), Chloramines (4 ppm), and THMs (80 ug/L).",
         "Standardized physical and chemical bounds required to prevent pathogenic disease transmission.",
         "Provides static safety thresholds; does not model complex multivariate interactions or provide automated predictions."),
        ("Kaggle Water Potability Benchmark Studies (2021-2024)",
         "Standard ML classification (SVM, Decision Trees, Random Forest, XGBoost) on 3,276 water quality samples.",
         "Confirmed published academic benchmark accuracy ranges between 64% and 68% due to low individual feature correlation.",
         "Evaluated only binary 0/1 potability in code; lacked domain feature engineering (CRI) and spatial GIS mapping."),
        ("Traditional Municipal Testing Systems",
         "Manual sample collection followed by agar-plate microbiological laboratory culturing.",
         "High biological certainty for confirming specific bacterial colonies (E. coli, Cholera).",
         "Requires 24-72 hour testing delay; unable to provide real-time early warning before public consumption.")
    ]

    for r_idx, row in enumerate(rows5, start=1):
        for c_idx, text in enumerate(row):
            cell = tbl5.cell(r_idx, c_idx)
            cell.text = text
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_BG_CARD if r_idx % 2 == 0 else COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_DARK_TEXT

    # Research Gap Card
    add_card(slide5, Inches(0.8), Inches(5.0), Inches(11.733), Inches(1.9), bg_color=RGBColor(240, 247, 255), border_color=COLOR_ACCENT_TEAL)
    rg_box = slide5.shapes.add_textbox(Inches(1.0), Inches(5.1), Inches(11.3), Inches(1.7))
    tf_rg = rg_box.text_frame
    tf_rg.word_wrap = True
    p_rg = tf_rg.paragraphs[0]
    p_rg.text = "RESEARCH GAP IDENTIFIED & OUR PROPOSED CONTRIBUTION"
    p_rg.font.size = Pt(11)
    p_rg.font.bold = True
    p_rg.font.color.rgb = COLOR_ACCENT_TEAL

    p_rg_body = tf_rg.add_paragraph()
    p_rg_body.text = (
        "• The Research Gap: Existing works treat water potability as an isolated tabular classification problem (predicting 0 or 1), "
        "ignoring spatial clustering, epidemiological pathogen translation, and actionable municipal response workflows.\n"
        "• Our Contribution: We bridge the 48-hour lab testing gap by converting rapid chemical sensor data into an integrated pipeline featuring: "
        "(1) Contamination Risk Index (CRI), (2) Robust OOP modular architecture, and (3) An interactive GIS surveillance dashboard with automated public health advisories."
    )
    p_rg_body.font.size = Pt(9.5)
    p_rg_body.font.color.rgb = COLOR_DARK_TEXT

    # =========================================================================
    # SLIDE 6: METHODOLOGY AND ARCHITECTURE
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Methodology & 5-Tier System Architecture")

    arch_tiers = [
        ("1. Data Ingestion Tier", "Kaggle Benchmark Dataset\n• 3,276 water samples\n• 9 WHO chemical features\n• 1 Potability Target (0/1)", COLOR_PRIMARY_NAVY),
        ("2. Preprocessing Tier", "Data Cleaning Module\n• Missingness: 1,434 NaNs\n• Median Imputation applied\n• 0 duplicates / 0 data loss", COLOR_ACCENT_TEAL),
        ("3. Feature Engineering", "Domain Risk Synthesizer\n• Contamination Risk Index (CRI)\n• WHO pH Safe Envelope Deviation\n• 3-Tier Quantile Risk Binning", COLOR_PRIMARY_NAVY),
        ("4. ML Model Engine", "Classifier Benchmarking\n• Random Forest (200 trees)\n• Logistic Regression\n• Stratified 80/20 Train-Test", COLOR_ACCENT_TEAL),
        ("5. GIS Surveillance", "Interactive Spatial Map\n• Folium Multi-layer Map\n• Pathogen Vector Mapping\n• Automated Municipal Alerts", COLOR_PRIMARY_NAVY)
    ]

    for i, (t_title, t_desc, accent) in enumerate(arch_tiers):
        left = Inches(0.8 + i * 2.38)
        top = Inches(1.55)
        add_card(slide6, left, top, Inches(2.25), Inches(2.9))

        header_pill = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.1), top + Inches(0.1), Inches(2.05), Inches(0.45))
        header_pill.fill.solid()
        header_pill.fill.fore_color.rgb = accent
        header_pill.line.fill.background()
        p_hp = header_pill.text_frame.paragraphs[0]
        p_hp.text = t_title
        p_hp.font.size = Pt(9.5)
        p_hp.font.bold = True
        p_hp.font.color.rgb = COLOR_WHITE
        p_hp.alignment = PP_ALIGN.CENTER

        t_box = slide6.shapes.add_textbox(left + Inches(0.1), top + Inches(0.65), Inches(2.05), Inches(2.1))
        tf = t_box.text_frame
        tf.word_wrap = True
        for idx, line in enumerate(t_desc.split("\n")):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_DARK_TEXT
            if idx == 0:
                p.font.bold = True

    # OOP Architecture Summary Card (Lower Half)
    add_card(slide6, Inches(0.8), Inches(4.65), Inches(11.733), Inches(2.35))
    oop_box = slide6.shapes.add_textbox(Inches(1.0), Inches(4.75), Inches(11.3), Inches(2.1))
    tf_oop = oop_box.text_frame
    tf_oop.word_wrap = True
    p_oop = tf_oop.paragraphs[0]
    p_oop.text = "OBJECT-ORIENTED SOFTWARE ARCHITECTURE (pipeline_oop.py)"
    p_oop.font.size = Pt(11)
    p_oop.font.bold = True
    p_oop.font.color.rgb = COLOR_PRIMARY_NAVY

    oop_classes = [
        ("DataLoader", "Ingests `water_potability.csv`, validates dataset shape (3,276x10), and verifies 0 duplicate entries."),
        ("DataCleaner", "Calculates feature medians and applies median imputation on pH, Sulfate, and Trihalomethanes."),
        ("FeatureEngineer", "Computes Contamination Risk Index (CRI), pH deviation, and quantile risk categories (Low, Med, High)."),
        ("ModelTrainer", "Performs 80/20 stratified split, StandardScaler normalization, trains RF and LR, and computes Gini importances."),
        ("Visualizer", "Renders and exports diagnostic artifacts: `feature_importance.png` and `correlation_heatmap.png`."),
        ("WaterborneDiseasePipeline", "Master orchestrator class composing all modules into an automated end-to-end execution workflow.")
    ]

    for cname, cdesc in oop_classes:
        p = tf_oop.add_paragraph()
        p.text = f"• {cname}: {cdesc}"
        p.font.size = Pt(9)
        p.font.color.rgb = COLOR_DARK_TEXT

    # =========================================================================
    # SLIDE 7: IMPLEMENTATION PROGRESS (REVIEW-2: 60% MILESTONE)
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Implementation Progress: Review-2 Milestone (60% Achieved)")

    # Left Column: Metrics & Implementation Status
    add_card(slide7, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.4))
    m_box = slide7.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
    tf_m = m_box.text_frame
    tf_m.word_wrap = True

    p = tf_m.paragraphs[0]
    p.text = "VERIFIED MODEL PERFORMANCE & MILESTONE STATUS"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY

    p = tf_m.add_paragraph()
    p.text = "Review-2 Requirement: Minimum 50% implementation expected.\nCurrent Project Status: ~60% Core ML Pipeline + GIS Fully Operational."
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_TEAL

    p = tf_m.add_paragraph()
    p.text = "\nBenchmarking Results (80/20 Stratified Split):"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY

    results = [
        ("Random Forest Classifier (Primary):", "62.96% Accuracy (Beats 61% baseline)"),
        ("Logistic Regression (Baseline):", "50.46% Accuracy (Linear model fails)"),
        ("Precision on Non-Potable (Class 0):", "0.67 (Reliable contamination alert)"),
        ("Recall on Non-Potable (Class 0):", "0.76 (Captures 76% of unsafe water)"),
        ("Top Contributing Feature (Gini):", "Sulfate (15.6%), pH (12.2%), Hardness (11.3%)"),
        ("Real Data Truth:", "63% matches published academic Kaggle benchmark (64-68%); proves zero data leakage vs fake 98% synthetic POC.")
    ]
    for lbl, val in results:
        p = tf_m.add_paragraph()
        p.text = f"• {lbl} {val}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK_TEXT

    # Right Column: Embed Diagnostic Artifacts if available
    add_card(slide7, Inches(6.8), Inches(1.55), Inches(5.733), Inches(5.4))
    p_img_title = slide7.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(0.4))
    p_img_title.text_frame.paragraphs[0].text = "DIAGNOSTIC ARTIFACTS (GENERATED PLOTS)"
    p_img_title.text_frame.paragraphs[0].font.size = Pt(12)
    p_img_title.text_frame.paragraphs[0].font.bold = True
    p_img_title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRIMARY_NAVY

    # Add feature importance image if exists
    fi_path = os.path.abspath("feature_importance.png")
    if os.path.exists(fi_path):
        slide7.shapes.add_picture(fi_path, Inches(7.0), Inches(2.2), width=Inches(5.3))
    else:
        fi_placeholder = slide7.shapes.add_textbox(Inches(7.0), Inches(2.5), Inches(5.3), Inches(2.0))
        fi_placeholder.text_frame.paragraphs[0].text = "[Feature Importance Plot Generated: feature_importance.png]"

    # Caption box
    cap_box = slide7.shapes.add_textbox(Inches(7.0), Inches(5.8), Inches(5.3), Inches(1.0))
    tf_cap = cap_box.text_frame
    tf_cap.word_wrap = True
    p_cap = tf_cap.paragraphs[0]
    p_cap.text = "Key Insight: Sulfate, pH, and Hardness lead Gini importance, but all features hover near the 9.1% uniform baseline, proving potability is an interdependent multi-parameter challenge."
    p_cap.font.size = Pt(8.5)
    p_cap.font.color.rgb = COLOR_MUTED_TEXT

    # =========================================================================
    # SLIDE 8: GIS SURVEILLANCE & REAL-WORLD WORKFLOW
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "GIS Spatial Outbreak Surveillance Engine (outbreak_risk_map.html)")

    gis_cards = [
        ("🔴 CRITICAL OUTBREAK ALERT (RED)",
         "• Station Condition: High Turbidity (>4.5 NTU) + Elevated Organic Carbon (>16 ppm) + High ML Risk (>65%).\n"
         "• Example Node: Station 102 (Sector 4 High-Density Settlement Pipeline).\n"
         "• Pathogen Vector Flagged: Vibrio cholerae (Cholera), Salmonella typhi (Typhoid).\n"
         "• Automated Municipal Protocol: 'Issue immediate boil-water advisory; accelerate secondary coagulation.'",
         COLOR_ALERT_RED),
        ("🟠 MODERATE RISK - ACTION REQUIRED (ORANGE)",
         "• Station Condition: Sub-therapeutic Chloramine disinfectant residual (<4.0 ppm) with moderate CRI (48-64%).\n"
         "• Example Node: Station 105 (Industrial Corridor Drainage Boundary).\n"
         "• Pathogen Vector Flagged: Bacterial re-growth in secondary distribution pipelines.\n"
         "• Automated Municipal Protocol: 'Dispatch emergency chlorine booster injection units.'",
         COLOR_ALERT_ORANGE),
        ("🟢 LOW RISK - POTABLE BASELINE (GREEN)",
         "• Station Condition: All parameters within WHO guideline limits, Composite Risk <48%.\n"
         "• Example Node: Station 101 (Central Water Treatment Plant Outlet).\n"
         "• Pathogen Vector Flagged: None detected / within baseline safety limits.\n"
         "• Automated Municipal Protocol: 'Maintain routine sensor calibration and regular distribution pressure.'",
         COLOR_ALERT_GREEN)
    ]

    for i, (title, desc, accent) in enumerate(gis_cards):
        top = Inches(1.55 + i * 1.85)
        add_card(slide8, Inches(0.8), top, Inches(11.733), Inches(1.7))

        accent_bar = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(0.15), Inches(1.7))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = accent
        accent_bar.line.fill.background()

        t_box = slide8.shapes.add_textbox(Inches(1.1), top + Inches(0.1), Inches(11.2), Inches(1.5))
        tf = t_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_PRIMARY_NAVY

        for line in desc.split("\n"):
            p = tf.add_paragraph()
            p.text = line
            p.font.size = Pt(9.5)
            p.font.color.rgb = COLOR_DARK_TEXT

    # =========================================================================
    # SLIDE 9: INDIVIDUAL CONTRIBUTION & WORK DIVISION
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "Individual Contribution & Team Responsibilities (3 Members)")

    members = [
        ("MEMBER 1: TEAM LEADER (Sriram)",
         "System Architecture & Data Engineering",
         [
             "Problem formulation & real-world clinical analysis (48-hr lab lag).",
             "Data loading, missing value profiling (1,434 NaNs), and Median Imputation logic (`01`).",
             "Object-Oriented modular class architecture in `pipeline_oop.py` (`DataLoader`, `DataCleaner`).",
             "Git repository management, code refactoring, and team workflow coordination."
         ],
         COLOR_PRIMARY_NAVY),
        ("MEMBER 2: TEAM MEMBER 2",
         "Feature Engineering & ML Benchmarking",
         [
             "Formulation of Contamination Risk Index (CRI) and WHO pH Safe Envelope Deviation (`02`).",
             "Model training: Random Forest Classifier (200 trees, depth 10, balanced class weights) (`03`).",
             "Logistic Regression baseline benchmarking & 80/20 stratified split evaluation (`03`).",
             "Gini feature importance extraction and correlation heatmap visualization (`04`)."
         ],
         COLOR_ACCENT_TEAL),
        ("MEMBER 3: TEAM MEMBER 3",
         "GIS Spatial Surveillance & Public Health Integration",
         [
             "GIS surveillance map engine implementation using Folium (`gis_outbreak_mapper.py`).",
             "Multi-layer basemap integration (CartoDB Voyager, Esri Street Map, Satellite View).",
             "Epidemiological pathogen vector mapping (*V. cholerae*, *S. typhi*, *Rotavirus*).",
             "Automated municipal action protocol generation and live interactive demonstration."
         ],
         COLOR_PRIMARY_NAVY)
    ]

    for i, (m_name, m_role, m_tasks, accent) in enumerate(members):
        left = Inches(0.8 + i * 3.98)
        top = Inches(1.55)
        add_card(slide9, left, top, Inches(3.78), Inches(5.4))

        # Member Header Pill
        pill = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.15), top + Inches(0.15), Inches(3.48), Inches(0.65))
        pill.fill.solid()
        pill.fill.fore_color.rgb = accent
        pill.line.fill.background()
        tf_p = pill.text_frame
        p1 = tf_p.paragraphs[0]
        p1.text = m_name
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_WHITE
        p1.alignment = PP_ALIGN.CENTER
        p2 = tf_p.add_paragraph()
        p2.text = m_role
        p2.font.size = Pt(8.5)
        p2.font.color.rgb = COLOR_BG_CARD
        p2.alignment = PP_ALIGN.CENTER

        # Tasks List
        t_box = slide9.shapes.add_textbox(left + Inches(0.15), top + Inches(0.9), Inches(3.48), Inches(4.3))
        tf_t = t_box.text_frame
        tf_t.word_wrap = True
        for t_idx, task in enumerate(m_tasks):
            p = tf_t.paragraphs[0] if t_idx == 0 else tf_t.add_paragraph()
            p.text = f"✔ {task}"
            p.font.size = Pt(9.5)
            p.font.color.rgb = COLOR_DARK_TEXT
            p.space_after = Pt(8)

    # =========================================================================
    # SLIDE 10: GITHUB REPOSITORY & LIVE CODE REPOSITORY
    # =========================================================================
    slide10 = prs.slides.add_slide(blank_layout)
    add_header(slide10, "GitHub Code Repository & Public Access")

    add_card(slide10, Inches(0.8), Inches(1.55), Inches(11.733), Inches(5.4))

    git_box = slide10.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(10.9), Inches(4.9))
    tf_git = git_box.text_frame
    tf_git.word_wrap = True

    p = tf_git.paragraphs[0]
    p.text = "OFFICIAL PROJECT CODE REPOSITORY (PUBLIC ACCESS)"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY

    p = tf_git.add_paragraph()
    p.text = "GitHub Repository Link: https://github.com/Sriram-012/waterborne-disease-prediction"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_TEAL

    p = tf_git.add_paragraph()
    p.text = "\nRepository Architecture & Tracked Code Files:"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY

    files_info = [
        ("01_data_loading_and_cleaning.py", "Data ingestion, missing value inspection, median imputation logic."),
        ("02_feature_engineering.py", "Contamination Risk Index (CRI) calculation and pH deviation formulation."),
        ("03_model_training.py", "Stratified 80/20 split, Random Forest and Logistic Regression training & scoring."),
        ("04_feature_importance.py", "Gini feature importance analysis, correlation heatmap export."),
        ("pipeline_oop.py", "Enterprise Object-Oriented pipeline comprising 6 dedicated modular classes."),
        ("gis_outbreak_mapper.py", "Folium GIS spatial engine with multi-layer maps, popups, and municipal alerts."),
        ("run_all.py", "Master execution script running all 5 stages and auto-launching the GIS map in browser."),
        ("README.md & requirements.txt", "Comprehensive setup instructions, dependencies, and execution guide.")
    ]

    for fname, fdesc in files_info:
        p = tf_git.add_paragraph()
        p.text = f"• {fname} — {fdesc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK_TEXT

    # =========================================================================
    # SLIDE 11: CONCLUSION & FUTURE ROADMAP (REMAINING 40%)
    # =========================================================================
    slide11 = prs.slides.add_slide(blank_layout)
    add_header(slide11, "Conclusion & Future Roadmap (Remaining 40%)")

    # Left: Conclusion
    add_card(slide11, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.4))
    c_box = slide11.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
    tf_c = c_box.text_frame
    tf_c.word_wrap = True

    p = tf_c.paragraphs[0]
    p.text = "PROJECT SUMMARY & REVIEW-2 CONCLUSIONS"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY

    conclusions = [
        "Rapid Early-Warning Feasibility: Demonstrated that rapid physical and chemical water parameters act as effective real-time proxies before 48-hour bacterial culture results arrive.",
        "Data Integrity & Zero Loss: Median imputation successfully handled 1,434 NaNs across pH, Sulfate, and Trihalomethanes without discarding any samples.",
        "Non-Linear Model Superiority: Random Forest (62.96% accuracy) significantly outperformed Logistic Regression (50.46%), proving water safety boundaries are non-linear.",
        "Spatial Public Health Decision Support: Successfully mapped risk nodes into Critical (Red), Moderate (Orange), and Safe (Green) tiers with automated municipal advisories.",
        "Production-Grade Software: Codebase completely structured under Object-Oriented design and tracked on public GitHub."
    ]
    for c in conclusions:
        p = tf_c.add_paragraph()
        p.text = f"✔ {c}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK_TEXT
        p.space_after = Pt(6)

    # Right: Roadmap for 40%
    add_card(slide11, Inches(6.8), Inches(1.55), Inches(5.733), Inches(5.4))
    r_box = slide11.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.1))
    tf_r = r_box.text_frame
    tf_r.word_wrap = True

    p = tf_r.paragraphs[0]
    p.text = "ROADMAP FOR REMAINING 40% (FINAL SUBMISSION)"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_TEAL

    roadmap = [
        ("Phase 1: K-Fold Stratified Cross-Validation", "Implement 5-Fold Stratified Cross-Validation to evaluate model variance across multiple independent data partitions."),
        ("Phase 2: Hyperparameter Tuning (GridSearchCV)", "Tune Random Forest hyperparameters (n_estimators, max_depth, min_samples_split) to optimize classification boundaries."),
        ("Phase 3: Class Imbalance Resampling (SMOTE)", "Apply Synthetic Minority Over-sampling Technique (SMOTE) to further improve recall on minority potable samples."),
        ("Phase 4: Streamlit Interactive Web Application", "Develop an interactive GUI allowing municipal operators to adjust chemical sliders and receive instant risk diagnostics."),
        ("Phase 5: Real-Time IoT Sensor Data Streaming", "Integrate MQTT/REST APIs to stream live sensor data directly into the GIS prediction engine.")
    ]
    for r_title, r_desc in roadmap:
        p = tf_r.add_paragraph()
        p.text = f"➔ {r_title}: {r_desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK_TEXT
        p.space_after = Pt(4)

    # =========================================================================
    # SLIDE 12: REFERENCES (IEEE & APA FORMAT)
    # =========================================================================
    slide12 = prs.slides.add_slide(blank_layout)
    add_header(slide12, "References (IEEE & APA Citations)")

    add_card(slide12, Inches(0.8), Inches(1.55), Inches(11.733), Inches(5.4))
    ref_box = slide12.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(11.1), Inches(5.1))
    tf_ref = ref_box.text_frame
    tf_ref.word_wrap = True

    p = tf_ref.paragraphs[0]
    p.text = "ACADEMIC & RESEARCH CITATIONS"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY

    references = [
        ("[1]", "World Health Organization. (2022). Guidelines for drinking-water quality: Fourth edition incorporating the first and second addenda. Geneva: World Health Organization. Licence: CC BY-NC-SA 3.0 IGO."),
        ("[2]", "Kadiwal, A. (2021). Water Potability Dataset: Water Quality Metrics for Potability Prediction. Kaggle Benchmark Dataset. https://www.kaggle.com/datasets/adityakadiwal/water-potability"),
        ("[3]", "Breiman, L. (2001). Random Forests. Machine Learning, 45(1), 5-32. https://doi.org/10.1023/A:1010933404324"),
        ("[4]", "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., ... & Duchesnay, E. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12, 2825-2830."),
        ("[5]", "Few, R., Lake, I., Hunter, P. R., & Tran, F. (2013). Floods, health and climate change: a strategic review. Tyndall Centre for Climate Change Research Working Paper 162."),
        ("[6]", "Saha, S., & Bhowmik, S. (2020). Water quality index prediction using machine learning algorithms: a comprehensive review. Environmental Monitoring and Assessment, 192(11), 1-28.")
    ]

    for num, ref in references:
        p = tf_ref.add_paragraph()
        p.text = f"{num} {ref}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK_TEXT
        p.space_after = Pt(8)

    # =========================================================================
    # SLIDE 13: THANK YOU SLIDE
    # =========================================================================
    slide13 = prs.slides.add_slide(blank_layout)

    # Full Background Banner
    bg_card = slide13.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg_card.fill.solid()
    bg_card.fill.fore_color.rgb = COLOR_PRIMARY_NAVY
    bg_card.line.fill.background()

    # Center Card
    center_card = slide13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.5), Inches(1.5), Inches(8.333), Inches(4.5))
    center_card.fill.solid()
    center_card.fill.fore_color.rgb = COLOR_WHITE
    center_card.line.fill.background()

    # Sticky Note Accent
    pin = slide13.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.36), Inches(1.3), Inches(0.6), Inches(0.6))
    pin.fill.solid()
    pin.fill.fore_color.rgb = COLOR_ACCENT_TEAL
    pin.line.fill.background()

    ty_box = slide13.shapes.add_textbox(Inches(2.8), Inches(2.2), Inches(7.733), Inches(3.5))
    tf_ty = ty_box.text_frame
    tf_ty.word_wrap = True

    p = tf_ty.paragraphs[0]
    p.text = "Thank You!"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_NAVY
    p.alignment = PP_ALIGN.CENTER

    p = tf_ty.add_paragraph()
    p.text = "Waterborne Disease Outbreak Risk Prediction System (ML + GIS)\nProject ID: PRJ_441  •  CSS7102 Review-2"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_ACCENT_TEAL
    p.alignment = PP_ALIGN.CENTER

    p = tf_ty.add_paragraph()
    p.text = "\nWe are now open for Questions, Feedback & Evaluation."
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_DARK_TEXT
    p.alignment = PP_ALIGN.CENTER

    p = tf_ty.add_paragraph()
    p.text = "GitHub: https://github.com/Sriram-012/waterborne-disease-prediction"
    p.font.size = Pt(10)
    p.font.color.rgb = COLOR_MUTED_TEXT
    p.alignment = PP_ALIGN.CENTER

    # Save presentation
    output_pptx = "PRJ_441_Presidency_University_Review2_Presentation.pptx"
    prs.save(output_pptx)
    print(f"[PPT Engine] [SUCCESS] Fully editable Review-2 presentation created: '{output_pptx}'")
    return output_pptx


if __name__ == '__main__':
    create_presentation()
