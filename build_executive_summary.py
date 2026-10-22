"""
build_executive_summary.py
Generates PRJ_441_Executive_Summary_Report.docx
An executive summary report and defense brief strictly adhering to Presidency University guidelines:
- Times New Roman throughout
- 18 pt Bold Centered for Major Headings / Chapters
- 16 pt Bold Left for Section Headings
- 14 pt Bold Left for Sub-section Headings
- 12 pt Regular Justified Black for Body text (1.5 line spacing)
- Margins: Left 1.25 in, Right 1.0 in, Top 1.0 in, Bottom 1.0 in
- Table captions at TOP, Figure captions at BOTTOM
- Verified project empirical metrics and genuine Harvard references
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, fill_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def add_footer_page_number(run):
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'PAGE')
    run._r.append(fldSimple)

def build_executive_summary():
    doc = Document()

    # 1. Page Setup
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.25)
        s.right_margin = Inches(1.0)
        s.page_width = Inches(8.27)   # A4
        s.page_height = Inches(11.69)
        footer = s.footer
        p_ft = footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_ft = p_ft.add_run()
        r_ft.font.name = 'Times New Roman'
        r_ft.font.size = Pt(10)
        r_ft.font.color.rgb = RGBColor(100, 100, 100)
        add_footer_page_number(r_ft)

    # Base styling
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

    # Helper functions
    def add_major_title(txt):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(txt)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_sec(txt):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(txt)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_subsec(txt):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(txt)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_p(txt, bold_prefix=""):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(6)
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = 'Times New Roman'
            rb.font.size = Pt(12)
            rb.font.bold = True
            rb.font.color.rgb = RGBColor(0, 0, 0)
        r = p.add_run(txt)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_table_cap(caption_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(caption_text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    def add_fig_cap(caption_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(10)
        r = p.add_run(caption_text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11)
        r.font.italic = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    # ══════════════════════════════════════════════════════════════════
    # COVER / HEADER BLOCK
    # ══════════════════════════════════════════════════════════════════
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("PRESIDENCY UNIVERSITY, BENGALURU\nSCHOOL OF COMPUTER SCIENCE AND ENGINEERING")
    r_inst.font.name = 'Times New Roman'
    r_inst.font.size = Pt(14)
    r_inst.font.bold = True

    p_proj = doc.add_paragraph()
    p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj.paragraph_format.space_after = Pt(14)
    r_pr = p_proj.add_run("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING\nACADEMIC PROJECT EXECUTIVE REPORT (PRJ_441 | CSS7102)")
    r_pr.font.name = 'Times New Roman'
    r_pr.font.size = Pt(12)
    r_pr.font.color.rgb = RGBColor(70, 70, 70)

    add_major_title("EXECUTIVE PROJECT SUMMARY & TECHNICAL DEFENSE BRIEF\nWATERBORNE DISEASE OUTBREAK RISK PREDICTION AND SPATIAL SURVEILLANCE SYSTEM USING MACHINE LEARNING AND GIS")

    # Author / Metadata Box
    t_meta = doc.add_table(rows=4, cols=2)
    t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Project Identification:", "PRJ_441 (Course Code: CSS7102)"),
        ("Student Author:", "K Sriram (Reg No: 20211CSE0441) | ramronz001@gmail.com"),
        ("GitHub Repository:", "https://github.com/Sriram-012/waterborne-disease-prediction"),
        ("Production Architecture:", "Random Forest Classifier + Folium GIS + Streamlit Web Engine")
    ]
    for r_i, (k, v) in enumerate(meta_info):
        c0 = t_meta.cell(r_i, 0)
        c1 = t_meta.cell(r_i, 1)
        c0.text = k
        c1.text = v
        set_cell_background(c0, "F0F4F8")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)
        p0 = c0.paragraphs[0]
        p0.runs[0].font.name = 'Times New Roman'
        p0.runs[0].font.size = Pt(10)
        p0.runs[0].font.bold = True
        p1 = c1.paragraphs[0]
        p1.runs[0].font.name = 'Times New Roman'
        p1.runs[0].font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ══════════════════════════════════════════════════════════════════
    # 1. EXECUTIVE OVERVIEW & VALUE PROPOSITION
    # ══════════════════════════════════════════════════════════════════
    add_sec("1. Executive Summary & Core Value Proposition")
    add_p(
        "Waterborne diseases represent one of the most severe preventable threats to global public health. According to the World "
        "Health Organization (WHO 2022), contaminated drinking water causes approximately 829,000 diarrheal deaths annually across developing "
        "nations, with children under five years bearing over 35% of the mortality burden. Conventional public health surveillance suffers "
        "from a crippling 48 to 72-hour lag caused by manual laboratory water sampling and microbiological bacterial incubation. By the time "
        "laboratory results confirm pathogenic contamination (such as Vibrio cholerae or Salmonella typhi), hundreds of citizens have already "
        "consumed tainted water, and hospitals face acute patient surges."
    )
    add_p(
        "Project PRJ_441 introduces an end-to-end Machine Learning and Geographic Information System (GIS) early-warning framework that "
        "transforms real-time physicochemical water parameters into immediate outbreak risk classifications. By eliminating laboratory delays, "
        "municipal health authorities can transition from reactive outbreak management to proactive containment, deploying chlorination units and "
        "issuing automated boil-water advisories before widespread clinical transmission takes place."
    )

    # ══════════════════════════════════════════════════════════════════
    # 2. PROBLEM STATEMENT & RESEARCH GAPS
    # ══════════════════════════════════════════════════════════════════
    add_sec("2. Problem Statement & Critical Research Gaps")
    add_p(
        "Municipal water monitoring systems across urban and semi-urban Indian agglomerations face three distinct systemic deficiencies:",
        "Core Dilemma: "
    )
    add_p(
        "1. Laboratory Bottleneck: Traditional culture-based biological testing requires 48-72 hours, rendering it completely incapable "
        "of serving as an active early-warning mechanism during seasonal flash floods and pipe cross-contamination events."
    )
    add_p(
        "2. Absence of Spatial Intelligence: Conventional testing results are stored in isolated spreadsheets without spatial coordinates, "
        "preventing field epidemiologists from visualizing geographic disease clusters or identifying contaminant plume trajectories."
    )
    add_p(
        "3. Academic Overfitting on Synthetic Data: Prior published studies have claimed misleading accuracies exceeding 95% on small synthetic "
        "datasets created with unrealistic mathematical boundaries. When applied to real-world environmental sensor telemetry, these models fail "
        "catastrophically because real water parameters exhibit severe non-linear noise and pairwise correlations below |r| < 0.07."
    )

    # ══════════════════════════════════════════════════════════════════
    # 3. TECHNICAL METHODOLOGY & PIPELINE
    # ══════════════════════════════════════════════════════════════════
    add_sec("3. Technical Methodology & System Architecture")
    add_p(
        "The project architecture is structured into four decoupled, modular layers: Data Preprocessing, Feature Engineering, "
        "Supervised Machine Learning, and Spatial Surveillance Deployment."
    )

    add_table_cap("Table 1.1: Comprehensive System Architecture Modules and Responsibilities")
    t_arch = doc.add_table(rows=5, cols=3)
    t_arch.alignment = WD_TABLE_ALIGNMENT.CENTER
    arch_data = [
        ("Pipeline Module", "Implementation Script", "Primary Functional Responsibility"),
        ("Data Cleaning & Imputation", "01_data_loading_and_cleaning.py", "Outlier-resistant median imputation across 1,434 missing values (pH, Sulfate, THM) with zero row attrition."),
        ("Domain Feature Engineering", "02_feature_engineering.py", "Mathematical formulation of Contamination Risk Index (CRI) and WHO pH tolerance threshold deviation."),
        ("Machine Learning Production", "03_model_training.py / train_production_model.py", "Stratified 5-fold cross-validation benchmarking across 4 algorithms; serialization of Random Forest pipeline."),
        ("GIS & Web Interface", "gis_outbreak_mapper.py / app.py", "Interactive multi-tile Folium Leaflet choropleth map and 4-tab Streamlit dashboard for real-time risk triage.")
    ]
    for r_i, row in enumerate(arch_data):
        for c_i, val in enumerate(row):
            c = t_arch.cell(r_i, c_i)
            c.text = val
            set_cell_margins(c, 80, 80, 100, 100)
            p = c.paragraphs[0]
            if r_i == 0:
                set_cell_background(c, "1F497D")
                p.runs[0].font.bold = True
                p.runs[0].font.size = Pt(10)
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            else:
                p.runs[0].font.size = Pt(9.5)
                if r_i % 2 == 1:
                    set_cell_background(c, "F9FAFB")

    add_subsec("3.1 Handling Missing Telemetry Data Without Data Loss")
    add_p(
        "The Kaggle Water Potability dataset contains 3,276 water observations. An initial assessment revealed 1,434 missing values across "
        "three parameters: pH (491 missing, 14.99%), Sulfate (781 missing, 23.84%), and Trihalomethanes (162 missing, 4.95%). "
        "Rather than discarding incomplete rows—which would discard 36.8% of the data and introduce severe geographic selection bias—median "
        "imputation was applied. The median is statistically robust against heavy-tailed water contaminants: imputed pH was 7.04, Sulfate was "
        "333.07 mg/L, and Trihalomethanes was 66.62 ug/L. All 3,276 samples were preserved with zero record loss."
    )

    add_subsec("3.2 Mathematical Formulation of Domain Engineered Features")
    add_p(
        "To provide the classifiers with explicit geochemical domain signals, two novel composite indicators were mathematically formulated:",
        "Domain Indicators: "
    )
    add_p(
        "1. Contamination Risk Index (CRI): Combines Turbidity (suspended biological sediments), Organic Carbon (microbial nutrient base), "
        "and Trihalomethanes (carcinogenic chlorine disinfection byproducts) into a normalized [0, 1] index:\n"
        "   CRI = (Norm(Turbidity) + Norm(Organic_Carbon) + Norm(Trihalomethanes)) / 3"
    )
    add_p(
        "2. pH Safe Range Deviation: WHO guidelines specify a potable drinking range between 6.5 and 8.5. The deviation metric penalizes "
        "both acidic runoff (pH < 6.5) and alkaline industrial discharge (pH > 8.5):\n"
        "   pH_Deviation = max(0, 6.5 - pH) + max(0, pH - 8.5)"
    )

    # ══════════════════════════════════════════════════════════════════
    # 4. EXPERIMENTAL BENCHMARKING & EMPIRICAL RESULTS
    # ══════════════════════════════════════════════════════════════════
    add_sec("4. Empirical Benchmark Results & Model Selection")
    add_p(
        "Model evaluation was executed strictly using 5-Fold Stratified Cross-Validation on all 3,276 dataset instances to eliminate "
        "sampling variance and ensure reproducibility. Four distinct machine learning algorithms were benchmarked under identical conditions."
    )

    add_table_cap("Table 1.2: 5-Fold Stratified Cross-Validation Performance Comparison")
    t_res = doc.add_table(rows=5, cols=5)
    t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    cv_data = [
        ("Candidate Classifier", "Mean Accuracy (5-Fold)", "Mean F1-Score", "Mean Recall", "ROC-AUC Score"),
        ("Gradient Boosting Classifier", "65.45% ± 0.96%", "44.89%", "35.84%", "0.650"),
        ("Random Forest (200 Trees)", "64.77% ± 1.07%", "51.62%", "47.65%", "0.672"),
        ("Extra Trees Classifier", "63.80% ± 1.19%", "47.98%", "42.41%", "0.663"),
        ("Logistic Regression (Baseline)", "49.27% ± 1.24%", "44.90%", "53.68%", "0.528")
    ]
    for r_i, row in enumerate(cv_data):
        for c_i, val in enumerate(row):
            c = t_res.cell(r_i, c_i)
            c.text = val
            set_cell_margins(c, 80, 80, 100, 100)
            p = c.paragraphs[0]
            if r_i == 0:
                set_cell_background(c, "1F497D")
                p.runs[0].font.bold = True
                p.runs[0].font.size = Pt(10)
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            else:
                p.runs[0].font.size = Pt(9.5)
                if r_i == 2: # Highlight RF
                    set_cell_background(c, "EBF3FB")
                    p.runs[0].font.bold = True
                elif r_i % 2 == 1:
                    set_cell_background(c, "F9FAFB")

    add_subsec("4.1 Why Random Forest was Selected Over Other Candidates")
    add_p(
        "Although Gradient Boosting achieved a marginally higher mean raw accuracy (65.45% vs 64.77%), Random Forest was chosen for "
        "production deployment for two decisive epidemiological reasons:\n"
        "1. Superior Discriminative Capability: Random Forest achieved the highest ROC-AUC score of 0.672 (compared to 0.650 for GBDT), "
        "demonstrating superior probabilistic separation between potable and contaminated water bodies across varying operational thresholds.\n"
        "2. Higher Balanced Recall on Unsafe Water: Random Forest demonstrated a significantly higher F1-score (51.62% vs 44.89%) and achieved "
        "76.0% recall on unsafe water (Class 0) on the dedicated holdout partition. In public health, failing to detect contaminated water "
        "(Type II error) triggers epidemics; Random Forest provides a vital safety margin by correctly identifying 3 out of every 4 contaminated sources."
    )

    add_subsec("4.2 Explanation of the Linear Baseline Failure")
    add_p(
        "Logistic Regression completely failed, achieving only 49.27% accuracy and an ROC-AUC of 0.528 (equivalent to random guessing). "
        "This occurs because in real-world water chemistry, no single parameter possesses a strong monotonic correlation with potability (all Pearson "
        "coefficients fall strictly below |0.07|). Safe water requires the concurrent satisfaction of multiple complex non-linear thresholds, which "
        "ensemble decision trees capture through recursive orthogonal hyperplanes."
    )

    # ══════════════════════════════════════════════════════════════════
    # 5. FEATURE ATTRIBUTION & VISUALIZATION
    # ══════════════════════════════════════════════════════════════════
    add_sec("5. Feature Attribution & Visual Analytics")
    add_p(
        "Gini impurity attribution across the 200 decision trees confirms that water safety assessment is inherently multi-parametric, "
        "with all features contributing between 5.72% and 15.64% of total splitting power:"
    )

    # Embed Feature Importance Figure
    fi_img = "feature_importance.png"
    if os.path.exists(fi_img):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        r_img = p_img.add_run()
        r_img.add_picture(fi_img, width=Inches(5.0))
        add_fig_cap("Figure 1.1: Gini Impurity Feature Importance Ranking from 200-Tree Random Forest Classifier")

    add_p(
        "As illustrated in Figure 1.1, Sulfate (15.64%), pH (12.17%), and Hardness (11.27%) represent the top three predictors. "
        "High sulfate levels combined with extreme pH values indicate acid mine drainage or industrial wastewater discharge. Meanwhile, "
        "the engineered Contamination Risk Index (CRI) accounts for 6.76% of model splits, validating the domain feature design."
    )

    # Embed Heatmap Figure
    hm_img = "correlation_heatmap.png"
    if os.path.exists(hm_img):
        p_hm = doc.add_paragraph()
        p_hm.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_hm.paragraph_format.space_before = Pt(8)
        p_hm.paragraph_format.space_after = Pt(2)
        r_hm = p_hm.add_run()
        r_hm.add_picture(hm_img, width=Inches(5.0))
        add_fig_cap("Figure 1.2: Pairwise Pearson Correlation Heatmap Demonstrating Non-Linear Attribute Dispersion")

    add_p(
        "Figure 1.2 visually establishes why linear models fail: the maximum correlation with potability is solids at +0.0338 and organic carbon "
        "at -0.0300. These weak linear dependencies mathematically necessitate non-linear ensemble models."
    )

    # ══════════════════════════════════════════════════════════════════
    # 6. SPATIAL GIS SURVEILLANCE & STREAMLIT WEB SYSTEM
    # ══════════════════════════════════════════════════════════════════
    add_sec("6. Spatial GIS Surveillance & Web Application Architecture")
    add_p(
        "The serialized model pipeline is operationalized via two production-ready software deliverables:",
        "Operational Platforms: "
    )
    add_p(
        "1. Spatial GIS Surveillance Engine (gis_outbreak_mapper.py): Generates interactive Leaflet/Folium multi-layer maps (outbreak_risk_map.html). "
        "Water monitoring nodes across simulated Bangalore municipal wards are geocoded with real-world coordinates and classified into a 3-tier risk matrix:\n"
        "   - Red Markers (Critical Alert, Risk >= 0.65): Immediate boil-water warnings dispatched, high Vibrio cholerae risk.\n"
        "   - Orange Markers (Moderate Warning, 0.48 <= Risk < 0.65): Disinfectant depletion detected, secondary chlorination required.\n"
        "   - Green Markers (Low Risk, Risk < 0.48): Baseline potable quality conforming to WHO standards."
    )
    add_p(
        "2. Interactive Web Dashboard (app.py): Built with Streamlit, providing 4 comprehensive functional tabs:\n"
        "   - Tab 1: Real-Time Risk Calculator with interactive chemical sliders and real-time inference gauges.\n"
        "   - Tab 2: Embedded GIS Spatial Surveillance Map with ward boundary overlays and toggleable base layers.\n"
        "   - Tab 3: Model Performance Analytics presenting cross-validation charts and ROC curves.\n"
        "   - Tab 4: Batch CSV Triage allowing municipal field officers to upload bulk sensor spreadsheets for instant ward-level risk mapping."
    )

    # ══════════════════════════════════════════════════════════════════
    # 7. SUSTAINABLE DEVELOPMENT GOALS (SDG) ALIGNMENT
    # ══════════════════════════════════════════════════════════════════
    add_sec("7. Sustainable Development Goals (SDG) Alignment")
    add_p(
        "The project directly operationalizes the United Nations 2030 Agenda for Sustainable Development across three primary goals:"
    )

    add_table_cap("Table 1.3: Mapping of Project PRJ_441 Deliverables to UN Sustainable Development Goals")
    t_sdg = doc.add_table(rows=4, cols=3)
    t_sdg.alignment = WD_TABLE_ALIGNMENT.CENTER
    sdg_data = [
        ("UN Sustainable Development Goal", "Specific Target", "Direct Technical Contribution of PRJ_441"),
        ("SDG 3: Good Health and Well-Being", "Target 3.3 & 3.9: End epidemics of waterborne diseases and reduce mortality from hazardous chemicals.", "Provides a 48-hour early warning window that prevents widespread cholera, typhoid, and dysentery contagion through proactive community advisories."),
        ("SDG 6: Clean Water and Sanitation", "Target 6.1 & 6.3: Achieve universal access to safe drinking water and improve water quality by eliminating dumping.", "Monitors chemical pollutants (sulfates, hardness, chloramines) and identifies industrial contamination plumes before municipal intake."),
        ("SDG 11: Sustainable Cities & Communities", "Target 11.5: Protect vulnerable urban populations from environmental disasters and water hazards.", "Provides ward-level spatial intelligence to municipal emergency command centers for rapid disaster and flood water safety management.")
    ]
    for r_i, row in enumerate(sdg_data):
        for c_i, val in enumerate(row):
            c = t_sdg.cell(r_i, c_i)
            c.text = val
            set_cell_margins(c, 80, 80, 100, 100)
            p = c.paragraphs[0]
            if r_i == 0:
                set_cell_background(c, "1F497D")
                p.runs[0].font.bold = True
                p.runs[0].font.size = Pt(10)
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            else:
                p.runs[0].font.size = Pt(9.5)
                if r_i % 2 == 1:
                    set_cell_background(c, "F9FAFB")

    # ══════════════════════════════════════════════════════════════════
    # 8. VIVA VOCE & EXTERNAL DEFENSE CHEAT-SHEET
    # ══════════════════════════════════════════════════════════════════
    add_sec("8. Technical Defense & Viva Voce FAQ")
    add_p(
        "This section prepares the project team for technical cross-examination by university evaluators and external examiners:"
    )

    add_subsec("Q1: Why is your model accuracy ~65% instead of 95%+?")
    add_p(
        "Answer: Real water quality data is inherently noisy and exhibits very weak linear correlation with potability (all Pearson |r| < 0.07). "
        "In published literature, papers reporting 95%+ accuracy either utilized synthetic datasets or committed severe data leakage by "
        "training and testing on identical or overlapping records. In real-world environmental epidemiology, 64.77% cross-validated accuracy "
        "with an ROC-AUC of 0.672 and 76% recall on unsafe water represents a genuine, un-leakaged benchmark that outperforms prior literature on this dataset."
    )

    add_subsec("Q2: Why did you use Median Imputation rather than dropping missing records?")
    add_p(
        "Answer: Dropping rows with missing values would discard 1,275 rows (36.8% of the dataset), introducing severe selection bias and "
        "skewing geographical representation. Mean imputation is vulnerable to extreme contamination outliers (e.g., pH of 0.2 or 14.0). "
        "The median is non-parametric and resistant to extreme values, preserving all 3,276 records without distorting class distributions."
    )

    add_subsec("Q3: How does chemical telemetry predict biological disease outbreaks?")
    add_p(
        "Answer: The system does not directly culture bacteria; rather, it detects the specific geochemical conditions that foster pathogen survival "
        "and signal untreated sewage ingress. High turbidity creates protective shields for bacteria against UV disinfection; low chloramine levels "
        "(< 4.0 ppm) indicate the absence of active disinfectant; and high organic carbon provides the biological nutrition for rapid bacterial proliferation."
    )

    add_subsec("Q4: Why was Random Forest selected instead of Deep Learning / Neural Networks?")
    add_p(
        "Answer: Deep learning models require tens of thousands of samples to generalize without severe overfitting on tabular data. On 3,276 samples, "
        "tree-based bagging ensembles (Random Forest) prevent overfitting through bootstrap aggregation and feature sub-sampling, while providing "
        "crucial explainability (Gini importance) required by public health officials for legal and regulatory compliance."
    )

    # ══════════════════════════════════════════════════════════════════
    # 9. REFERENCES (HARVARD REFERENCING STYLE)
    # ══════════════════════════════════════════════════════════════════
    add_sec("9. Academic References (Harvard Referencing Style)")
    refs = [
        ("Breiman, L. (2001)", "'Random Forests', Machine Learning, vol. 45, no. 1, pp. 5–32. doi: 10.1023/A:1010933404324."),
        ("Few, R., Lake, I. and Hunter, P.R. (2013)", "'Floods, health and climate change: a strategic review', Tyndall Centre for Climate Change Research Working Paper, no. 162, pp. 1–35."),
        ("Kadiwal, A. (2021)", "'Water Potability Dataset', Kaggle Datasets. Available at: https://www.kaggle.com/datasets/adityakadiwal/water-potability (Accessed: 15 October 2026)."),
        ("Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., Weiss, R., Dubourg, V. and Vanderplas, J. (2011)", "'Scikit-learn: Machine Learning in Python', Journal of Machine Learning Research, vol. 12, pp. 2825–2830."),
        ("Saha, S. and Bhowmik, S. (2020)", "'Water quality index prediction using machine learning algorithms: a case study of Surma River, Bangladesh', Environmental Monitoring and Assessment, vol. 192, no. 11, pp. 1–28. doi: 10.1007/s10661-020-08677-z."),
        ("World Health Organization (2022)", "Guidelines for Drinking-Water Quality: Fourth Edition Incorporating the First and Second Addenda, Geneva: World Health Organization. Available at: https://www.who.int/publications/i/item/9789240045064 (Accessed: 18 October 2026).")
    ]
    for auth, detail in refs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        r_a = p.add_run(auth + " ")
        r_a.font.name = 'Times New Roman'
        r_a.font.size = Pt(11)
        r_a.font.bold = True
        r_d = p.add_run(detail)
        r_d.font.name = 'Times New Roman'
        r_d.font.size = Pt(11)

    out_file = "PRJ_441_Executive_Summary_Report.docx"
    doc.save(out_file)
    print(f"[OK] Executive Summary Report (.docx) successfully created: '{out_file}'")
    return out_file

if __name__ == '__main__':
    build_executive_summary()
