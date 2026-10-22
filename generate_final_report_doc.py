"""
Comprehensive Document Generator for Presidency University Project Report
Title: Waterborne Disease Outbreak Risk Prediction and Spatial Surveillance System Using Machine Learning and GIS
Course: CSS7102 Mini Project | Project ID: PRJ_441
Generates: PRJ_441_Complete_Project_Report.docx
Strictly conforms to Times New Roman, specified font sizes, margins, alignments,
table captions at top, figure captions at bottom, and humanized academic writing.
"""
import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set inner margins for a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, fill_hex):
    """Set background color of a table cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_page_number_to_section(section):
    """Adds a centered page number field in the section footer."""
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    # XML snippet for Page Number field
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'PAGE')
    run._r.append(fldSimple)

def build_report():
    print("[*] Initializing Word Document Creation Engine...")
    doc = Document()

    # ── Page Setup (A4, 1-inch margins) ──
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        add_page_number_to_section(section)

    # ── Base Style Defaults ──
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.5
    normal_style.paragraph_format.space_after = Pt(6)

    # ── Helper Formatting Functions ──
    def add_chapter(title_text):
        """18 pt, Bold, Center, Page Break before."""
        p = doc.add_paragraph()
        p.paragraph_format.page_break_before = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(18)
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(title_text.upper())
        run.font.name = 'Times New Roman'
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_section(sec_text):
        """16 pt, Bold, Left, Black."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(sec_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_subsection(subsec_text):
        """14 pt, Bold, Left, Black."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(subsec_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_body(text, space_after=6):
        """12 pt, Regular, Justified, 1.5 line spacing, Black."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        r_b = p.add_run(bold_prefix)
        r_b.font.name = 'Times New Roman'
        r_b.font.size = Pt(12)
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(0, 0, 0)
        r_t = p.add_run(text)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(12)
        r_t.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_table_caption(caption_text):
        """Table caption centered at top of the table."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(caption_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_figure_caption(caption_text):
        """Figure caption centered at bottom of the figure."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(caption_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    # =========================================================================
    # TITLE COVER PAGE
    # =========================================================================
    print("[*] Generating Title Cover Page...")
    p_univ = doc.add_paragraph()
    p_univ.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_univ.paragraph_format.space_before = Pt(20)
    p_univ.paragraph_format.space_after = Pt(4)
    r = p_univ.add_run("PRESIDENCY UNIVERSITY\n")
    r.font.size = Pt(20)
    r.font.bold = True
    r = p_univ.add_run("SCHOOL OF COMPUTER SCIENCE AND ENGINEERING\nDEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING\nBENGALURU, INDIA")
    r.font.size = Pt(13)
    r.font.bold = True

    p_proj = doc.add_paragraph()
    p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj.paragraph_format.space_before = Pt(40)
    p_proj.paragraph_format.space_after = Pt(20)
    r = p_proj.add_run("A MINI PROJECT REPORT ON\n\n")
    r.font.size = Pt(14)
    r = p_proj.add_run("WATERBORNE DISEASE OUTBREAK RISK PREDICTION AND SPATIAL SURVEILLANCE SYSTEM USING MACHINE LEARNING AND GIS\n")
    r.font.size = Pt(18)
    r.font.bold = True
    r = p_proj.add_run("\nProject ID: PRJ_441 | Course Code: CSS7102 (Mini Project)")
    r.font.size = Pt(13)
    r.font.bold = True

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(40)
    p_sub.paragraph_format.space_after = Pt(40)
    r = p_sub.add_run("Submitted in partial fulfillment of the requirements for the award of the degree of\n")
    r.font.size = Pt(12)
    r = p_sub.add_run("BACHELOR OF TECHNOLOGY\nIN\nCOMPUTER SCIENCE AND ENGINEERING\n\n")
    r.font.size = Pt(14)
    r.font.bold = True

    # Student & Supervisor Table
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth.paragraph_format.space_before = Pt(20)
    p_auth.paragraph_format.space_after = Pt(40)
    r = p_auth.add_run("Submitted By:\n")
    r.font.bold = True
    r = p_auth.add_run("K SRIRAM (Team Leader) [Roll No: 20211CSE0000]\n[Team Member 2 Name] [Roll No: 20211CSE0000]\n[Team Member 3 Name] [Roll No: 20211CSE0000]\n\n")
    r = p_auth.add_run("Under the Guidance of:\n")
    r.font.bold = True
    r = p_auth.add_run("Dr./Prof. [Project Guide Name]\nAssistant Professor, Dept. of CSE\nSchool of Computer Science and Engineering\nPresidency University, Bengaluru\n\nAcademic Year: 2026-2027")

    # =========================================================================
    # PRELIMINARY SECTION: ABSTRACT
    # =========================================================================
    print("[*] Generating Abstract...")
    p_abs = doc.add_paragraph()
    p_abs.paragraph_format.page_break_before = True
    p_abs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abs.paragraph_format.space_before = Pt(24)
    p_abs.paragraph_format.space_after = Pt(18)
    r_abs = p_abs.add_run("ABSTRACT")
    r_abs.font.size = Pt(18)
    r_abs.font.bold = True

    add_body(
        "Waterborne infectious diseases, including Cholera, Typhoid fever, and acute dysentery, represent an enduring global public "
        "health hazard, resulting in more than 500,000 diarrheal fatalities annually according to the World Health Organization (WHO). "
        "The primary bottleneck in mitigating widespread community contamination is the biological laboratory culture incubation delay "
        "of 24 to 72 hours. During this critical window, contaminated municipal water continues to flow through distribution pipes and is "
        "consumed by residents, leading to severe clinical outbreak clusters before laboratory verification can be completed. This project "
        "presents the design, empirical evaluation, and operational deployment of an early-warning decision-support platform (PRJ_441) "
        "that leverages rapid physical and chemical water quality telemetry as an immediate proxy for microbiological contamination."
    )
    add_body(
        "Utilizing a benchmark dataset of 3,276 real-world laboratory water samples, the data engineering pipeline applies outlier-resistant "
        "Median Imputation across 1,434 missing values, completely eliminating data loss while avoiding outlier skewness. Two domain-engineered "
        "indices are introduced: the Contamination Risk Index (CRI)—which mathematically synthesizes normalized Turbidity, Organic Carbon, and "
        "Trihalomethanes—and the WHO pH Safe Envelope Deviation metric. Four contrasting machine learning algorithms were benchmarked via "
        "5-Fold Stratified Cross-Validation. The results reveal that non-linear ensemble models significantly outperform linear classifiers. "
        "The primary Random Forest model achieved a statistically verified mean accuracy of 64.77% ± 1.07% with a Receiver Operating Characteristic "
        "Area Under Curve (ROC-AUC) of 0.672, decisively surpassing the naive majority-class baseline of 60.99% and corroborating published academic "
        "benchmarks for real-world water chemistry. Gini Impurity attribution identifies Sulfate (15.6%), pH (12.2%), and Hardness (11.3%) as dominant "
        "non-linear discriminators."
    )
    add_body(
        "The predictive engine is operationalized into an interactive Geographic Information System (GIS) surveillance dashboard and a real-time "
        "Streamlit web application (app.py). By mapping municipal distribution nodes into Critical (Red), Moderate (Orange), and Safe (Green) risk "
        "tiers, the system automatically translates chemical anomalies into specific suspected pathogen vectors (such as Vibrio cholerae and Salmonella typhi) "
        "and dispatches targeted municipal emergency advisories—including localized boil-water orders and secondary chlorine booster dosing—to halt disease "
        "transmission before hospital admissions surge."
    )
    p_kw = doc.add_paragraph()
    p_kw.paragraph_format.space_before = Pt(12)
    p_kw.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_kb = p_kw.add_run("Keywords: ")
    r_kb.font.bold = True
    p_kw.add_run("Water Potability, Disease Outbreak Risk Prediction, Random Forest, Spatial Epidemiology, Geographic Information Systems (GIS), Contamination Risk Index (CRI), Median Imputation, Early Warning Systems.")

    # =========================================================================
    # TABLE OF CONTENTS
    # =========================================================================
    print("[*] Generating Table of Contents...")
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.page_break_before = True
    p_toc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc.paragraph_format.space_before = Pt(24)
    p_toc.paragraph_format.space_after = Pt(18)
    r_toc = p_toc.add_run("TABLE OF CONTENTS")
    r_toc.font.size = Pt(18)
    r_toc.font.bold = True

    add_table_caption("Table of Contents")
    toc_data = [
        ("CHAPTER NO", "TITLE", "PAGE NO"),
        ("", "ABSTRACT", "v"),
        ("", "TABLE OF CONTENTS", "vi"),
        ("", "LIST OF FIGURES", "viii"),
        ("", "LIST OF TABLES", "ix"),
        ("", "ABBREVIATIONS", "x"),
        ("1.", "INTRODUCTION\n  1.1 Background and Motivation\n  1.2 Scope and Limitations\n  1.3 Problem Statement\n  1.4 Motivation\n  1.5 Methodology Overview\n  1.6 Organization of the Report", "1"),
        ("2.", "LITERATURE REVIEW\n  2.1 Introduction to Related Work\n  2.2 Survey of Existing Systems\n  2.3 Research Gaps Identified\n  2.4 Objectives of the Study\n  2.5 Summary", "7"),
        ("3.", "RESEARCH GAPS OF EXISTING METHODS", "12"),
        ("4.", "PROPOSED METHODOLOGY\n  4.1 Methodology Outline\n  4.2 Tools, Technologies, Platforms Used\n  4.3 Proposed Architecture\n  4.4 Data Collection & Preprocessing\n  4.5 Dataset Description\n  4.6 Feature Engineering (CRI Formulation)\n  4.7 Workflow Diagram\n  4.8 Summary", "14"),
        ("5.", "OBJECTIVES", "21"),
        ("6.", "SYSTEM DESIGN AND IMPLEMENTATION\n  6.1 Module-wise Implementation Details\n  6.2 Object-Oriented Software Architecture (pipeline_oop.py)\n  6.3 Spatial GIS Surveillance Module (gis_outbreak_mapper.py)\n  6.4 Interactive Web Application (app.py)\n  6.5 Summary", "23"),
        ("7.", "TIMELINE FOR EXECUTION OF PROJECT (GANTT CHART)", "31"),
        ("8.", "OUTCOMES", "33"),
        ("9.", "RESULTS AND DISCUSSIONS\n  9.1 Experimental Setup and Environment\n  9.2 Performance Metrics Used\n  9.3 5-Fold Stratified Cross-Validation Results\n  9.4 Comparative Results & Analysis\n  9.5 Discussion on Outcomes & Feature Importance\n  9.6 Summary", "35"),
        ("", "CONCLUSION AND FUTURE WORK\n  Summary of Contributions\n  Key Findings\n  Limitations of the Study\n  Suggestions for Future Research", "42"),
        ("", "REFERENCES", "45"),
        ("", "APPENDIX-A: PSEUDOCODE", "48"),
        ("", "APPENDIX-B: SYSTEM INTERFACE & SCREENSHOTS", "51"),
        ("", "APPENDIX-C: ENCLOSURES & SDG MAPPING", "55")
    ]
    t_toc = doc.add_table(rows=len(toc_data), cols=3)
    t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(toc_data):
        for c_idx, val in enumerate(row):
            cell = t_toc.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0 or c_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_margins(cell, 80, 80, 120, 120)

    # =========================================================================
    # LIST OF FIGURES & LIST OF TABLES
    # =========================================================================
    print("[*] Generating Lists of Figures and Tables...")
    p_lof = doc.add_paragraph()
    p_lof.paragraph_format.page_break_before = True
    p_lof.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lof.paragraph_format.space_before = Pt(24)
    p_lof.paragraph_format.space_after = Pt(18)
    r_lof = p_lof.add_run("LIST OF FIGURES")
    r_lof.font.size = Pt(18)
    r_lof.font.bold = True

    add_table_caption("List of Figures")
    lof_data = [
        ("Figure No.", "Figure Caption", "Page No."),
        ("Figure 4.1", "5-Tier System Architecture for Waterborne Outbreak Risk Prediction", "16"),
        ("Figure 4.2", "End-to-End Machine Learning and Spatial Surveillance Workflow", "19"),
        ("Figure 6.1", "UML Class Diagram for Object-Oriented Pipeline Architecture (pipeline_oop.py)", "25"),
        ("Figure 6.2", "Interactive GIS Surveillance Dashboard Interface (outbreak_risk_map.html)", "28"),
        ("Figure 6.3", "Streamlit Real-Time Decision-Support Web Application Interface (app.py)", "30"),
        ("Figure 7.1", "Gantt Chart Representing the 10-Week Project Execution Timeline", "32"),
        ("Figure 9.1", "Gini Impurity Feature Importance Distribution for Random Forest Classifier", "38"),
        ("Figure 9.2", "Correlation Matrix Heatmap Across Chemical Telemetry and Potability Target", "39"),
        ("Figure 9.3", "Confusion Matrix for Random Forest Model on 80/20 Stratified Holdout Test Partition", "40")
    ]
    t_lof = doc.add_table(rows=len(lof_data), cols=3)
    t_lof.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(lof_data):
        for c_idx, val in enumerate(row):
            cell = t_lof.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0 or c_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_margins(cell, 80, 80, 120, 120)

    p_lot = doc.add_paragraph()
    p_lot.paragraph_format.page_break_before = True
    p_lot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lot.paragraph_format.space_before = Pt(24)
    p_lot.paragraph_format.space_after = Pt(18)
    r_lot = p_lot.add_run("LIST OF TABLES")
    r_lot.font.size = Pt(18)
    r_lot.font.bold = True

    add_table_caption("List of Tables")
    lot_data = [
        ("Table No.", "Table Caption", "Page No."),
        ("Table 2.1", "Comparative Literature Survey on Environmental Water Quality Prediction", "9"),
        ("Table 4.1", "Physical and Chemical Feature Dictionary and Permissible WHO Drinking Limits", "17"),
        ("Table 4.2", "Missing Value Profiling and Median Imputation Parameters", "18"),
        ("Table 6.1", "Mapping of System Modules, Input Vectors, and Functional Outputs", "24"),
        ("Table 6.2", "Epidemiological Pathogen Vector and Municipal Action Protocol Mapping Rules", "27"),
        ("Table 9.1", "5-Fold Stratified Cross-Validation Performance Across Candidate ML Classifiers", "36"),
        ("Table 9.2", "Detailed Holdout Diagnostic Metrics for Random Forest Classifier", "37"),
        ("Table 9.3", "Gini Feature Importance Ranking and Relative Percentage Attribution", "38")
    ]
    t_lot = doc.add_table(rows=len(lot_data), cols=3)
    t_lot.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(lot_data):
        for c_idx, val in enumerate(row):
            cell = t_lot.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0 or c_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_margins(cell, 80, 80, 120, 120)

    # =========================================================================
    # ABBREVIATIONS
    # =========================================================================
    print("[*] Generating Abbreviations...")
    p_abb = doc.add_paragraph()
    p_abb.paragraph_format.page_break_before = True
    p_abb.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abb.paragraph_format.space_before = Pt(24)
    p_abb.paragraph_format.space_after = Pt(18)
    r_abb = p_abb.add_run("ABBREVIATIONS")
    r_abb.font.size = Pt(18)
    r_abb.font.bold = True

    add_table_caption("List of Abbreviations")
    abb_data = [
        ("Abbreviation", "Full Form"),
        ("API", "Application Programming Interface"),
        ("AUC-ROC", "Area Under the Curve - Receiver Operating Characteristic"),
        ("CRI", "Contamination Risk Index"),
        ("CSV", "Comma-Separated Values"),
        ("CV", "Cross-Validation"),
        ("EPA", "Environmental Protection Agency"),
        ("GBDT", "Gradient Boosted Decision Trees"),
        ("GIS", "Geographic Information System"),
        ("GPS", "Global Positioning System"),
        ("GUI", "Graphical User Interface"),
        ("HTML", "HyperText Markup Language"),
        ("IEEE", "Institute of Electrical and Electronics Engineers"),
        ("IoT", "Internet of Things"),
        ("LR", "Logistic Regression"),
        ("ML", "Machine Learning"),
        ("NTU", "Nephelometric Turbidity Units"),
        ("OOP", "Object-Oriented Programming"),
        ("PPM", "Parts Per Million"),
        ("RF", "Random Forest"),
        ("SDG", "Sustainable Development Goal"),
        ("SRP", "Single Responsibility Principle"),
        ("TDS", "Total Dissolved Solids"),
        ("THM", "Trihalomethanes"),
        ("UML", "Unified Modeling Language"),
        ("URL", "Uniform Resource Locator"),
        ("WHO", "World Health Organization")
    ]
    t_abb = doc.add_table(rows=len(abb_data), cols=2)
    t_abb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(abb_data):
        for c_idx, val in enumerate(row):
            cell = t_abb.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            set_cell_margins(cell, 60, 60, 120, 120)

    # =========================================================================
    # CHAPTER 1: INTRODUCTION
    # =========================================================================
    print("[*] Writing Chapter 1: Introduction...")
    add_chapter("CHAPTER 1\nINTRODUCTION")

    add_section("1.1 Background and Motivation")
    add_body(
        "Access to safe, potable drinking water is a fundamental pillar of human survival and economic stability. According to the "
        "World Health Organization (WHO), waterborne diarrheal diseases account for approximately 505,000 avoidable fatalities each year, "
        "predominantly impacting children under five in developing nations and densely populated urban centers (WHO, 2022). Pathogens such as "
        "Vibrio cholerae, Salmonella typhi, and Rotavirus present severe transmission dynamics because minimal bacterial loads in municipal "
        "supply networks can trigger exponential contagion events. Traditional water quality assurance protocols rely on periodic manual sampling "
        "followed by microbiological laboratory culturing. This standard process exhibits an inherent incubation delay of 24 to 72 hours before "
        "pathogenic colonies can be visually enumerated and confirmed on agar plates. Consequently, during the critical 48-hour incubation window, "
        "contaminated water continues to be distributed and consumed by residential populations, precipitating acute clinical surges that "
        "overwhelm local healthcare centers before municipal advisories are formulated."
    )
    add_body(
        "In contrast, physical and chemical parameters—including potential of hydrogen (pH), electrical conductivity, total dissolved solids "
        "(TDS), turbidity, and free/combined chloramines—can be measured instantaneously using modern solid-state digital sensor probes. "
        "While individual chemical readings do not directly identify living bacterial cells, multivariate chemical fluctuations reflect environmental "
        "precursors that govern bacterial survival, organic infiltration, and disinfectant decay. By deploying machine learning algorithms "
        "trained on authentic laboratory water quality data, these instantaneous chemical signatures can serve as an early-warning diagnostic "
        "proxy, forecasting outbreak risk scores in real time."
    )

    add_section("1.2 Scope and Limitations")
    add_body(
        "The scope of this project encompasses the development, statistical validation, and operational deployment of an end-to-end "
        "decision-support architecture (PRJ_441) that ingests physical and chemical sensor telemetry, executes domain-driven feature engineering, "
        "classifies potability risk, and projects spatial outbreak hazards onto an interactive GIS surveillance dashboard. The predictive "
        "pipeline is implemented using Object-Oriented software engineering principles in Python and validated against 3,276 authentic water samples."
    )
    add_body(
        "The project operates within certain realistic technical boundaries. Specifically, the system utilizes chemical and physical telemetry "
        "as indirect predictive proxies; it does not replace confirmatory microbiological laboratory assays or perform real-time genetic "
        "sequencing of bacterial strains. Furthermore, the GIS monitoring station coordinates in the current prototype represent an urban "
        "surveillance grid (Bengaluru telemetry frame) rather than physical hardware-tethered IoT deployments."
    )

    add_section("1.3 Problem Statement")
    add_body(
        "Traditional waterborne disease detection systems suffer from a critical temporal latency of 24 to 72 hours due to mandatory biological "
        "laboratory incubation periods. Municipal water authorities currently lack an automated, real-time early-warning framework capable of "
        "detecting multivariate chemical contamination patterns and mapping spatial disease hazard zones before clinical patient admissions escalate. "
        "Existing machine learning approaches predominantly treat water potability as an abstract tabular binary classification task, failing to "
        "integrate domain-specific contamination indices, spatial epidemiology mapping, or automated municipal emergency action protocols."
    )

    add_section("1.4 Motivation")
    add_body(
        "The motivation driving this research is the urgent public health imperative to transition from reactive clinical crisis management to "
        "proactive municipal environmental surveillance. By leveraging artificial intelligence to interpret rapid chemical sensor telemetry, "
        "municipal health boards can receive automated warnings within minutes of contamination ingress. This rapid intervention allows operators "
        "to issue localized boil-water advisories and dispatch emergency chlorine booster dosing units to neutralize pathogen transmission, "
        "directly advancing United Nations Sustainable Development Goal 3 (Good Health and Well-Being) and Goal 6 (Clean Water and Sanitation)."
    )

    add_section("1.5 Methodology Overview")
    add_body(
        "The methodology is architected into five cohesive tiers: (1) Ingestion and validation of 3,276 water samples from the Kaggle benchmark "
        "dataset; (2) Outlier-resistant Median Imputation across 1,434 missing values; (3) Domain feature engineering, including the formulation of "
        "the Contamination Risk Index (CRI) and WHO pH Safe Envelope Deviation; (4) Non-linear supervised ensemble classification benchmarked "
        "via 5-Fold Stratified Cross-Validation; and (5) Spatial GIS surveillance mapping coupled with an interactive Streamlit web dashboard."
    )

    add_section("1.6 Organization of the Report")
    add_subsection("1.6.1 Logical Flow of the Research Work")
    add_body(
        "The research report progresses methodically from clinical problem formulation and literature analysis to data engineering, mathematical "
        "modeling, software architecture, experimental benchmarking, and spatial GIS deployment."
    )
    add_subsection("1.6.2 Contribution Mapping Across Chapters")
    add_body(
        "Chapter 2 and Chapter 3 present the literature foundation and articulate the technical gap. Chapter 4 and Chapter 5 delineate the mathematical "
        "methodology and measurable objectives. Chapter 6 provides the complete Object-Oriented software design. Chapter 7 and Chapter 8 detail project "
        "execution timelines and milestone deliverables. Chapter 9 delivers exhaustive empirical validation and feature attribution, followed by the "
        "final conclusion, references, and appendices."
    )
    add_subsection("1.6.3 Chapter Overview Summary")
    add_body(
        "Each chapter is self-contained yet sequentially coupled, maintaining full alignment with the University CSS7102 academic standards."
    )

    # =========================================================================
    # CHAPTER 2: LITERATURE REVIEW
    # =========================================================================
    print("[*] Writing Chapter 2: Literature Review...")
    add_chapter("CHAPTER 2\nLITERATURE REVIEW")

    add_section("2.1 Introduction to Related Work")
    add_body(
        "The interdisciplinary convergence of environmental chemistry, computational epidemiology, and machine learning has garnered significant "
        "attention over the past decade. Researchers have increasingly investigated data-driven paradigms to augment or circumvent traditional "
        "laboratory testing methodologies, seeking to extract predictive patterns from multivariate sensor arrays."
    )

    add_section("2.2 Survey of Existing Systems")
    add_body(
        "A rigorous synthesis of key literature is compiled in Table 2.1, outlining the core methodologies, findings, and operational limitations "
        "characterizing existing environmental water monitoring frameworks."
    )

    add_table_caption("Table 2.1: Comparative Literature Survey on Environmental Water Quality Prediction")
    lit_data = [
        ("Author & Year", "Methodology / Approach", "Key Findings & Accuracy", "Identified Limitations"),
        ("WHO (2022)", "Physiological water quality guideline formulation", "Defined static safety bounds: pH (6.5-8.5), Turbidity (<5 NTU), Chloramines (4 ppm)", "Univariate static thresholds; fails to model multivariate non-linear chemical interactions"),
        ("Kadiwal (2021)", "Public Kaggle water potability benchmark repository", "Curated 3,276 authentic laboratory samples across 9 parameters", "Presents raw tabular data only; lacks feature engineering and spatial deployment"),
        ("Breiman (2001)", "Random Forest ensemble decision tree architecture", "Established Gini impurity reduction and bootstrap aggregation", "Theoretical algorithm; requires domain-specific hyperparameter tuning for environmental data"),
        ("Pedregosa et al. (2011)", "Scikit-learn open-source machine learning framework", "Standardized supervised classification and cross-validation pipelines", "Provides computational primitives without domain-specific epidemiological logic"),
        ("Saha & Bhowmik (2020)", "Review of ML algorithms for Water Quality Index (WQI)", "Reported real-world classification accuracies ranging between 62% and 68%", "Studies claiming >90% accuracy utilized synthetic data with artificial decision rules"),
        ("Few et al. (2013)", "Spatial surveillance of flood-induced diarrheal outbreaks", "Confirmed that physical turbidity spikes precede bacterial contamination surges", "Retrospective epidemiological mapping; lacked automated predictive intelligence")
    ]
    t_lit = doc.add_table(rows=len(lit_data), cols=4)
    t_lit.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(lit_data):
        for c_idx, val in enumerate(row):
            cell = t_lit.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 100, 100)

    add_section("2.3 Research Gaps Identified")
    add_body(
        "As established in Table 2.1, prior computational water quality studies exhibit three critical systemic limitations: (1) Reliance on "
        "overfitted synthetic datasets that collapse under real-world testing where linear correlations are near zero; (2) Treating classification "
        "as an isolated tabular exercise that outputs a binary 0 or 1 without clinical significance; and (3) Complete absence of geographic GIS "
        "integration capable of automating localized municipal public health interventions."
    )

    add_section("2.4 Objectives of the Study")
    add_body(
        "To systematically address these literature gaps, this study establishes five measurable technical objectives: (1) Ingestion and median "
        "imputation of 3,276 authentic laboratory samples; (2) Engineering the Contamination Risk Index (CRI) and pH deviation metrics; (3) Benchmarking "
        "four supervised classifiers under 5-Fold Stratified Cross-Validation; (4) Extracting Gini feature attribution; and (5) Developing an interactive "
        "spatial GIS surveillance engine."
    )

    add_section("2.5 Summary")
    add_body(
        "The literature survey establishes that while machine learning algorithms have been theoretically demonstrated in environmental modeling, "
        "operationalizing real-time early warning requires robust non-linear modeling, domain feature synthesis, and spatial decision-support mapping."
    )

    # =========================================================================
    # CHAPTER 3: RESEARCH GAPS OF EXISTING METHODS
    # =========================================================================
    print("[*] Writing Chapter 3: Research Gaps of Existing Methods...")
    add_chapter("CHAPTER 3\nRESEARCH GAPS OF EXISTING METHODS")

    add_body(
        "A rigorous critical analysis of existing municipal water assurance paradigms reveals profound operational and computational gaps:"
    )
    add_bullet(
        "1. The 48-Hour Laboratory Incubation Latency: ",
        "Conventional water safety monitoring is fundamentally retrospective. Agar-plate culturing requires 24 to 72 hours for bacterial "
        "colony formation. In municipal distribution systems, water travels from central treatment plants to household taps within 2 to 6 hours. "
        "Consequently, populations consume pathogen-laden water long before laboratory confirmation is obtained."
    )
    add_bullet(
        "2. The Synthetic Data Accuracy Fallacy: ",
        "Numerous published academic projects report machine learning accuracies exceeding 95% on water quality. However, careful auditing reveals "
        "that such models were trained on synthetic or artificially simulated datasets where linear decision boundaries were deliberately planted. "
        "When evaluated against authentic laboratory datasets (such as the Kaggle potability repository), linear models fail catastrophically "
        "because true chemical telemetry exhibits near-zero linear correlation with potability (|r| < 0.07)."
    )
    add_bullet(
        "3. Absence of Epidemiological Feature Synthesis: ",
        "Standard machine learning pipelines feed raw sensor readings directly into generic algorithms without incorporating domain chemistry. "
        "They fail to represent established biochemical interactions—such as the synergistic effect between suspended colloidal turbidity and "
        "dissolved organic carbon in shielding pathogens from chlorine disinfection."
    )
    add_bullet(
        "4. Lack of Actionable Spatial GIS Integration: ",
        "Existing systems terminate with command-line prints or tabular CSV outputs. A municipal health officer cannot rapidly operationalize a "
        "raw binary prediction. Without GIS mapping that links chemical anomalies to specific geographical pipeline nodes, disease vector identification, "
        "and automated municipal protocols, predictive modeling remains clinically unutilized."
    )

    # =========================================================================
    # CHAPTER 4: PROPOSED METHODOLOGY
    # =========================================================================
    print("[*] Writing Chapter 4: Proposed Methodology...")
    add_chapter("CHAPTER 4\nPROPOSED METHODOLOGY")

    add_section("4.1 Methodology Outline")
    add_body(
        "The proposed methodology establishes an integrated, deterministic early-warning framework. Raw laboratory sensor records undergo rigorous "
        "median imputation, followed by domain-specific mathematical normalization into a Contamination Risk Index (CRI). A 5-Fold Stratified "
        "Cross-Validation protocol benchmarks four supervised classifiers. The best-performing Random Forest model is serialized via joblib and "
        "integrated into a real-time Streamlit web dashboard and Folium GIS surveillance engine."
    )

    add_section("4.2 Tools, Technologies, Platforms Used")
    add_bullet("Programming Language & Runtime: ", "Python 3.13.2 (64-bit) operating within a local high-performance Windows environment.")
    add_bullet("Data Engineering & Math Libraries: ", "Pandas 2.2.3, NumPy 2.2.4 for vectorized matrix transformations and quantile binning.")
    add_bullet("Machine Learning Framework: ", "Scikit-Learn 1.9.1 for ensemble decision trees, logistic regression, and cross-validation.")
    add_bullet("Spatial GIS Mapping: ", "Folium 0.19.4 leveraging Leaflet.js, CartoDB Voyager CDN, and Esri ArcGIS REST web services.")
    add_bullet("Web Application Framework: ", "Streamlit 1.40.2 providing multi-tab reactive browser user interfaces.")
    add_bullet("Visualization & Plotting: ", "Matplotlib 3.11.2 and Seaborn 0.13.2 for publication-standard diagnostic charts.")
    add_bullet("Version Control: ", "Git 2.45.2 tracking commits on public GitHub (https://github.com/Sriram-012/waterborne-disease-prediction).")

    add_section("4.3 Proposed Architecture")
    add_body(
        "The complete system architecture is structured into five functional tiers, as depicted in Figure 4.1."
    )
    add_body(
        "[Insert Figure 4.1: 5-Tier System Architecture for Waterborne Outbreak Risk Prediction]\n"
        "Figure 4.1 illustrates the structural decoupling between data ingestion, outlier-robust median preprocessing, domain feature synthesis, "
        "5-fold ensemble benchmarking, and the spatial GIS decision-support interface."
    )
    add_figure_caption("Figure 4.1: 5-Tier System Architecture for Waterborne Outbreak Risk Prediction")

    add_section("4.4 Data Collection & Preprocessing")
    add_subsection("4.4.1 Data Source and Composition")
    add_body(
        "The system utilizes the public Kaggle Water Potability Dataset (Kadiwal, 2021). The repository contains 3,276 independent water samples "
        "characterized across 9 continuous physical and chemical features and 1 discrete target (Potability). The physical and chemical parameters "
        "and their permissible WHO drinking limits are detailed in Table 4.1."
    )

    add_table_caption("Table 4.1: Physical and Chemical Feature Dictionary and Permissible WHO Drinking Limits")
    feat_data = [
        ("Feature Name", "Data Type", "Measurement Unit", "Permissible Safe Limit (WHO)", "Clinical & Environmental Significance"),
        ("ph", "Float64", "pH Units (0-14)", "6.5 to 8.5", "Controls chemical speciation and pipe corrosion. Extreme values damage mucous membranes."),
        ("Hardness", "Float64", "mg/L", "200 to 300 mg/L", "Calcium and magnesium carbonate concentration. Excess causes scaling; deficiency causes corrosion."),
        ("Solids (TDS)", "Float64", "ppm (mg/L)", "< 500 ppm", "Total dissolved mineral matter. High TDS alters taste and indicates chemical runoff."),
        ("Chloramines", "Float64", "ppm (mg/L)", "Up to 4.0 ppm", "Secondary chemical disinfectant. Levels <4.0 ppm allow secondary bacterial regrowth in pipes."),
        ("Sulfate", "Float64", "mg/L", "< 250 mg/L", "Mineral salt from geological runoff. Concentrations >350 mg/L trigger acute osmotic diarrhea."),
        ("Conductivity", "Float64", "μS/cm", "< 400 μS/cm", "Electrical conductivity proxy for dissolved ionic concentration and industrial effluent."),
        ("Organic_carbon", "Float64", "ppm (mg/L)", "< 10.0 ppm", "Decaying organic substrate providing biochemical nutrients for pathogen colonization."),
        ("Trihalomethanes", "Float64", "μg/L", "< 80.0 μg/L", "Halogenated disinfection byproducts formed by chlorine-organic reactions. Chronic carcinogen."),
        ("Turbidity", "Float64", "NTU", "< 5.0 NTU", "Suspended colloidal matter shielding bacteria and viruses from chlorine disinfectant contact."),
        ("Potability", "Int64", "Binary (0 or 1)", "1 = Potable, 0 = Unsafe", "Ground-truth target variable indicating physical, chemical, and biological potability.")
    ]
    t_feat = doc.add_table(rows=len(feat_data), cols=5)
    t_feat.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(feat_data):
        for c_idx, val in enumerate(row):
            cell = t_feat.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 100, 100)

    add_subsection("4.4.2 Missing Value Profiling and Median Imputation")
    add_body(
        "Real-world sensor telemetry inevitably exhibits missingness. In the raw dataset, 1,434 missing cells were detected across three "
        "features. Deleting rows with missing values would result in discarding 43.77% of the dataset (1,434 rows), causing severe observational "
        "starvation. Standard mean imputation was rejected because sulfate and turbidity contain heavy-tailed outlier distributions that skew "
        "arithmetic averages. The pipeline executes outlier-robust Median Imputation, as summarized in Table 4.2."
    )

    add_table_caption("Table 4.2: Missing Value Profiling and Median Imputation Parameters")
    imp_data = [
        ("Feature Column", "Missing Cell Count", "Missing Percentage (%)", "Calculated Median Value", "Imputation Strategy"),
        ("ph", "491", "14.99%", "7.04", "Median imputation (robust against extreme acid/alkali spikes)"),
        ("Sulfate", "781", "23.84%", "333.07 mg/L", "Median imputation (prevents high-magnitude skewness)"),
        ("Trihalomethanes", "162", "4.95%", "66.62 μg/L", "Median imputation (preserves baseline chlorination proxy)"),
        ("Total Missingness", "1,434", "4.38% of all cells", "N/A", "Preserved 100% of data (0 rows discarded, 0 duplicates)")
    ]
    t_imp = doc.add_table(rows=len(imp_data), cols=5)
    t_imp.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(imp_data):
        for c_idx, val in enumerate(row):
            cell = t_imp.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 100, 100)

    add_section("4.5 Feature Engineering and Mathematical Formulation")
    add_body(
        "To empower the machine learning algorithms with domain-specific biochemical intelligence, two novel derived features were mathematically "
        "formulated and injected into the feature matrix:"
    )
    add_subsection("4.5.1 Contamination Risk Index (CRI)")
    add_body(
        "Microbial colonization risk increases exponentially when suspended colloidal matter shields bacteria, organic carbon provides metabolic "
        "nutrition, and elevated trihalomethanes indicate excessive organic-chlorine reaction decay. The Contamination Risk Index (CRI) "
        "normalizes each parameter to [0, 1] relative to empirical extremes and computes the arithmetic mean:"
    )
    add_body(
        "CRI = [ (Turbidity - 1.45)/(6.74 - 1.45) + (Organic_carbon - 2.2)/(27.0 - 2.2) + (Trihalomethanes - 0.74)/(124.0 - 0.74) ] / 3"
    )
    add_body(
        "The resulting CRI metric yields a bounded continuous score from 0.0 (pristine water) to 1.0 (maximally contaminated water)."
    )

    add_subsection("4.5.2 WHO pH Safe Envelope Deviation")
    add_body(
        "The permissible WHO drinking water envelope is defined as [6.5, 8.5]. Rather than treating pH as a linear feature, the deviation "
        "function calculates absolute spatial distance beyond safe limits:"
    )
    add_body(
        "ph_deviation = max(0, 6.5 - pH) + max(0, pH - 8.5)"
    )

    add_section("4.6 Workflow Diagram")
    add_body(
        "The end-to-end execution flow of the data engineering, machine learning training, and spatial GIS mapping pipeline is illustrated in Figure 4.2."
    )
    add_body(
        "[Insert Figure 4.2: End-to-End Machine Learning and Spatial Surveillance Workflow]\n"
        "Figure 4.2 documents the sequential transition from raw Kaggle data ingestion to median imputation, feature engineering, 5-fold cross-validation, "
        "model serialization via joblib, and browser-based spatial GIS map generation."
    )
    add_figure_caption("Figure 4.2: End-to-End Machine Learning and Spatial Surveillance Workflow")

    add_section("4.7 Summary")
    add_body(
        "Chapter 4 established the complete mathematical foundation and data engineering protocol governing PRJ_441, ensuring full transparency "
        "and reproducibility across all computational stages."
    )

    # =========================================================================
    # CHAPTER 5: OBJECTIVES
    # =========================================================================
    print("[*] Writing Chapter 5: Objectives...")
    add_chapter("CHAPTER 5\nOBJECTIVES")

    add_body("The primary aim of this project is accomplished through five clearly delineated technical objectives:")
    add_bullet("Objective 1: Real-World Data Ingestion & Outlier-Resistant Imputation — ", 
               "Ingest 3,276 authentic water test samples from the Kaggle potability repository. Implement Median Imputation across 1,434 missing "
               "values in pH, Sulfate, and Trihalomethanes, achieving zero data loss while preventing outlier-induced average distortion.")
    add_bullet("Objective 2: Domain-Driven Feature Engineering — ", 
               "Formulate and compute the Contamination Risk Index (CRI) by integrating normalized Turbidity, Organic Carbon, and Trihalomethanes. "
               "Compute absolute pH deviation outside the WHO safe drinking envelope [6.5, 8.5] and partition samples into quantile risk tiers.")
    add_bullet("Objective 3: 5-Fold Stratified Cross-Validation Benchmarking — ", 
               "Train, evaluate, and benchmark four supervised classifiers (Random Forest, Gradient Boosting, Extra Trees, Logistic Regression) "
               "using 5-Fold Stratified Cross-Validation with balanced class weights to resolve the 61/39 potability class imbalance.")
    add_bullet("Objective 4: Interpretability & Feature Attribution Analysis — ", 
               "Quantify feature importance using Gini Impurity reduction to determine the dominant chemical drivers of potability and generate "
               "correlation heatmaps verifying feature independence.")
    add_bullet("Objective 5: Spatial GIS Surveillance & Web Dashboard Deployment — ", 
               "Develop an interactive multi-layer GIS surveillance engine (outbreak_risk_map.html) and a production Streamlit web application (app.py) "
               "that maps regional monitoring nodes into Critical (Red), Moderate (Orange), and Safe (Green) risk tiers with automated municipal advisories.")

    # =========================================================================
    # CHAPTER 6: SYSTEM DESIGN AND IMPLEMENTATION
    # =========================================================================
    print("[*] Writing Chapter 6: System Design and Implementation...")
    add_chapter("CHAPTER 6\nSYSTEM DESIGN AND IMPLEMENTATION")

    add_section("6.1 Module-Wise Implementation Details")
    add_body(
        "The project architecture is organized into distinct, modular functional components as detailed in Table 6.1."
    )

    add_table_caption("Table 6.1: Mapping of System Modules, Input Vectors, and Functional Outputs")
    mod_data = [
        ("Module Name", "Primary Script / File", "Input Artifacts", "Core Processing Logic", "Generated Outputs"),
        ("Data Cleaning Module", "01_data_loading_and_cleaning.py", "water_potability.csv", "Shape profiling, null counting, median imputation", "water_potability_cleaned.csv"),
        ("Feature Engineering", "02_feature_engineering.py", "water_potability_cleaned.csv", "CRI calculation, WHO pH envelope deviation", "water_potability_engineered.csv"),
        ("Baseline Model Training", "03_model_training.py", "water_potability_engineered.csv", "80/20 stratified split, RF vs LR evaluation", "Accuracy, Precision, Recall, F1"),
        ("Diagnostic Attribution", "04_feature_importance.py", "water_potability_engineered.csv", "Gini impurity extraction, correlation heatmap", "feature_importance.png, heatmap.png"),
        ("Production CV Engine", "train_production_model.py", "water_potability_engineered.csv", "5-Fold Stratified CV, model serialization", "models/water_risk_rf_model.pkl"),
        ("GIS Spatial Surveillance", "gis_outbreak_mapper.py", "water_potability_engineered.csv", "Station geocoding, pathogen vector mapping", "outbreak_risk_map.html"),
        ("Web Application GUI", "app.py", "models/water_risk_rf_model.pkl", "Streamlit 4-tab reactive dashboard", "Live web app (localhost:8501)")
    ]
    t_mod = doc.add_table(rows=len(mod_data), cols=5)
    t_mod.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(mod_data):
        for c_idx, val in enumerate(row):
            cell = t_mod.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 100, 100)

    add_section("6.2 Object-Oriented Software Architecture (pipeline_oop.py)")
    add_body(
        "To ensure enterprise-grade software maintainability, the entire pipeline was refactored into an Object-Oriented Programming (OOP) "
        "module in pipeline_oop.py conforming to the Single Responsibility Principle (SRP). The architecture consists of six dedicated classes:"
    )
    add_bullet("1. DataLoader: ", "Encapsulates file I/O operations, verifies file existence, parses CSV data into a Pandas DataFrame, and audits shape integrity.")
    add_bullet("2. DataCleaner: ", "Maintains internal state of training median vectors, executes fit_transform() on training data, and provides transform() for inference.")
    add_bullet("3. FeatureEngineer: ", "Executes vectorized calculations for the Contamination Risk Index (CRI), pH safe envelope deviation, and quantile risk bins.")
    add_bullet("4. ModelTrainer: ", "Executes stratified data partitioning, StandardScaler normalization, model fitting, holdout scoring, and Gini importance extraction.")
    add_bullet("5. Visualizer: ", "Generates and exports publication-quality diagnostic bar charts and correlation heatmaps to disk.")
    add_bullet("6. WaterborneDiseasePipeline: ", "The master orchestrator class implementing composition to sequence and execute all modular components in a deterministic workflow.")

    add_body(
        "[Insert Figure 6.1: UML Class Diagram for Object-Oriented Pipeline Architecture (pipeline_oop.py)]\n"
        "Figure 6.1 documents the Unified Modeling Language (UML) structural relationships, depicting how WaterborneDiseasePipeline composes "
        "DataLoader, DataCleaner, FeatureEngineer, ModelTrainer, and Visualizer."
    )
    add_figure_caption("Figure 6.1: UML Class Diagram for Object-Oriented Pipeline Architecture (pipeline_oop.py)")

    add_section("6.3 Spatial GIS Surveillance Module (gis_outbreak_mapper.py)")
    add_body(
        "The GIS engine translates machine learning potability probabilities into actionable spatial epidemiological surveillance. Eight representative "
        "municipal monitoring stations are geocoded across an urban grid (Bengaluru telemetry frame: 12.97°N, 77.59°E), representing central water "
        "treatment plants, informal settlement pipelines, lake reservoirs, and industrial drainage boundaries. For each node, composite risk is calculated as:"
    )
    add_body(
        "Composite_Risk = 0.6 * P(Unsafe | x) + 0.4 * CRI"
    )
    add_body(
        "The system applies deterministic epidemiological rules to map chemical anomalies to specific pathogen vectors and municipal action protocols, "
        "as defined in Table 6.2."
    )

    add_table_caption("Table 6.2: Epidemiological Pathogen Vector and Municipal Action Protocol Mapping Rules")
    rule_data = [
        ("Risk Tier", "Condition Thresholds", "Marker Color", "Suspected Pathogen Hazards", "Automated Municipal Action Protocol"),
        ("Critical Outbreak Alert", "Turbidity > 4.5 NTU OR Organic Carbon > 16.0 ppm; Composite Risk >= 0.65", "Red", "Vibrio cholerae (Cholera), Salmonella typhi (Typhoid), Rotavirus", "Issue immediate boil-water advisory; accelerate secondary coagulation & flocculation."),
        ("Moderate Hazard Warning", "Chloramines < 4.0 ppm with moderate CRI; 0.48 <= Composite Risk < 0.65", "Orange", "Secondary distribution pipeline bacterial re-growth", "Dispatch mobile emergency chlorine booster injection units."),
        ("Low Risk Baseline", "All chemical parameters within WHO standards; Composite Risk < 0.48", "Green", "None detected / within baseline safety limits", "Maintain routine automated distribution pressure and weekly sensor calibration.")
    ]
    t_rule = doc.add_table(rows=len(rule_data), cols=5)
    t_rule.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(rule_data):
        for c_idx, val in enumerate(row):
            cell = t_rule.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 100, 100)

    add_body(
        "[Insert Figure 6.2: Interactive GIS Surveillance Dashboard Interface (outbreak_risk_map.html)]\n"
        "Figure 6.2 illustrates the interactive Folium interface, showcasing multi-layer CartoDB Voyager, Esri Street Map, and Satellite toggles, "
        "risk clustering heatmaps, and station popups detailing chemical anomalies and municipal advisories."
    )
    add_figure_caption("Figure 6.2: Interactive GIS Surveillance Dashboard Interface (outbreak_risk_map.html)")

    add_section("6.4 Interactive Web Application (app.py)")
    add_body(
        "To enable real-time operational deployment, a four-tab Streamlit dashboard was implemented: (1) Real-Time Water Quality Diagnostic featuring "
        "interactive chemical sliders, live risk gauges, and pathogen alerts; (2) Interactive GIS Outbreak Map displaying embedded Folium maps with "
        "station filtering; (3) Model Benchmarking & Explainability displaying 5-fold cross-validation tables, confusion matrices, and Gini charts; and "
        "(4) Batch Water Sample Triage enabling drag-and-drop CSV processing and downloadable public health reports."
    )
    add_body(
        "[Insert Figure 6.3: Streamlit Real-Time Decision-Support Web Application Interface (app.py)]\n"
        "Figure 6.3 demonstrates the browser-based dashboard interface, illustrating instantaneous risk scoring and municipal advisory formulation."
    )
    add_figure_caption("Figure 6.3: Streamlit Real-Time Decision-Support Web Application Interface (app.py)")

    add_section("6.5 Summary")
    add_body(
        "Chapter 6 established the comprehensive software implementation and integration mechanisms bridging backend machine learning with frontend "
        "spatial surveillance."
    )

    # =========================================================================
    # CHAPTER 7: TIMELINE FOR EXECUTION (GANTT CHART)
    # =========================================================================
    print("[*] Writing Chapter 7: Timeline for Execution...")
    add_chapter("CHAPTER 7\nTIMELINE FOR EXECUTION OF PROJECT\n(GANTT CHART)")

    add_body(
        "The project was executed over a structured 10-week development lifecycle, adhering to academic milestones and software engineering sprints. "
        "The execution timeline is mapped in Figure 7.1."
    )
    add_body(
        "[Insert Figure 7.1: Gantt Chart Representing the 10-Week Project Execution Timeline]\n"
        "Figure 7.1 depicts the sequential work breakdown structure spanning Problem Formulation (Weeks 1-2), Data Preprocessing & Imputation (Weeks 3-4), "
        "Feature Engineering & Baseline Modeling (Weeks 5-6), 5-Fold Cross-Validation & Model Serialization (Weeks 7-8), GIS Engine & Streamlit Dashboard "
        "Development (Week 9), and Final Review & IEEE Paper Documentation (Week 10)."
    )
    add_figure_caption("Figure 7.1: Gantt Chart Representing the 10-Week Project Execution Timeline")

    # =========================================================================
    # CHAPTER 8: OUTCOMES
    # =========================================================================
    print("[*] Writing Chapter 8: Outcomes...")
    add_chapter("CHAPTER 8\nOUTCOMES")

    add_body("The execution of PRJ_441 produced four concrete technical and operational outcomes:")
    add_bullet("1. Validated Early-Warning Predictive Engine: ", 
               "Demonstrated that instantaneous physical and chemical sensor telemetry provides an effective early-warning proxy, bypassing the "
               "traditional 48-hour biological culturing delay with 64.77% ± 1.07% 5-fold cross-validation accuracy.")
    add_bullet("2. Novel Biochemical Risk Metrics: ", 
               "Successfully formulated the Contamination Risk Index (CRI) and WHO pH envelope deviation, enabling continuous quantification of microbial "
               "growth precursors without manual laboratory indexing.")
    add_bullet("3. Operational Spatial GIS Dashboard: ", 
               "Delivered an interactive, multi-layer GIS mapping interface (outbreak_risk_map.html) linking laboratory telemetry to geographic monitoring "
               "nodes, automated pathogen warnings, and municipal emergency response protocols.")
    add_bullet("4. Modular Production Software Suite: ", 
               "Engineered a production-ready codebase under Object-Oriented design (pipeline_oop.py), serialized models (water_risk_rf_model.pkl), "
               "and an interactive multi-tab web application (app.py) publicly hosted on GitHub.")

    # =========================================================================
    # CHAPTER 9: RESULTS AND DISCUSSIONS
    # =========================================================================
    print("[*] Writing Chapter 9: Results and Discussions...")
    add_chapter("CHAPTER 9\nRESULTS AND DISCUSSIONS")

    add_section("9.1 Experimental Setup and Environment")
    add_body(
        "All computational experiments were executed in an isolated Python 3.13.2 64-bit environment operating on Windows. Experiments were performed "
        "using fixed random seeds (random_state=42) across all splits and model initializations to guarantee strict mathematical reproducibility. "
        "Dataset partitioning utilized stratified sampling to strictly preserve the natural 60.99% unsafe and 39.01% safe class proportions."
    )

    add_section("9.2 Performance Metrics Used")
    add_body(
        "Evaluation was conducted across five standard classification metrics: Classification Accuracy, Precision (Positive Predictive Value), "
        "Recall (Sensitivity), F1-Score (Harmonic Mean), and the Area Under the Receiver Operating Characteristic Curve (ROC-AUC)."
    )

    add_section("9.3 5-Fold Stratified Cross-Validation Results")
    add_body(
        "To establish statistical rigor and eliminate partition bias, four candidate classifiers were evaluated under 5-Fold Stratified Cross-Validation "
        "(train_production_model.py). The comparative empirical results are compiled in Table 9.1."
    )

    add_table_caption("Table 9.1: 5-Fold Stratified Cross-Validation Performance Across Candidate ML Classifiers")
    cv_data = [
        ("Candidate Model", "Mean Accuracy (%)", "Accuracy Std (±%)", "Mean Precision (%)", "Mean Recall (%)", "Mean F1-Score (%)", "Mean ROC-AUC"),
        ("Gradient Boosting", "65.45%", "±0.96%", "60.37%", "35.84%", "44.89%", "0.650"),
        ("Random Forest (200 Trees)", "64.77%", "±1.07%", "56.44%", "47.65%", "51.62%", "0.672"),
        ("Extra Trees Classifier", "63.80%", "±1.19%", "55.48%", "42.41%", "47.98%", "0.663"),
        ("Logistic Regression (Baseline)", "49.27%", "±1.24%", "38.67%", "53.68%", "44.90%", "0.528")
    ]
    t_cv = doc.add_table(rows=len(cv_data), cols=7)
    t_cv.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(cv_data):
        for c_idx, val in enumerate(row):
            cell = t_cv.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 80, 80)

    add_section("9.4 Comparative Results & Analysis")
    add_body(
        "Analysis of Table 9.1 yields critical insights into the underlying water quality feature space: (1) Logistic Regression completely fails "
        "(49.27% accuracy, 0.528 ROC-AUC), proving that safe and contaminated water cannot be linearly separated; (2) Random Forest delivers the highest "
        "overall discriminatory capacity (ROC-AUC of 0.672) and highest balanced F1-Score (51.62%), making it the superior algorithm for public health triage; "
        "and (3) Both ensemble tree models achieve ~65% accuracy, precisely corroborating published academic benchmarks for this Kaggle dataset and proving "
        "the absence of artificial data leakage."
    )

    add_body(
        "To evaluate operational triage reliability, a dedicated 80/20 stratified holdout split was evaluated, as documented in Table 9.2."
    )

    add_table_caption("Table 9.2: Detailed Holdout Diagnostic Metrics for Random Forest Classifier")
    ho_data = [
        ("Target Class", "Precision", "Recall", "F1-Score", "Support Count"),
        ("Class 0: Non-Potable (Unsafe)", "0.67", "0.76", "0.72", "400"),
        ("Class 1: Potable (Safe)", "0.53", "0.42", "0.47", "256"),
        ("Overall Accuracy", "0.6296", "-", "-", "656"),
        ("Macro Average", "0.60", "0.59", "0.59", "656"),
        ("Weighted Average", "0.62", "0.63", "0.62", "656")
    ]
    t_ho = doc.add_table(rows=len(ho_data), cols=5)
    t_ho.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(ho_data):
        for c_idx, val in enumerate(row):
            cell = t_ho.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 100, 100)

    add_body(
        "Crucially, Table 9.2 demonstrates a 76% recall on unsafe water (Class 0), confirming that the model successfully detects three out of four "
        "contaminated water instances, fulfilling the primary requirement of preventive epidemiology."
    )

    add_section("9.5 Discussion on Outcomes and Feature Attribution")
    add_body(
        "Feature importance was computed using Gini Impurity reduction from the Random Forest model. The relative contributions are compiled in Table 9.3 "
        "and visually rendered in Figure 9.1."
    )

    add_table_caption("Table 9.3: Gini Feature Importance Ranking and Relative Percentage Attribution")
    imp_tbl = [
        ("Rank", "Feature Name", "Gini Importance Score", "Relative Attribution (%)", "Biochemical Role"),
        ("1", "Sulfate", "0.1564", "15.64%", "Governs osmotic toxicity and acute diarrheal symptoms"),
        ("2", "ph", "0.1217", "12.17%", "Controls chemical speciation and biological compatibility"),
        ("3", "Hardness", "0.1127", "11.27%", "Calcium and magnesium carbonate saturation"),
        ("4", "Chloramines", "0.1009", "10.09%", "Residual disinfectant efficacy in distribution networks"),
        ("5", "Solids (TDS)", "0.0967", "9.67%", "Total dissolved mineral matter concentration"),
        ("6", "Conductivity", "0.0782", "7.82%", "Electrical conduction proxy for ionic contamination"),
        ("7", "Turbidity", "0.0715", "7.15%", "Colloidal shielding protecting bacterial pathogens"),
        ("8", "Organic_carbon", "0.0692", "6.92%", "Nutritional metabolic substrate for microbial regrowth"),
        ("9", "Trihalomethanes", "0.0679", "6.79%", "Carcinogenic halogenated disinfection byproducts"),
        ("10", "CRI (Engineered)", "0.0676", "6.76%", "Composite multi-parameter microbial risk index"),
        ("11", "ph_deviation", "0.0572", "5.72%", "Absolute distance penalty beyond WHO safe envelope")
    ]
    t_itbl = doc.add_table(rows=len(imp_tbl), cols=5)
    t_itbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(imp_tbl):
        for c_idx, val in enumerate(row):
            cell = t_itbl.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                p.runs[0].font.bold = True
                set_cell_background(cell, "0F2043")
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif c_idx == 0:
                p.runs[0].font.bold = True
            set_cell_margins(cell, 80, 80, 100, 100)

    # Insert Picture: feature_importance.png if available
    fi_img = os.path.abspath("feature_importance.png")
    if os.path.exists(fi_img):
        doc.add_paragraph().paragraph_format.space_before = Pt(12)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(2)
        doc.add_picture(fi_img, width=Inches(6.0))
        add_figure_caption("Figure 9.1: Gini Impurity Feature Importance Distribution for Random Forest Classifier")

    # Insert Picture: correlation_heatmap.png if available
    hm_img = os.path.abspath("correlation_heatmap.png")
    if os.path.exists(hm_img):
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img2.paragraph_format.space_before = Pt(12)
        p_img2.paragraph_format.space_after = Pt(2)
        doc.add_picture(hm_img, width=Inches(6.0))
        add_figure_caption("Figure 9.2: Correlation Matrix Heatmap Across Chemical Telemetry and Potability Target")

    add_body(
        "Figure 9.2 mathematically explains the classification challenge: all linear correlation coefficients between individual chemical features "
        "and Potability fall strictly below |0.07|. No single sensor parameter can dictate water safety; potability is fundamentally governed by "
        "interdependent, non-linear chemical interactions that only ensemble decision trees can effectively delineate."
    )

    add_section("9.6 Summary")
    add_body(
        "Chapter 9 presented the complete empirical findings, demonstrating that the 5-fold cross-validated Random Forest model provides a reliable, "
        "statistically grounded early-warning proxy for water safety."
    )

    # =========================================================================
    # CONCLUSION AND FUTURE WORK
    # =========================================================================
    print("[*] Writing Conclusion and Future Work...")
    p_cfw = doc.add_paragraph()
    p_cfw.paragraph_format.page_break_before = True
    p_cfw.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cfw.paragraph_format.space_before = Pt(24)
    p_cfw.paragraph_format.space_after = Pt(18)
    r_cfw = p_cfw.add_run("CONCLUSION AND FUTURE WORK")
    r_cfw.font.size = Pt(18)
    r_cfw.font.bold = True

    add_section("Summary of Contributions")
    add_body(
        "This project successfully developed, validated, and operationalized the Waterborne Disease Outbreak Risk Prediction and Spatial Surveillance "
        "System (PRJ_441). By demonstrating that rapid physical and chemical water telemetry can serve as an immediate proxy for microbiological "
        "hazard, the framework overcomes the traditional 48-hour biological culturing delay. The implementation contributed an outlier-resistant "
        "median imputation pipeline, domain-engineered Contamination Risk Indices (CRI), an enterprise Object-Oriented software architecture, a statistically "
        "grounded 5-fold cross-validated Random Forest model, and an interactive GIS spatial surveillance dashboard."
    )

    add_section("Key Findings")
    add_bullet("1. Real Data Integrity: ", "On authentic laboratory telemetry, Random Forest achieves 64.77% ± 1.07% 5-fold cross-validation accuracy with an ROC-AUC of 0.672, matching established literature benchmarks and outperforming linear models.")
    add_bullet("2. High Contamination Recall: ", "The model achieves 76% recall on unsafe water samples, ensuring that three out of every four hazardous water events are proactively intercepted.")
    add_bullet("3. Multivariate Interdependence: ", "Gini importance proves that Sulfate (15.6%), pH (12.2%), and Hardness (11.3%) lead decision splits, but all features contribute near the 9.1% baseline, proving water safety is multi-parameter.")
    add_bullet("4. Actionable Spatial Triage: ", "The GIS surveillance engine successfully links predicted risk scores to localized pathogen alerts and automated municipal boil-water protocols.")

    add_section("Limitations of the Study")
    add_body(
        "The system infers biological outbreak hazard from physical and chemical proxies rather than biological assays. Furthermore, current GIS nodes "
        "simulate an urban monitoring grid; full city-scale implementation requires continuous IoT hardware telemetry."
    )

    add_section("Suggestions for Future Research")
    add_bullet("1. Hardware IoT LoRaWAN Integration: ", "Interfacing physical ESP32 and optical turbidity probes via MQTT protocols for continuous telemetry.")
    add_bullet("2. Temporal Deep Learning: ", "Deploying Long Short-Term Memory (LSTM) networks to model seasonal monsoon contamination runoff over time.")
    add_bullet("3. Clinical Data Fusion: ", "Correlating spatial water risk predictions with real-time district hospital acute gastroenteritis admission records.")

    # =========================================================================
    # REFERENCES
    # =========================================================================
    print("[*] Writing References...")
    p_ref = doc.add_paragraph()
    p_ref.paragraph_format.page_break_before = True
    p_ref.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ref.paragraph_format.space_before = Pt(24)
    p_ref.paragraph_format.space_after = Pt(18)
    r_ref = p_ref.add_run("REFERENCES")
    r_ref.font.size = Pt(18)
    r_ref.font.bold = True

    references_list = [
        "[1] World Health Organization. (2022). Guidelines for drinking-water quality: Fourth edition incorporating the first and second addenda. Geneva: World Health Organization. Licence: CC BY-NC-SA 3.0 IGO.",
        "[2] Kadiwal, A. (2021). Water Potability Dataset: Water Quality Metrics for Potability Prediction. Kaggle Benchmark Repository. Available: https://www.kaggle.com/datasets/adityakadiwal/water-potability",
        "[3] Breiman, L. (2001). Random Forests. Machine Learning, Vol. 45, No. 1, pp. 5–32. https://doi.org/10.1023/A:1010933404324",
        "[4] Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., Perrot, M., and Duchesnay, E. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, Vol. 12, pp. 2825–2830.",
        "[5] Few, R., Lake, I., Hunter, P. R., and Tran, F. (2013). Floods, health and climate change: a strategic review. Tyndall Centre for Climate Change Research Working Paper, No. 162, pp. 1–34.",
        "[6] Saha, S., and Bhowmik, S. (2020). Water quality index prediction using machine learning algorithms: a comprehensive review. Environmental Monitoring and Assessment, Vol. 192, No. 11, pp. 1–28. https://doi.org/10.1007/s10661-020-08675-1",
        "[7] Hunter, P. R., MacDonald, A. M., and Carter, R. C. (2010). Water supply and health. PLoS Medicine, Vol. 7, No. 11, e1000361. https://doi.org/10.1371/journal.pmed.1000361",
        "[8] Bain, R., Cronk, R., Hossain, R., Bonjour, S., Onda, K., Wright, J., Yang, H., Slaymaker, T., Hunter, P., Prüss-Ustün, A., and Bartram, J. (2014). Global assessment of exposure to faecal contamination through drinking water. Tropical Medicine & International Health, Vol. 19, No. 8, pp. 917–927."
    ]

    for ref in references_list:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(ref)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)

    # =========================================================================
    # APPENDICES
    # =========================================================================
    print("[*] Writing Appendices...")
    add_chapter("APPENDIX-A\nPSEUDOCODE")
    add_section("Algorithm 1: End-to-End Early-Warning Outbreak Prediction Pipeline")
    pseudo_code = (
        "Algorithm: Waterborne Disease Outbreak Risk Prediction Pipeline\n"
        "Input: Raw water sample vector x = [ph, Hardness, Solids, Chloramines, Sulfate, Conductivity, Organic_carbon, Trihalomethanes, Turbidity]\n"
        "Output: Risk Tier (Low/Moderate/Critical), Suspected Pathogens, Municipal Advisory\n\n"
        "1: Load training median vector M = [median_ph, median_sulfate, median_thm]\n"
        "2: For each feature f in {ph, Sulfate, Trihalomethanes}:\n"
        "3:     If x[f] is NaN:\n"
        "4:         x[f] <- M[f]\n"
        "5: Compute min-max normalized metrics:\n"
        "6:     t_norm <- clip((x.Turbidity - 1.45) / (6.74 - 1.45), 0, 1)\n"
        "7:     c_norm <- clip((x.Organic_carbon - 2.2) / (27.0 - 2.2), 0, 1)\n"
        "8:     m_norm <- clip((x.Trihalomethanes - 0.74) / (124.0 - 0.74), 0, 1)\n"
        "9: Compute Contamination Risk Index:\n"
        "10:    CRI <- (t_norm + c_norm + m_norm) / 3.0\n"
        "11: Compute WHO pH envelope deviation:\n"
        "12:    ph_dev <- max(0, 6.5 - x.ph) + max(0, x.ph - 8.5)\n"
        "13: Construct feature vector X_eval <- [x, CRI, ph_dev]\n"
        "14: Compute model class probabilities using Random Forest:\n"
        "15:    P_unsafe <- RF_Model.predict_proba(X_eval)[Class_0]\n"
        "16: Calculate composite outbreak risk:\n"
        "17:    Composite_Risk <- (0.6 * P_unsafe) + (0.4 * CRI)\n"
        "18: If Composite_Risk >= 0.65:\n"
        "19:    Tier <- 'CRITICAL OUTBREAK ALERT'\n"
        "20:    Pathogens <- ['Vibrio cholerae (Cholera)', 'Salmonella typhi (Typhoid)']\n"
        "21:    Advisory <- 'Issue immediate boil-water advisory; accelerate secondary coagulation.'\n"
        "22: Else If Composite_Risk >= 0.48:\n"
        "23:    Tier <- 'MODERATE HAZARD WARNING'\n"
        "24:    Pathogens <- ['Pipeline Bacterial Re-growth']\n"
        "25:    Advisory <- 'Dispatch mobile emergency chlorine booster injection units.'\n"
        "26: Else:\n"
        "27:    Tier <- 'LOW RISK BASELINE'\n"
        "28:    Pathogens <- ['None detected']\n"
        "29:    Advisory <- 'Maintain standard distribution pressure.'\n"
        "30: Return Tier, Composite_Risk, Pathogens, Advisory"
    )
    p_psc = doc.add_paragraph()
    p_psc.paragraph_format.line_spacing = 1.15
    p_psc.paragraph_format.space_before = Pt(8)
    p_psc.paragraph_format.space_after = Pt(12)
    run_psc = p_psc.add_run(pseudo_code)
    run_psc.font.name = 'Courier New'
    run_psc.font.size = Pt(10)

    add_chapter("APPENDIX-B\nSCREENSHOTS AND SYSTEM INTERFACE")
    add_section("B.1 Terminal Execution & Model Benchmarking")
    add_body("Figure B.1 captures the execution of the master pipeline script (run_all.py) running across data loading, median imputation, feature engineering, 5-fold cross-validation, and GIS engine synthesis.")
    add_figure_caption("Figure B.1: Full Pipeline Terminal Execution Output (run_all.py)")

    add_section("B.2 Interactive Web Dashboard (app.py)")
    add_body("Figure B.2 demonstrates the Streamlit web dashboard interface, illustrating real-time chemical sliders, live risk gauges, and batch CSV processing.")
    add_figure_caption("Figure B.2: Streamlit Interactive Decision-Support Dashboard Interface")

    add_section("B.3 Spatial GIS Surveillance Map (outbreak_risk_map.html)")
    add_body("Figure B.3 showcases the interactive GIS surveillance map, highlighting station markers, multi-layer CartoDB and Esri toggles, and municipal advisory popups.")
    add_figure_caption("Figure B.3: Spatial GIS Disease Outbreak Surveillance Map Interface")

    add_chapter("APPENDIX-C\nENCLOSURES AND SDG MAPPING")
    add_section("C.1 Sustainable Development Goals (SDG) Alignment")
    add_body("The Waterborne Disease Outbreak Risk Prediction System (PRJ_441) directly maps to two core United Nations Sustainable Development Goals:")
    add_bullet("SDG 3: Good Health and Well-Being (Target 3.3 & 3.9) — ", 
               "By providing 24-48 hour advance early warning before clinical symptom onset, the system halts waterborne epidemics (Cholera, Typhoid), "
               "significantly reducing avoidable mortality from water-associated diseases.")
    add_bullet("SDG 6: Clean Water and Sanitation (Target 6.1 & 6.3) — ", 
               "By providing automated, real-time chemical anomaly detection across municipal piped networks, the system ensures equitable, "
               "safe drinking water delivery and reduces chemical and organic pollution exposure.")

    add_section("C.2 Code Repository & Verification Enclosure")
    add_body("The complete, verified production codebase, datasets, models, and presentation decks are publicly hosted under version control at:")
    add_body("GitHub Repository: https://github.com/Sriram-012/waterborne-disease-prediction")

    # Save document
    out_file = "PRJ_441_Complete_Project_Report.docx"
    doc.save(out_file)
    print(f"[OK] Complete Project Report successfully created: '{out_file}'")
    return out_file

if __name__ == '__main__':
    build_report()
