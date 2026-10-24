"""
generate_report_1.py
Creates Report_1_Certificates_and_Front_Matter.docx based on the exact template of media_1792801392422.docx
Preserves the exact university header images, signatures tables, layout, font sizes, margins, and page breaks.
Adapts all student details, registration numbers, and project title:
- Project Title: "WATERBORNE DISEASE OUTBREAK RISK PREDICTION AND SPATIAL SURVEILLANCE SYSTEM USING MACHINE LEARNING AND GIS"
- Students:
  * K Sriram (20231COM0002)
  * Hemraj N A (20231COM0055)
  * Nagbhooshan (20231COM0029)
- Degree: Bachelor of Technology in Computer Engineering
- Department: Presidency School of Computer Science and Engineering, Presidency University, Bengaluru
"""

import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def generate_report_1():
    user_dir = r"C:\Users\LENOVO\.gemini\antigravity\brain\38e0aa45-c2da-4a11-a610-4a454595a8fa\.user_uploaded"
    src_file = os.path.join(user_dir, "media_1792801392422.docx")
    
    if not os.path.exists(src_file):
        raise FileNotFoundError(f"Source file {src_file} not found.")

    doc = docx.Document(src_file)

    # Dictionary of paragraph indices to replacement run configurations:
    # index: (text, size_pt, bold, alignment, line_spacing)
    updates = {
        1: (
            "“WATERBORNE DISEASE OUTBREAK RISK PREDICTION AND SPATIAL SURVEILLANCE SYSTEM USING MACHINE LEARNING AND GIS”",
            20, True, WD_ALIGN_PARAGRAPH.CENTER, 1.15
        ),
        6: ("K SRIRAM - 20231COM0002", 18, True, WD_ALIGN_PARAGRAPH.CENTER, 1.15),
        7: ("HEMRAJ N A - 20231COM0055", 18, True, WD_ALIGN_PARAGRAPH.CENTER, 1.15),
        8: ("NAGBHOOSHAN - 20231COM0029", 18, True, WD_ALIGN_PARAGRAPH.CENTER, 1.15),
        25: (
            "Certified that this report “Waterborne Disease Outbreak Risk Prediction and Spatial Surveillance System Using Machine Learning and GIS” is a bonafide work of “Mr. K Sriram (20231COM0002)”, “Mr. Hemraj N A (20231COM0055)”, “Mr. Nagbhooshan (20231COM0029)”, who have successfully carried out the project work and submitted the report for partial fulfilment of the requirements for the award of the degree of BACHELOR OF TECHNOLOGY in COMPUTER ENGINEERING during the academic year 2026-27.",
            12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5
        ),
        51: (
            "This is to certify that Mr. K Sriram (20231COM0002), Mr. Hemraj N A (20231COM0055), and Mr. Nagbhooshan (20231COM0029) have been subjected to the Viva-voce Examination on 28-11-2026 at Presidency University, Bengaluru.",
            12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5
        ),
        59: (
            "We the students of final year B.Tech in COMPUTER SCIENCE AND ENGINEERING, at Presidency University, Bengaluru, named K SRIRAM (20231COM0002), HEMRAJ N A (20231COM0055), and NAGBHOOSHAN (20231COM0029), hereby declare that the project work titled “Waterborne Disease Outbreak Risk Prediction and Spatial Surveillance System Using Machine Learning and GIS” has been independently carried out by us and submitted in partial fulfillment for the award of the degree of B.Tech in COMPUTER ENGINEERING during the academic year of 2026-27. Further, the matter embodied in the project has not been submitted previously by anybody for the award of any Degree or Diploma to any other institution.",
            12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5
        ),
        61: ("K Sriram                              20231COM0002                           K Sriram", 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.15),
        62: ("Hemraj N A                            20231COM0055                           Hemraj N A", 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.15),
        63: ("Nagbhooshan                           20231COM0029                           Nagbhooshan", 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.15),
        75: ("K SRIRAM", 12, True, WD_ALIGN_PARAGRAPH.RIGHT, 1.15),
        76: ("HEMRAJ N A", 12, True, WD_ALIGN_PARAGRAPH.RIGHT, 1.15),
        77: ("NAGBHOOSHAN", 12, True, WD_ALIGN_PARAGRAPH.RIGHT, 1.15),
    }

    for idx, (text_val, pt_size, is_bold, alignment, line_sp) in updates.items():
        if idx < len(doc.paragraphs):
            p = doc.paragraphs[idx]
            p.text = ""  # Clear existing runs cleanly
            r = p.add_run(text_val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(pt_size)
            r.font.bold = is_bold
            r.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = alignment
            p.paragraph_format.line_spacing = line_sp

    out_file = "Report_1_Certificates_and_Front_Matter.docx"
    doc.save(out_file)
    print(f"[OK] Report 1 created successfully: '{out_file}' (Size: {os.path.getsize(out_file)} bytes)")
    return out_file

if __name__ == "__main__":
    generate_report_1()
