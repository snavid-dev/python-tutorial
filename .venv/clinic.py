import docx
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


# --- Helper Functions for English Formatting ---
def format_run(run, size=11, bold=False):
    run.font.name = 'Times New Roman'  # Standard for English research papers
    run.font.size = Pt(size)
    run.font.bold = bold


def add_paragraph(doc, text, bold=False, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.alignment = alignment
    run = p.add_run(text)
    format_run(run, size, bold)
    return p


def create_english_research_crf():
    doc = docx.Document()

    # Page Margins (Standard A4)
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    # --- Footer ---
    section = doc.sections[0]
    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_line = footer_p.add_run("\n_________________________________________________________________________________\n")
    run_line.font.size = Pt(10)

    run_ft = footer_p.add_run("International Research CRF - Pain Management Department")
    format_run(run_ft, size=9)

    # ================= HEADER =================
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("CLINICAL RESEARCH CASE REPORT FORM (CRF)")
    format_run(run_title, size=16, bold=True)

    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle_p.add_run("Pain Management Center – Herat / Iran Clinic")
    format_run(run_sub, size=12, bold=True)

    doc.add_paragraph()

    # ================= SECTION 1: Demographics =================
    add_paragraph(doc, "1. Patient Demographics & Risk Factors", bold=True, size=12)

    table_demo = doc.add_table(rows=4, cols=4)
    table_demo.style = 'Table Grid'

    # Data structure: (Label 1, Input 1, Label 2, Input 2)
    demo_data = [
        ("File ID / Code:", "", "Visit Date:", ""),
        ("Patient Name/Initials:", "", "Gender:", ""),
        ("Age:", "", "Height (cm) / Weight (kg):", ""),
        ("Occupation:", "", "BMI / Smoking History:", "")
    ]

    for r_idx, row_data in enumerate(demo_data):
        row = table_demo.rows[r_idx]

        # Col 0: Label
        p = row.cells[0].paragraphs[0]
        format_run(p.add_run(row_data[0]), bold=True)

        # Col 2: Label
        p = row.cells[2].paragraphs[0]
        format_run(p.add_run(row_data[2]), bold=True)

    # Ethical Consent
    p_consent = doc.add_paragraph()
    r_con = p_consent.add_run("\n[ ] Informed Research Consent Signed by Patient")
    format_run(r_con, bold=True)

    doc.add_paragraph()

    # ================= SECTION 2: Baseline Assessment =================
    add_paragraph(doc, "2. Baseline Clinical Assessment", bold=True, size=12)

    # History
    p_hist = doc.add_paragraph()
    r_h1 = p_hist.add_run("Symptom Duration (Months): _______   History of Related Surgery: [ ] Yes  [ ] No")
    format_run(r_h1)

    doc.add_paragraph()
    add_paragraph(doc, "Confirmed Diagnosis (ICD-10): ______________________________________________________")

    # Scores
    table_scores = doc.add_table(rows=2, cols=2)
    table_scores.style = 'Table Grid'

    # Row 1
    p = table_scores.rows[0].cells[0].paragraphs[0]
    format_run(p.add_run("Baseline VAS Pain Score (0-10): ______"), bold=True)

    p = table_scores.rows[0].cells[1].paragraphs[0]
    format_run(p.add_run("Functional Score (WOMAC / ODI): ______"), bold=True)

    # Row 2 - Comorbidities
    p = table_scores.rows[1].cells[0].paragraphs[0]
    format_run(p.add_run("Comorbidities: [ ] Diabetes  [ ] HTN  [ ] Autoimmune"), size=10)

    p = table_scores.rows[1].cells[1].paragraphs[0]
    format_run(p.add_run("Medications: [ ] NSAIDs  [ ] Corticosteroids  [ ] Anticoagulants"), size=10)

    doc.add_paragraph()

    # ================= SECTION 3: Intervention Details =================
    add_paragraph(doc, "3. Intervention Technical Details", bold=True, size=12)

    p_type = doc.add_paragraph()
    r_type = p_type.add_run("Procedure Type:  [ ] PRP  [ ] Gel (HA)  [ ] Corticosteroid  [ ] Prolotherapy  [ ] Other")
    format_run(r_type)

    # Technical Grid
    table_tech = doc.add_table(rows=2, cols=2)
    table_tech.style = 'Table Grid'

    # Cell 0,0: Material
    p = table_tech.rows[0].cells[0].paragraphs[0]
    format_run(p.add_run("Material Specs (Brand/Kit):"), bold=True)
    format_run(p.add_run("\n__________________________\nVol (ml): ______"))

    # Cell 0,1: Guidance
    p = table_tech.rows[0].cells[1].paragraphs[0]
    format_run(p.add_run("Guidance Method:"), bold=True)
    p.add_run("\n[ ] Blind (Landmark)\n[ ] Ultrasound Guided\n[ ] Fluoroscopy (C-Arm)")

    # Cell 1,0: Site
    p = table_tech.rows[1].cells[0].paragraphs[0]
    format_run(p.add_run("Injection Site (Side/Level): _______________"))

    # Cell 1,1: Approach
    p = table_tech.rows[1].cells[1].paragraphs[0]
    format_run(p.add_run("Approach / Technique: _______________"))

    doc.add_page_break()

    # ================= SECTION 4: Follow-up =================
    title_fu = doc.add_paragraph()
    title_fu.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_fu = title_fu.add_run("4. Standardized Follow-up Protocol")
    format_run(run_fu, size=14, bold=True)

    doc.add_paragraph()

    # Headers
    headers = [
        "Timepoint",
        "Date",
        "VAS Pain\n(0-10)",
        "Functional Imp.\n(%)",
        "Adverse Events",
        "Patient Satisfaction\n(1-5 Likert)"
    ]

    table_fu = doc.add_table(rows=6, cols=6)
    table_fu.style = 'Table Grid'

    # Set Headers
    for i, h in enumerate(headers):
        cell = table_fu.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h = p.add_run(h)
        format_run(run_h, bold=True, size=9)
        # Gray background
        tc_pr = cell._element.get_or_add_tcPr()
        shd = docx.oxml.shared.OxmlElement("w:shd")
        shd.set(docx.oxml.ns.qn("w:fill"), "E7E7E7")
        tc_pr.append(shd)

    # Intervals
    intervals = ["Baseline (Injection)", "2 Weeks", "1 Month", "3 Months", "6 Months"]

    for idx, interval in enumerate(intervals):
        row = table_fu.rows[idx + 1]
        cell_label = row.cells[0]
        p = cell_label.paragraphs[0]
        format_run(p.add_run(interval), bold=True, size=10)

    doc.add_paragraph()

    # ================= SECTION 5: Outcomes & Signatures =================
    add_paragraph(doc, "5. Safety & Conclusion", bold=True)

    p_ae = doc.add_paragraph()
    r_ae = p_ae.add_run(
        "[ ] No Adverse Events    [ ] Infection    [ ] Flare-up    [ ] Nerve Injury    [ ] Other: _______")
    format_run(r_ae)

    doc.add_paragraph()
    add_paragraph(doc, "Physician / Investigator Notes:")
    doc.add_paragraph("_" * 80)
    doc.add_paragraph("_" * 80)

    # Signatures
    doc.add_paragraph()
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.style = 'Table Grid'
    # Remove borders for signature block look
    sig_table.rows[0].cells[0].border = None
    sig_table.rows[0].cells[1].border = None

    p_sig1 = sig_table.rows[0].cells[0].paragraphs[0]
    p_sig1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_run(p_sig1.add_run("\n_______________________\nInvestigator Signature"), bold=True)

    p_sig2 = sig_table.rows[0].cells[1].paragraphs[0]
    p_sig2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_run(p_sig2.add_run("\n_______________________\nSupervisor / Head of Clinic"), bold=True)

    # Save
    file_path = "Pain_Clinic_Research_CRF_English.docx"
    doc.save(file_path)
    print(f"File created successfully: {file_path}")


if __name__ == "__main__":
    create_english_research_crf()
