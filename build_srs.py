import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, hex_color):
    """Sets background color of a table cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Sets cell padding (in twips: 20 twips = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_borders(cell, top="CBD5E1", bottom="CBD5E1", left="CBD5E1", right="CBD5E1", sz="4"):
    """Sets cell borders with given color and size."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{top}"/><w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{left}"/><w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{bottom}"/><w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{right}"/></w:tcBorders>')
    tcPr.append(borders)

def build_srs_document(output_path="Software_Requirements_Specification_SRS.docx"):
    doc = Document()

    # Set Margins to 0.75 in
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    HEX_PRIMARY = "1E3A8A"      # Deep Navy
    HEX_SECONDARY = "2563EB"    # Blue
    HEX_ACCENT = "0D9488"       # Teal
    HEX_HEADER_BG = "F1F5F9"    # Slate 100
    HEX_ALT_BG = "F8FAFC"       # Slate 50
    HEX_BORDER = "CBD5E1"       # Slate 300
    
    COLOR_PRIMARY = RGBColor(30, 58, 138)
    COLOR_SECONDARY = RGBColor(37, 99, 235)
    COLOR_MUTED = RGBColor(100, 116, 139)
    COLOR_TEXT = RGBColor(15, 23, 42)

    def add_h1(num_str, title_str):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        
        r_num = h.add_run(f"{num_str}. ")
        r_num.font.name = "Calibri"
        r_num.font.size = Pt(13)
        r_num.font.bold = True
        r_num.font.color.rgb = COLOR_SECONDARY

        r_title = h.add_run(title_str)
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(13)
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_PRIMARY
        return h

    def add_h2(num_str, title_str):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(3)
        h.paragraph_format.keep_with_next = True
        
        r_num = h.add_run(f"{num_str} ")
        r_num.font.name = "Calibri"
        r_num.font.size = Pt(11)
        r_num.font.bold = True
        r_num.font.color.rgb = COLOR_SECONDARY

        r_title = h.add_run(title_str)
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(11)
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_PRIMARY
        return h

    def add_p(text, bold_prefix=None, space_after=6, italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = "Calibri"
            rb.font.size = Pt(10)
            rb.font.bold = True
            rb.font.color.rgb = COLOR_PRIMARY
        rt = p.add_run(text)
        rt.font.name = "Calibri"
        rt.font.size = Pt(10)
        rt.font.color.rgb = COLOR_TEXT
        rt.font.italic = italic
        return p

    def add_bullet(bold_prefix, text):
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.space_before = Pt(1)
        bp.paragraph_format.space_after = Pt(2)
        r_bt = bp.add_run(bold_prefix)
        r_bt.font.name = "Calibri"
        r_bt.font.size = Pt(9.5)
        r_bt.font.bold = True
        r_bt.font.color.rgb = COLOR_PRIMARY
        r_bc = bp.add_run(text)
        r_bc.font.name = "Calibri"
        r_bc.font.size = Pt(9.5)
        r_bc.font.color.rgb = COLOR_TEXT
        return bp

    # -------------------------------------------------------------
    # COVER / TITLE BLOCK
    # -------------------------------------------------------------
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(0)
    p_inst.paragraph_format.space_after = Pt(2)
    r_dept = p_inst.add_run("DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING\n<NAME OF INSTITUTE / UNIVERSITY, CITY>")
    r_dept.font.name = "Calibri"
    r_dept.font.size = Pt(11)
    r_dept.font.bold = True
    r_dept.font.color.rgb = COLOR_MUTED

    p_srs_title = doc.add_paragraph()
    p_srs_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_srs_title.paragraph_format.space_before = Pt(8)
    p_srs_title.paragraph_format.space_after = Pt(2)
    r_srs = p_srs_title.add_run("SOFTWARE REQUIREMENTS SPECIFICATION (SRS)")
    r_srs.font.name = "Calibri"
    r_srs.font.size = Pt(18)
    r_srs.font.bold = True
    r_srs.font.color.rgb = COLOR_PRIMARY

    p_proj_title = doc.add_paragraph()
    p_proj_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj_title.paragraph_format.space_before = Pt(0)
    p_proj_title.paragraph_format.space_after = Pt(10)
    r_pt = p_proj_title.add_run("Project Title: AI-Based Resume Screening and Candidate Shortlisting System\nStandard Academic Compliance: IEEE Std 830-1998")
    r_pt.font.name = "Calibri"
    r_pt.font.size = Pt(11)
    r_pt.font.bold = True
    r_pt.font.color.rgb = COLOR_SECONDARY

    # Split Submission Block Table
    sub_table = doc.add_table(rows=1, cols=2)
    sub_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sub_table.autofit = False
    col_w = [Inches(3.5), Inches(3.5)]

    cell_l, cell_r = sub_table.rows[0].cells[0], sub_table.rows[0].cells[1]
    cell_l.width = col_w[0]
    cell_r.width = col_w[1]
    set_cell_background(cell_l, HEX_ALT_BG)
    set_cell_background(cell_r, HEX_ALT_BG)
    set_cell_margins(cell_l, top=100, bottom=100, left=140, right=140)
    set_cell_margins(cell_r, top=100, bottom=100, left=140, right=140)
    set_cell_borders(cell_l, top=HEX_PRIMARY, bottom=HEX_PRIMARY, left=HEX_PRIMARY, right=HEX_PRIMARY, sz="6")
    set_cell_borders(cell_r, top=HEX_PRIMARY, bottom=HEX_PRIMARY, left=HEX_PRIMARY, right=HEX_PRIMARY, sz="6")

    pl = cell_l.paragraphs[0]
    pl.paragraph_format.space_after = Pt(0)
    rl = pl.add_run("SUBMITTED BY:\n1. Khush Jain (Roll No: [Roll No: XXXXXX])\nDegree: Bachelor of Technology (B.Tech)\nAcademic Year: 2026 - 2027")
    rl.font.name = "Calibri"
    rl.font.size = Pt(9.5)
    rl.font.color.rgb = COLOR_TEXT

    pr = cell_r.paragraphs[0]
    pr.paragraph_format.space_after = Pt(0)
    rr = pr.add_run("UNDER THE GUIDANCE OF:\nProf. / Dr. <Supervisor Name>\n<Designation / Title>\nDepartment of Computer Science & Engg.\nStatus: 3rd Year Project Specification\nDate of Submission: 01-11-2026")
    rr.font.name = "Calibri"
    rr.font.size = Pt(9.5)
    rr.font.color.rgb = COLOR_TEXT

    # Revision & Approval History
    p_rev = doc.add_paragraph()
    p_rev.paragraph_format.space_before = Pt(12)
    p_rev.paragraph_format.space_after = Pt(3)
    r_rev = p_rev.add_run("Document Revision & Approval History")
    r_rev.font.name = "Calibri"
    r_rev.font.size = Pt(10.5)
    r_rev.font.bold = True
    r_rev.font.color.rgb = COLOR_PRIMARY

    rev_data = [
        ("Version", "Date", "Description / Major Changes", "Prepared By", "Reviewed By"),
        ("0.1", "15-Sep-2026", "Initial Draft & Scope Definition", "Project Team", "Project Guide"),
        ("1.0", "30-Sep-2026", "Final 3rd Year Mid-Term SRS Baseline", "Project Team", "HOD / Committee")
    ]
    t_rev = doc.add_table(rows=len(rev_data), cols=5)
    t_rev.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_rev.autofit = False
    rev_widths = [Inches(0.9), Inches(1.1), Inches(2.6), Inches(1.2), Inches(1.2)]

    for row_idx, row_vals in enumerate(rev_data):
        row = t_rev.rows[row_idx]
        for col_idx, val in enumerate(row_vals):
            cell = row.cells[col_idx]
            cell.width = rev_widths[col_idx]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_borders(cell, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if row_idx == 0:
                set_cell_background(cell, HEX_PRIMARY)
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                set_cell_background(cell, HEX_ALT_BG if row_idx % 2 == 1 else "FFFFFF")
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)
                r.font.color.rgb = COLOR_TEXT

    # Guidance Note Box
    p_guidance = doc.add_paragraph()
    p_guidance.paragraph_format.space_before = Pt(8)
    p_guidance.paragraph_format.space_after = Pt(12)
    p_guidance.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_gn = p_guidance.add_run("Guidance Note: This Software Requirements Specification document conforms to the IEEE 830-1998 standard recommended practice. All sections are strictly designed for 3rd-year university lab evaluation, project reviews, and system design baselines.")
    r_gn.font.name = "Calibri"
    r_gn.font.size = Pt(8.5)
    r_gn.font.italic = True
    r_gn.font.color.rgb = COLOR_MUTED

    # -------------------------------------------------------------
    # 1. INTRODUCTION
    # -------------------------------------------------------------
    add_h1("1", "Introduction")

    add_h2("1.1", "Purpose")
    add_p(
        "The purpose of this Software Requirements Specification (SRS) document is to provide a complete, clear, and unambiguous definition "
        "of the functionality, external interfaces, performance targets, and constraints of the proposed system: "
        "'AI-Based Resume Screening and Candidate Shortlisting System'. This specification serves as a binding reference agreement "
        "between student developers, project evaluators, academic supervisors, and prospective enterprise end users during the 3rd-year engineering curriculum."
    )

    add_h2("1.2", "Scope of the System")
    add_p(
        "The proposed application is designed to solve the acute bottleneck of manual resume screening by delivering an end-to-end "
        "automated Natural Language Processing (NLP) and Machine Learning (ML) platform. Key boundary guidelines include:"
    )
    add_bullet("In-Scope Capabilities: ", "Centralized multi-format document ingestion (PDF, Word DOCX, TXT), automated contact and entity extraction, target Job Description parsing, TF-IDF semantic vectorization, skill-gap analysis, Gradient Boosting shortlist classification (90.67% accuracy), explainable AI insights (strengths, red flags, recommendations), and dynamic CSV leaderboard generation.")
    add_bullet("Out-of-Scope Boundaries: ", "Hardware fabrication, enterprise ERP database synchronization, live external Applicant Tracking System (ATS) commercial API sync, and live paid video interviewing integrations.")
    add_bullet("Expected Benefits: ", "Reduction of initial candidate screening turnaround time by over 85%, elimination of subconscious recruiter bias, guaranteed evaluation consistency across 30,000+ applicants, and instantaneous sub-second rank ordering.")

    add_h2("1.3", "Definitions, Acronyms, and Abbreviations")
    def_data = [
        ("Category", "Term / Acronym", "Definition / Standard Context"),
        ("Standard", "SRS", "Software Requirements Specification (IEEE Std 830-1998 compliant)."),
        ("Architecture", "API", "Application Programming Interface for data interchange via JSON/HTTPS."),
        ("Architecture", "REST", "Representational State Transfer stateless software architectural style."),
        ("Database", "CRUD", "Create, Read, Update, Delete basic data manipulation operations."),
        ("Security", "DPDP Act 2023", "Digital Personal Data Protection Act 2023 governing data privacy in India."),
        ("AI / ML", "GBDT", "Gradient Boosted Decision Trees — ensemble machine learning classifier."),
        ("AI / ML", "TF-IDF", "Term Frequency-Inverse Document Frequency numerical statistic for text relevance."),
        ("AI / ML", "Cosine Similarity", "Metric measuring the cosine of the angle between two multi-dimensional vectors."),
        ("Document NLP", "PyPDF / docx", "Python libraries for byte-level text extraction from PDF and Word files."),
        ("UI Framework", "Streamlit", "Open-source Python reactive web framework for real-time data applications.")
    ]
    t_def = doc.add_table(rows=len(def_data), cols=3)
    t_def.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_def.autofit = False
    def_widths = [Inches(1.4), Inches(1.8), Inches(3.8)]
    for row_idx, row_vals in enumerate(def_data):
        row = t_def.rows[row_idx]
        for col_idx, val in enumerate(row_vals):
            cell = row.cells[col_idx]
            cell.width = def_widths[col_idx]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_borders(cell, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if row_idx == 0:
                set_cell_background(cell, HEX_PRIMARY)
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                set_cell_background(cell, HEX_ALT_BG if row_idx % 2 == 1 else "FFFFFF")
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)
                r.font.color.rgb = COLOR_TEXT
                if col_idx == 1:
                    r.font.bold = True

    add_h2("1.4", "References")
    add_bullet("[1] ", "IEEE Std 830-1998: IEEE Recommended Practice for Software Requirements Specifications.")
    add_bullet("[2] ", "Sommerville, Ian. (2015). Software Engineering (10th Edition), Pearson Education.")
    add_bullet("[3] ", "Pedregosa, F., et al. (2011). 'Scikit-learn: Machine Learning in Python.' Journal of Machine Learning Research, 12, 2825–2830.")
    add_bullet("[4] ", "Friedman, J. H. (2001). 'Greedy function approximation: A gradient boosting machine.' The Annals of Statistics, 29(5), 1189–1232.")
    add_bullet("[5] ", "Salton, G., & Buckley, C. (1988). 'Term-weighting approaches in automatic text retrieval.' Information Processing & Management, 24(5), 513–523.")

    add_h2("1.5", "Document Overview")
    add_p(
        "The remainder of this SRS is partitioned into structural sections: Section 2 outlines product perspective, target user personas, "
        "operational context, and design constraints. Section 3 itemizes detailed functional requirements organized into 5 logical modules. "
        "Section 4 specifies external interfaces (UI, Hardware, Software, Network). Section 5 defines non-functional attributes (Performance, Security, Reliability). "
        "Section 6 provides UML analysis models, DFDs, data dictionary schemas, and sequence flows. Finally, Section 7 provides the academic review and sign-off."
    )

    # -------------------------------------------------------------
    # 2. OVERALL DESCRIPTION
    # -------------------------------------------------------------
    add_h1("2", "Overall Description")

    add_h2("2.1", "Product Perspective")
    add_p(
        "The project is structured as a self-contained multi-tier system comprising a presentation layer (Streamlit Web UI), "
        "an ingestion and NLP parsing tier (PyPDF, python-docx, Regex), an ML inference pipeline (Scikit-Learn Gradient Boosting), "
        "and a serialized model/reporting persistence tier. It operates independently without requiring heavy external cloud database dependencies."
    )

    diag_path = "architecture_diagram.png"
    if os.path.exists(diag_path):
        dp = doc.add_paragraph()
        dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dp.paragraph_format.space_before = Pt(4)
        dp.paragraph_format.space_after = Pt(2)
        drun = dp.add_run()
        drun.add_picture(diag_path, width=Inches(6.6))
        
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(8)
        crun = cp.add_run("Figure 2.1: Data Flow Diagram (DFD Level 0 Context Model) & 4-Tier System Architecture")
        crun.font.name = "Calibri"
        crun.font.size = Pt(8.5)
        crun.font.italic = True
        crun.font.color.rgb = COLOR_MUTED

    add_h2("2.2", "User Classes and Characteristics")
    user_classes = [
        ("User Role / Persona", "Technical Proficiency", "Core Responsibilities & Rights"),
        ("System Administrator / Evaluator", "High (DevOps / Python ML)", "Model verification, retraining pipeline execution, hyperparameters tuning, benchmark auditing."),
        ("Recruiter / Hiring Specialist", "Intermediate (General HR / Web)", "Job description entry, single/batch resume upload, leaderboard analysis, report CSV export."),
        ("Candidate / Applicant (Indirect)", "Basic", "Resume submission in standard PDF/DOCX format, reviewing skill feedback and competency matching."),
        ("Academic Evaluator / Guide", "Advanced", "Reviewing architectural models, validating accuracy metrics, and assessing code compliance.")
    ]
    t_uc = doc.add_table(rows=len(user_classes), cols=3)
    t_uc.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_uc.autofit = False
    uc_widths = [Inches(1.8), Inches(1.8), Inches(3.4)]
    for row_idx, row_vals in enumerate(user_classes):
        row = t_uc.rows[row_idx]
        for col_idx, val in enumerate(row_vals):
            cell = row.cells[col_idx]
            cell.width = uc_widths[col_idx]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_borders(cell, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if row_idx == 0:
                set_cell_background(cell, HEX_PRIMARY)
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                set_cell_background(cell, HEX_ALT_BG if row_idx % 2 == 1 else "FFFFFF")
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)
                r.font.color.rgb = COLOR_TEXT
                if col_idx == 0:
                    r.font.bold = True

    add_h2("2.3", "Operating Environment")
    add_bullet("Client Requirements: ", "Modern web browser (Google Chrome >= v110, Mozilla Firefox >= v108, Microsoft Edge >= v110, Apple Safari >= v16); minimum display resolution 1280x720.")
    add_bullet("Server Runtime: ", "Windows 10/11 or Linux (Ubuntu 22.04 LTS / Debian 12); Python 3.10+ (Current runtime: Python 3.13.3) with Streamlit web server daemon on Port 8501.")
    add_bullet("Storage & Artifact Tier: ", "Local filesystem storage containing pre-trained models ('best_model.pkl', 'scaler.pkl', 'metadata.pkl') and tabular historical screening dataset ('ai_resume_screening.csv', 30,000 rows).")

    add_h2("2.4", "Design and Implementation Constraints")
    add_bullet("Academic Budget Constraint: ", "100% open-source software stack utilizing Scikit-learn, Python, and local execution with zero recurring cloud API subscription costs.")
    add_bullet("Sub-Second Latency Constraint: ", "Inference execution per candidate profile must terminate within 150 milliseconds to enable dynamic real-time slider adjustments.")
    add_bullet("Data Privacy (DPDP Act 2023): ", "No personally identifiable information (PII) such as phone numbers or home addresses are stored permanently; resumes are parsed in-memory and discarded post-inference.")

    add_h2("2.5", "Assumptions and Dependencies")
    add_bullet("Assumptions: ", "Uploaded candidate resumes adhere to standard textual PDF or Word format (not raster image scans); target Job Descriptions contain at least 15 words.")
    add_bullet("Dependencies: ", "Availability of Scikit-Learn 1.9.1, NumPy 2.3.2, Pandas 2.3.2, PyPDF 6.21.4, python-docx 1.2.0, and Plotly 6.3.0.")

    # -------------------------------------------------------------
    # 3. FUNCTIONAL REQUIREMENTS
    # -------------------------------------------------------------
    add_h1("3", "System Features and Functional Requirements")

    modules_reqs = [
        ("3.1 Module 1: Resume Ingestion & Multi-Format Parsing", [
            ("FR-1.1", "Multi-Format Resume Upload", "Accepts .pdf, .docx, and .txt files up to 10MB via drag-and-drop", "High"),
            ("FR-1.2", "Dual-Engine Text Extraction", "Extracts raw text via PyPDF with fallback to pdfplumber / python-docx", "High"),
            ("FR-1.3", "Candidate Entity Parsing", "Extracts candidate Name, Email, Phone, and LinkedIn/GitHub URLs via Regex", "High"),
            ("FR-1.4", "Skills & Experience Extraction", "Matches against 100+ technical taxonomy; computes years of experience and education level", "High")
        ]),
        ("3.2 Module 2: Job Description Parsing & Semantic Matching", [
            ("FR-2.1", "Job Description Ingestion", "Text area input supporting custom recruiter JDs with preset role templates", "High"),
            ("FR-2.2", "Direct Skill Overlap Computation", "Categorizes candidate skills into Matched, Missing, and Bonus Skills", "High"),
            ("FR-2.3", "TF-IDF Vectorization", "Vectorizes resume and JD corpus using unigram/bigram n-gram tokenization", "Medium"),
            ("FR-2.4", "Cosine Similarity Scoring", "Calculates semantic text relevance score ranging from 0.0% to 100.0%", "High")
        ]),
        ("3.3 Module 3: Machine Learning Shortlisting & Predictive Inference", [
            ("FR-3.1", "Feature Vector Assembly", "Assembles 6 standardized features matching exact metadata training order", "High"),
            ("FR-3.2", "StandardScaler Normalization", "Transforms candidate feature vector using pre-fitted training scaler", "High"),
            ("FR-3.3", "Gradient Boosting Inference", "Executes trained GBDT model returning binary prediction (Shortlisted/Not)", "High"),
            ("FR-3.4", "Probability Calibration & Thresholding", "Outputs confidence score (0-100%) with dynamic recruiter threshold slider", "Medium")
        ]),
        ("3.4 Module 4: Explainable AI, Insights & Candidate Feedback", [
            ("FR-4.1", "Strengths Identification", "Detects competitive advantages (e.g., high GitHub activity, advanced degree)", "Medium"),
            ("FR-4.2", "Red Flag Detection", "Surfaces potential hiring liabilities (e.g., severe experience deficit)", "Medium"),
            ("FR-4.3", "Actionable Improvement Advice", "Generates tailored recommendations to bridge identified skill gaps", "Low")
        ]),
        ("3.5 Module 5: Leaderboard, Analytics Dashboard & Export Engine", [
            ("FR-5.1", "Single Candidate Deep-Dive View", "Interactive Plotly doughnut charts, score breakdown meters, and raw text preview", "High"),
            ("FR-5.2", "Batch Candidate Leaderboard", "Uploads multiple resumes, sorts candidates by match score, and assigns ranks", "High"),
            ("FR-5.3", "CSV Shortlist Export", "Generates downloadable 'candidate_shortlist_report.csv' with all metrics", "Medium"),
            ("FR-5.4", "Hypothetical Profile Simulator", "Real-time slider controls allowing recruiters to simulate what-if scenarios", "Medium")
        ])
    ]

    for mod_title, req_list in modules_reqs:
        doc.add_paragraph()
        h_mod = doc.add_paragraph()
        h_mod.paragraph_format.space_before = Pt(8)
        h_mod.paragraph_format.space_after = Pt(3)
        h_mod.paragraph_format.keep_with_next = True
        r_mod = h_mod.add_run(mod_title)
        r_mod.font.name = "Calibri"
        r_mod.font.size = Pt(10.5)
        r_mod.font.bold = True
        r_mod.font.color.rgb = COLOR_PRIMARY

        t_req = doc.add_table(rows=len(req_list) + 1, cols=4)
        t_req.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_req.autofit = False
        req_widths = [Inches(1.0), Inches(2.2), Inches(3.0), Inches(0.8)]

        # Header
        headers = ["Req ID", "Feature Description", "Input / Validation Rule", "Priority"]
        for c_idx, h_text in enumerate(headers):
            cell = t_req.rows[0].cells[c_idx]
            cell.width = req_widths[c_idx]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_borders(cell, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
            set_cell_background(cell, HEX_PRIMARY)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(h_text)
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for r_idx, row_vals in enumerate(req_list, start=1):
            row = t_req.rows[r_idx]
            for c_idx, val in enumerate(row_vals):
                cell = row.cells[c_idx]
                cell.width = req_widths[c_idx]
                set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
                set_cell_borders(cell, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
                set_cell_background(cell, HEX_ALT_BG if r_idx % 2 == 1 else "FFFFFF")
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)
                r.font.color.rgb = COLOR_TEXT
                if c_idx == 0:
                    r.font.bold = True
                if c_idx == 3:
                    if val == "High":
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(185, 28, 28)

    # -------------------------------------------------------------
    # 4. EXTERNAL INTERFACES
    # -------------------------------------------------------------
    add_h1("4", "External Interface Requirements")

    add_h2("4.1", "User Interfaces (UI)")
    add_p(
        "The web interface is engineered via Streamlit with custom CSS styling inspired by modern design systems (Plus Jakarta Sans font, "
        "dark gradient hero container, card elevation, and responsive flex pill tags for skills). Key interface workspaces include:"
    )
    add_bullet("1. Single Candidate Screening Workspace: ", "Split screen layout with resume file uploader & JD input on the left; real-time gauge meter, shortlisted verdict badge, radar score breakdown, and explainability tabs on the right.")
    add_bullet("2. Batch Leaderboard Workspace: ", "Multi-file drag-and-drop pane evaluating 50+ resumes simultaneously against a selected JD, dynamically generating a ranked table with shortlisting badges, match scores, and a CSV download button.")
    add_bullet("3. Hypothetical Profile Simulator: ", "Interactive slider panel with 6 real-time controls (Experience, Match %, Degree, Projects, Resume Words, GitHub Commits) with instant recalculation.")
    add_bullet("4. Model Architecture & System Insights: ", "Diagnostic view presenting training distributions (30k rows), correlation heatmaps, confusion matrices, and ROC-AUC curves.")

    add_h2("4.2", "Hardware Interfaces")
    add_bullet("Client Hardware: ", "Standard personal computer or laptop equipped with input peripherals (keyboard, mouse/trackpad) and display with at least 720p resolution.")
    add_bullet("Server Hardware: ", "Standard x86_64 CPU workstation (Min 4GB RAM, Intel Core i5 / AMD Ryzen 5 or cloud equivalent). The inference pipeline is completely CPU-optimized and does not require dedicated GPU hardware.")

    add_h2("4.3", "Software Interfaces")
    add_bullet("Operating System: ", "Cross-platform compatibility across Windows 10/11, macOS, and Linux (Ubuntu 20.04+).")
    add_bullet("Python Runtime: ", "Python 3.10 to 3.13 (Active: Python 3.13.3) with standard site-packages.")
    add_bullet("Core Libraries: ", "Streamlit (1.49.1), Scikit-Learn (1.9.1), Pandas (2.3.2), NumPy (2.3.2), Plotly (6.3.0), PyPDF (6.21.4), python-docx (1.2.0), pdfplumber (0.11.7).")
    add_bullet("Model Storage: ", "Standard Joblib serialization for 'best_model.pkl' (Gradient Boosting), 'scaler.pkl', and 'metadata.pkl'.")

    add_h2("4.4", "Communication Interfaces")
    add_bullet("Web Protocol: ", "HTTP / HTTPS over standard ports (default development port: 8501).")
    add_bullet("Session Communication: ", "WebSocket protocol (ws:// / wss://) utilized by Streamlit for bi-directional reactive updates between client browser and backend Python engine.")
    add_bullet("Data Ingestion Format: ", "Multipart form-data encoding for binary document file uploads (PDF, DOCX).")

    # -------------------------------------------------------------
    # 5. NON-FUNCTIONAL REQUIREMENTS
    # -------------------------------------------------------------
    add_h1("5", "Non-Functional Requirements (NFRs)")

    nfr_data = [
        ("Attribute", "Target Metric / Standard", "Verification & Testing Method"),
        ("5.1 Performance", "Single resume parsing & inference latency < 120 ms; Batch processing of 25 resumes < 3.5s; Page load < 1.5s.", "Execution profiling using Python time module and Apache Benchmark load testing."),
        ("5.2 Security & Privacy", "Zero persistent storage of PII data; In-memory stream buffer disposal post-inference; Strict DPDP Act 2023 compliance.", "Code inspection, static vulnerability scan, and memory heap leak verification."),
        ("5.3 Reliability & Availability", "System uptime >= 99.0% during academic evaluation; Automatic error recovery on encrypted or malformed PDF documents.", "Fuzz testing with corrupted and non-standard PDF/DOCX files; unhandled exception monitoring."),
        ("5.4 Usability", "Intuitive navigation with zero training required; Visual feedback badges (green/red) and interactive explainability cards.", "System Usability Scale (SUS) survey administered to academic peers (Target SUS > 85)."),
        ("5.5 Portability & Maintainability", "Full cross-browser compatibility (Chrome, Edge, Firefox, Safari); Fully decoupled modular code structure.", "Cross-browser testing in Developer Tools; PEP 8 style guide compliance.")
    ]
    t_nfr = doc.add_table(rows=len(nfr_data), cols=3)
    t_nfr.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_nfr.autofit = False
    nfr_widths = [Inches(1.8), Inches(3.2), Inches(2.0)]
    for row_idx, row_vals in enumerate(nfr_data):
        row = t_nfr.rows[row_idx]
        for col_idx, val in enumerate(row_vals):
            cell = row.cells[col_idx]
            cell.width = nfr_widths[col_idx]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_borders(cell, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if row_idx == 0:
                set_cell_background(cell, HEX_PRIMARY)
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                set_cell_background(cell, HEX_ALT_BG if row_idx % 2 == 1 else "FFFFFF")
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)
                r.font.color.rgb = COLOR_TEXT
                if col_idx == 0:
                    r.font.bold = True

    # -------------------------------------------------------------
    # 6. SYSTEM DESIGN & ANALYSIS MODELS
    # -------------------------------------------------------------
    add_h1("6", "System Design and Analysis Models")

    add_h2("6.1", "UML Use Case Diagram & Actor Specification")
    add_p(
        "The Use Case model maps primary actors (Recruiter/HR, System Administrator/Faculty Evaluator) "
        "to functional interactions within the system boundary:"
    )

    uc_diag_path = "srs_use_case_diagram.png"
    if os.path.exists(uc_diag_path):
        dp = doc.add_paragraph()
        dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dp.paragraph_format.space_before = Pt(4)
        dp.paragraph_format.space_after = Pt(2)
        drun = dp.add_run()
        drun.add_picture(uc_diag_path, width=Inches(6.4))
        
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(6)
        crun = cp.add_run("Figure 6.1: UML Use Case Diagram for AI Resume Screening System")
        crun.font.name = "Calibri"
        crun.font.size = Pt(8.5)
        crun.font.italic = True
        crun.font.color.rgb = COLOR_MUTED

    add_h2("6.2", "Data Flow Diagram (DFD Level 0 & Level 1 Modular Details)")
    add_bullet("DFD Level 0 (Context): ", "External HR Recruiter supplies Job Description and Resumes. The central 0.0 AI Resume Screening System processes candidate credentials against criteria and returns shortlisted status, skill gap metrics, and ranked CSV records.")
    add_bullet("DFD Level 1.1 Ingestion: ", "File upload stream is sanitized and routed to PyPDF/python-docx parsers producing raw unicode text.")
    add_bullet("DFD Level 1.2 Entity Extraction: ", "Regex matches candidate metadata (Name, Email, Phone, URLs) and skills taxonomy dictionary filters 100+ domain skills.")
    add_bullet("DFD Level 1.3 Semantic Matching: ", "TF-IDF Vectorizer computes cosine similarity against tokenized JD text.")
    add_bullet("DFD Level 1.4 ML Classification: ", "Normalized 6-feature vector is evaluated by Gradient Boosting model returning class probability.")

    add_h2("6.3", "Feature Schema & Data Dictionary")
    add_p("The ML screening engine relies on 6 predictor features and 1 ground-truth target variable from the 30,000-record dataset:")
    
    schema_data = [
        ("Feature Name", "Data Type", "Value Range / Categories", "Engineering & Normalization Role"),
        ("years_experience", "Integer (Numeric)", "0 to 15+ years", "Scaled via StandardScaler; measures cumulative professional tenure."),
        ("skills_match_score", "Float (Continuous)", "0.0% to 100.0%", "Derived via TF-IDF cosine similarity & direct skill overlap."),
        ("education_level", "Ordinal (Integer)", "High School=0, Bachelors=1, Masters=2, PhD=3", "Mapped categorical academic qualification tier."),
        ("project_count", "Integer (Discrete)", "0 to 10+ projects", "Quantifies completed software/engineering portfolio deliverables."),
        ("resume_length", "Integer (Numeric)", "100 to 1,500+ words", "Proxy for content depth and resume detail richness."),
        ("github_activity", "Integer (Numeric)", "0 to 500+ commits", "Measures open-source developer engagement and coding contributions."),
        ("shortlisted (Target)", "Categorical Binary", "'Yes' (1) / 'No' (0)", "Supervised ground-truth label for Gradient Boosting classifier.")
    ]
    t_sch = doc.add_table(rows=len(schema_data), cols=4)
    t_sch.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sch.autofit = False
    sch_widths = [Inches(1.6), Inches(1.4), Inches(2.0), Inches(2.0)]
    for row_idx, row_vals in enumerate(schema_data):
        row = t_sch.rows[row_idx]
        for col_idx, val in enumerate(row_vals):
            cell = row.cells[col_idx]
            cell.width = sch_widths[col_idx]
            set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
            set_cell_borders(cell, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if row_idx == 0:
                set_cell_background(cell, HEX_PRIMARY)
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                set_cell_background(cell, HEX_ALT_BG if row_idx % 2 == 1 else "FFFFFF")
                r = p.add_run(val)
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)
                r.font.color.rgb = COLOR_TEXT
                if col_idx == 0:
                    r.font.bold = True

    add_h2("6.4", "UML Sequence Diagram: End-to-End Evaluation Workflow")
    add_p(
        "The chronological interaction sequence between user, Streamlit controller, NLP parser, semantic matcher, "
        "and Gradient Boosting ML model is detailed below:"
    )

    seq_diag_path = "srs_sequence_diagram.png"
    if os.path.exists(seq_diag_path):
        dp = doc.add_paragraph()
        dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dp.paragraph_format.space_before = Pt(4)
        dp.paragraph_format.space_after = Pt(2)
        drun = dp.add_run()
        drun.add_picture(seq_diag_path, width=Inches(6.4))
        
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(8)
        crun = cp.add_run("Figure 6.2: UML Sequence Diagram for End-to-End Candidate Screening & Shortlisting")
        crun.font.name = "Calibri"
        crun.font.size = Pt(8.5)
        crun.font.italic = True
        crun.font.color.rgb = COLOR_MUTED

    # -------------------------------------------------------------
    # 7. ACADEMIC REVIEW & SIGN-OFF
    # -------------------------------------------------------------
    add_h1("7", "Academic Review & Sign-Off")
    add_p(
        "This Software Requirements Specification document has been submitted, presented, and reviewed by the academic supervisory panel "
        "in compliance with the B.Tech Computer Science & Engineering 3rd-Year project evaluation criteria:"
    )

    sign_table = doc.add_table(rows=2, cols=3)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_table.autofit = False
    s_col_w = [Inches(2.3), Inches(2.3), Inches(2.4)]

    sign_headers = ["Project Coordinator / Guide", "Internal Examiner", "Head of Department (HOD)"]
    for idx, sh in enumerate(sign_headers):
        c = sign_table.rows[0].cells[idx]
        c.width = s_col_w[idx]
        set_cell_background(c, HEX_PRIMARY)
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        set_cell_borders(c, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(sh)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    sign_bodies = [
        "Signature: ____________________\n\nName: Prof. / Dr. <Supervisor>\nDate: 01-11-2026\nStatus: Approved / Baseline",
        "Signature: ____________________\n\nName: <Internal Examiner>\nDate: ________________________\nStatus: Recommended",
        "Signature: ____________________\n\nName: Dr. <HOD Name>\nDate: ________________________\nStatus: Final Acceptance"
    ]
    for idx, sb in enumerate(sign_bodies):
        c = sign_table.rows[1].cells[idx]
        c.width = s_col_w[idx]
        set_cell_background(c, HEX_ALT_BG)
        set_cell_margins(c, top=80, bottom=80, left=80, right=80)
        set_cell_borders(c, top=HEX_BORDER, bottom=HEX_BORDER, left=HEX_BORDER, right=HEX_BORDER)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(sb)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.color.rgb = COLOR_TEXT

    doc.save(output_path)
    print(f"SRS Document successfully built and saved to: {output_path}")

if __name__ == "__main__":
    build_srs_document("Software_Requirements_Specification_SRS.docx")
    build_srs_document("SRS.docx")
