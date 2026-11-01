import matplotlib.pyplot as plt
import matplotlib.patches as patches

def generate_uml_diagrams():
    # 1. Use Case Diagram
    fig, ax = plt.subplots(figsize=(10, 7), dpi=300)
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7)
    ax.axis('off')

    c_primary = '#1E3A8A'
    c_secondary = '#2563EB'
    c_accent = '#0D9488'
    c_text = '#0F172A'
    c_muted = '#475569'

    ax.text(5, 6.7, 'UML Use Case Diagram: AI Resume Screening System', 
            fontsize=13, fontweight='bold', ha='center', color=c_primary)

    # Boundary Box
    boundary = patches.FancyBboxPatch((2.2, 0.4), 5.6, 6.0,
                                     boxstyle="round,pad=0.05,rounding_size=0.15",
                                     edgecolor='#64748B', facecolor='#FFFFFF', lw=1.5)
    ax.add_patch(boundary)
    ax.text(5.0, 6.15, '«System Boundary» AI Resume Platform', fontsize=9.5, fontweight='bold', ha='center', color=c_primary)

    # Actor 1: Recruiter / HR (Left)
    ax.plot([1.0], [4.8], marker='o', markersize=14, color=c_secondary)
    ax.plot([1.0, 1.0], [4.6, 4.0], color=c_secondary, lw=2.5) # body
    ax.plot([0.5, 1.5], [4.4, 4.4], color=c_secondary, lw=2.5) # arms
    ax.plot([1.0, 0.6], [4.0, 3.4], color=c_secondary, lw=2.5) # left leg
    ax.plot([1.0, 1.4], [4.0, 3.4], color=c_secondary, lw=2.5) # right leg
    ax.text(1.0, 3.1, 'Recruiter / HR\n(Hiring Specialist)', fontsize=9, fontweight='bold', ha='center', color=c_primary)

    # Actor 2: System Administrator / Evaluator (Right)
    ax.plot([9.0], [4.8], marker='o', markersize=14, color=c_accent)
    ax.plot([9.0, 9.0], [4.6, 4.0], color=c_accent, lw=2.5)
    ax.plot([8.5, 9.5], [4.4, 4.4], color=c_accent, lw=2.5)
    ax.plot([9.0, 8.6], [4.0, 3.4], color=c_accent, lw=2.5)
    ax.plot([9.0, 9.4], [4.0, 3.4], color=c_accent, lw=2.5)
    ax.text(9.0, 3.1, 'Faculty Evaluator /\nAdministrator', fontsize=9, fontweight='bold', ha='center', color='#065F46')

    # Use cases (Ovals)
    use_cases = [
        ("UC-1: Upload Single / Batch Resumes", 5.4),
        ("UC-2: Input Target Job Description (JD)", 4.6),
        ("UC-3: Semantic Match & Skill Gap Analysis", 3.8),
        ("UC-4: ML Shortlist Prediction & Probability", 3.0),
        ("UC-5: Review Explainability Strengths & Gaps", 2.2),
        ("UC-6: Export Ranked Leaderboard (CSV)", 1.4),
        ("UC-7: Interactive Profile Simulation", 0.7)
    ]

    for uc_text, uc_y in use_cases:
        ellipse = patches.Ellipse((5.0, uc_y), 4.8, 0.55, edgecolor=c_secondary, facecolor='#EFF6FF', lw=1.2)
        ax.add_patch(ellipse)
        ax.text(5.0, uc_y, uc_text, fontsize=8, fontweight='bold', ha='center', va='center', color=c_text)

        # Lines to Recruiter (UC-1 to UC-6)
        if uc_y >= 1.4:
            ax.plot([1.5, 2.6], [4.2, uc_y], color='#94A3B8', linestyle='-', lw=1.2)

        # Lines to Admin / Evaluator (UC-4, UC-5, UC-6, UC-7)
        if uc_y <= 3.8:
            ax.plot([7.4, 8.5], [uc_y, 4.2], color='#94A3B8', linestyle='-', lw=1.2)

    plt.tight_layout()
    plt.savefig('srs_use_case_diagram.png', bbox_inches='tight', dpi=300)
    plt.close()

    # 2. Sequence Diagram
    fig, ax = plt.subplots(figsize=(10, 7.5), dpi=300)
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')

    ax.text(5, 9.6, 'UML Sequence Diagram: End-to-End Candidate Shortlisting', 
            fontsize=13, fontweight='bold', ha='center', color=c_primary)

    lifelines = [
        ("User / Recruiter", 1.2),
        ("Streamlit Web UI", 3.4),
        ("NLP Parser Engine", 5.6),
        ("Semantic Matcher", 7.6),
        ("Gradient Boosting ML", 9.2)
    ]

    for name, x in lifelines:
        box = patches.FancyBboxPatch((x - 0.75, 8.7), 1.5, 0.55,
                                    boxstyle="round,pad=0.05,rounding_size=0.08",
                                    edgecolor=c_primary, facecolor='#DBEAFE', lw=1.2)
        ax.add_patch(box)
        ax.text(x, 8.97, name, fontsize=8, fontweight='bold', ha='center', color=c_primary)
        # Lifeline dashed vertical line
        ax.plot([x, x], [8.7, 0.8], color='#94A3B8', linestyle='--', lw=1.2)

    # Sequence steps (y coordinates descending)
    steps = [
        (8.2, 1.2, 3.4, "1: Upload Resume (PDF/DOCX) & paste JD", "->"),
        (7.4, 3.4, 5.6, "2: parse_resume(file_bytes)", "->"),
        (6.7, 5.6, 5.6, "3: Extract text, skills, exp, edu", "self"),
        (6.0, 5.6, 3.4, "4: Return Candidate Profile Dict", "-->"),
        (5.3, 3.4, 7.6, "5: evaluate_match(resume_text, jd_text)", "->"),
        (4.6, 7.6, 7.6, "6: Compute TF-IDF & Skill Gap", "self"),
        (3.9, 7.6, 3.4, "7: Return match score & skill arrays", "-->"),
        (3.2, 3.4, 9.2, "8: predict_candidate_shortlist(features)", "->"),
        (2.5, 9.2, 9.2, "9: Scale features & run GBDT inference", "self"),
        (1.8, 9.2, 3.4, "10: Return probability & shortlist verdict", "-->"),
        (1.1, 3.4, 1.2, "11: Display Radar Chart, Verdict, & CSV Report", "-->")
    ]

    for y, x1, x2, msg, style in steps:
        if style == "self":
            # Self call loop
            ax.annotate('', xy=(x1, y - 0.25), xytext=(x1 + 0.5, y),
                        arrowprops=dict(arrowstyle="->", color=c_secondary, lw=1.2, connectionstyle="arc3,rad=-0.5"))
            ax.text(x1 + 0.6, y - 0.1, msg, fontsize=7.2, color=c_text, va='center')
        elif style == "-->":
            ax.annotate('', xy=(x2, y), xytext=(x1, y),
                        arrowprops=dict(arrowstyle="->", color='#059669', lw=1.2, linestyle='dashed'))
            ax.text((x1 + x2)/2, y + 0.12, msg, fontsize=7.2, ha='center', color='#047857')
        else:
            ax.annotate('', xy=(x2, y), xytext=(x1, y),
                        arrowprops=dict(arrowstyle="->", color=c_secondary, lw=1.2))
            ax.text((x1 + x2)/2, y + 0.12, msg, fontsize=7.2, ha='center', color=c_primary)

    plt.tight_layout()
    plt.savefig('srs_sequence_diagram.png', bbox_inches='tight', dpi=300)
    plt.close()
    print("UML Diagrams created successfully!")

if __name__ == "__main__":
    generate_uml_diagrams()
