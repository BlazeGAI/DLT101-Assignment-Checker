from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def check_word_1(doc):
    checklist_data = {
        "Grading Criteria": [
            "Is the font Times New Roman, 12pt?",
            "Is line spacing set to double?",
            "Are margins set to 1 inch on all sides?",
            "Is the title centered and not bold?",
            "Are paragraphs properly indented?",
            "Are there at least 3 paragraphs?",
            "Is there a References page?",
            "Are in-text citations properly formatted?"
        ],
        "Completed": []
    }

    def is_correct_font(paragraph):
        """Check if the paragraph uses Times New Roman, 12pt font."""
        paragraph_font = paragraph.style.font.name
        paragraph_size = paragraph.style.font.size

        for run in paragraph.runs:
            run_font = run.font.name or paragraph_font  # Use run font if set, otherwise fallback to paragraph
            run_size = run.font.size or paragraph_size  # Use run size if set, otherwise fallback to paragraph
            
            if run_font != 'Times New Roman' or run_size != Pt(12):
                return False
        return True

    # Check if all paragraphs follow the correct font rule
    correct_font = all(is_correct_font(p) for p in doc.paragraphs if p.text.strip())
    checklist_data["Completed"].append("Yes" if correct_font else "No")

    # Check if line spacing is set to double (2.0)
    correct_spacing = all(
        p.paragraph_format.line_spacing in [None, 2.0] for p in doc.paragraphs if p.text.strip()
    )
    checklist_data["Completed"].append("Yes" if correct_spacing else "No")

    # Check if margins are set to 1 inch on all sides
    correct_margins = all(
        section.left_margin.inches == 1 and
        section.right_margin.inches == 1 and
        section.top_margin.inches == 1 and
        section.bottom_margin.inches == 1
        for section in doc.sections
    )
    checklist_data["Completed"].append("Yes" if correct_margins else "No")

    # Check if title is centered and not bold
    title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
    title_centered = (
        title_paragraph and 
        title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
        not any(run.bold for run in title_paragraph.runs)
    )
    checklist_data["Completed"].append("Yes" if title_centered else "No")

    # Count body paragraphs and detect References section
    body_paragraphs = []
    found_references = False
    header_done = False
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if not header_done and len(text) > 100:
            header_done = True
        if text.lower() == 'references':
            found_references = True
            continue
        if header_done and not found_references and len(text) > 100 and not text.lower().startswith('in conclusion'):
            body_paragraphs.append(text)

    # Indentation check (placeholder, real indentation check can be added if needed)
    proper_indentation = True  
    checklist_data["Completed"].append("Yes" if proper_indentation else "No")

    # Check if there are at least 3 body paragraphs
    sufficient_paragraphs = len(body_paragraphs) >= 3
    checklist_data["Completed"].append("Yes" if sufficient_paragraphs else "No")

    # Check if References section exists and contains at least one reference
    has_references = any(p.text.strip().lower() == 'references' for p in doc.paragraphs)
    references_content = any(
        has_references and '(' in p.text and ')' in p.text for p in doc.paragraphs
    )
    checklist_data["Completed"].append("Yes" if (has_references and references_content) else "No")

    # Check for in-text citations
    has_citations = any(
        '(' in p and ')' in p and any(str(year) in p for year in range(1900, 2025))
        for p in body_paragraphs
    )
    checklist_data["Completed"].append("Yes" if has_citations else "No")

    return checklist_data
