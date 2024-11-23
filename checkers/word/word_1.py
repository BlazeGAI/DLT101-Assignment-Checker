from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

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

    # Font check remains the same
    correct_font = True
    for paragraph in doc.paragraphs:
        if paragraph.style.name not in ['Title', 'Heading 1']:
            for run in paragraph.runs:
                if run.font.name != 'Times New Roman' or run.font.size != Pt(12):
                    correct_font = False
                    break
        if not correct_font:
            break
    checklist_data["Completed"].append("Yes" if correct_font else "No")

    # Line spacing check remains the same
    correct_spacing = all(
        paragraph.paragraph_format.line_spacing == 2.0
        for paragraph in doc.paragraphs
        if paragraph.text.strip()
    )
    checklist_data["Completed"].append("Yes" if correct_spacing else "No")

    # Margins check remains the same
    sections = doc.sections
    correct_margins = all(
        section.left_margin.inches == 1 and
        section.right_margin.inches == 1 and
        section.top_margin.inches == 1 and
        section.bottom_margin.inches == 1
        for section in sections
    )
    checklist_data["Completed"].append("Yes" if correct_margins else "No")

    # Title check remains the same
    title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
    title_centered = (title_paragraph and 
                     title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
                     not any(run.bold for run in title_paragraph.runs))
    checklist_data["Completed"].append("Yes" if title_centered else "No")

    # Improved paragraph and indentation checking
    body_paragraphs = []
    in_references = False
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        
        # Skip empty paragraphs
        if not text:
            continue
            
        # Check for references section
        if text.lower() == 'references':
            in_references = True
            continue
            
        # Skip references section entries
        if in_references:
            continue
            
        # Skip header information (first few lines)
        if i < 5:  # Assuming first 5 lines are header
            continue
            
        # Skip conclusion
        if text.lower().startswith('in conclusion'):
            continue
            
        # Count as body paragraph if it's substantial (more than 50 characters)
        if len(text) > 50:
            body_paragraphs.append(p)

    # Check for proper indentation
    proper_indentation = all(
        (p.paragraph_format.first_line_indent is not None and
         p.paragraph_format.first_line_indent >= Pt(36))  # 0.5 inch in points
        for p in body_paragraphs
    )
    checklist_data["Completed"].append("Yes" if proper_indentation else "No")

    # Check number of paragraphs
    sufficient_paragraphs = len(body_paragraphs) >= 3
    checklist_data["Completed"].append("Yes" if sufficient_paragraphs else "No")

    # Improved references page check
    has_references = False
    references_content = False
    for p in doc.paragraphs:
        if p.text.strip().lower() == 'references':
            has_references = True
        # Check if there's actually content in the references section
        elif has_references and p.text.strip():
            references_content = True
            break
    checklist_data["Completed"].append("Yes" if (has_references and references_content) else "No")

    # Improved citation check
    has_citations = any(
        ('(' in p.text and ')' in p.text and 
         any(str(year) in p.text for year in range(1900, 2025)))  # Look for years in parentheses
        for p in body_paragraphs
    )
    checklist_data["Completed"].append("Yes" if has_citations else "No")

    return checklist_data
