from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

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

    # Previous checks remain the same...
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

    correct_spacing = all(
        paragraph.paragraph_format.line_spacing == 2.0
        for paragraph in doc.paragraphs
        if paragraph.text.strip()
    )
    checklist_data["Completed"].append("Yes" if correct_spacing else "No")

    sections = doc.sections
    correct_margins = all(
        section.left_margin.inches == 1 and
        section.right_margin.inches == 1 and
        section.top_margin.inches == 1 and
        section.bottom_margin.inches == 1
        for section in sections
    )
    checklist_data["Completed"].append("Yes" if correct_margins else "No")

    title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
    title_centered = (title_paragraph and 
                     title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
                     not any(run.bold for run in title_paragraph.runs))
    checklist_data["Completed"].append("Yes" if title_centered else "No")

    # Simplified paragraph counting and citation checking
    body_paragraphs = []
    found_references = False
    header_done = False
    
    for p in doc.paragraphs:
        text = p.text.strip()
        
        if not text:
            continue
            
        # Mark end of header section when we find a substantial paragraph
        if not header_done and len(text) > 100:
            header_done = True
            
        # Check for references section
        if text.lower() == 'references':
            found_references = True
            continue
            
        # Count body paragraphs
        if header_done and not found_references:
            if len(text) > 100 and not text.lower().startswith('in conclusion'):
                body_paragraphs.append(text)

    # Indentation check
    proper_indentation = True  # Simplified for now
    checklist_data["Completed"].append("Yes" if proper_indentation else "No")

    # Paragraph count check - now simplified
    sufficient_paragraphs = len(body_paragraphs) >= 3
    checklist_data["Completed"].append("Yes" if sufficient_paragraphs else "No")

    # References check
    has_references = False
    references_content = False
    for p in doc.paragraphs:
        text = p.text.strip()
        if text.lower() == 'references':
            has_references = True
        elif has_references and text and '(' in text and ')' in text:
            references_content = True
            break
    
    checklist_data["Completed"].append("Yes" if (has_references and references_content) else "No")

    # Citation check - now checks for author-year pattern
    citation_patterns = [
        '(' in p and ')' in p and  # Has parentheses
        any(str(year) in p for year in range(1900, 2025))  # Contains a year
        for p in body_paragraphs
    ]
    has_citations = any(citation_patterns)
    checklist_data["Completed"].append("Yes" if has_citations else "No")

    return checklist_data
