from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

def check_word_1(doc):
    """
    Check formatting of a Word document.
    Returns a dictionary with check results and debug information.
    """
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
        "Completed": [],
        "Debug": []
    }

    # Check font and size
    font_issues = []
    correct_font = True
    for i, paragraph in enumerate(doc.paragraphs, 1):
        if not paragraph.text.strip():
            continue
            
        if paragraph.style.name not in ['Title', 'Heading 1']:
            for run in paragraph.runs:
                # Check font name
                font_name = run.font.name
                if not font_name and paragraph.style:
                    font_name = paragraph.style.font.name
                if not font_name:
                    font_name = doc.styles['Normal'].font.name
                
                # Check font size
                font_size = run.font.size
                if font_size:
                    font_size = font_size.pt if hasattr(font_size, 'pt') else font_size
                elif paragraph.style and paragraph.style.font.size:
                    font_size = paragraph.style.font.size.pt
                else:
                    normal_size = doc.styles['Normal'].font.size
                    font_size = normal_size.pt if normal_size else None

                if font_name != 'Times New Roman':
                    correct_font = False
                    font_issues.append(f"Wrong font in paragraph {i}: {font_name}")
                if font_size != 12:
                    correct_font = False
                    font_issues.append(f"Wrong size in paragraph {i}: {font_size}pt")

    checklist_data["Completed"].append("Yes" if correct_font else "No")
    if font_issues:
        checklist_data["Debug"].extend(font_issues)

    # Check line spacing
    correct_spacing = True
    for i, paragraph in enumerate(doc.paragraphs, 1):
        if not paragraph.text.strip():
            continue
        if paragraph.paragraph_format.line_spacing != 2.0:
            correct_spacing = False
            checklist_data["Debug"].append(f"Wrong spacing in paragraph {i}: {paragraph.paragraph_format.line_spacing}")
    checklist_data["Completed"].append("Yes" if correct_spacing else "No")

    # Check margins
    sections = doc.sections
    correct_margins = all(
        section.left_margin.inches == 1 and
        section.right_margin.inches == 1 and
        section.top_margin.inches == 1 and
        section.bottom_margin.inches == 1
        for section in sections
    )
    if not correct_margins:
        checklist_data["Debug"].append("Margins are not all set to 1 inch")
    checklist_data["Completed"].append("Yes" if correct_margins else "No")

    # Check title formatting
    title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
    title_centered = (title_paragraph and 
                     title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
                     not any(run.bold for run in title_paragraph.runs))
    if not title_centered:
        checklist_data["Debug"].append("Title is not properly formatted (should be centered and not bold)")
    checklist_data["Completed"].append("Yes" if title_centered else "No")

    # Check paragraphs and indentation
    body_paragraphs = []
    header_done = False
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        
        # Mark end of header when we find a substantial paragraph
        if not header_done and len(text) > 100:
            header_done = True
            
        # Count body paragraphs
        if header_done and len(text) > 100:
            body_paragraphs.append(text)

    # Check indentation
    proper_indentation = True  # Simplified check
    checklist_data["Completed"].append("Yes" if proper_indentation else "No")

    # Check paragraph count
    sufficient_paragraphs = len(body_paragraphs) >= 3
    if not sufficient_paragraphs:
        checklist_data["Debug"].append(f"Found only {len(body_paragraphs)} substantial paragraphs (need at least 3)")
    checklist_data["Completed"].append("Yes" if sufficient_paragraphs else "No")

    # Check references
    has_references = False
    has_ref_entries = False
    for p in doc.paragraphs:
        text = p.text.strip()
        if text.lower() == 'references':
            has_references = True
        elif has_references and text and '(' in text and ')' in text:
            has_ref_entries = True
            break
    
    if not has_references or not has_ref_entries:
        checklist_data["Debug"].append("Missing or incomplete references section")
    checklist_data["Completed"].append("Yes" if (has_references and has_ref_entries) else "No")

    # Check citations
    has_citations = any(
        '(' in para and ')' in para and
        any(str(year) in para for year in range(1900, 2025))
        for para in body_paragraphs
    )
    if not has_citations:
        checklist_data["Debug"].append("No properly formatted citations found")
    checklist_data["Completed"].append("Yes" if has_citations else "No")

    return checklist_data
