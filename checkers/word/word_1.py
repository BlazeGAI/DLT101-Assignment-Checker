from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

def check_word_1(doc):
    checklist_data = {
        "Grading Criteria": [
            "I the font Times New Roman, 12pt?",
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

    # Check font
    correct_font = True
    for paragraph in doc.paragraphs:
        # Skip checking font for Title or Heading 1 styles
        if paragraph.style.name not in ['Title', 'Heading 1']:
            for run in paragraph.runs:
                if run.font.name != 'Times New Roman' or run.font.size != Pt(12):
                    correct_font = False
                    break  # Exit loop if a mismatch is found
        if not correct_font:
            break  # Exit outer loop if a mismatch is found
    
    checklist_data["Completed"].append("Yes" if correct_font else "No")


    # Check line spacing
    correct_spacing = all(
        paragraph.paragraph_format.line_spacing == 2.0
        for paragraph in doc.paragraphs
        if paragraph.text.strip()
    )
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
    checklist_data["Completed"].append("Yes" if correct_margins else "No")

    # Check title centering
    title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
    title_centered = (title_paragraph and 
                     title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
                     not any(run.bold for run in title_paragraph.runs))
    checklist_data["Completed"].append("Yes" if title_centered else "No")

    # Simpler paragraph counting
    body_paragraphs = []
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        
        # Skip empty paragraphs
        if not text:
            continue
            
        # Skip if it's the title or header info (first few paragraphs)
        if i < 3:
            continue
            
        # Skip if it's references or conclusion
        if text.lower().startswith(('references', 'works cited', 'in conclusion')):
            continue
            
        # Count substantial paragraphs
        if len(text) > 50:
            body_paragraphs.append(text)
            print(f"Found paragraph: {text[:50]}...")  # Debug print

    print(f"Total body paragraphs found: {len(body_paragraphs)}")  # Debug print
    sufficient_paragraphs = len(body_paragraphs) >= 3
    checklist_data["Completed"].append("Yes" if sufficient_paragraphs else "No")

    # Check paragraph indentation
    proper_indentation = all(
        p.paragraph_format.first_line_indent is not None and
        p.paragraph_format.first_line_indent >= Pt(36)  # 0.5 inch in points
        for p in doc.paragraphs
        if len(p.text.strip()) > 0 and p != doc.paragraphs[0]  # Skip title
    )
    checklist_data["Completed"].append("Yes" if proper_indentation else "No")

    # Check for Works Cited
    has_works_cited = any(
        "References" in p.text.lower()
        for p in doc.paragraphs
    )
    checklist_data["Completed"].append("Yes" if has_works_cited else "No")

    # Check for in-text citations (basic check for parentheses patterns)
    has_citations = any(
        '(' in p.text and ')' in p.text
        for p in doc.paragraphs
    )
    checklist_data["Completed"].append("Yes" if has_citations else "No")

    return checklist_data
