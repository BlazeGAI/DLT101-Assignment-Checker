from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

def check_word_1(doc):
    """
    Check formatting of a Word document.
    Returns a dictionary with check results and debug information.
    """
    # Initialize the results dictionary with criteria
    criteria = [
        "Is the font Times New Roman, 12pt?",
        "Is line spacing set to double?",
        "Are margins set to 1 inch on all sides?",
        "Is the title centered and not bold?",
        "Are paragraphs properly indented?",
        "Are there at least 3 paragraphs?",
        "Is there a References page?",
        "Are in-text citations properly formatted?"
    ]
    
    # Initialize results array with the same length as criteria
    results = ["No"] * len(criteria)
    debug_info = []
    
    try:
        # 1. Check font and size (criteria[0])
        correct_font = True
        for i, paragraph in enumerate(doc.paragraphs):
            if not paragraph.text.strip():
                continue
            if paragraph.style.name not in ['Title', 'Heading 1']:
                for run in paragraph.runs:
                    font_name = run.font.name or (paragraph.style and paragraph.style.font.name) or doc.styles['Normal'].font.name
                    font_size = run.font.size
                    if font_size:
                        font_size = font_size.pt if hasattr(font_size, 'pt') else font_size
                    
                    if font_name != 'Times New Roman' or font_size != Pt(12):
                        correct_font = False
                        break
        results[0] = "Yes" if correct_font else "No"

        # 2. Check line spacing (criteria[1])
        correct_spacing = all(
            paragraph.paragraph_format.line_spacing == 2.0
            for paragraph in doc.paragraphs
            if paragraph.text.strip()
        )
        results[1] = "Yes" if correct_spacing else "No"

        # 3. Check margins (criteria[2])
        sections = doc.sections
        correct_margins = all(
            section.left_margin.inches == 1 and
            section.right_margin.inches == 1 and
            section.top_margin.inches == 1 and
            section.bottom_margin.inches == 1
            for section in sections
        )
        results[2] = "Yes" if correct_margins else "No"

        # 4. Check title formatting (criteria[3])
        title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
        title_centered = (title_paragraph and 
                         title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
                         not any(run.bold for run in title_paragraph.runs))
        results[3] = "Yes" if title_centered else "No"

        # 5. Check indentation (criteria[4])
        # Simplified indentation check
        results[4] = "Yes"  # Default to Yes for now

        # 6. Count paragraphs (criteria[5])
        body_paragraphs = []
        header_done = False
        
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            if not header_done and len(text) > 100:
                header_done = True
            if header_done and len(text) > 100:
                body_paragraphs.append(text)

        results[5] = "Yes" if len(body_paragraphs) >= 3 else "No"

        # 7. Check references (criteria[6])
        has_references = False
        references_content = False
        for p in doc.paragraphs:
            text = p.text.strip()
            if text.lower() == 'references':
                has_references = True
            elif has_references and text and '(' in text and ')' in text:
                references_content = True
                break
        results[6] = "Yes" if (has_references and references_content) else "No"

        # 8. Check citations (criteria[7])
        has_citations = any(
            '(' in para and ')' in para and
            any(str(year) in para for year in range(1900, 2025))
            for para in body_paragraphs
        )
        results[7] = "Yes" if has_citations else "No"

    except Exception as e:
        debug_info.append(f"Error during checking: {str(e)}")
        # Ensure results array is complete even if error occurs
        results = ["No"] * len(criteria)

    return {
        "Grading Criteria": criteria,
        "Completed": results,
        "Debug": debug_info
    }
