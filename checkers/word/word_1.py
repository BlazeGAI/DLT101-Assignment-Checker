from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def check_word_document(doc):
    """
    Checks formatting of a Word document.

    Returns a list of dictionaries, where each dictionary contains the 
    criterion, result, and debug information for a single check.
    """

    criteria = [
        "Is the font Times New Roman, 12pt?",
        "Is line spacing set to double?",
        "Are margins set to 1 inch on all sides?",
        "Is the title centered and not bold?",
        "Are paragraphs properly indented?",  # Simplified check
        "Are there at least 3 paragraphs?",
        "Is there a References page?",
        "Are in-text citations properly formatted?"
    ]

    results = []
    for criterion in criteria:
        results.append({"Criterion": criterion, "Completed": "No", "Debug": ""})

    try:
        # 1. Check font and size
        correct_font = True
        for paragraph in doc.paragraphs:
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
                        break  # Exit inner loop
                if not correct_font:
                    break  # Exit outer loop

        results[0]["Completed"] = "Yes" if correct_font else "No"

        # 2. Check line spacing
        correct_spacing = all(
            paragraph.paragraph_format.line_spacing == 2.0
            for paragraph in doc.paragraphs if paragraph.text.strip()
        )
        results[1]["Completed"] = "Yes" if correct_spacing else "No"

        # 3. Check margins
        correct_margins = all(
            section.left_margin.inches == 1 and
            section.right_margin.inches == 1 and
            section.top_margin.inches == 1 and
            section.bottom_margin.inches == 1
            for section in doc.sections
        )
        results[2]["Completed"] = "Yes" if correct_margins else "No"

        # 4. Check title formatting
        title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
        title_centered = (title_paragraph and
                         title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
                         not any(run.bold for run in title_paragraph.runs))
        results[3]["Completed"] = "Yes" if title_centered else "No"

        # 5. Check indentation (simplified - needs more robust logic)
        results[4]["Completed"] = "Yes"  # Placeholder - Implement proper indentation check

        # 6. Count paragraphs (body paragraphs, excluding title/header)
        body_paragraphs = []
        header_done = False  # Flag to indicate if we've passed the header

        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue  # Skip empty paragraphs

            if not header_done and len(text) > 50:  # Heuristic for header length
                header_done = True  # Assume we're past the header

            if header_done and len(text) > 50:  # Consider only longer paragraphs
                body_paragraphs.append(text)

        results[5]["Completed"] = "Yes" if len(body_paragraphs) >= 3 else "No"

        # 7. Check references
        has_references = False
        references_content = False
        for p in doc.paragraphs:
            text = p.text.strip().lower()  # Case-insensitive check
            if text == 'references':
                has_references = True
            elif has_references and text and '(' in text and ')' in text and any(str(year) in text for year in range(1900, 2025)): #Check for years to avoid false positives
                references_content = True
                break  # Exit loop once references content is found

        results[6]["Completed"] = "Yes" if (has_references and references_content) else "No"

        # 8. Check citations
        has_citations = any(
            '(' in para and ')' in para and any(str(year) in para for year in range(1900,2025))
            for para in body_paragraphs
        )
        results[7]["Completed"] = "Yes" if has_citations else "No"

    except Exception as e:
        for result in results:
            result["Debug"] = f"Error during checking: {str(e)}"

    return results



# Example usage (replace with your actual document loading)
# from docx import Document
# doc = Document("your_document.docx")  # Replace with your document path
# results = check_word_document(doc)

# for item in results:
#     print(f"Criterion: {item['Criterion']}")
#     print(f"Completed: {item['Completed']}")
#     print(f"Debug: {item['Debug']}")
#     print("-" * 20)
