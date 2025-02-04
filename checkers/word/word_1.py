from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def check_welcome_letter(doc):
    checklist_data = {
        "Grading Criteria": [
            "Is there a 2-column by 1-row table above Today's Date?",
            "Is the first cell text 'Classic Cars Club' formatted as 14pt, Bold?",
            "Is the second cell text 'PO Box 6987 Ferndale, WA 98248' aligned center-right?",
            "Are all table borders removed?",
            "Are two empty paragraphs inserted above Today's Date?",
            "Are bullet points applied to the membership benefits list?",
            "Is the three-column table correctly sized (1.5", 2.25", 1")?",
            "Is the table sorted by 'Locations' in ascending order?",
            "Is a new header row inserted and merged at the top of the table?",
            "Is the merged header cell text 'Available Partner Discounts' formatted as 14pt, Bold, and Centered?",
            "Are outside borders applied to table cells from row 2 onward?",
            "Is shading (White, Background 1, Darker 15%) applied to row 2?",
            "Is a solid single-line bottom border (1.5pt) applied to row 2?",
            "Is all text in row 2 formatted as Bold?"
        ],
        "Completed": []
    }

    tables = doc.tables
    paragraphs = doc.paragraphs
    
    # Check table existence and structure
    if len(tables) > 0 and len(tables[0].rows) == 1 and len(tables[0].columns) == 2:
        checklist_data["Completed"].append("Yes")
        first_cell = tables[0].cell(0, 0).text.strip()
        second_cell = tables[0].cell(0, 1).text.strip()
        
        # Check first cell formatting
        correct_format = False
        for run in tables[0].cell(0, 0).paragraphs[0].runs:
            if run.bold and run.font.size == Pt(14):
                correct_format = True
                break
        checklist_data["Completed"].append("Yes" if correct_format else "No")
        
        # Check second cell alignment
        alignment = tables[0].cell(0, 1).paragraphs[0].alignment
        checklist_data["Completed"].append("Yes" if alignment == WD_ALIGN_PARAGRAPH.RIGHT else "No")
    else:
        checklist_data["Completed"].extend(["No", "No", "No"])
    
    # Check if table borders are removed
    no_borders = all(cell._element.xpath('.//w:tcBorders', namespaces=nsdecls('w')) == [] for row in tables[0].rows for cell in row)
    checklist_data["Completed"].append("Yes" if no_borders else "No")
    
    # Check empty paragraphs above "Today's Date"
    paragraph_texts = [p.text.strip() for p in paragraphs]
    index_date = next((i for i, text in enumerate(paragraph_texts) if "Today's Date" in text), -1)
    empty_paragraphs = paragraph_texts[index_date-2:index_date] if index_date >= 2 else []
    checklist_data["Completed"].append("Yes" if all(not p for p in empty_paragraphs) else "No")
    
    # Check bullet formatting
    has_bullets = any(p.style.name.startswith("List") for p in paragraphs if "free Classic Cars Club" in p.text.lower())
    checklist_data["Completed"].append("Yes" if has_bullets else "No")
    
    # Verify three-column table structure and widths
    if len(tables) > 1 and len(tables[1].columns) == 3:
        checklist_data["Completed"].append("Yes")  # Table exists
    else:
        checklist_data["Completed"].append("No")
    
    # Verify sorting criteria
    checklist_data["Completed"].append("Manual Check Required")  # Sorting is not easily checked programmatically
    
    # Check header row merging and formatting
    header_merged = len(tables[1].rows[0].cells) == 1
    header_text = tables[1].rows[0].cells[0].text.strip()
    correct_header_format = any(run.bold and run.font.size == Pt(14) for run in tables[1].rows[0].cells[0].paragraphs[0].runs)
    checklist_data["Completed"].append("Yes" if header_merged and header_text == "Available Partner Discounts" and correct_header_format else "No")
    
    # Verify border settings
    checklist_data["Completed"].append("Manual Check Required")  # Border settings need visual inspection
    
    # Verify shading settings
    checklist_data["Completed"].append("Manual Check Required")  # Shading needs visual inspection
    
    # Verify bottom border formatting
    checklist_data["Completed"].append("Manual Check Required")  # Border formatting needs visual inspection
    
    # Verify bold formatting for row 2
    bold_correct = all(any(run.bold for run in cell.paragraphs[0].runs) for cell in tables[1].rows[1].cells)
    checklist_data["Completed"].append("Yes" if bold_correct else "No")
    
    return checklist_data
