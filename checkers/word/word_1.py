from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

def get_target_table(doc, expected_rows, expected_cols):
    """
    Returns the first table in the document with the specified number of rows and columns.
    """
    for table in doc.tables:
        if len(table.rows) == expected_rows and len(table.columns) == expected_cols:
            return table
    return None

def cell_borders_removed(cell):
    """
    Checks the cell's XML to see if any border is defined with a value other than 'none' or 'nil'.
    If no borders exist or all are set to 'none'/'nil', returns True.
    """
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    tcPr = cell._element.find('.//w:tcPr', ns)
    if tcPr is not None:
        # Look for a direct tcBorders element
        borders = tcPr.findall('.//w:tcBorders', ns)
        if borders:
            for border in borders:
                # Check each child border (top, left, bottom, right, etc.)
                for b in border:
                    val = b.get(qn('w:val'))
                    if val not in (None, 'none', 'nil'):
                        return False
    return True

def get_grid_span(cell):
    """
    Returns the gridSpan value (an integer) for a cell if it is merged horizontally.
    Defaults to 1 if gridSpan is not set.
    """
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    tcPr = cell._element.find('.//w:tcPr', ns)
    if tcPr is not None:
        gridSpan = tcPr.find('.//w:gridSpan', ns)
        if gridSpan is not None:
            val = gridSpan.get(qn('w:val'))
            try:
                return int(val)
            except (TypeError, ValueError):
                return 1
    return 1

def check_word_1(doc):
    checklist_data = {
        "Grading Criteria": [
            "Is there a 2-column by 1-row table above Today's Date?",
            "Is the first cell text 'Classic Cars Club' formatted as 14pt, Bold?",
            "Is the second cell text 'PO Box 6987 Ferndale, WA 98248' aligned center-right?",
            "Are all table borders removed?",
            "Are two empty paragraphs inserted above Today's Date?",
            "Are bullet points applied to the membership benefits list?",
            "Is the three-column table correctly sized (1.5\", 2.25\", 1\")?",
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

    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Find the 2-column by 1-row table (assumed to be above "Today's Date")
    table_2col = get_target_table(doc, expected_rows=1, expected_cols=2)
    checklist_data["Completed"].append("Yes" if table_2col else "No")
    
    if table_2col:
        # Check first cell: text must equal 'Classic Cars Club' and all runs (with text) must be Bold and 14pt.
        cell1 = table_2col.cell(0, 0)
        cell1_text = cell1.text.strip()
        if cell1_text == "Classic Cars Club":
            if all(run.bold and run.font.size == Pt(14)
                   for para in cell1.paragraphs
                   for run in para.runs if run.text.strip()):
                checklist_data["Completed"].append("Yes")
            else:
                checklist_data["Completed"].append("No")
        else:
            checklist_data["Completed"].append("No")
        
        # Check second cell: text must equal 'PO Box 6987 Ferndale, WA 98248' and its first paragraph must be right aligned.
        cell2 = table_2col.cell(0, 1)
        cell2_text = cell2.text.strip()
        if cell2_text == "PO Box 6987 Ferndale, WA 98248":
            if cell2.paragraphs and cell2.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                checklist_data["Completed"].append("Yes")
            else:
                checklist_data["Completed"].append("No")
        else:
            checklist_data["Completed"].append("No")
        
        # Check if all borders are effectively removed in table_2col.
        borders_ok = True
        for row in table_2col.rows:
            for cell in row.cells:
                if not cell_borders_removed(cell):
                    borders_ok = False
                    break
            if not borders_ok:
                break
        checklist_data["Completed"].append("Yes" if borders_ok else "No")
    else:
        checklist_data["Completed"].extend(["No", "No", "No"])
    
    # Check for two empty paragraphs immediately above "Today's Date"
    paragraph_texts = [p.text.strip() for p in doc.paragraphs]
    index_date = next((i for i, text in enumerate(paragraph_texts) if "Today's Date" in text), -1)
    if index_date == -1 or index_date < 2:
        checklist_data["Completed"].append("No")
    else:
        empty_above = paragraph_texts[index_date-2:index_date]
        checklist_data["Completed"].append("Yes" if all(not p for p in empty_above) else "No")
    
    # Detect bullet points by checking for numbering properties in paragraphs.
    bullet_found = False
    for p in doc.paragraphs:
        if p._p.find('.//w:numPr', ns) is not None and p.text.strip():
            bullet_found = True
            break
    checklist_data["Completed"].append("Yes" if bullet_found else "No")
    
    # Find the three-column table (assumed to have at least one header row and one data row)
    table_3col = None
    for table in doc.tables:
        if len(table.columns) == 3 and len(table.rows) >= 2:
            table_3col = table
            break
    checklist_data["Completed"].append("Yes" if table_3col else "No")
    
    # Sorting criteria requires manual verification.
    checklist_data["Completed"].append("Manual Check Required")
    
    # Verify the header row in the three-column table: the first cell should span all columns,
    # contain the text 'Available Partner Discounts', and its first paragraph should be centered,
    # Bold, and 14pt.
    if table_3col and len(table_3col.rows) > 0:
        header_row = table_3col.rows[0]
        first_cell = header_row.cells[0]
        span = get_grid_span(first_cell)
        if span == len(table_3col.columns):
            header_text = first_cell.text.strip()
            if header_text == "Available Partner Discounts":
                header_para = first_cell.paragraphs[0] if first_cell.paragraphs else None
                if header_para and header_para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    if all(run.bold and run.font.size == Pt(14)
                           for run in header_para.runs if run.text.strip()):
                        checklist_data["Completed"].append("Yes")
                    else:
                        checklist_data["Completed"].append("No")
                else:
                    checklist_data["Completed"].append("No")
            else:
                checklist_data["Completed"].append("No")
        else:
            checklist_data["Completed"].append("No")
    else:
        checklist_data["Completed"].append("No")
    
    # Append manual check markers if any criteria remain unchecked.
    while len(checklist_data["Completed"]) < len(checklist_data["Grading Criteria"]):
        checklist_data["Completed"].append("Manual Check Required")
    
    return checklist_data

# Example usage:
# doc = Document("path_to_document.docx")
# results = check_word_1(doc)
# print(results)
