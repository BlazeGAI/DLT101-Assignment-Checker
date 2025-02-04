from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT

def check_word_1(doc):
    """
    Checks a Word document for the instructions shown in the Classic Cars Club
    'Enhancing a Welcome Letter' exercise. It looks for:

    1) A 2-column by 1-row table at the top with:
       - "Classic Cars Club" in cell 1 (14 pt, Bold)
       - "PO Box 6987 Ferndale, WA 98248" in cell 2
       - Cell 1 center-left alignment, Cell 2 center-right alignment
       - All borders removed

    2) Two empty paragraphs above "Today's Date"

    3) A three-column table (centered) with column widths:
       - Column 1: 1.5"
       - Column 2: 2.25"
       - Column 3: 1"

    4) Bulleted membership perks paragraphs

    5) Sorting by 'Locations' in ascending order with a header row

    6) A new row at the top of the table with merged cells containing:
       - "Available Partner Discounts" in 14 pt, Bold, center-aligned

    7) From row 2 onward, a solid single line (½ pt) “Outside Borders”

    8) White, Background 1, Darker 15% shading in row 2

    9) A bottom border (1½ pt, solid) on row 2

    10) Bold font formatting for text in row 2
    """

    checklist_data = {
        "Grading Criteria": [
            "Top table (2-col x 1-row) found",
            "Correct text in top table cells",
            "Cell 1 text is 14 pt Bold and center-left aligned",
            "Cell 2 text is center-right aligned",
            "All borders removed from top table",
            "Two empty paragraphs above 'Today's Date'",
            "Three-column main table with correct widths (1.5, 2.25, 1.0) centered",
            "Bulleted membership perks paragraphs present",
            "Main table sorted by 'Locations' in ascending order (header row)",
            "New top row merged and labeled 'Available Partner Discounts' (14 pt, Bold, centered)",
            "From row 2 onward: outside borders single line (½ pt)",
            "Row 2 shading: White, Background 1, Darker 15%",
            "Row 2 bottom border: solid 1½ pt",
            "Row 2 text is Bold"
        ],
        "Completed": []
    }

    # 1) Check the top table (2-column x 1-row) and its properties
    top_table_ok = False
    top_table_text_ok = False
    cell1_format_ok = False
    cell2_align_ok = False
    top_table_no_borders = False

    if doc.tables:
        # We assume the first table is the top table
        top_table = doc.tables[0]
        # Check rows/cols
        if len(top_table.rows) == 1 and len(top_table.columns) == 2:
            top_table_ok = True
            
            # Check text in cells
            cell1_text = top_table.cell(0, 0).text.strip()
            cell2_text = top_table.cell(0, 1).text.strip()
            if cell1_text == "Classic Cars Club" and "Ferndale" in cell2_text:
                top_table_text_ok = True

            # Check formatting of cell1 text (14 pt, Bold, center-left)
            cell1_run_format_is_14_bold = False
            for p in top_table.cell(0, 0).paragraphs:
                for r in p.runs:
                    size_ok = (r.font.size == Pt(14))
                    bold_ok = (r.bold is True)
                    if size_ok and bold_ok:
                        cell1_run_format_is_14_bold = True
                        break
                if cell1_run_format_is_14_bold:
                    break
            # "Center-left" typically means left alignment (horizontal), so we check alignment == LEFT
            if cell1_run_format_is_14_bold and top_table.cell(0, 0).paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.LEFT:
                cell1_format_ok = True

            # Check cell2 alignment for "center-right" effect (usually right alignment)
            if top_table.cell(0, 1).paragraphs and \
               top_table.cell(0, 1).paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                cell2_align_ok = True

            # Check if the table has borders removed
            no_border_detected = True
            for row in top_table.rows:
                for cell in row.cells:
                    tc_pr = cell._tc.get_or_add_tcPr()
                    for brd_tag in ['w:top', 'w:left', 'w:bottom', 'w:right']:
                        brd = tc_pr.find(qn('w:tcBorders'))
                        if brd is not None and brd.find(qn('w:' + brd_tag)) is not None:
                            no_border_detected = False
                            break
                    if not no_border_detected:
                        break
                if not no_border_detected:
                    break
            if no_border_detected:
                top_table_no_borders = True

    checklist_data["Completed"].append("Yes" if top_table_ok else "No")
    checklist_data["Completed"].append("Yes" if top_table_text_ok else "No")
    checklist_data["Completed"].append("Yes" if cell1_format_ok else "No")
    checklist_data["Completed"].append("Yes" if cell2_align_ok else "No")
    checklist_data["Completed"].append("Yes" if top_table_no_borders else "No")

    # 2) Check for two empty paragraphs above "Today's Date"
    two_paragraphs_ok = False
    idx_of_date = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "Today's Date":
            idx_of_date = i
            break
    if idx_of_date is not None and idx_of_date >= 2:
        # Check if the two paragraphs above are empty
        if not doc.paragraphs[idx_of_date - 1].text.strip() and not doc.paragraphs[idx_of_date - 2].text.strip():
            two_paragraphs_ok = True

    checklist_data["Completed"].append("Yes" if two_paragraphs_ok else "No")

    # 3) Check for a three-column table with correct widths, centered
    main_table_ok = False
    correct_widths_ok = False
    table_centered_ok = False

    if len(doc.tables) > 1:
        main_table = doc.tables[1]
        if len(main_table.columns) == 3:
            main_table_ok = True

            # Check column widths (approximate match)
            col_widths = [col.width.inches if col.width else 0 for col in main_table.columns]
            desired = [1.5, 2.25, 1.0]
            tolerance = 0.05
            width_matches = [
                abs(col_widths[i] - desired[i]) < tolerance
                for i in range(len(desired))
            ]
            correct_widths_ok = all(width_matches)

            # Check if table is centered
            if main_table.alignment == WD_TABLE_ALIGNMENT.CENTER:
                table_centered_ok = True

    checklist_data["Completed"].append("Yes" if main_table_ok else "No")
    checklist_data["Completed"].append("Yes" if correct_widths_ok else "No")
    checklist_data["Completed"].append("Yes" if table_centered_ok else "No")

    # 4) Check for bulleted membership perks paragraphs
    bullet_points_found = False
    bullet_keywords = [
        "Free entry to local and regional shows",
        "30% entry discount",
        "25% discount on merchandise",
        "A free Classic Cars Club plaque",
        "A free Classic Cars Club license plate frame"
    ]
    bullet_hits = 0
    for p in doc.paragraphs:
        text_low = p.text.lower()
        if any(k.lower() in text_low for k in bullet_keywords):
            # If style or numbering suggests a bullet
            if p.style and "list" in p.style.name.lower():
                bullet_hits += 1
    if bullet_hits >= 3:
        bullet_points_found = True

    checklist_data["Completed"].append("Yes" if bullet_points_found else "No")

    # 5) Check table sorting by 'Locations' in ascending order with header row
    sorting_ok = False
    if len(doc.tables) > 1:
        table = doc.tables[1]
        header_cells = table.rows[0].cells
        locations_col_idx = None
        for i, cell in enumerate(header_cells):
            if "Locations" in cell.text:
                locations_col_idx = i
                break
        if locations_col_idx is not None:
            # Collect the column entries (excluding the header row)
            col_texts = [row.cells[locations_col_idx].text.strip() for row in table.rows[1:]]
            if col_texts == sorted(col_texts):
                sorting_ok = True

    checklist_data["Completed"].append("Yes" if sorting_ok else "No")

    # 6) New row at the top of the main table, merged, "Available Partner Discounts" 14 pt Bold, center
    top_row_merged_ok = False
    top_row_text_format_ok = False
    if len(doc.tables) > 1:
        table = doc.tables[1]
        if len(table.rows) > 1:
            first_row = table.rows[0]
            # Check if the first row is merged into 1 cell
            if len(first_row.cells) == 1:
                top_row_merged_ok = True
                merged_text = first_row.cells[0].text.strip()
                if merged_text == "Available Partner Discounts":
                    row_paras = first_row.cells[0].paragraphs
                    if row_paras:
                        run_14_bold = False
                        center_align = (row_paras[0].alignment == WD_ALIGN_PARAGRAPH.CENTER)
                        for r in row_paras[0].runs:
                            if r.bold and r.font.size == Pt(14):
                                run_14_bold = True
                                break
                        if run_14_bold and center_align:
                            top_row_text_format_ok = True

    checklist_data["Completed"].append("Yes" if top_row_merged_ok else "No")
    checklist_data["Completed"].append("Yes" if top_row_text_format_ok else "No")

    # 7) From row 2 onward: outside borders single line (½ pt)
    outside_borders_ok = False
    if len(doc.tables) > 1:
        table = doc.tables[1]
        border_good_count = 0
        total_cells = 0
        for r_idx in range(1, len(table.rows)):
            row = table.rows[r_idx]
            for cell in row.cells:
                total_cells += 1
                tc_pr = cell._tc.get_or_add_tcPr()
                borders = tc_pr.find(qn('w:tcBorders'))
                if borders is not None:
                    # Look for single line, ½ pt (sz="8") on any side
                    good_sides = 0
                    for side_tag in ['w:top', 'w:left', 'w:bottom', 'w:right']:
                        side = borders.find(qn('w:' + side_tag))
                        if side is not None:
                            if side.get(qn('w:val')) == 'single' and side.get(qn('w:sz')) == '8':
                                good_sides += 1
                    if good_sides > 0:
                        border_good_count += 1
        if total_cells > 0 and border_good_count == total_cells:
            outside_borders_ok = True

    checklist_data["Completed"].append("Yes" if outside_borders_ok else "No")

    # 8) Row 2 shading, 9) Row 2 bottom border 1½ pt, 10) Bold text in row 2
    shading_ok = False
    bottom_border_ok = False
    row2_bold_ok = False
    if len(doc.tables) > 1:
        table = doc.tables[1]
        if len(table.rows) > 1:
            second_row = table.rows[1]

            # Check shading (D9D9D9 is common for "White, Background 1, Darker 15%")
            shading_cells = 0
            for cell in second_row.cells:
                tc_pr = cell._tc.get_or_add_tcPr()
                shd = tc_pr.find(qn('w:shd'))
                if shd is not None:
                    fill_val = shd.get(qn('w:fill'))
                    if fill_val and fill_val.lower() == "d9d9d9":
                        shading_cells += 1
            if shading_cells == len(second_row.cells):
                shading_ok = True

            # Bottom border 1½ pt (w:sz="24")
            border_count = 0
            for cell in second_row.cells:
                tc_pr = cell._tc.get_or_add_tcPr()
                borders = tc_pr.find(qn('w:tcBorders'))
                if borders is not None:
                    bottom_side = borders.find(qn('w:bottom'))
                    if bottom_side is not None:
                        if bottom_side.get(qn('w:val')) == 'single' and bottom_side.get(qn('w:sz')) == '24':
                            border_count += 1
            if border_count == len(second_row.cells):
                bottom_border_ok = True

            # Bold text in row 2
            bold_count = 0
            total_runs = 0
            for cell in second_row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        total_runs += 1
                        if r.bold:
                            bold_count += 1
            if total_runs > 0 and bold_count == total_runs:
                row2_bold_ok = True

    checklist_data["Completed"].append("Yes" if shading_ok else "No")
    checklist_data["Completed"].append("Yes" if bottom_border_ok else "No")
    checklist_data["Completed"].append("Yes" if row2_bold_ok else "No")

    return checklist_data
