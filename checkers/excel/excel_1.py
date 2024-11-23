from openpyxl.styles import Alignment, Font, PatternFill

def check_excel_1(workbook):
    sheet = workbook.active
    checklist_data = {
     "Grading Criteria": [
         "Are there 6 columns A-F?",
         "Are there exactly 10 rows of data in the dataset?",
         "Are the first 5 column headers named ID, First Name, Last Name, Date of Birth, Hometown?",
         "Does the last column have a meaningful header and consistent data?",
         "Are the columns banded color?",
         "Are the headers in row 1 bolded?",
         "Are all borders and a thick outside border applied to the table?",
         "Is the ChatGPT hyperlink centered in cell A13?",
         "Are cells A13:F13 merged in row 13?",
         "Does cell A13 have a background color?"
            ],
    "Completed": []
        }

    # Check number of columns (should be 6)
    num_columns = sheet.max_column
    checklist_data["Completed"].append("Yes" if num_columns == 6 else "No")

    # Check number of rows of data (should be 10)
    data_rows = sum(1 for row in range(2, 12) if any(sheet.cell(row=row, column=col).value for col in range(1, 7)))
    checklist_data["Completed"].append("Yes" if data_rows == 10 else "No")

    # Check first 5 column headers
    expected_headers = ["ID", "First Name", "Last Name", "Date of Birth", "Hometown"]
    headers = [sheet.cell(row=1, column=i).value for i in range(1, 6)]
    headers_match = all(a == b for a, b in zip(headers, expected_headers))
    checklist_data["Completed"].append("Yes" if headers_match else "No")

    # Check if last column has a header and data
    last_col_header = sheet.cell(row=1, column=6).value
    last_col_has_data = any(sheet.cell(row=i, column=6).value for i in range(2, 12))
    checklist_data["Completed"].append("Yes" if last_col_header and last_col_has_data else "No")

    # Check for banded colors
    has_banded_colors = False
    for row in range(2, 12):  # Check rows 2-11
        if row % 2 == 0:
            if (sheet.cell(row=row, column=1).fill.start_color.index != 
                sheet.cell(row=row-1, column=1).fill.start_color.index):
                has_banded_colors = True
                break
    checklist_data["Completed"].append("Yes" if has_banded_colors else "No")

    # Check if headers are bold
    headers_bold = all(
        sheet.cell(row=1, column=col).font.bold
        for col in range(1, 7)
    )
    checklist_data["Completed"].append("Yes" if headers_bold else "No")

    # Check borders
    all_borders_applied = all(
        sheet.cell(row=row, column=col).border is not None
        for row in range(1, 12)
        for col in range(1, 7)
    )
    checklist_data["Completed"].append("Yes" if all_borders_applied else "No")

    # Check ChatGPT link alignment in A13
    center_aligned = sheet['A13'].alignment.horizontal == 'center'
    checklist_data["Completed"].append("Yes" if center_aligned else "No")

    # Check merged cells A13:F13
    merged_in_row_13 = any("A13:F13" in str(range) for range in sheet.merged_cells.ranges)
    checklist_data["Completed"].append("Yes" if merged_in_row_13 else "No")

    # Check background color in A13
    background_color_present = (sheet['A13'].fill is not None and 
                              sheet['A13'].fill.fill_type is not None)
    checklist_data["Completed"].append("Yes" if background_color_present else "No")

    return checklist_data
