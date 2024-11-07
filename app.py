import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, PatternFill
import re

st.title("Excel Assignment Checker")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file with openpyxl
        workbook = load_workbook(uploaded_file)
        sheet_names = workbook.sheetnames
        sheet = workbook["Alumni"] if "Alumni" in workbook.sheetnames else None
        
        # Initialize checklist data
        checklist_data = {
            "Grading Criteria": [
                "Is 'Alumni' sheet the first sheet?",
                "Do columns match the expected names and order?",
                "Does cell A1 contain 'ID' in the header?",
                "Does ID column contain unique numerical identifiers starting from 1001?",
                "Is Graduation Year calculation formula in column G?",
                "Is Income Earned calculation formula in column I?",
                "Is Accounting format applied to Income Earned (I2:I32) with no decimals?",
                "Are columns rearranged in the correct order?",
                "Are rows styled differently based on Experience?",
                "Is the table sorted by Salary?",
                "Are columns A, F, G, and H center-aligned?",
                "Is total 'Salary' in cell H33 and set to bold?",
                "Is average 'Salary' in H34 and set to bold?",
                "Is total 'Income Earned' in I33 and set to bold?",
                "Is average 'Income Earned' in I34 and set to bold?",
                "Are headers in row 1 bolded?",
                "Are all borders and thick outside border applied?",
                "Is ChatGPT link present in cell A35 and centered?",
                "Is row 35 filled with a background color?"
            ],
            "Completed": []
        }
        
        # 1. Check if "Alumni" is the first sheet
        checklist_data["Completed"].append("Yes" if sheet and sheet_names[0] == "Alumni" else "No")

        # 2. Check if columns match the expected names and order
        expected_columns = [
            "ID", "First Name", "Last Name", "Bachelor's Degree",
            "Current Profession", "Graduation Year", "Experience",
            "Salary", "Income Earned"
        ]
        if sheet:
            actual_columns = [cell.value for cell in sheet[1][:len(expected_columns)]]
            columns_match = actual_columns == expected_columns
            checklist_data["Completed"].append("Yes" if columns_match else "No")
        else:
            checklist_data["Completed"].append("No")

        # 3. Check if cell A1 contains "ID" in the header
        a1_value = sheet["A1"].value if sheet else ""
        contains_id_in_header = "ID" in str(a1_value).upper()
        checklist_data["Completed"].append("Yes" if contains_id_in_header else "No")

        # 4. Verify ID column contains unique numbers starting from 1001
        id_column_valid = False
        if sheet:
            id_values = [
                sheet.cell(row=row, column=1).value for row in range(2, 32)
            ]
            numeric_ids = pd.to_numeric(id_values, errors='coerce').dropna()
            id_column_valid = numeric_ids.is_unique and (numeric_ids >= 1001).all()
        checklist_data["Completed"].append("Yes" if id_column_valid else "No")

        # 5. Check if Graduation Year calculation formula is in column G
        grad_year_formula = all(
            sheet.cell(row=row, column=7).data_type == 'f'
            for row in range(2, 32)
        )
        checklist_data["Completed"].append("Yes" if grad_year_formula else "No")

        # 6. Check if Income Earned calculation formula is in column I
        income_earned_formula = all(
            sheet.cell(row=row, column=9).data_type == 'f'
            for row in range(2, 32)
        )
        checklist_data["Completed"].append("Yes" if income_earned_formula else "No")

        # 7. Check Accounting format with no decimals in Income Earned (I2:I32)
        accounting_format = all(
            sheet.cell(row=row, column=9).number_format in ["$#,##0", "$#,##0;[Red]$-#,##0", "Accounting"]
            for row in range(2, 32)
        )
        checklist_data["Completed"].append("Yes" if accounting_format else "No")

        # 8. Check if columns are rearranged in the correct order (already covered in expected columns check)
        checklist_data["Completed"].append("Yes" if columns_match else "No")

        # 9. Check if rows are styled differently based on Experience
        different_styles = any(
            sheet.cell(row=row, column=7).fill != sheet.cell(row=row+1, column=7).fill
            for row in range(2, 31)
        )
        checklist_data["Completed"].append("Yes" if different_styles else "No")

        # 10. Check if table is sorted by Salary
        salary_values = [sheet.cell(row=row, column=8).value for row in range(2, 32)]
        salary_sorted = sorted(salary_values) == salary_values
        checklist_data["Completed"].append("Yes" if salary_sorted else "No")

        # 11. Check center alignment for columns A, F, G, and H
        aligned_columns = all(
            sheet.cell(row=row, column=col).alignment.horizontal == 'center'
            for col in [1, 6, 7, 8]  # Columns A (ID), F (Graduation Year), G (Experience), H (Salary)
            for row in range(2, 32)
        )
        checklist_data["Completed"].append("Yes" if aligned_columns else "No")

        # 12-15. Check bold and formula presence in cells H33, H34, I33, I34
        bold_cells = {
            "H33": sheet["H33"], "H34": sheet["H34"], "I33": sheet["I33"], "I34": sheet["I34"]
        }
        for cell_name, cell in bold_cells.items():
            is_bold_and_formula = cell.font.bold and cell.data_type == 'f'
            checklist_data["Completed"].append("Yes" if is_bold_and_formula else "No")

        # 16. Check if headers in row 1 are bolded
        headers_bold = all(
            sheet.cell(row=1, column=col).font.bold
            for col in range(1, len(expected_columns) + 1)
        )
        checklist_data["Completed"].append("Yes" if headers_bold else "No")

        # 17. Check if all borders and thick outside border are applied
        all_borders_applied = all(
            sheet.cell(row=row, column=col).border is not None
            for row in range(1, 32)
            for col in range(1, len(expected_columns) + 1)
        )
        checklist_data["Completed"].append("Yes" if all_borders_applied else "No")

        # 18. Check if ChatGPT link is present in A35 and centered
        a35_value = sheet["A35"].value if sheet else ""
        link_present = bool(re.match(r"https://chatgpt.com/share/\w+", str(a35_value)))
        centered = sheet["A35"].alignment.horizontal == "center" if sheet["A35"].alignment else False
        checklist_data["Completed"].append("Yes" if link_present and centered else "No")

        # 19. Check if row 35 has a background color
        background_color_present = (
            sheet["A35"].fill is not None and sheet["A35"].fill.fill_type is not None
        )
        checklist_data["Completed"].append("Yes" if background_color_present else "No")

        # Display checklist table
        st.subheader("Checklist Results")
        checklist_df = pd.DataFrame(checklist_data)
        st.table(checklist_df)

    except Exception as e:
        st.error(f"An error occurred: {e}")
