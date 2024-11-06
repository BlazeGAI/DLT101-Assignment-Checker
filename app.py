import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

st.title("Alumni Sheet Checker")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file with openpyxl
        workbook = load_workbook(uploaded_file)
        sheet_names = workbook.sheetnames
        alumni_sheet_present = (sheet_names[0] == "Alumni")
        
        # Initialize checklist data
        checklist_data = {
            "Grading Criteria": [
                "Is 'Alumni' sheet the first sheet?",
                "Do columns match the expected names and order?",
                "Is ID column added with unique four-digit numbers starting from 1001?",
                "Is Graduation Year calculation formula in column G?",
                "Is Income Earned calculation formula in column I?",
                "Is Accounting format applied to Income Earned with no decimals?",
                "Are columns rearranged in the correct order?",
                "Is the table sorted by Graduation Year?",
                "Are rows styled differently based on Graduation Year?",
                "Is the table sorted by Salary?",
                "Are numeric columns center-aligned?",
                "Is total 'Salary' in cell H33 and set to bold?",
                "Is average 'Salary' in H34 and set to bold?",
                "Is total 'Income Earned' in I33 and set to bold?",
                "Is average 'Income Earned' in I34 and set to bold?",
                "Are headers in row 1 bolded?",
                "Are all borders and thick outside border applied?",
                "Is ChatGPT link merged and centered in row 35?",
                "Is row 35 filled with a background color?"
            ],
            "Completed": []
        }

        # Load the Alumni sheet
        sheet = workbook[sheet_names[0]] if alumni_sheet_present else workbook[sheet_names[0]]
        
        # Determine the last row and last column with data
        max_row = sheet.max_row
        max_column = sheet.max_column

        # Load data into DataFrame up to the last filled cell
        data = sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_column, values_only=True)
        alumni_df = pd.DataFrame(data)

        # Set headers and drop any fully empty rows
        alumni_df.columns = alumni_df.iloc[0]  # Set the header row
        alumni_df = alumni_df.drop(0).reset_index(drop=True)  # Remove header row from data

        # Drop any completely empty rows or columns to avoid alignment issues
        alumni_df = alumni_df.dropna(how='all', axis=0).dropna(how='all', axis=1)

        # Expected columns in the final order
        expected_columns = [
            "ID", "First Name", "Last Name", "Bachelor's Degree",
            "Current Profession", "Graduation Year", "Experience",
            "Salary", "Income Earned"
        ]

        # Check if "Alumni" is the first sheet
        checklist_data["Completed"].append("Yes" if alumni_sheet_present else "No")
        
        # Check column order and names
        columns_match = (alumni_df.columns.tolist() == expected_columns)
        checklist_data["Completed"].append("Yes" if columns_match else "No")

        # Check if ID column contains unique four-digit numbers, starting from 1001 or higher
        id_values = pd.to_numeric(alumni_df['ID'], errors='coerce')
        id_column_valid = id_values.apply(lambda x: 1001 <= x <= 9999).all() and id_values.is_unique
        checklist_data["Completed"].append("Yes" if id_column_valid else "No")
        
        # Check if Graduation Year calculation formula is in column G (Experience)
        graduation_year_formula_present = all(
            sheet.cell(row=row, column=7).data_type == 'f'  # 'f' indicates a formula
            for row in range(2, max_row + 1)  # Rows with data
        )
        checklist_data["Completed"].append("Yes" if graduation_year_formula_present else "No")

        # Check if Income Earned calculation formula is in column I
        income_earned_formula_present = all(
            sheet.cell(row=row, column=9).data_type == 'f'
            for row in range(2, max_row + 1)
        )
        checklist_data["Completed"].append("Yes" if income_earned_formula_present else "No")

        # Check Accounting format with no decimals in Income Earned column (column I)
        accounting_format = all(
            sheet.cell(row=row, column=9).number_format in ["$#,##0", "$#,##0;[Red]$-#,##0", "Accounting"]
            for row in range(2, max_row + 1)
        )
        checklist_data["Completed"].append("Yes" if accounting_format else "No")
        
        # Check column order
        checklist_data["Completed"].append("Yes" if columns_match else "No")

        # Check sorting by Graduation Year, ignoring non-numeric values
        graduation_year_values = pd.to_numeric(alumni_df['Graduation Year'], errors='coerce').dropna()
        graduation_year_sorted = graduation_year_values.is_monotonic_increasing
        checklist_data["Completed"].append("Yes" if graduation_year_sorted else "No")

        # Check different row styles based on Graduation Year
        different_styles = any(
            sheet.cell(row=row, column=1).fill != sheet.cell(row=row+1, column=1).fill
            for row in range(2, max_row - 1)
        )
        checklist_data["Completed"].append("Yes" if different_styles else "No")

        # Check sorting by Salary
        salary_values = pd.to_numeric(alumni_df['Salary'], errors='coerce').dropna()
        salary_sorted = salary_values.is_monotonic_increasing
        checklist_data["Completed"].append("Yes" if salary_sorted else "No")

        # Check center alignment for numeric columns
        numeric_columns_aligned = all(
            sheet.cell(row=row, column=col).alignment.horizontal == 'center'
            for col in [1, 7, 8, 9]  # ID, Experience, Salary, Income Earned columns
            for row in range(2, max_row + 1)
        )
        checklist_data["Completed"].append("Yes" if numeric_columns_aligned else "No")

        # Check total Salary in H33 is bold and contains a formula
        total_salary_bold = sheet['H33'].font.bold if sheet['H33'].data_type == 'f' else False
        checklist_data["Completed"].append("Yes" if total_salary_bold else "No")

        # Check average Salary in H34 is bold and contains a formula
        average_salary_bold = sheet['H34'].font.bold if sheet['H34'].data_type == 'f' else False
        checklist_data["Completed"].append("Yes" if average_salary_bold else "No")

        # Check total Income Earned in I33 is bold and contains a formula
        total_income_bold = sheet['I33'].font.bold if sheet['I33'].data_type == 'f' else False
        checklist_data["Completed"].append("Yes" if total_income_bold else "No")

        # Check average Income Earned in I34 is bold and contains a formula
        average_income_bold = sheet['I34'].font.bold if sheet['I34'].data_type == 'f' else False
        checklist_data["Completed"].append("Yes" if average_income_bold else "No")

        # Check if headers are bold
        headers_bold = all(
            sheet.cell(row=1, column=col).font.bold
            for col in range(1, len(expected_columns) + 1)
        )
        checklist_data["Completed"].append("Yes" if headers_bold else "No")

        # Check borders and thick outside border
        all_borders_applied = all(
            sheet.cell(row=row, column=col).border is not None
            for row in range(1, max_row + 1)
            for col in range(1, len(expected_columns) + 1)
        )
        checklist_data["Completed"].append("Yes" if all_borders_applied else "No")

        # Check if cells in row 35 are merged and center-aligned
        merged_in_row_35 = any("A35" in str(range) for range in sheet.merged_cells.ranges)
        center_aligned = sheet['A35'].alignment.horizontal == 'center' if merged_in_row_35 else False
        checklist_data["Completed"].append("Yes" if merged_in_row_35 and center_aligned else "No")

        # Check if row 35 has a background color
        background_color_present = sheet['A35'].fill is not None and sheet['A35'].fill.fill_type is not None
        checklist_data["Completed"].append("Yes" if background_color_present else "No")

        # Display checklist table
        st.subheader("Checklist Results")
        checklist_df = pd.DataFrame(checklist_data)
        st.table(checklist_df)

    except Exception as e:
        st.error(f"An error occurred: {e}")
