import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

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
                "Is Graduation Year calculation formula in column H?",
                "Is Income Earned calculation formula in column I?",
                "Is Accounting format applied to Income Earned with no decimals?",
                "Are columns rearranged in the correct order?",
                "Is the table sorted by Graduation Year?",
                "Are rows styled differently based on Graduation Year?",
                "Is the table sorted by Salary?",
                "Are numeric columns center-aligned?",
                "Is total 'Salary' in cell H32 and set to bold?",
                "Is average 'Salary' in H33 and set to bold?",
                "Is total 'Income Earned' in I32 and set to bold?",
                "Is average 'Income Earned' in I33 and set to bold?",
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
        
        # Check if Graduation Year calculation formula is in column H
        graduation_year_formula_present = all(
            sheet.cell(row=row, column=7).data_type == 'f'  # 'f' indicates a formula
            for row in range(2, max_row)  # Rows with data
        )
        checklist_data["Completed"].append("Yes" if graduation_year_formula_present else "No")

        # Check if Income Earned calculation formula is in column I
        income_earned_formula_present = all(
            sheet.cell(row=row, column=9).data_type == 'f'
            for row in range(2, max_row)
        )
        checklist_data["Completed"].append("Yes" if income_earned_formula_present else "No")

        # Check Accounting format with no decimals in Income Earned column
        accounting_format = all(
            sheet.cell(row=row, column=9).number_format in ["$#,##0", "$#,##0;[Red]$-#,##0", "Accounting"]
            for row in range(2, max_row)
        )
        checklist_data["Completed"].append("Yes" if accounting_format else "No")
        
        # Check column order
        checklist_data["Completed"].append("Yes" if columns_match else "No")

        # Check sorting by Graduation Year, ignoring non-numeric values
        graduation_year_values = pd.to_numeric(alumni_df['Graduation Year'], errors='coerce').dropna()
        graduation_year_sorted = graduation_year_values.is_monotonic_increasing
        checklist_data["Completed"].append("Yes" if graduation_year_sorted else "No")

        # Continue with the rest of the checklist items as before
        # ...

        # Display checklist table
        st.subheader("Checklist Results")
        checklist_df = pd.DataFrame(checklist_data)
        st.table(checklist_df)

    except Exception as e:
        st.error(f"An error occurred: {e}")
