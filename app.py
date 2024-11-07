import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

st.title("Excel Assignment Checker")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file with openpyxl
        workbook = load_workbook(uploaded_file)
        sheet_names = workbook.sheetnames
        alumni_sheet_present = "Alumni" in sheet_names
        is_first_sheet = (sheet_names[0] == "Alumni")

        # Initialize checklist data
        checklist_data = {
            "Grading Criteria": [
                "Is 'Alumni' sheet the first sheet?",
                "Do columns match the expected names and order?",
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
                "Is ChatGPT link merged and centered in row 35?",
                "Is row 35 filled with a background color?"
            ],
            "Completed": []
        }

        # Add to checklist whether "Alumni" sheet is first
        checklist_data["Completed"].append("Yes" if is_first_sheet else "No")

        # Load the "Alumni" sheet if present, otherwise load the first available sheet
        if alumni_sheet_present:
            alumni_sheet = workbook["Alumni"]
        else:
            alumni_sheet = workbook[sheet_names[0]]

        # Determine the last row and last column with data in the loaded sheet
        max_row = alumni_sheet.max_row
        max_column = alumni_sheet.max_column

        # Load data into DataFrame up to the last filled cell
        data = alumni_sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_column, values_only=True)
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

        # Check column order and names
        columns_match = (alumni_df.columns.tolist() == expected_columns)
        checklist_data["Completed"].append("Yes" if columns_match else "No")

        # Proceed only if "ID" column is present
        if "ID" in alumni_df.columns:
            # Check if ID column contains unique numerical identifiers starting from 1001
            id_values = pd.to_numeric(alumni_df['ID'], errors='coerce')
            id_values = id_values.dropna()  # Remove any NaN values
            are_unique = id_values.is_unique
            all_above_1001 = (id_values >= 1001).all()
            id_column_valid = are_unique and all_above_1001
            checklist_data["Completed"].append("Yes" if id_column_valid else "No")
        else:
            checklist_data["Completed"].append("No")

        # Check if Graduation Year calculation formula is in column G
        graduation_year_formula_present = all(
            alumni_sheet.cell(row=row, column=7).data_type == 'f'
            for row in range(2, 33)
        )
        checklist_data["Completed"].append("Yes" if graduation_year_formula_present else "No")

        # Check if Income Earned calculation formula is in column I
        income_earned_formula_present = all(
            alumni_sheet.cell(row=row, column=9).data_type == 'f'
            for row in range(2, 33)
        )
        checklist_data["Completed"].append("Yes" if income_earned_formula_present else "No")

        # Check Accounting format with no decimals in Income Earned column (I2:I32)
        try:
            accounting_format = True
            for row in range(2, 33):
                cell_format = alumni_sheet.cell(row=row, column=9).number_format
                is_valid_format = (
                    cell_format in ['_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)',
                                  '$#,##0',
                                  'Accounting',
                                  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'] or
                    '$' in cell_format
                )

                if not is_valid_format:
                    accounting_format = False
                    break

            checklist_data["Completed"].append("Yes" if accounting_format else "No")

        except Exception as e:
            print(f"Error checking accounting format: {e}")
            checklist_data["Completed"].append("No")

        # Continue with the rest of the checks as before, using alumni_sheet for cell operations...

        # Display checklist table
        st.subheader("Checklist Results")
        checklist_df = pd.DataFrame(checklist_data)
        st.table(checklist_df)

        # Calculate percentage complete and points
        total_yes = checklist_data["Completed"].count("Yes")
        total_items = len(checklist_data["Completed"])
        percentage_complete = (total_yes / total_items) * 100
        points = (total_yes / total_items) * 20

        # Create two columns for displaying scores
        col1, col2 = st.columns(2)

        # Display percentage in first column
        with col1:
            if percentage_complete == 100:
                st.success(f"Completion Score: {percentage_complete:.1f}%")
            elif percentage_complete >= 80:
                st.warning(f"Completion Score: {percentage_complete:.1f}%")
            else:
                st.error(f"Completion Score: {percentage_complete:.1f}%")

        # Display points in second column
        with col2:
            if points == 20:
                st.success(f"Points: {points:.1f}/20")
            elif points >= 16:
                st.warning(f"Points: {points:.1f}/20")
            else:
                st.error(f"Points: {points:.1f}/20")

    except Exception as e:
        st.error(f"An error occurred: {e}")
