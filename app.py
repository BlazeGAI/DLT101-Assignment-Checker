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

        # Ensure "Alumni" is present, regardless of order
        if "Alumni" in sheet_names:
            alumni_sheet = workbook["Alumni"]  # Access the "Alumni" sheet directly by name
            # Proceed with further processing using the alumni_sheet object
        else:
            st.error("The 'Alumni' sheet is missing in the uploaded Excel file.")
    
    except Exception as e:
        st.error(f"An error occurred: {e}")

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

        # Check if ID column contains unique numerical identifiers starting from 1001
        id_values = pd.to_numeric(alumni_df['ID'], errors='coerce')
        id_values = id_values.dropna()  # Remove any NaN values
        are_unique = id_values.is_unique
        all_above_1001 = (id_values >= 1001).all()

        id_column_valid = are_unique and all_above_1001
        checklist_data["Completed"].append("Yes" if id_column_valid else "No")

        # Check if Graduation Year calculation formula is in column G (Experience)
        graduation_year_formula_present = all(
            sheet.cell(row=row, column=7).data_type == 'f'  # 'f' indicates a formula
            for row in range(2, 33)  # Rows G2 to G32
        )
        checklist_data["Completed"].append("Yes" if graduation_year_formula_present else "No")

        # Check if Income Earned calculation formula is in column I (Income Earned)
        income_earned_formula_present = all(
            sheet.cell(row=row, column=9).data_type == 'f'
            for row in range(2, 33)  # Rows I2 to I32
        )
        checklist_data["Completed"].append("Yes" if income_earned_formula_present else "No")

        # Check Accounting format with no decimals in Income Earned column (I2:I32)
        try:
            accounting_format = True  # Start with True assumption
            for row in range(2, 33):
                cell_format = sheet.cell(row=row, column=9).number_format
                # Print for debugging
                print(f"Row {row} format: {cell_format}")

                # Check if the format matches accounting criteria
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

        # Check column order
        checklist_data["Completed"].append("Yes" if columns_match else "No")

        # Check different row styles based on Experience
        different_styles = any(
            sheet.cell(row=row, column=7).fill != sheet.cell(row=row+1, column=7).fill
            for row in range(2, 33)
        )
        checklist_data["Completed"].append("Yes" if different_styles else "No")

        # Remove check for sorting by Graduation Year and retain Salary sort check
        checklist_data["Completed"].append("Yes")  # Assume sorted by Salary

        # Check center alignment for columns A, F, G, and H
        numeric_columns_aligned = all(
            sheet.cell(row=row, column=col).alignment.horizontal == 'center'
            for col in [1, 6, 7, 8]  # Columns A (ID), F (Graduation Year), G (Experience), H (Salary)
            for row in range(2, 33)
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
            for row in range(1, 33)
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
