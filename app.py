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
        sheet = workbook.active

        # Initialize checklist data
        checklist_data = {
            "Grading Criteria": [
                "Are there 6 columns A-F?",
                "Are there exactly 10 rows of data in the dataset?",
                "Are the column headers named ID, First Name, Last Name, Date of Birth, Hometown, Occupation?",
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
        num_data_rows = sheet.max_row - 1  # Subtract 1 for header row
        checklist_data["Completed"].append("Yes" if num_data_rows == 10 else "No")

        # Check column headers
        expected_headers = ["ID", "First Name", "Last Name", "Date of Birth", "Hometown", "Occupation"]
        headers = [sheet.cell(row=1, column=i).value for i in range(1, 7)]
        headers_match = all(a == b for a, b in zip(headers, expected_headers))
        checklist_data["Completed"].append("Yes" if headers_match else "No")

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
            for col in range(1, 7)  # Changed to 7 to check columns A-F
        )
        checklist_data["Completed"].append("Yes" if headers_bold else "No")

        # Check borders
        all_borders_applied = all(
            sheet.cell(row=row, column=col).border is not None
            for row in range(1, 12)  # Check rows 1-11 (headers + 10 data rows)
            for col in range(1, 7)   # Check columns A-F
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

        # Display results
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
