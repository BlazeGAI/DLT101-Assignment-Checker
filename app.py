import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, numbers

st.title("Alumni Sheet Checker")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file with openpyxl
        sheet_names = workbook.sheetnames
        alumni_sheet_present = (sheet_names[0] == "Alumni")
        
        # Initialize checklist data
        checklist_data = {
            "Grading Criteria": [
                "Is 'Alumni' sheet the first sheet?",
                "Do columns match the expected names and order?",
                "Is 'Income Earned' in Accounting format with no decimals?",
                "Are all numeric columns center-aligned?",
                "Are header row and one summary row bolded?",
                "Are all borders applied with a thick outside border?",
                "Are total and average calculations present for 'Salary' and 'Income Earned'?",
                "Is the ChatGPT link included in row 35?"
            ],
            "Completed": []
        }

        # Check sheet name and order
        checklist_data["Completed"].append("Yes" if alumni_sheet_present else "No")
        
        # Load the sheet based on the name or default to first sheet if name mismatch
        sheet = workbook[sheet_names[0] if alumni_sheet_present else sheet_names[0]]
        
        # Read the sheet into a DataFrame for column checking and data analysis
        alumni_df = pd.DataFrame(sheet.values)
        alumni_df.columns = alumni_df.iloc[0]  # Set the header row
        alumni_df = alumni_df.drop(0)  # Remove the header row from data
        
        # Expected column names
        expected_columns = [
            "ID", "First Name", "Last Name", "Bachelor's Degree", 
            "Current Profession", "Graduation Year", "Experience", 
            "Salary", "Income Earned"
        ]
        
        # Check column names and order
        columns_match = (alumni_df.columns.tolist() == expected_columns)
        checklist_data["Completed"].append("Yes" if columns_match else "No")
        
        # Check Accounting format for "Income Earned"
        income_earned_column = expected_columns.index("Income Earned") + 1  # 1-based index
        accounting_format = all(
            sheet.cell(row=row, column=income_earned_column).number_format == numbers.FORMAT_ACCOUNTING
            for row in range(2, sheet.max_row)  # Skipping header row
            if sheet.cell(row=row, column=income_earned_column).value is not None
        )
        checklist_data["Completed"].append("Yes" if accounting_format else "No")
        
        # Check center alignment for numeric columns
        numeric_columns = ["ID", "Graduation Year", "Experience", "Salary", "Income Earned"]
        numeric_columns_indices = [expected_columns.index(col) + 1 for col in numeric_columns]  # 1-based indices
        center_aligned = all(
            sheet.cell(row=row, column=col).alignment.horizontal == 'center'
            for col in numeric_columns_indices
            for row in range(2, sheet.max_row)  # Skipping header row
            if sheet.cell(row=row, column=col).value is not None
        )
        checklist_data["Completed"].append("Yes" if center_aligned else "No")
        
        # Check bold formatting for header row and one summary row
        header_bold = all(
            sheet.cell(row=1, column=col).font.bold
            for col in range(1, len(expected_columns) + 1)
        )
        
        # Check if one row of totals is bolded (e.g., second to last or last row of data)
        summary_bold = any(
            all(sheet.cell(row=row, column=col).font.bold for col in range(1, len(expected_columns) + 1))
            for row in range(sheet.max_row - 2, sheet.max_row)  # Checking last two rows for bold formatting
        )
        
        checklist_data["Completed"].append("Yes" if header_bold and summary_bold else "No")
        
        # Check borders (including thick outside border around the entire table)
        thin_border = Side(border_style="thin", color="000000")
        thick_border = Side(border_style="thick", color="000000")
        all_borders = all(
            sheet.cell(row=row, column=col).border.top == thin_border and
            sheet.cell(row=row, column=col).border.bottom == thin_border and
            sheet.cell(row=row, column=col).border.left == thin_border and
            sheet.cell(row=row, column=col).border.right == thin_border
            for row in range(2, sheet.max_row + 1)
            for col in range(1, len(expected_columns) + 1)
        )
        outer_borders = (
            sheet.cell(row=1, column=1).border.top == thick_border and
            sheet.cell(row=sheet.max_row, column=1).border.bottom == thick_border and
            sheet.cell(row=1, column=1).border.left == thick_border and
            sheet.cell(row=1, column=len(expected_columns)).border.right == thick_border
        )
        borders_applied = all_borders and outer_borders
        checklist_data["Completed"].append("Yes" if borders_applied else "No")
        
        # Check for total and average rows at the bottom
        summary_present = False
        if alumni_df.iloc[-5:].isnull().values.any():
            summary_rows = alumni_df.iloc[-5:].dropna(how="all").reset_index(drop=True)
            if len(summary_rows) >= 2:
                summary_present = True
        checklist_data["Completed"].append("Yes" if summary_present else "No")
        
        # Check for ChatGPT link specifically in row 35
        chatgpt_link_row = 35
        if chatgpt_link_row <= sheet.max_row:
            chatgpt_link_cell = sheet.cell(row=chatgpt_link_row, column=1)
            link_present = 'ChatGPT Link' in str(chatgpt_link_cell.value or "")
        else:
            link_present = False
        checklist_data["Completed"].append("Yes" if link_present else "No")
        
        # Display checklist table
        st.subheader("Checklist Results")
        checklist_df = pd.DataFrame(checklist_data)
        st.table(checklist_df)
        
    except Exception as e:
        st.error(f"An error occurred: {e}")
