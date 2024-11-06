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
                "Does ID column contain unique numerical identifiers starting from 1001?",
                # Other checklist items...
            ],
            "Completed": []
        }

        # Load the Alumni sheet
        sheet = workbook[sheet_names[0]] if alumni_sheet_present else workbook[sheet_names[0]]
        
        # Convert the sheet into a DataFrame for easier manipulation
        data = sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column, values_only=True)
        alumni_df = pd.DataFrame(data)
        
        # Set headers and drop any fully empty rows
        alumni_df.columns = alumni_df.iloc[0]  # Set the header row
        alumni_df = alumni_df.drop(0).reset_index(drop=True)  # Remove header row from data
        
        # DEBUG: Print the ID column to see its contents (excluding header)
        id_values = pd.to_numeric(alumni_df['ID'][1:], errors='coerce')
        st.write("ID Values:", id_values)  # Debug output to see the IDs
        
        # Check if ID column contains unique numerical identifiers starting from 1001
        id_column_valid = id_values.apply(lambda x: x >= 1001).all() and id_values.is_unique
        checklist_data["Completed"].append("Yes" if id_column_valid else "No")

        # DEBUG: Print the validation result for the ID column
        st.write("ID Column Validity Check:", id_column_valid)  # Debug output to see result
        
        # Display checklist table
        st.subheader("Checklist Results")
        checklist_df = pd.DataFrame(checklist_data)
        st.table(checklist_df)

    except Exception as e:
        st.error(f"An error occurred: {e}")
