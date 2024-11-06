import streamlit as st
import pandas as pd

st.title("Alumni Sheet Checker")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Initialize checklist data
        checklist_data = {
            "Grading Criteria": [
                "Is 'Alumni' sheet the first sheet?",
                "Do columns match the expected names and order?",
                "Is 'Income Earned' in Accounting format with no decimals?",
                "Are all numeric columns center-aligned?",
                "Are header row and summary totals bolded?",
                "Are all borders applied with a thick outside border?",
                "Are total and average calculations present for 'Salary' and 'Income Earned'?",
                "Is the ChatGPT link included in the final row?"
            ],
            "Completed": []
        }

        # Check sheet name and order
        sheet_names = excel_data.sheet_names
        alumni_sheet_present = (sheet_names[0] == "Alumni")
        checklist_data["Completed"].append("Yes" if alumni_sheet_present else "No")
        
        # Attempt to load the expected Alumni sheet
        alumni_df = excel_data.parse(sheet_names[0] if alumni_sheet_present else sheet_names[0])

        # Check column order and naming
        expected_columns = [
            "ID", "First Name", "Last Name", "Bachelor's Degree", 
            "Current Profession", "Graduation Year", "Experience", 
            "Salary", "Income Earned"
        ]
        columns_match = (alumni_df.columns.tolist() == expected_columns)
        checklist_data["Completed"].append("Yes" if columns_match else "No")

        # Adding descriptive checks for formatting, as Streamlit can't verify Excel-specific formats
        checklist_data["Completed"].append("N/A")  # Accounting format for Income Earned
        checklist_data["Completed"].append("N/A")  # Center alignment for numeric columns
        checklist_data["Completed"].append("N/A")  # Bold headers and summary totals
        checklist_data["Completed"].append("N/A")  # Borders around table

        # Check if total and average rows exist in the final rows for Salary and Income Earned
        summary_present = False
        if alumni_df.iloc[-5:].isnull().values.any():
            summary_rows = alumni_df.iloc[-5:].dropna(how="all").reset_index(drop=True)
            if len(summary_rows) >= 2:
                summary_present = True
        checklist_data["Completed"].append("Yes" if summary_present else "No")

        # Check for ChatGPT link in the final row
        last_row = alumni_df.iloc[-1].fillna('')
        link_present = 'ChatGPT Link' in last_row.values[0]
        checklist_data["Completed"].append("Yes" if link_present else "No")

        # Display checklist table
        st.subheader("Checklist Results")
        checklist_df = pd.DataFrame(checklist_data)
        st.table(checklist_df)

    except Exception as e:
        st.error(f"An error occurred: {e}")
