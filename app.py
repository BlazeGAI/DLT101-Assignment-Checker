import streamlit as st
import pandas as pd

st.title("Alumni Sheet Checker")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load the Excel file
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Display sheet names
        st.subheader("Sheet Information")
        sheet_names = excel_data.sheet_names
        st.write("Available Sheets:", sheet_names)
        
        # Check for expected "Alumni" sheet as first sheet
        alumni_sheet_present = (sheet_names[0] == "Alumni")
        st.write(f"Is 'Alumni' sheet the first sheet? {'Yes' if alumni_sheet_present else 'No'}")
        
        # Attempt to load "Alumni" or the first sheet
        if alumni_sheet_present:
            alumni_df = excel_data.parse("Alumni")
        else:
            alumni_df = excel_data.parse(sheet_names[0])
        
        # Display first few rows of the sheet for user review
        st.subheader("Preview of Loaded Sheet")
        st.write(alumni_df.head())
        
        # Expected column order and names
        expected_columns = [
            "ID", "First Name", "Last Name", "Bachelor's Degree", 
            "Current Profession", "Graduation Year", "Experience", 
            "Salary", "Income Earned"
        ]
        
        # Column Order and Naming Check
        st.subheader("Column Naming and Order Check")
        columns_match = (alumni_df.columns.tolist() == expected_columns)
        st.write(f"Do columns match the expected names and order? {'Yes' if columns_match else 'No'}")
        if not columns_match:
            st.write("Expected Columns:", expected_columns)
            st.write("Current Columns:", alumni_df.columns.tolist())
        
        # Formatting Checks
        st.subheader("Formatting Checks")
        
        # Check if "Income Earned" is in Accounting format
        # Streamlit cannot directly check Excel formatting, so we describe the expected format.
        st.write("**Expected:** 'Income Earned' should be in Accounting format with no decimals.")
        
        # Check for center alignment
        st.write("**Expected:** All numeric columns should be center-aligned.")
        
        # Check for bold formatting in headers (describe)
        st.write("**Expected:** Header row and summary totals should be bolded.")
        
        # Borders check (describe)
        st.write("**Expected:** All borders applied, with a thick outside border around the entire table.")
        
        # Summary Calculations
        st.subheader("Summary Calculations Check")
        
        # Look for summary rows for Salary and Income Earned
        if alumni_df.iloc[-5:].isnull().values.any():
            totals_and_averages = alumni_df.iloc[-5:].dropna(how="all").reset_index(drop=True)
            if len(totals_and_averages) >= 2:
                total_row = totals_and_averages.iloc[0]
                average_row = totals_and_averages.iloc[1]
                
                st.write("Total and Average Rows found in last 5 rows.")
                st.write("**Total Row:**")
                st.write(total_row)
                st.write("**Average Row:**")
                st.write(average_row)
            else:
                st.write("Total and average calculations are missing or not positioned correctly.")
        else:
            st.write("Total and average calculations are missing.")

        # Check ChatGPT link in last row
        st.subheader("ChatGPT Link Check")
        last_row = alumni_df.iloc[-1].fillna('')
        link_present = 'ChatGPT Link' in last_row.values[0]
        st.write(f"Is ChatGPT link included in the final row? {'Yes' if link_present else 'No'}")
        
    except Exception as e:
        st.error(f"An error occurred: {e}")
