import streamlit as st
from utils.display import display_results
from checkers.excel.excel_1 import check_excel_1
from checkers.excel.excel_2 import check_excel_2
from checkers.excel.excel_3 import check_excel_3
from openpyxl import load_workbook

st.title("Excel Assignment Checker")

# Create three columns for file uploaders
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Excel Assignment 1")
    excel_1_file = st.file_uploader("Upload Excel_1", type=["xlsx"], key="excel_1")

with col2:
    st.subheader("Excel Assignment 2")
    excel_2_file = st.file_uploader("Upload Excel_2", type=["xlsx"], key="excel_2")

with col3:
    st.subheader("Excel Assignment 3")
    excel_3_file = st.file_uploader("Upload Excel_3", type=["xlsx"], key="excel_3")

# Check files and display results
if excel_1_file:
    try:
        workbook = load_workbook(excel_1_file)
        checklist_data = check_excel_1(workbook)
        st.subheader("Excel Assignment 1 Results")
        display_results(checklist_data)
    except Exception as e:
        st.error(f"An error occurred with Excel Assignment 1: {str(e)}")

if excel_2_file:
    try:
        workbook = load_workbook(excel_2_file)
        checklist_data = check_excel_2(workbook)
        st.subheader("Excel Assignment 2 Results")
        display_results(checklist_data)
    except Exception as e:
        st.error(f"An error occurred with Excel Assignment 2: {str(e)}")

if excel_3_file:
    try:
        workbook = load_workbook(excel_3_file)
        checklist_data = check_excel_3(workbook)
        st.subheader("Excel Assignment 3 Results")
        display_results(checklist_data)
    except Exception as e:
        st.error(f"An error occurred with Excel Assignment 3: {str(e)}")
