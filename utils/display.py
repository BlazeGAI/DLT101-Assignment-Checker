import streamlit as st
import pandas as pd

def display_results(checklist_data):
    # Calculate scores
    total_yes = checklist_data["Completed"].count("Yes")
    total_items = len(checklist_data["Completed"])
    percentage_complete = (total_yes / total_items) * 100
    points = (total_yes / total_items) * 20

    # Display scores
    col1, col2 = st.columns(2)
    
    with col1:
        if percentage_complete == 100:
            st.success(f"Completion Score: {percentage_complete:.1f}%")
        elif percentage_complete >= 80:
            st.warning(f"Completion Score: {percentage_complete:.1f}%")
        else:
            st.error(f"Completion Score: {percentage_complete:.1f}%")

    with col2:
        if points == 20:
            st.success(f"Points: {points:.1f}/20")
        elif points >= 16:
            st.warning(f"Points: {points:.1f}/20")
        else:
            st.error(f"Points: {points:.1f}/20")

    # Display checklist
    st.subheader("Detailed Checklist")
    checklist_df = pd.DataFrame(checklist_data)
    st.table(checklist_df)
