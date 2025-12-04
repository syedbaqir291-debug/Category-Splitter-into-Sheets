import streamlit as st
import pandas as pd
import io

# ===========================
# App Header
# ===========================
st.set_page_config(
    page_title="Categories Splitter for Excel - OMAC Developer",
    page_icon="📊",
    layout="wide"
)

st.markdown(
    """
    <h1 style='text-align: center; color: #4B8BBE;'>📂 Categories Splitter for Excel</h1>
    <p style='text-align: center; color: #555;'>Upload your Excel, select category column, and get separate sheets automatically</p>
    <hr>
    """, unsafe_allow_html=True
)

# ===========================
# Upload Excel File
# ===========================
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # ===========================
        # Select Sheet
        # ===========================
        selected_sheet = st.selectbox("Select the sheet to process", sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        # ===========================
        # Select Category Column
        # ===========================
        col_option = st.selectbox("Select the category column", df.columns)
        
        # ===========================
        # Function to Split Excel
        # ===========================
        def split_to_sheets(df, category_col):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                categories = df[category_col].dropna().unique()
                for cat in categories:
                    df_cat = df[df[category_col] == cat]
                    sheet_name_safe = str(cat)[:31]  # Excel max 31 chars
                    df_cat.to_excel(writer, sheet_name=sheet_name_safe, index=False)
            processed_data = output.getvalue()
            return processed_data
        
        # ===========================
        # Button to Run
        # ===========================
        if st.button("Split Excel by Category"):
            result_bytes = split_to_sheets(df, col_option)
            st.success("✅ Excel split successfully!")
            
            st.download_button(
                label="📥 Download Split Excel",
                data=result_bytes,
                file_name=f"Split_{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
