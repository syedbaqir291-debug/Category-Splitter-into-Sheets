import streamlit as st
import pandas as pd
from io import BytesIO

# Page config
st.set_page_config(page_title="Categories Splitter for Excel - OMAC Developer", layout="wide")

st.markdown("""
    <h1 style='text-align: center; color: #1f77b4;'>Categories Splitter for Excel</h1>
    <h4 style='text-align: center; color: #555;'>OMAC Developer</h4>
""", unsafe_allow_html=True)

st.write("---")

# Upload Excel
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read Excel to list available sheets
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select sheet to work on", xls.sheet_names)
    
    if sheet_name:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Ask for category column
        col_option = st.selectbox("Select column which has categories", df.columns)
        
        if col_option:
            st.write("Preview of your data:")
            st.dataframe(df.head(5))
            
            # Split into sheets based on category
            def split_to_sheets(df, category_col):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for cat in df[category_col].dropna().unique():
                        df_cat = df[df[category_col] == cat]
                        sheet_name_safe = str(cat)[:31]  # Excel max 31 chars
                        df_cat.to_excel(writer, sheet_name=sheet_name_safe, index=False)
                    writer.save()
                processed_data = output.getvalue()
                return processed_data
            
            if st.button("Split Excel by Category"):
                result_bytes = split_to_sheets(df, col_option)
                st.success("Excel split successfully!")
                
                st.download_button(
                    label="Download Updated Excel",
                    data=result_bytes,
                    file_name="Categories_Split.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
