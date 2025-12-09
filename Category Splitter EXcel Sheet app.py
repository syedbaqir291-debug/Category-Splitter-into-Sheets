import streamlit as st
import pandas as pd
import io

# ===========================
# LOGIN SECTION (ADDED)
# ===========================
st.set_page_config(
    page_title="Login - Categories Splitter",
    page_icon="🔐",
    layout="wide"
)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown(
        """
        <h2 style='text-align:center;'>🔐 Login Required</h2>
        <hr>
        """,
        unsafe_allow_html=True
    )

    username = st.text_input("User Name")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == "Anum" and password == "Pakistan@1947":
            st.session_state.authenticated = True
            st.success("✅ Login successful")
            st.rerun()
        else:
            st.error("❌ Invalid User Name or Password")

    st.stop()

# ===========================
# App Header (UNCHANGED)
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
                    sheet_name_safe = str(cat)[:31]
                    df_cat.to_excel(writer, sheet_name=sheet_name_safe, index=False)
            return output.getvalue()
        
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
