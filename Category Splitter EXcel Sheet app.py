import streamlit as st
import pandas as pd
import io

st.title("📊 Excel Category Splitter - OMAC Developer")

# Step 1: Upload Excel file
uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded:
    # Convert uploaded file to BytesIO
    uploaded_bytes = uploaded.read()
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes), engine="openpyxl")

    # Step 2: Select sheet
    sheet_name = st.selectbox("Select sheet to process:", xls.sheet_names)

    df = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=sheet_name, engine="openpyxl")

    st.write("### Preview of selected sheet")
    st.dataframe(df.head())

    # Step 3: Select category column
    col = st.selectbox("Select the column that contains categories:", df.columns)

    if st.button("Generate Category-wise Sheets"):
        unique_cats = df[col].dropna().unique()

        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine="openpyxl")

        for cat in unique_cats:
            filtered = df[df[col] == cat]
            clean_name = str(cat)[:31].replace("/", "-")
            filtered.to_excel(writer, sheet_name=clean_name, index=False)

        writer.save()
        output.seek(0)

        st.success("Excel file created!")

        st.download_button(
            label="⬇ Download Category-wise Excel",
            data=output,
            file_name="Category_Split.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
