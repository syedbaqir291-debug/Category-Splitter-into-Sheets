import streamlit as st
import pandas as pd
import io

st.title("📊 Excel Category Splitter - OMAC Developer")

# Step 1: Upload Excel file
uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded:
    # Read excel and list sheets
    xls = pd.ExcelFile(uploaded)
    sheet_name = st.selectbox("Select sheet to process:", xls.sheet_names)

    df = pd.read_excel(uploaded, sheet_name=sheet_name)

    st.write("### Preview of selected sheet")
    st.dataframe(df.head())

    # Step 2: Ask user to select category column
    col = st.selectbox("Select the column that contains categories:", df.columns)

    if st.button("Generate Category-wise Sheets"):
        unique_cats = df[col].dropna().unique()

        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine="openpyxl")

        for cat in unique_cats:
            # Filter rows for this category
            filtered = df[df[col] == cat]

            # Clean sheet name (Excel doesn't allow some chars)
            clean_name = str(cat)[:31].replace("/", "-")

            filtered.to_excel(writer, sheet_name=clean_name, index=False)

        writer.save()
        output.seek(0)

        st.success("Excel file generated successfully!")

        st.download_button(
            label="⬇ Download Category-wise Excel",
            data=output,
            file_name="Category_Wise_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
