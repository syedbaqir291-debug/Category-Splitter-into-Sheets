import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Category Splitter", layout="wide")
st.title("📊 Excel Category Splitter into Sheets - OMAC Developer")

# Step 1: Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    # Step 2: Select the sheet to work on
    sheet_name = st.selectbox("Select Sheet to Process", xls.sheet_names)

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    st.write("### Preview of Selected Sheet")
    st.dataframe(df.head())

    # Step 3: Select the category column
    category_col = st.selectbox("Select the Column that Contains Categories", df.columns)

    if st.button("Split by Categories and Download"):
        unique_categories = df[category_col].dropna().unique()

        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine="openpyxl")

        for cat in unique_categories:
            filtered_df = df[df[category_col] == cat]

            # Clean sheet name (Excel allows max 31 chars)
            sheet_name_clean = str(cat)[:31].replace("/", "-")
            filtered_df.to_excel(writer, sheet_name=sheet_name_clean, index=False)

        writer.save()
        output.seek(0)

        st.success(f"✅ Excel file created with {len(unique_categories)} sheets!")
        st.download_button(
            label="⬇ Download Category-wise Excel",
            data=output,
            file_name="Category_Split.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
