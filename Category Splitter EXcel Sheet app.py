import streamlit as st
import xlwings as xw
import tempfile
import os

st.title("Excel Category Splitter with VBA")

# Step 1: Upload Excel
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    # Step 2: Ask for sheet name
    wb = xw.Book(tmp_path)
    sheet_names = [s.name for s in wb.sheets]
    selected_sheet = st.selectbox("Select the sheet to work on", sheet_names)

    # Step 3: Ask for category column (letter)
    category_column = st.text_input("Enter the column letter with categories (e.g., M)")

    # Button to run VBA
    if st.button("Run VBA Macro"):
        if not category_column:
            st.warning("Please enter the category column letter.")
        else:
            macro_code = f'''
Sub SplitByCategory_Custom()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim catCol As Long
    Dim catVal As String
    Dim newWS As Worksheet
    Dim r As Long
    Dim categoryDict As Object

    Set ws = ThisWorkbook.Sheets("{selected_sheet}")
    catCol = Columns("{category_column}").Column
    
    lastRow = ws.Cells(ws.Rows.Count, catCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Set categoryDict = CreateObject("Scripting.Dictionary")
    
    ' Collect unique categories
    For r = 2 To lastRow
        catVal = ws.Cells(r, catCol).Value
        If Len(Trim(catVal)) > 0 Then
            If Not categoryDict.Exists(catVal) Then
                categoryDict.Add catVal, 1
            End If
        End If
    Next r
    
    ' Create sheets and copy data
    Dim key As Variant
    For Each key In categoryDict.Keys
        On Error Resume Next
        Set newWS = ThisWorkbook.Sheets(key)
        On Error GoTo 0
        
        If newWS Is Nothing Then
            Set newWS = ThisWorkbook.Sheets.Add
            newWS.Name = Left(key, 31)
        Else
            newWS.Cells.Clear
        End If
        
        ws.Rows(1).Copy Destination:=newWS.Rows(1)
        
        Dim pasteRow As Long
        pasteRow = 2
        For r = 2 To lastRow
            If ws.Cells(r, catCol).Value = key Then
                ws.Rows(r).Copy Destination:=newWS.Rows(pasteRow)
                pasteRow = pasteRow + 1
            End If
        Next r
        
        Set newWS = Nothing
    Next key
End Sub
'''

            # Add VBA module and run macro
            xw.apps.active.api.VBE.ActiveVBProject.VBComponents.Add(1).CodeModule.AddFromString(macro_code)
            wb.macro("SplitByCategory_Custom")()
            
            # Save updated workbook
            updated_path = tmp_path.replace(".xlsx", "_updated.xlsx")
            wb.save(updated_path)
            wb.close()
            st.success("✅ VBA Macro executed successfully!")

            # Provide download
            with open(updated_path, "rb") as f:
                st.download_button("Download Updated Excel", f, file_name="Updated_Categories.xlsx")

            # Cleanup temp files
            os.remove(tmp_path)
            os.remove(updated_path)
