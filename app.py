import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Filter Excel by Assignments")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Load workbook with openpyxl to read headers from row 1 and columns D to O
    import openpyxl

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active  # or specify sheet name if needed

    # Extract headers from row 1, columns D to O (D=4, O=15)
    headers = [ws.cell(row=1, column=col).value for col in range(4, 16)]

    # Columns to copy (D,G,H,M,N,O) correspond to 4,7,8,13,14,15
    cols_to_copy = [4, 7, 8, 13, 14, 15]

    # Extract selected headers for those columns
    selected_headers = [ws.cell(row=1, column=col).value for col in cols_to_copy]

    # Extract data starting from row 2 for those columns
    data = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_data = [row[col - 1].value for col in cols_to_copy]  # zero-index fix
        data.append(row_data)

    # Create DataFrame
    df = pd.DataFrame(data, columns=selected_headers)

    # Verify 'Assignments' column exists in selected_headers
    if "Assignments" not in selected_headers:
        st.error("The column 'Assignments' is not found in the selected columns.")
    else:
        # Show unique values from "Assignments"
        assignments_options = sorted(df["Assignments"].dropna().unique())
        selected_assignments = st.multiselect("Select Assignments to filter", assignments_options)

        if selected_assignments:
            filtered_df = df[df["Assignments"].isin(selected_assignments)]
        else:
            filtered_df = df

        st.write("Filtered data:")
        st.dataframe(filtered_df)

        # Function to save filtered_df to Excel in memory
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='FilteredData')
                # Write headers from row 1 in columns D to O as Excel does (optional)
                # But here pandas writes headers in first row by default.
            output.seek(0)
            return output

        excel_data = to_excel(filtered_df)

        st.download_button(
            label="Download filtered Excel",
            data=excel_data,
            file_name="filtered_assignments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
