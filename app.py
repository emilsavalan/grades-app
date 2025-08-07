import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

st.title("Filter Excel by Assignments")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active  # or specify sheet explicitly

    # Columns indexes for D,G,H,M,N,O
    cols_to_copy = [4, 7, 8, 13, 14, 15]

    # Get headers from row 1 for these columns
    selected_headers = [ws.cell(row=1, column=col).value for col in cols_to_copy]

    # Read data starting from row 2 to last row for these columns
    data = []
    for row in range(2, ws.max_row + 1):
        row_values = [ws.cell(row=row, column=col).value for col in cols_to_copy]
        data.append(row_values)

    df = pd.DataFrame(data, columns=selected_headers)

    st.write("Extracted headers:", selected_headers)
    st.write("Sample data:", df.head())

    # Find column with header "Assignments" (case insensitive)
    assignments_col = None
    for col_name in selected_headers:
        if col_name and col_name.strip().lower() == "assignments":
            assignments_col = col_name
            break

    if assignments_col is None:
        st.error("The column 'Assignments' was not found in the selected columns.")
    else:
        assignments_options = sorted(df[assignments_col].dropna().unique())
        selected_assignments = st.multiselect("Select Assignments to filter", assignments_options)

        if selected_assignments:
            filtered_df = df[df[assignments_col].isin(selected_assignments)]
        else:
            filtered_df = df

        st.dataframe(filtered_df)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='FilteredData')
            output.seek(0)
            return output

        excel_data = to_excel(filtered_df)

        st.download_button(
            label="Download filtered Excel",
            data=excel_data,
            file_name="filtered_assignments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
