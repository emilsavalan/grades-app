import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

st.title("Filter Excel by Assignments")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active  # or specify sheet explicitly
    
    # Columns indexes for D,G,H,M,N,O (1-indexed: D=4, G=7, H=8, M=13, N=14, O=15)
    cols_to_copy = [4, 7, 8, 13, 14, 15]
    
    # Get title from row 1, column D
    title_cell = ws.cell(row=1, column=4).value  # D1
    st.write("Title from D1:", title_cell)
    
    # Get actual headers from row 2 for these columns
    raw_headers = [ws.cell(row=2, column=col).value for col in cols_to_copy]
    
    # Clean up headers (remove None values and ensure they're strings)
    selected_headers = []
    for i, header in enumerate(raw_headers):
        if header is None or header == "":
            # Use column letter as fallback for empty headers
            col_letter = openpyxl.utils.get_column_letter(cols_to_copy[i])
            unique_header = f"Column_{col_letter}"
        else:
            unique_header = str(header).strip()
        
        # Handle duplicates by adding a suffix
        original_header = unique_header
        counter = 1
        while unique_header in selected_headers:
            unique_header = f"{original_header}_{counter}"
            counter += 1
        
        selected_headers.append(unique_header)
    
    # First, find which column index corresponds to "Assignments"
    assignments_col_index = None
    assignments_excel_col = None
    for i, header in enumerate(raw_headers):
        if header and "assignment" in str(header).lower():
            assignments_col_index = i
            assignments_excel_col = cols_to_copy[i]
            break
    
    # Read data starting from row 3, filtering for "riyaziyyat" in Assignments column
    data = []
    filtered_rows_count = 0
    total_rows_count = 0
    
    for row in range(3, ws.max_row + 1):  # Start from row 3 since row 2 has headers
        total_rows_count += 1
        row_values = [ws.cell(row=row, column=col).value for col in cols_to_copy]
        
        # Check if this row should be included (contains "riyaziyyat" in assignments column)
        if assignments_col_index is not None:
            assignments_value = row_values[assignments_col_index]
            if assignments_value and isinstance(assignments_value, str):
                # Check for both "riyaziyyat" and "Riyyaziyyat" (case insensitive)
                if "riyaziyyat" in assignments_value.lower():
                    data.append(row_values)
                    filtered_rows_count += 1
        else:
            # If assignments column not found, include all rows
            data.append(row_values)
    
    # Create DataFrame with unique headers
    df = pd.DataFrame(data, columns=selected_headers)
    
    st.write(f"Total rows processed: {total_rows_count}")
    st.write(f"Rows containing 'riyaziyyat': {filtered_rows_count}")
    st.write(f"Rows copied to new file: {len(df)}")
    
    st.write("Extracted headers:", selected_headers)
    st.write("Raw headers from Excel:", raw_headers)
    
    # Display sample data (handle potential display issues)
    try:
        st.write("Sample data:")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Error displaying data: {e}")
        st.write("DataFrame shape:", df.shape)
        st.write("DataFrame columns:", df.columns.tolist())
    
    # Find column with header "Assignments" (case insensitive)
    assignments_col = None
    for col_name in selected_headers:
        if col_name and "assignment" in col_name.lower():
            assignments_col = col_name
            break
    
    if assignments_col is None:
        st.error("The column 'Assignments' was not found in the selected columns.")
        st.write("Available columns:", selected_headers)
    else:
        st.success(f"Found assignments column: {assignments_col}")
        
        # Get unique assignments, filtering out None/empty values
        assignments_series = df[assignments_col].dropna()
        assignments_series = assignments_series[assignments_series != ""]
        assignments_options = sorted(assignments_series.astype(str).unique())
        
        if len(assignments_options) == 0:
            st.warning("No assignments found in the assignments column.")
        else:
            selected_assignments = st.multiselect(
                "Select Assignments to filter", 
                assignments_options,
                help="Select one or more assignments to filter the data"
            )
            
            if selected_assignments:
                # Convert to string for comparison to handle mixed types
                mask = df[assignments_col].astype(str).isin(selected_assignments)
                filtered_df = df[mask]
                st.write(f"Filtered data ({len(filtered_df)} rows):")
            else:
                filtered_df = df
                st.write(f"All data ({len(filtered_df)} rows):")
            
            # Display filtered data
            try:
                st.dataframe(filtered_df)
            except Exception as e:
                st.error(f"Error displaying filtered data: {e}")
            
            # Excel download function
            def to_excel(df):
                output = BytesIO()
                try:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='FilteredData')
                    output.seek(0)
                    return output
                except Exception as e:
                    st.error(f"Error creating Excel file: {e}")
                    return None
            
            # Create download button
            excel_data = to_excel(filtered_df)
            if excel_data:
                st.download_button(
                    label="Download filtered Excel",
                    data=excel_data,
                    file_name="filtered_assignments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Display some statistics
            st.subheader("Data Statistics")
            st.write(f"Total rows in original data: {len(df)}")
            st.write(f"Total rows after filtering: {len(filtered_df)}")
            if assignments_col in df.columns:
                st.write(f"Unique assignments found: {len(assignments_options)}")