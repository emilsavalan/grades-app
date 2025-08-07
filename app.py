import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

# Set page config to wide mode
st.set_page_config(
    page_title="Filter Excel by Assignments",
    layout="wide"  # This makes the app use the full width of the browser
)

# Custom CSS to make it even wider if needed
st.markdown("""
    <style>
    .main .block-container {
        max-width: 95%;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    </style>
    """, unsafe_allow_html=True)

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
            
            # Display filtered data with custom height and width
            try:
                st.dataframe(
                    filtered_df, 
                    use_container_width=True,  # Make it use full width
                    height=800  # Set height to show more rows (approximately 40-50 rows)
                )
            except Exception as e:
                st.error(f"Error displaying filtered data: {e}")
            
            # Check for duplicates in Email Address column
            email_col = None
            for col_name in selected_headers:
                if col_name and "email" in col_name.lower():
                    email_col = col_name
                    break
            
            if email_col is None:
                st.warning("Email Address column not found. Download is disabled.")
                allow_download = False
            else:
                # Find duplicates
                duplicates_mask = filtered_df.duplicated(subset=[email_col], keep=False)
                duplicated_df = filtered_df[duplicates_mask]
                
                if len(duplicated_df) > 0:
                    st.warning(f"⚠️ Found {len(duplicated_df)} rows with duplicate Email Addresses!")
                    
                    # Group duplicates by email address
                    duplicate_groups = duplicated_df.groupby(email_col)
                    
                    st.subheader("Duplicate Email Addresses - Select one row from each group:")
                    
                    selected_indices = []
                    final_df = filtered_df[~duplicates_mask].copy()  # Start with non-duplicates
                    
                    for email, group in duplicate_groups:
                        st.write(f"**Email: {email}** ({len(group)} duplicates)")
                        
                        # Display the duplicate group with highlighting
                        st.dataframe(
                            group.style.highlight_max(axis=0, color='lightcoral'),
                            use_container_width=True
                        )
                        
                        # Let user select which row to keep
                        options = []
                        for idx, row in group.iterrows():
                            # Create a summary of the row for selection
                            summary_parts = []
                            for col in group.columns[:3]:  # Show first 3 columns as summary
                                val = str(row[col])[:30] if row[col] is not None else "None"
                                summary_parts.append(f"{col}: {val}")
                            summary = " | ".join(summary_parts)
                            options.append((idx, f"Row {idx}: {summary}"))
                        
                        selected_idx = st.selectbox(
                            f"Select row to keep for {email}:",
                            options=[opt[0] for opt in options],
                            format_func=lambda x: next(opt[1] for opt in options if opt[0] == x),
                            key=f"select_{email}"
                        )
                        
                        if selected_idx is not None:
                            selected_indices.append(selected_idx)
                            final_df = pd.concat([final_df, group.loc[[selected_idx]]])
                    
                    # Check if all duplicates are resolved
                    if len(selected_indices) == len(duplicate_groups):
                        st.success("✅ All duplicates resolved! You can now download the file.")
                        allow_download = True
                        final_filtered_df = final_df.reset_index(drop=True)
                        st.write(f"Final data ({len(final_filtered_df)} rows after removing duplicates):")
                        st.dataframe(final_filtered_df, use_container_width=True, height=400)
                    else:
                        st.error("❌ Please select one row from each duplicate group before downloading.")
                        allow_download = False
                        final_filtered_df = filtered_df
                else:
                    st.success("✅ No duplicate Email Addresses found!")
                    allow_download = True
                    final_filtered_df = filtered_df
            
            # Excel download function
            def to_excel(df, title):
                output = BytesIO()
                try:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Write the dataframe starting from row 2, column A (no empty columns)
                        df.to_excel(writer, index=False, sheet_name='FilteredData', startrow=1, startcol=0)
                        
                        # Get the workbook and worksheet to add the title
                        workbook = writer.book
                        worksheet = writer.sheets['FilteredData']
                        
                        # Add the title in cell A1
                        if title:
                            worksheet.cell(row=1, column=1, value=title)
                    
                    output.seek(0)
                    return output
                except Exception as e:
                    st.error(f"Error creating Excel file: {e}")
                    return None
            
            # Create download button (only if duplicates are resolved)
            if allow_download:
                excel_data = to_excel(final_filtered_df, title_cell)
                if excel_data:
                    st.download_button(
                        label="Download filtered Excel",
                        data=excel_data,
                        file_name="filtered_assignments.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("❌ Cannot download: Please resolve all duplicate Email Addresses first.")
            
            # Display some statistics
            st.subheader("Data Statistics")
            st.write(f"Total rows in original data: {len(df)}")
            st.write(f"Total rows after filtering: {len(filtered_df)}")
            if allow_download and 'final_filtered_df' in locals():
                st.write(f"Final rows after duplicate removal: {len(final_filtered_df)}")
            if assignments_col in df.columns:
                st.write(f"Unique assignments found: {len(assignments_options)}")