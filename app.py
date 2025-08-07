import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Filter App", layout="centered")
st.title("📋 Filter Excel by 'Assignments'")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Read Excel into DataFrame
        file_bytes = BytesIO(uploaded_file.read())
        df = pd.read_excel(file_bytes, engine='openpyxl')

        # Detect 'Assignments' column
        if "Assignments" not in df.columns:
            st.error("❌ The column 'Assignments' was not found in the Excel file.")
        else:
            # Get unique filter options
            options = sorted(df["Assignments"].dropna().unique())

            # Ask user to select one or more
            selected = st.multiselect("Select assignment(s):", options)

            # Filter based on selection
            if selected:
                filtered_df = df[df["Assignments"].isin(selected)]
                st.success(f"✅ Showing {len(filtered_df)} row(s) for selected assignments.")
                st.dataframe(filtered_df)
                
                # Excel download
                def to_excel(dataframe):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        dataframe.to_excel(writer, index=False)
                    output.seek(0)
                    return output

                st.download_button(
                    label="📥 Download Filtered Excel",
                    data=to_excel(filtered_df),
                    file_name="filtered_assignments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("☝️ Please select one or more assignments to display data.")

    except Exception as e:
        st.error(f"⚠️ Error reading file: {e}")
