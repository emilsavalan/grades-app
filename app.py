# app.py

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Processor", layout="centered")
st.title("� Excel Upload and Processor")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Original Data", df)

    # Example manipulation
    df["New Column"] = df[df.columns[0]] * 2

    st.write("### Modified Data", df)

    # Download logic
    def to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False)
        processed_data = output.getvalue()
        return 


