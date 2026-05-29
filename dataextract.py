import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Column Extractor", layout="wide")

st.title("Excel Column Extractor")

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        st.success(f"Loaded {len(df)} rows and {len(df.columns)} columns")

        st.subheader("Available Columns")
        st.write(list(df.columns))

        selected_cols = st.multiselect(
            "Select columns for output file",
            options=df.columns.tolist()
        )

        if selected_cols:
            output_df = df[selected_cols]

            st.subheader("Preview")
            st.dataframe(output_df.head())

            buffer = BytesIO()

            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                output_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Output"
                )

            st.download_button(
                label="Download Output Excel",
                data=buffer.getvalue(),
                file_name="filtered_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
