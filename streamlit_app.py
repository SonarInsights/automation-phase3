import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Pivot Table Generator with Download")

# Upload file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    try:
        sheet_name = st.text_input("Sheet name", value="Sheet1")
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        st.success("File loaded successfully!")
        st.write("Preview:", df.head())

        # Pilih field
        rows = st.multiselect("Rows", df.columns)
        columns = st.multiselect("Columns", df.columns)
        values = st.multiselect("Values", df.columns)
        aggfunc = st.selectbox("Aggregation Function", ["sum", "count", "mean"], index=0)

        # Optional filter
        filter_col = st.selectbox("Filter column (optional)", [""] + list(df.columns))
        filter_val = None
        if filter_col:
            filter_val = st.selectbox(f"Filter value for '{filter_col}'", [""] + list(df[filter_col].astype(str).unique()))

        pivot = None
        if st.button("Generate Pivot"):
            # Filter data
            if filter_col and filter_val:
                df = df[df[filter_col].astype(str) == filter_val]

            # Buat pivot
            pivot = pd.pivot_table(
                df,
                index=rows,
                columns=columns if columns else None,
                values=values,
                aggfunc=aggfunc,
                fill_value=0
            )
            st.subheader("Pivot Table Result")
            st.dataframe(pivot)

            # Simpan ke Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                pivot.to_excel(writer, sheet_name='Pivot Result')
            output.seek(0)

            # Tombol download
            st.download_button(
                label="📥 Download Pivot as Excel",
                data=output,
                file_name="pivot_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
