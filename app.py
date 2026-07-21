import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Jorhat Excel Parser", layout="wide")

st.title("Upload and Parse Excel File")

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls"]
)


def process_file(file):
    df = pd.read_excel(file)

    final_df = pd.DataFrame()

    # Transaction Date
    if 'Date & Timing' in df.columns:
        final_df['Transaction Date'] = (
            pd.to_datetime(
                df['Date & Timing'],
                format='%d/%m/%Y %H:%M:%S',
                errors='coerce'
            )
            .dt.strftime('%B %d, %Y')
        )
    date = final_df['Transaction Date'].iloc[0] if not final_df['Transaction Date'].empty else None
    # Empty Columns
    final_df['Transaction Type'] = ''

    # Scheme
    if 'Scheme/From Scheme' in df.columns:
        final_df['Scheme Name'] = df['Scheme/From Scheme']

    # Client Name
    if 'Primary Holder Name' in df.columns:
        final_df['Client Name'] = df['Primary Holder Name']

    # Platform
    final_df['Platform'] = 'NSE INVEST'

    # PAN / Client Code
    if 'Client Code' in df.columns:
        final_df['IIN / PAN'] = df['Client Code']

    # Folio
    if 'Folio No' in df.columns:
        final_df['Folio No'] = df['Folio No']

    final_df['Initiator'] = ''
    if 'Mode of Payment' in df.columns:
        final_df['Mode of Payment'] = df['Mode of Payment']

    # Amount
    if 'Amount/Units' in df.columns:
        final_df['Amount'] = (
            df['Amount/Units']
            .astype(str)
            .str.extract(r'([\d,.]+)', expand=False)
            .str.replace(",", "", regex=False)
            .astype(float)
            .apply(lambda x: f"{x:,.0f}")
        )

    final_df['SOA'] = ''
    final_df['Mint'] = ''

    desired_columns = [
        'Initiator',
        'Transaction Date',
        'Client Name',
        'Transaction Type',
        'Platform',
        'IIN / PAN',
        'Folio No',
        'Scheme Name',
        'Mode of Payment',
        'Amount',
        'SOA',
        'Mint'
    ]

    for col in desired_columns:
        if col not in final_df.columns:
            final_df[col] = np.nan

    final_df = final_df[desired_columns]

    return final_df,date


if uploaded_file is not None:

    try:
        output_df, date = process_file(uploaded_file)

        st.success("✅ File Parsed Successfully")

        st.subheader("Preview")
        st.dataframe(output_df, use_container_width=True)

        buffer = BytesIO()

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            output_df.to_excel(writer, index=False)

        st.download_button(
            label="📥 Download Parsed Excel",
            data=buffer.getvalue(),
            file_name=f"Parsed_Output_{date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
