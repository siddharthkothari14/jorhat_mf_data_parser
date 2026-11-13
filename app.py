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
    final_df_sip = pd.DataFrame()
    final_df_stp = pd.DataFrame()

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
    final_df_sip['Transaction Date'] = final_df['Transaction Date']
    final_df_stp['Transaction Date'] = final_df['Transaction Date']

    # Scheme
    if 'Scheme/From Scheme' in df.columns:
        final_df['Scheme Name'] = df['Scheme/From Scheme']
        final_df_sip['Scheme'] = df['Scheme/From Scheme']
        final_df_stp['Scheme'] = ''
    if {'Action', 'Scheme/From Scheme', 'To Scheme'}.issubset(df.columns):
        switch_rows = df['Action'].eq('SWITCH')
        final_df_stp.loc[switch_rows, 'Scheme'] = (
            df.loc[switch_rows, 'Scheme/From Scheme']
            + ' to '
            + df.loc[switch_rows, 'To Scheme']
        )
    # Client Name
    if 'Primary Holder Name' in df.columns:
        final_df['Client Name'] = df['Primary Holder Name']
        final_df_sip['Client Name'] = df['Primary Holder Name']
        final_df_stp['Client Name'] = df['Primary Holder Name']

    # Platform
    final_df['Platform'] = 'NSE INVEST'
    final_df_sip['Platform'] = 'NSE INVEST'
    final_df_stp['Platform'] = 'NSE INVEST'

    # PAN / Client Code
    if 'Client Code' in df.columns:
        final_df['IIN / PAN'] = df['Client Code']
        final_df_sip['IIN / PAN'] = df['Client Code']
        final_df_stp['IIN / CAN'] = df['Client Code']

    # Folio
    if 'Folio No' in df.columns:
        final_df['Folio No'] = df['Folio No']
        final_df_sip['Folio No'] = df['Folio No']
        final_df_stp['Folio No'] = df['Folio No']

    final_df['Initiator'] = ''
    final_df_sip['Initiator'] = ''
    final_df_stp['Initiator'] = ''
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
    final_df_sip['Amount'] = final_df['Amount']
    final_df_stp['STP Amount'] = final_df['Amount']

    final_df['SOA'] = ''
    final_df['Mint'] = ''
    final_df_sip['SOA'] = ''
    final_df_sip['Mint'] = ''
    final_df_sip['Daily/Monthly'] = ''
    final_df_sip['First SIP Date'] = ''
    final_df_sip['Second SIP Date'] = ''
    final_df_stp['SOA'] = ''
    final_df_stp['Mint'] = ''
    final_df_stp['Weekly / Monthly'] = ''
    final_df_stp['STP Tenure'] = ''
    final_df_stp['First STP Date'] = ''
    final_df_stp['Second STP Date'] = ''

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
    desired_columns_sip = [
        'Initiator',
        'Transaction Date',
        'Client Name',
        'Platform',
        'IIN / PAN',
        'Folio No',
        'Scheme',
        'Amount',
        'Daily/Monthly',
        'First SIP Date',
        'SOA',
        'Mint',
        'Second SIP Date'
    ]
    desired_columns_stp = [
        'Initiator',
        'Transaction Date',
        'Client Name',
        'Platform',
        'IIN / CAN',
        'Folio No',
        'Scheme',
        'Weekly / Monthly',
        'STP Tenure',
        'STP Amount',
        'First STP Date',
        'SOA',
        'Mint',
        'Second STP Date'
    ]

    def apply_schema(df, columns):
        for col in columns:
            if col not in df.columns:
                df[col] = np.nan

        return df[columns]

    final_df = apply_schema(final_df, desired_columns)
    final_df_sip = apply_schema(final_df_sip, desired_columns_sip)
    final_df_stp = apply_schema(final_df_stp, desired_columns_stp)

    return final_df, final_df_sip, final_df_stp, date


if uploaded_file is not None:

    try:
        output_df, output_df_sip, output_df_stp, date = process_file(uploaded_file)

        st.success("✅ File Parsed Successfully")

        st.subheader("Preview")
        st.dataframe(output_df, use_container_width=True)

        def make_excel_download(df, file_name):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            return buffer.getvalue()

        st.download_button(
            label="📥 Download Parsed Excel (Main)",
            data=make_excel_download(output_df, f"Parsed_Output_{date}.xlsx"),
            file_name=f"Parsed_Output_{date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="📥 Download SIP Excel",
            data=make_excel_download(output_df_sip, f"Parsed_SIP_{date}.xlsx"),
            file_name=f"Parsed_SIP_{date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="📥 Download STP Excel",
            data=make_excel_download(output_df_stp, f"Parsed_STP_{date}.xlsx"),
            file_name=f"Parsed_STP_{date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
