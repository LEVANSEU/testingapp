import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
import re

st.set_page_config(layout="wide")
st.markdown("""
    <style>
        body, .main, .block-container {
            background-color: white !important;
            color: #222 !important;
            font-family: 'Segoe UI', sans-serif;
        }
        h1, h2, h3, h4, h5, h6, .stMarkdown, .stText, .stTextLabelWrapper, label {
            color: #222 !important;
        }
        .stFileUploader, .stTextInput, .stSelectbox, .stRadio, .stButton, .stDataFrame,
        .stTextInput input, .stSelectbox div[data-baseweb="select"],
        .stSelectbox div[data-baseweb="select"] *,
        .stRadio div[role="radiogroup"] label,
        .stRadio div[role="radiogroup"] label * {
            background-color: #f5f5f5 !important;
            color: #222 !important;
            border-radius: 10px;
            font-size: 14px !important;
        }
        .stFileUploader {
            max-width: 600px !important;
            margin: 0 auto !important;
        }
        .stButton>button {
            background-color: #4CAF50;
            color: white !important;
            font-weight: bold;
            border: none;
            border-radius: 8px;
            padding: 6px 14px;
            font-size: 14px;
        }
        .stButton>button:hover {
            background-color: #45a049;
        }
        .summary-header {
            display: flex;
            font-weight: bold;
            margin-top: 1em;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #999;
            text-align: center;
            background-color: #f0f0f0;
            border-radius: 8px;
            color: #222 !important;
        }
        .summary-header div {
            flex: 1;
            padding: 0.5rem;
        }
        .number-cell {
            text-align: right !important;
            font-variant-numeric: tabular-nums;
            padding-right: 1rem;
            font-weight: bold;
            color: #222;
        }
        .stSelectbox > div, .stRadio > div, .stTextInput > div, .stButton > div {
            background-color: #f5f5f5 !important;
        }
        .stSelectbox > div *,
        .stRadio > div *,
        .stTextInput > div * {
            color: #222 !important;
        }
        .stRadio div[role="radiogroup"] label span {
            font-weight: bold !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("Excel გენერატორი")

report_file = st.file_uploader("📄 ატვირთე ანგარიშფაქტურების ფაილი (report.xlsx)", type=["xlsx"])
statement_files = st.file_uploader("📄 ატვირთე საბანკო ამონაწერის ფაილები (statement.xlsx)", type=["xlsx"], accept_multiple_files=True)

if report_file and statement_files:
    purchases_df = pd.read_excel(report_file, sheet_name='Grid')

    # Read and merge all statement files
    bank_dfs = []
    for file in statement_files:
        df = pd.read_excel(file)
        bank_dfs.append(df)
    bank_df = pd.concat(bank_dfs, ignore_index=True)

    purchases_df['დასახელება'] = purchases_df['გამყიდველი'].astype(str).apply(lambda x: re.sub(r'^\(\d+\)\s*', '', x).strip())
    purchases_df['საიდენტიფიკაციო კოდი'] = purchases_df['გამყიდველი'].apply(lambda x: ''.join(re.findall(r'\d', str(x)))[:11])
    bank_df['P'] = bank_df.iloc[:, 15].astype(str).str.strip()
    bank_df['Amount'] = pd.to_numeric(bank_df.iloc[:, 3], errors='coerce').fillna(0)

    missing_ids = bank_df[~bank_df['P'].isin(purchases_df['საიდენტიფიკაციო კოდი'])].copy()

    with st.expander("📌 ამონაწერის ჩანაწერები, სადაც საიდენტიფიკაციო კოდი ვერ მოიძებნა"):
        if not missing_ids.empty:
            missing_ids["ახალი კოდი"] = ""
            for i, row in missing_ids.iterrows():
                new_code = st.text_input(f"ჩაწერე საიდენტიფიკაციო კოდი ჩანაწერისთვის ({row['P']}):", key=f"missing_{i}")
                if new_code:
                    bank_df.at[i, 'P'] = new_code

    wb = Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet(title="ანგარიშფაქტურები კომპანიით")
    ws1.append(['დასახელება', 'საიდენტიფიკაციო კოდი', 'ანგარიშფაქტურების ჯამი', 'ჩარიცხული თანხა', 'სხვაობა'])

    company_summaries = []

    for company_id, group in purchases_df.groupby('საიდენტიფიკაციო კოდი'):
        company_name = group['დასახელება'].iloc[0]
        unique_invoices = group.groupby('სერია №')['ღირებულება დღგ და აქციზის ჩათვლით'].sum().reset_index()
        company_invoice_sum = unique_invoices['ღირებულება დღგ და აქციზის ჩათვლით'].sum()

        paid_sum = bank_df[bank_df["P"] == str(company_id)]["Amount"].sum()
        difference = company_invoice_sum - paid_sum

        ws1.append([company_name, company_id, company_invoice_sum, paid_sum, difference])
        company_summaries.append((company_name, company_id, company_invoice_sum, paid_sum, difference))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.markdown("### 📋 კომპანიების ჩამონათვალი")
    st.markdown("""
    <div class='summary-header'>
        <div style='flex: 2;'>დასახელება</div>
        <div style='flex: 2;'>საიდენტიფიკაციო კოდი</div>
        <div style='flex: 1.5;'>ინვოისების ჯამი</div>
        <div style='flex: 1.5;'>ჩარიცხვა</div>
        <div style='flex: 1.5;'>სხვაობა</div>
    </div>
    """, unsafe_allow_html=True)

    for name, company_id, invoice_sum, paid_sum, difference in company_summaries:
        col1, col2, col3, col4, col5 = st.columns([2, 2, 1.5, 1.5, 1.5])
        with col1:
            st.markdown(name)
        with col2:
            if st.button(f"{company_id}", key=f"id_{company_id}"):
                st.session_state['selected_company'] = company_id
        with col3:
            st.markdown(f"<div class='number-cell'>{invoice_sum:,.2f}</div>", unsafe_allow_html=True)
        with col4:
            st.markdown(f"<div class='number-cell'>{paid_sum:,.2f}</div>", unsafe_allow_html=True)
        with col5:
            st.markdown(f"<div class='number-cell'>{difference:,.2f}</div>", unsafe_allow_html=True)

    st.download_button(
        label="⬇️ ჩამოტვირთე Excel ფაილი",
        data=output,
        file_name="საბოლოო_ფაილი.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
