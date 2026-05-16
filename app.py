{\rtf1\ansi\ansicpg950\cocoartf2868
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 import streamlit as st\
import pandas as pd\
from datetime import datetime\
from io import BytesIO\
\
st.set_page_config(page_title="\uc0\u34277 \u21697 \u39511 \u25910  App", layout="wide")\
\
st.title("\uc0\u55357 \u56458  \u34277 \u21697 \u39511 \u25910  App")\
st.caption("\uc0\u34277 \u21697 \u25475 \u30908  / \u25163 \u21205 \u36664 \u20837  / \u20027 \u27284 \u23565 \u29031  / \u39511 \u25910 \u32000 \u37636 \u21295 \u20986 ")\
\
# \uc0\u21021 \u22987 \u21270 \u36039 \u26009 \
if "records" not in st.session_state:\
    st.session_state.records = []\
\
st.sidebar.header("\uc0\u55357 \u56514  \u34277 \u21697 \u20027 \u27284 \u19978 \u20659 ")\
master_file = st.sidebar.file_uploader(\
    "\uc0\u19978 \u20659 \u34277 \u21697 \u20027 \u27284  Excel / CSV",\
    type=["xlsx", "csv"]\
)\
\
master_df = None\
\
if master_file is not None:\
    if master_file.name.endswith(".csv"):\
        master_df = pd.read_csv(master_file)\
    else:\
        master_df = pd.read_excel(master_file)\
\
    st.sidebar.success("\uc0\u20027 \u27284 \u24050 \u19978 \u20659 ")\
    st.sidebar.dataframe(master_df.head())\
\
st.header("\uc0\u10133  \u26032 \u22686 \u39511 \u25910 \u36039 \u26009 ")\
\
col1, col2, col3 = st.columns(3)\
\
with col1:\
    barcode = st.text_input("\uc0\u26781 \u30908  / \u34277 \u21697 \u20195 \u30908 ")\
    receive_date = st.date_input("\uc0\u39511 \u25910 \u26085 \u26399 ", datetime.today())\
\
with col2:\
    batch_no = st.text_input("\uc0\u25209 \u34399 ")\
    expire_date = st.date_input("\uc0\u25928 \u26399 ")\
\
with col3:\
    quantity = st.number_input("\uc0\u25976 \u37327 ", min_value=0, step=1)\
    supplier = st.text_input("\uc0\u20379 \u25033 \u21830 ")\
\
drug_name = ""\
drug_code = ""\
\
if master_df is not None and barcode:\
    possible_cols = master_df.columns.tolist()\
\
    code_col = st.sidebar.selectbox("\uc0\u20027 \u27284 \u65306 \u34277 \u21697 \u20195 \u30908 \u27396 \u20301 ", possible_cols)\
    name_col = st.sidebar.selectbox("\uc0\u20027 \u27284 \u65306 \u34277 \u21697 \u21517 \u31281 \u27396 \u20301 ", possible_cols)\
\
    match = master_df[master_df[code_col].astype(str) == str(barcode)]\
\
    if not match.empty:\
        drug_code = str(match.iloc[0][code_col])\
        drug_name = str(match.iloc[0][name_col])\
        st.success(f"\uc0\u25214 \u21040 \u34277 \u21697 \u65306 \{drug_name\}")\
    else:\
        st.warning("\uc0\u20027 \u27284 \u25214 \u19981 \u21040 \u27492 \u20195 \u30908 \u65292 \u21487 \u25163 \u21205 \u36664 \u20837 ")\
\
drug_code = st.text_input("\uc0\u34277 \u21697 \u20195 \u30908 ", value=drug_code)\
drug_name = st.text_input("\uc0\u34277 \u21697 \u21517 \u31281 ", value=drug_name)\
\
note = st.text_area("\uc0\u20633 \u35387 ")\
\
if st.button("\uc0\u9989  \u26032 \u22686 \u39511 \u25910 \u32000 \u37636 "):\
    if not drug_code or not drug_name:\
        st.error("\uc0\u35531 \u36664 \u20837 \u34277 \u21697 \u20195 \u30908 \u33287 \u34277 \u21697 \u21517 \u31281 ")\
    else:\
        st.session_state.records.append(\{\
            "\uc0\u39511 \u25910 \u26178 \u38291 ": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),\
            "\uc0\u39511 \u25910 \u26085 \u26399 ": receive_date.strftime("%Y-%m-%d"),\
            "\uc0\u34277 \u21697 \u20195 \u30908 ": drug_code,\
            "\uc0\u34277 \u21697 \u21517 \u31281 ": drug_name,\
            "\uc0\u25209 \u34399 ": batch_no,\
            "\uc0\u25928 \u26399 ": expire_date.strftime("%Y-%m-%d"),\
            "\uc0\u25976 \u37327 ": quantity,\
            "\uc0\u20379 \u25033 \u21830 ": supplier,\
            "\uc0\u26781 \u30908 ": barcode,\
            "\uc0\u20633 \u35387 ": note\
        \})\
        st.success("\uc0\u24050 \u26032 \u22686 \u39511 \u25910 \u32000 \u37636 ")\
\
st.header("\uc0\u55357 \u56523  \u20170 \u26085 \u39511 \u25910 \u32000 \u37636 ")\
\
if st.session_state.records:\
    df = pd.DataFrame(st.session_state.records)\
    st.dataframe(df, use_container_width=True)\
\
    output = BytesIO()\
    with pd.ExcelWriter(output, engine="openpyxl") as writer:\
        df.to_excel(writer, index=False, sheet_name="\uc0\u39511 \u25910 \u32000 \u37636 ")\
\
    st.download_button(\
        label="\uc0\u11015 \u65039  \u21295 \u20986  Excel",\
        data=output.getvalue(),\
        file_name=f"\uc0\u34277 \u21697 \u39511 \u25910 \u32000 \u37636 _\{datetime.now().strftime('%Y%m%d_%H%M')\}.xlsx",\
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"\
    )\
\
    if st.button("\uc0\u55357 \u56785 \u65039  \u28165 \u31354 \u30446 \u21069 \u32000 \u37636 "):\
        st.session_state.records = []\
        st.rerun()\
else:\
    st.info("\uc0\u30446 \u21069 \u23578 \u28961 \u39511 \u25910 \u32000 \u37636 ")}