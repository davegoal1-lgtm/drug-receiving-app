import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="藥品驗收 App",
    page_icon="💊",
    layout="wide"
)

st.title("💊 藥品驗收 App")
st.caption("藥品主檔查詢、驗收紀錄建立、Excel 匯出")

if "records" not in st.session_state:
    st.session_state.records = []

st.sidebar.header("📁 上傳藥品主檔")

master_file = st.sidebar.file_uploader(
    "請上傳藥品主檔 Excel",
    type=["xlsx", "xls"]
)

if master_file is not None:
    try:
        df_master = pd.read_excel(master_file)
        st.sidebar.success("藥品主檔上傳成功")

        st.subheader("📋 藥品主檔預覽")
        st.dataframe(df_master, use_container_width=True)

        columns = df_master.columns.tolist()

        st.sidebar.header("欄位設定")
        code_col = st.sidebar.selectbox("藥品代碼欄位", columns)
        name_col = st.sidebar.selectbox("藥品名稱欄位", columns)
        spec_col = st.sidebar.selectbox("規格/劑量欄位", columns)
        barcode_col = st.sidebar.selectbox("條碼欄位（可選）", ["無"] + columns)

        st.divider()

        st.subheader("🔎 藥品查詢與驗收")

        search_code = st.text_input("請輸入或掃描藥品代碼 / 條碼")

        result = pd.DataFrame()

        if search_code:
            search_code = str(search_code).strip()

            code_match = df_master[
                df_master[code_col].astype(str).str.contains(search_code, case=False, na=False)
            ]

            if barcode_col != "無":
                barcode_match = df_master[
                    df_master[barcode_col].astype(str).str.contains(search_code, case=False, na=False)
                ]
                result = pd.concat([code_match, barcode_match]).drop_duplicates()
            else:
                result = code_match

            if result.empty:
                st.warning("查無此藥品，請確認代碼或條碼")
            else:
                st.success(f"找到 {len(result)} 筆符合資料")
                st.dataframe(result, use_container_width=True)

                selected_index = st.selectbox(
                    "選擇要驗收的藥品",
                    result.index,
                    format_func=lambda x: f"{result.loc[x, code_col]} - {result.loc[x, name_col]}"
                )

                selected_drug = result.loc[selected_index]

                col1, col2, col3 = st.columns(3)

                with col1:
                    receive_qty = st.number_input("驗收數量", min_value=0, step=1)

                with col2:
                    batch_no = st.text_input("批號")

                with col3:
                    exp_date = st.date_input("效期")

                supplier = st.text_input("供應商 / 廠商")
                note = st.text_area("備註")

                if st.button("➕ 加入驗收紀錄"):
                    record = {
                        "驗收時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "藥品代碼": selected_drug[code_col],
                        "藥品名稱": selected_drug[name_col],
                        "規格/劑量": selected_drug[spec_col],
                        "驗收數量": receive_qty,
                        "批號": batch_no,
                        "效期": exp_date.strftime("%Y-%m-%d"),
                        "供應商": supplier,
                        "備註": note
                    }

                    st.session_state.records.append(record)
                    st.success("已加入驗收紀錄")

        st.divider()

        st.subheader("📦 今日驗收紀錄")

        if st.session_state.records:
            df_records = pd.DataFrame(st.session_state.records)
            st.dataframe(df_records, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_records.to_excel(writer, index=False, sheet_name="驗收紀錄")

            excel_data = output.getvalue()

            st.download_button(
                label="⬇️ 下載驗收紀錄 Excel",
                data=excel_data,
                file_name=f"藥品驗收紀錄_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if st.button("🗑 清空目前紀錄"):
                st.session_state.records = []
                st.rerun()
        else:
            st.info("尚無驗收紀錄")

    except Exception as e:
        st.error("讀取藥品主檔時發生錯誤")
        st.exception(e)

else:
    st.info("請先從左側上傳藥品主檔 Excel")
    st.markdown("""
    ### 藥品主檔建議欄位
    - 藥品代碼
    - 藥品名稱
    - 規格/劑量
    - 條碼
    - 廠牌
    - 包裝單位
    """)
