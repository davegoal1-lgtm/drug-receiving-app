import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from PIL import Image
import re
from difflib import SequenceMatcher

st.set_page_config(page_title="藥品採購驗收 OCR App", page_icon="💊", layout="wide")

st.title("💊 藥品採購驗收 OCR App")
st.caption("採購單匯入、選擇採購單號、整張明細顯示、收貨單OCR核對、驗收紀錄匯出")

if "po_df" not in st.session_state:
    st.session_state.po_df = None

if "selected_po" not in st.session_state:
    st.session_state.selected_po = None

if "ocr_text" not in st.session_state:
    st.session_state.ocr_text = ""

if "records" not in st.session_state:
    st.session_state.records = []


def clean_text(x):
    return str(x).replace(" ", "").replace("\n", "").lower()


def similarity(a, b):
    return SequenceMatcher(None, clean_text(a), clean_text(b)).ratio()


def extract_lot(text):
    patterns = [
        r"批號[:：\s]*([A-Za-z0-9\-]+)",
        r"LOT[:：\s]*([A-Za-z0-9\-]+)",
        r"Lot[:：\s]*([A-Za-z0-9\-]+)",
        r"Batch[:：\s]*([A-Za-z0-9\-]+)"
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(1)
    return ""


def extract_expiry(text):
    patterns = [
        r"有效日期[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"有效期限[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"效期[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"EXP[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"Exp[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})"
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(1)
    return ""


def extract_quantity(text):
    patterns = [
        r"出貨數量[:：\s]*([0-9]+)",
        r"採購數量[:：\s]*([0-9]+)",
        r"數量[:：\s]*([0-9]+)",
        r"QTY[:：\s]*([0-9]+)",
        r"Qty[:：\s]*([0-9]+)"
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return int(m.group(1))
    return 0


def get_received_total(po_no, item_code):
    total = 0
    for r in st.session_state.records:
        if str(r["採購單號"]) == str(po_no) and str(r["品號"]) == str(item_code):
            total += int(r["本次驗收數量"])
    return total


def make_status(order_qty, received_qty):
    if received_qty == 0:
        return "待驗收"
    elif received_qty < order_qty:
        return "部分到貨"
    elif received_qty == order_qty:
        return "已完成"
    else:
        return "超收異常"


def find_best_match(ocr_text, po_items):
    best_row = None
    best_score = 0

    for _, row in po_items.iterrows():
        name = str(row["藥名"])
        code = str(row["品號"])

        score_name = similarity(ocr_text, name)
        score_code = similarity(ocr_text, code)
        score = max(score_name, score_code)

        if clean_text(name) in clean_text(ocr_text):
            score = 1.0

        if clean_text(code) in clean_text(ocr_text):
            score = 1.0

        if score > best_score:
            best_score = score
            best_row = row

    return best_row, best_score


tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "① 匯入採購單",
    "② 選擇採購單",
    "③ 收貨單 OCR",
    "④ 核對驗收",
    "⑤ 匯出紀錄"
])


with tab1:
    st.subheader("① 匯入採購單 Excel")

    po_file = st.file_uploader("上傳採購單 Excel", type=["xlsx", "xls"])

    if po_file is not None:
        raw_df = pd.read_excel(po_file)
        st.markdown("### 原始資料預覽")
        st.dataframe(raw_df.head(20), use_container_width=True)

        cols = raw_df.columns.tolist()

        st.markdown("### 請指定欄位")
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            po_col = st.selectbox("採購單號欄位", cols)

        with c2:
            item_col = st.selectbox("品號欄位", cols)

        with c3:
            name_col = st.selectbox("藥名欄位", cols)

        with c4:
            qty_col = st.selectbox("採購數量欄位", cols)

        supplier_col = st.selectbox("供應商欄位（可選）", ["無"] + cols)

        if st.button("✅ 載入 BB 採購單"):
            df = raw_df.copy()
            df = df[df[po_col].astype(str).str.startswith("BB")].copy()

            if df.empty:
                st.error("找不到 BB 開頭的採購單號")
            else:
                df = df.rename(columns={
                    po_col: "採購單號",
                    item_col: "品號",
                    name_col: "藥名",
                    qty_col: "採購數量"
                })

                if supplier_col != "無":
                    df = df.rename(columns={supplier_col: "供應商"})
                else:
                    df["供應商"] = ""

                df["採購數量"] = pd.to_numeric(df["採購數量"], errors="coerce").fillna(0).astype(int)
                df["已驗收數量"] = 0
                df["狀態"] = "待驗收"

                st.session_state.po_df = df[[
                    "採購單號", "品號", "藥名", "採購數量", "供應商", "已驗收數量", "狀態"
                ]]

                st.success("BB 採購單已成功載入")


with tab2:
    st.subheader("② 選擇採購單號並顯示整張採購單")

    if st.session_state.po_df is None:
        st.warning("請先到 Tab 1 匯入採購單")
    else:
        po_list = st.session_state.po_df["採購單號"].astype(str).unique().tolist()

        selected_po = st.selectbox("選擇採購單號", po_list)

        st.session_state.selected_po = selected_po

        selected_df = st.session_state.po_df[
            st.session_state.po_df["採購單號"].astype(str) == str(selected_po)
        ].copy()

        updated_rows = []
        for _, row in selected_df.iterrows():
            total_received = get_received_total(row["採購單號"], row["品號"])
            row["已驗收數量"] = total_received
            row["狀態"] = make_status(int(row["採購數量"]), total_received)
            updated_rows.append(row)

        selected_df = pd.DataFrame(updated_rows)

        st.markdown(f"### 採購單明細：{selected_po}")
        st.dataframe(selected_df, use_container_width=True)

        st.info("之後收貨單 OCR 只會比對這張採購單內的品項。")


with tab3:
    st.subheader("③ 收貨單 OCR")

    ocr_file = st.file_uploader("上傳收貨單圖片", type=["png", "jpg", "jpeg"])

    if ocr_file is not None:
        image = Image.open(ocr_file)
        st.image(image, caption="收貨單圖片", use_container_width=True)

        if st.button("🔍 開始 OCR 辨識"):
            try:
                import pytesseract
                text = pytesseract.image_to_string(image, lang="eng+chi_tra")
                st.session_state.ocr_text = text
                st.success("OCR 辨識完成")
            except Exception as e:
                st.error("OCR 辨識失敗")
                st.warning("Mac 本機請先安裝：brew install tesseract tesseract-lang")
                st.code(str(e))

    if st.session_state.ocr_text:
        st.markdown("### OCR 原始文字")
        st.text_area("辨識結果", value=st.session_state.ocr_text, height=260)

        st.markdown("### 自動抓取")
        c1, c2, c3 = st.columns(3)

        with c1:
            st.info(f"批號：{extract_lot(st.session_state.ocr_text) or '未抓到'}")

        with c2:
            st.info(f"效期：{extract_expiry(st.session_state.ocr_text) or '未抓到'}")

        with c3:
            qty = extract_quantity(st.session_state.ocr_text)
            st.info(f"數量：{qty if qty else '未抓到'}")


with tab4:
    st.subheader("④ 核對驗收")

    if st.session_state.po_df is None:
        st.warning("請先匯入採購單")
    elif st.session_state.selected_po is None:
        st.warning("請先到 Tab 2 選擇採購單號")
    else:
        selected_po = st.session_state.selected_po

        po_items = st.session_state.po_df[
            st.session_state.po_df["採購單號"].astype(str) == str(selected_po)
        ].copy()

        st.markdown(f"### 目前核對採購單：{selected_po}")

        matched_row = None
        match_score = 0

        if st.session_state.ocr_text:
            matched_row, match_score = find_best_match(st.session_state.ocr_text, po_items)

        if matched_row is not None and match_score >= 0.30:
            st.success(
                f"系統建議對應：{matched_row['品號']}｜{matched_row['藥名']}｜相似度 {match_score:.2f}"
            )
            default_code = str(matched_row["品號"])
        else:
            st.warning("尚無明確對應品項，請手動選擇")
            default_code = str(po_items.iloc[0]["品號"])

        item_options = [
            f"{row['品號']}｜{row['藥名']}｜採購量:{row['採購數量']}"
            for _, row in po_items.iterrows()
        ]

        default_index = 0
        for i, opt in enumerate(item_options):
            if opt.startswith(default_code):
                default_index = i
                break

        selected_item_text = st.selectbox(
            "選擇本次驗收品項",
            item_options,
            index=default_index
        )

        selected_item_code = selected_item_text.split("｜")[0]
        selected_row = po_items[
            po_items["品號"].astype(str) == str(selected_item_code)
        ].iloc[0]

        auto_qty = extract_quantity(st.session_state.ocr_text)
        auto_lot = extract_lot(st.session_state.ocr_text)
        auto_expiry = extract_expiry(st.session_state.ocr_text)

        c1, c2, c3 = st.columns(3)

        with c1:
            receive_qty = st.number_input(
                "本次驗收數量",
                min_value=0,
                step=1,
                value=auto_qty if auto_qty else 0
            )

        with c2:
            lot_no = st.text_input("批號", value=auto_lot)

        with c3:
            expiry_date = st.text_input("效期", value=auto_expiry)

        note = st.text_area("備註")

        already_received = get_received_total(selected_po, selected_item_code)
        after_total = already_received + int(receive_qty)
        status = make_status(int(selected_row["採購數量"]), after_total)

        st.markdown("### 驗收後狀態")
        st.write(f"採購數量：{int(selected_row['採購數量'])}")
        st.write(f"已驗收數量：{already_received}")
        st.write(f"本次驗收數量：{int(receive_qty)}")
        st.write(f"驗收後累計：{after_total}")
        st.write(f"狀態：**{status}**")

        if status == "超收異常":
            st.error("⚠️ 驗收數量超過採購數量，請確認")

        if st.button("✅ 確認驗收"):
            if receive_qty <= 0:
                st.error("本次驗收數量需大於 0")
            else:
                st.session_state.records.append({
                    "驗收時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "採購單號": selected_po,
                    "品號": selected_row["品號"],
                    "藥名": selected_row["藥名"],
                    "採購數量": int(selected_row["採購數量"]),
                    "本次驗收數量": int(receive_qty),
                    "累計驗收數量": after_total,
                    "批號": lot_no,
                    "效期": expiry_date,
                    "供應商": selected_row["供應商"],
                    "狀態": status,
                    "備註": note,
                    "OCR原文": st.session_state.ocr_text
                })

                st.success("已完成本次驗收")
                st.rerun()


with tab5:
    st.subheader("⑤ 匯出紀錄")

    if not st.session_state.records:
        st.info("目前尚無驗收紀錄")
    else:
        result_df = pd.DataFrame(st.session_state.records)
        st.dataframe(result_df, use_container_width=True)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, sheet_name="驗收紀錄")

            if st.session_state.po_df is not None:
                po_export = st.session_state.po_df.copy()

                updated_rows = []
                for _, row in po_export.iterrows():
                    total_received = get_received_total(row["採購單號"], row["品號"])
                    row["已驗收數量"] = total_received
                    row["狀態"] = make_status(int(row["採購數量"]), total_received)
                    updated_rows.append(row)

                pd.DataFrame(updated_rows).to_excel(
                    writer,
                    index=False,
                    sheet_name="採購單狀態"
                )

        st.download_button(
            label="📥 下載驗收 Excel",
            data=output.getvalue(),
            file_name=f"藥品驗收紀錄_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("🗑 清空驗收紀錄"):
            st.session_state.records = []
            st.rerun()
