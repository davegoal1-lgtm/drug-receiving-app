# app.py
import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
from PIL import Image
import re
from difflib import SequenceMatcher
import base64
import json
from openai import OpenAI

st.set_page_config(page_title="藥品 HIS 驗收 OCR App", page_icon="💊", layout="wide")

st.title("💊 藥品 HIS 驗收 OCR App")
st.caption("藥品主檔｜採購單 Excel/OCR 匯入｜收貨單 OCR｜HIS 驗收欄位｜Excel 匯出")

for key, default in {
    "master_df": None,
    "po_df": None,
    "po_ocr_text": "",
    "po_ocr_df": None,
    "selected_po": None,
    "delivery_ocr_text": "",
    "records": []
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def clean_text(x):
    return str(x).replace(" ", "").replace("\n", "").lower()


def similarity(a, b):
    return SequenceMatcher(None, clean_text(a), clean_text(b)).ratio()


def extract_by_patterns(text, patterns):
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(1)
    return ""


def extract_lot(text):
    return extract_by_patterns(text, [
        r"批號[:：\s]*([A-Za-z0-9\-]+)",
        r"LOT[:：\s]*([A-Za-z0-9\-]+)",
        r"Batch[:：\s]*([A-Za-z0-9\-]+)"
    ])


def extract_expiry(text):
    return extract_by_patterns(text, [
        r"有效日期[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"有效期限[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"效期[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"EXP[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})"
    ])


def extract_quantity(text):
    value = extract_by_patterns(text, [
        r"出貨數量[:：\s]*([0-9]+)",
        r"驗收數量[:：\s]*([0-9]+)",
        r"數量[:：\s]*([0-9]+)",
        r"QTY[:：\s]*([0-9]+)"
    ])
    return int(value) if value else 0


def extract_delivery_no(text):
    return extract_by_patterns(text, [
        r"出貨單號[:：\s]*([A-Za-z0-9\-]+)",
        r"收貨單號[:：\s]*([A-Za-z0-9\-]+)",
        r"送貨單號[:：\s]*([A-Za-z0-9\-]+)"
    ])
def extract_delivery_qty(text):
    patterns = [
        r"出貨數量[:：\s]*([0-9]+)",
        r"數量[:：\s]*([0-9]+)",
        r"QTY[:：\s]*([0-9]+)",
        r"([0-9]{1,5})\s*(SET|VIAL|AMP|TAB|CAP|BOX|BAG|BT|PK|支|盒|瓶|袋)"
    ]

    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return int(m.group(1))

    return 0


def extract_delivery_expiry(text):
    patterns = [
        r"有效日期[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"有效期限[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"效期[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"EXP[:：\s]*([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})",
        r"([0-9]{4}[/-][0-9]{1,2}[/-][0-9]{1,2})"
    ]

    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(1)

    return ""


def extract_delivery_lot(text):
    patterns = [
        r"批號[:：\s]*([A-Za-z0-9\-]+)",
        r"LOT[:：\s]*([A-Za-z0-9\-]+)",
        r"Lot[:：\s]*([A-Za-z0-9\-]+)",
        r"Batch[:：\s]*([A-Za-z0-9\-]+)"
    ]

    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(1)

    return ""


def match_delivery_drug_from_master(ocr_text):
    if "master_df" not in st.session_state:
        return None, 0

    if st.session_state["master_df"] is None:
        return None, 0

    master_df = st.session_state["master_df"].copy()

    best_row = None
    best_score = 0

    ocr_clean = clean_text(ocr_text)

    for _, row in master_df.iterrows():
        candidates = [
            row.get("品號", ""),
            row.get("標準品名", ""),
            row.get("品名", ""),
            row.get("學名", ""),
            row.get("中文藥名", ""),
            row.get("別名1", ""),
            row.get("別名2", "")
        ]

        for c in candidates:
            c = str(c).strip()
            if not c or c == "nan":
                continue

            if clean_text(c) in ocr_clean:
                score = 1.0
            else:
                score = similarity(ocr_text, c)

            if score > best_score:
                best_score = score
                best_row = row

    return best_row, best_score

def ai_parse_delivery_image(image_file):
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

    image_bytes = image_file.getvalue()
    image_b64 = base64.b64encode(image_bytes).decode("utf-8")

    prompt = """
你是醫院藥品驗收單 OCR 專家。
請從收貨單/藥品購買憑證圖片中擷取資料。

只回傳 JSON，不要解釋。

欄位：
drug_name: 藥品名稱
quantity: 出貨數量，數字
lot: 批號
expiry: 有效日期或效期，格式 YYYY/MM/DD
delivery_no: 收貨單號或出貨單號
vendor: 廠商名稱

如果看不到就填空字串或 0。
"""

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {
                        "type": "input_image",
                        "image_url": f"data:image/jpeg;base64,{image_b64}"
                    }
                ]
            }
        ]
    )

    text = response.output_text.strip()

    text = text.replace("```json", "").replace("```", "").strip()

    try:
        return json.loads(text)
    except Exception:
        return {
            "drug_name": "",
            "quantity": 0,
            "lot": "",
            "expiry": "",
            "delivery_no": "",
            "vendor": "",
            "raw": text
        }


def match_ai_drug_name_to_master(drug_name):
    if "master_df" not in st.session_state:
        return None, 0

    if st.session_state["master_df"] is None:
        return None, 0

    master_df = st.session_state["master_df"].copy()

    best_row = None
    best_score = 0

    for _, row in master_df.iterrows():
        candidates = [
            row.get("標準品名", ""),
            row.get("品名", ""),
            row.get("學名", ""),
            row.get("中文藥名", ""),
            row.get("別名1", ""),
            row.get("別名2", "")
        ]

        for c in candidates:
            c = str(c).strip()

            if not c or c == "nan":
                continue

            score = similarity(drug_name, c)

            if clean_text(c) in clean_text(drug_name) or clean_text(drug_name) in clean_text(c):
                score = max(score, 1.0)

            if score > best_score:
                best_score = score
                best_row = row

    return best_row, best_score




def ocr_image(image):
    import pytesseract
    return pytesseract.image_to_string(image, lang="eng+chi_tra")


def get_received_total(po_no, item_code):
    return sum(
        int(r["驗收數量"])
        for r in st.session_state.records
        if str(r["採購單號"]) == str(po_no) and str(r["品號"]) == str(item_code)
    )


def make_status(order_qty, received_qty):
    if received_qty == 0:
        return "待驗收"
    if received_qty < order_qty:
        return "部分到貨"
    if received_qty == order_qty:
        return "已完成"
    return "超收異常"


def find_best_match(ocr_text, po_items):
    best_row = None
    best_score = 0

    for _, row in po_items.iterrows():
        candidates = [
            row.get("品號", ""),
            row.get("藥名", ""),
            row.get("標準藥名", ""),
            row.get("學名", ""),
            row.get("別名1", ""),
            row.get("別名2", "")
        ]

        score = 0
        for c in candidates:
            if str(c).strip():
                if clean_text(c) in clean_text(ocr_text):
                    score = max(score, 1.0)
                else:
                    score = max(score, similarity(ocr_text, c))

        if score > best_score:
            best_score = score
            best_row = row

    return best_row, best_score


def parse_po_ocr(text):
    po_no = extract_by_patterns(text, [
        r"(BB[0-9]{6,})"
    ])

    supplier = extract_by_patterns(text, [
        r"(.+?股份有限公司)",
        r"(.+?有限公司)"
    ])

    if "master_df" not in st.session_state:
        st.error("請先載入藥品主檔")
        return pd.DataFrame()

    if st.session_state["master_df"] is None:
        st.error("請先載入藥品主檔")
        return pd.DataFrame()

    if st.session_state["master_df"].empty:
        st.error("藥品主檔是空的")
        return pd.DataFrame()

    master_df = st.session_state["master_df"].copy()
    master_df["品號"] = master_df["品號"].astype(str).str.strip()

    clean_ocr = re.sub(r"\s+", " ", text)

    items = []

    for _, row in master_df.iterrows():
        code = str(row["品號"]).strip()

        if not code:
            continue

        code_match = re.search(re.escape(code), clean_ocr, re.IGNORECASE)

        if not code_match:
            continue

        segment = clean_ocr[code_match.end():code_match.end() + 180]

        qty_match = re.search(
            r"([0-9]{1,5})\s*(SET|VIAL|AMP|TAB|CAP|盒|支|瓶)",
            segment,
            re.IGNORECASE
        )

        if qty_match:
            qty = int(qty_match.group(1))
        else:
            nums = re.findall(r"\b[0-9]{1,5}\b", segment)
            nums = [int(n) for n in nums if int(n) not in [1, 2, 3, 4, 5, 30, 100, 2000]]

            if nums:
                qty = nums[-1]
            else:
                qty = 0

        items.append({
            "採購單號": po_no,
            "品號": code,
            "品名": row["標準藥名"],
            "標準品名": row["標準藥名"],
            "學名": row.get("學名", ""),
            "別名1": row.get("別名1", ""),
            "別名2": row.get("別名2", ""),
            "採購數量": qty,
            "廠商": supplier,
            "已驗收數量": 0,
            "狀態": "待驗收"
        })

    df = pd.DataFrame(items)

    if df.empty:
        st.warning("採購單 OCR 沒抓到任何主檔品號")
        return df

    df = df.drop_duplicates(subset=["品號"])

    return df
tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "⓪ 藥品主檔",
    "① 匯入採購單",
    "② 選擇採購單",
    "③ 收貨單 OCR",
    "④ HIS 驗收確認",
    "⑤ 匯出"
])

with tab0:
    st.subheader("⓪ 藥品主檔")
    uploaded_master = st.file_uploader(
    "上傳藥品主檔 Excel",
    type=["xlsx", "xls", "csv"]
)

if uploaded_master is not None:

    if uploaded_master.name.endswith(".csv"):
        uploaded_master_df = pd.read_csv(uploaded_master)
    else:
        uploaded_master_df = pd.read_excel(uploaded_master)

    st.dataframe(uploaded_master_df.head(20), use_container_width=True)

    cols = uploaded_master_df.columns.tolist()

    c1, c2, c3 = st.columns(3)

    with c1:
        code_col = st.selectbox("品號欄位", cols)

    with c2:
        name_col = st.selectbox("標準藥名欄位", cols)

    with c3:
        generic_col = st.selectbox("學名欄位", ["無"] + cols)

    c4, c5 = st.columns(2)

    with c4:
        alias1_col = st.selectbox("別名1欄位", ["無"] + cols)

    with c5:
        alias2_col = st.selectbox("別名2欄位", ["無"] + cols)

    if st.button("載入藥品主檔"):
        master_df = uploaded_master_df.copy()

        master_df["品號"] = master_df[code_col].astype(str).str.strip()
        master_df["標準藥名"] = master_df[name_col].astype(str)

        if generic_col != "無":
            master_df["學名"] = master_df[generic_col].astype(str)
        else:
            master_df["學名"] = ""

        if alias1_col != "無":
            master_df["別名1"] = master_df[alias1_col].astype(str)
        else:
            master_df["別名1"] = ""

        if alias2_col != "無":
            master_df["別名2"] = master_df[alias2_col].astype(str)
        else:
            master_df["別名2"] = ""

        st.session_state["master_df"] = master_df

        st.success("藥品主檔載入完成")

    if st.session_state.master_df is not None:
        st.dataframe(st.session_state.master_df, use_container_width=True)

with tab1:
    st.subheader("① 匯入採購單")

    import_mode = st.radio(
        "選擇匯入方式",
        ["Excel 匯入", "圖片 OCR 匯入"],
        horizontal=True
    )

    if import_mode == "Excel 匯入":
        po_file = st.file_uploader("上傳採購單 Excel", type=["xlsx", "xls"], key="po_excel")

        if po_file:
            raw = pd.read_excel(po_file)
            st.dataframe(raw.head(20), use_container_width=True)
            cols = raw.columns.tolist()

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                po_col = st.selectbox("採購單號欄位", cols)
            with c2:
                item_col = st.selectbox("品號欄位", cols)
            with c3:
                drug_col = st.selectbox("採購單藥名欄位", cols)
            with c4:
                qty_col = st.selectbox("採購數量欄位", cols)

            supplier_col = st.selectbox("廠商欄位", ["無"] + cols)

            if st.button("載入 BB 採購單"):
                df = raw.copy()
                df = df[df[po_col].astype(str).str.startswith("BB")].copy()

                if df.empty:
                    st.error("找不到 BB 開頭採購單")
                else:
                    df = df.rename(columns={
                        po_col: "採購單號",
                        item_col: "品號",
                        drug_col: "藥名",
                        qty_col: "採購數量"
                    })
                    df["品號"] = df["品號"].astype(str).str.strip()
                    df["採購數量"] = pd.to_numeric(df["採購數量"], errors="coerce").fillna(0).astype(int)

                    if supplier_col != "無":
                        df = df.rename(columns={supplier_col: "廠商"})
                    else:
                        df["廠商"] = ""

                    if st.session_state.master_df is not None:
                        df = df.merge(st.session_state.master_df, on="品號", how="left")
                        df["標準藥名"] = df["標準藥名"].fillna(df["藥名"])
                        df["學名"] = df["學名"].fillna("")
                        df["別名1"] = df["別名1"].fillna("")
                        df["別名2"] = df["別名2"].fillna("")
                    else:
                        df["標準藥名"] = df["藥名"]
                        df["學名"] = ""
                        df["別名1"] = ""
                        df["別名2"] = ""

                    df["已驗收數量"] = 0
                    df["狀態"] = "待驗收"

                    st.session_state.po_df = df[[
                        "採購單號", "品號", "藥名", "標準藥名", "學名",
                        "別名1", "別名2", "採購數量", "廠商", "已驗收數量", "狀態"
                    ]]
                    st.success("Excel 採購單已載入")

    else:
        po_img = st.file_uploader("上傳採購單圖片", type=["png", "jpg", "jpeg"], key="po_img")

        if po_img:
            image = Image.open(po_img)
            st.image(image, caption="採購單圖片", use_container_width=True)

            if st.button("開始辨識採購單 OCR"):
                try:
                    text = ocr_image(image)
                    st.session_state.po_ocr_text = text
                    st.session_state.po_ocr_df = parse_po_ocr(text)
                    st.success("採購單 OCR 完成")
                except Exception as e:
                    st.error("採購單 OCR 失敗")
                    st.warning("Streamlit Cloud 請確認 packages.txt 已包含 tesseract-ocr 與 tesseract-ocr-chi-tra")
                    st.code(str(e))

        if st.session_state.po_ocr_text:
            st.text_area("採購單 OCR 原文", value=st.session_state.po_ocr_text, height=220)

        if st.session_state.po_ocr_df is not None:
            st.markdown("### OCR 解析出的採購單明細")
            edited_df = st.data_editor(
                st.session_state.po_ocr_df,
                use_container_width=True,
                num_rows="dynamic"
            )

            if st.button("確認載入 OCR 採購單"):
                if edited_df.empty:
                    st.error("沒有可載入的採購單明細")
                else:
                    st.session_state.po_df = edited_df.copy()
                    st.success("OCR 採購單已載入")

with tab2:
    st.subheader("② 選擇採購單")
    if st.session_state.po_df is None:
        st.warning("請先匯入採購單")
    else:
        po_list = st.session_state.po_df["採購單號"].astype(str).unique().tolist()
        selected_po = st.selectbox("選擇採購單號", po_list)
        st.session_state.selected_po = selected_po

        df = st.session_state.po_df[st.session_state.po_df["採購單號"].astype(str) == selected_po].copy()

        rows = []
        for _, row in df.iterrows():
            received = get_received_total(row["採購單號"], row["品號"])
            row["已驗收數量"] = received
            row["狀態"] = make_status(int(row["採購數量"]), received)
            rows.append(row)

        st.dataframe(pd.DataFrame(rows), use_container_width=True)

with tab3:
    st.subheader("③ 收貨單 OCR")

    ocr_file = st.file_uploader(
        "上傳收貨單圖片",
        type=["png", "jpg", "jpeg"],
        key="delivery_img"
    )

    if ocr_file:
        image = Image.open(ocr_file)
        st.image(image, caption="收貨單", use_container_width=True)

        if st.button("開始收貨單 OCR"):
            try:
                st.session_state.delivery_ocr_text = ocr_image(image)
                st.success("收貨單 OCR 完成")
            except Exception as e:
                st.error("收貨單 OCR 失敗")
                st.warning("請確認 packages.txt 與 requirements.txt 設定正確")
                st.code(str(e))

    if st.session_state.delivery_ocr_text:
        text = st.session_state.delivery_ocr_text

        st.text_area(
            "收貨單 OCR 原文",
            value=text,
            height=250
        )

        matched_drug, drug_score = match_delivery_drug_from_master(text)

        if matched_drug is not None and drug_score >= 0.35:
            st.success(
                f"辨識藥品：{matched_drug['品號']}｜{matched_drug['標準品名']}｜相似度 {drug_score:.2f}"
            )
        else:
            st.warning("收貨單 OCR 尚未明確對應到藥品主檔")

        st.info(f"批號：{extract_delivery_lot(text) or '未抓到'}")
        st.info(f"有效日期：{extract_delivery_expiry(text) or '未抓到'}")
        st.info(f"出貨數量：{extract_delivery_qty(text) or '未抓到'}")
        st.info(f"收貨單號：{extract_delivery_no(text) or '未抓到'}")

with tab4:
    st.subheader("④ HIS 驗收確認")

    if st.session_state.po_df is None or st.session_state.selected_po is None:
        st.warning("請先匯入並選擇採購單")
    else:
        po_items = st.session_state.po_df[
            st.session_state.po_df["採購單號"].astype(str) == str(st.session_state.selected_po)
        ].copy()

        matched_row, score = find_best_match(st.session_state.delivery_ocr_text, po_items) if st.session_state.delivery_ocr_text else (None, 0)

delivery_matched_drug, delivery_drug_score = match_delivery_drug_from_master(st.session_state.delivery_ocr_text)

if delivery_matched_drug is not None and delivery_drug_score >= 0.35:
    matched_code = str(delivery_matched_drug.get("品號", ""))

    po_match = po_items[
        po_items["品號"].astype(str) == matched_code
    ]

    if not po_match.empty:
        matched_row = po_match.iloc[0]
        score = delivery_drug_score
        

        if matched_row is not None and score >= 0.30:
            st.success(f"建議品項：{matched_row['品號']}｜{matched_row['標準藥名']}｜相似度 {score:.2f}")
            default_code = str(matched_row["品號"])
        else:
            default_code = str(po_items.iloc[0]["品號"])
            st.warning("請手動選擇品項")

        options = [
            f"{r['品號']}｜{r['標準藥名']}｜採購量:{r['採購數量']}"
            for _, r in po_items.iterrows()
        ]

        idx = next((i for i, opt in enumerate(options) if opt.startswith(default_code)), 0)
        selected_item = st.selectbox("選擇驗收品項", options, index=idx)
        selected_code = selected_item.split("｜")[0]
        row = po_items[po_items["品號"].astype(str) == selected_code].iloc[0]

        st.markdown("### HIS 紅色必填欄位")

        c1, c2, c3 = st.columns(3)
        with c1:
            his_po = st.text_input("採購單號", value=str(row["採購單號"]))
        with c2:
            his_code = st.text_input("品號", value=str(row["品號"]))
        with c3:
            his_supplier = st.text_input("廠商", value=str(row["廠商"]))

        st.text_input("藥名", value=str(row["標準藥名"]), disabled=True)

        c4, c5, c6 = st.columns(3)
        with c4:
            delivery_no = st.text_input("收貨單號", value=extract_delivery_no(st.session_state.delivery_ocr_text))
        with c5:
            lot_no = st.text_input("批號", value=extract_lot(st.session_state.delivery_ocr_text))
        with c6:
            expiry = st.text_input("有效日期", value=extract_expiry(st.session_state.delivery_ocr_text))

        c7, c8, c9 = st.columns(3)
        with c7:
            receive_qty = st.number_input(
                "驗收數量",
                min_value=0,
                step=1,
                value=extract_quantity(st.session_state.delivery_ocr_text) or 0
            )
        with c8:
            inspector = st.text_input("驗收人")
        with c9:
            inspect_date = st.date_input("驗收日期", value=date.today())

        note = st.text_area("備註")

        already = get_received_total(his_po, his_code)
        after = already + int(receive_qty)
        status = make_status(int(row["採購數量"]), after)

        st.write(f"採購數量：{int(row['採購數量'])}")
        st.write(f"已驗收數量：{already}")
        st.write(f"驗收後累計：{after}")
        st.write(f"狀態：**{status}**")

        if status == "超收異常":
            st.error("⚠️ 超過採購數量，請確認")

        if st.button("確認寫入驗收紀錄"):
            if receive_qty <= 0:
                st.error("驗收數量需大於 0")
            elif not lot_no or not expiry:
                st.error("批號與有效日期為必填")
            else:
                st.session_state.records.append({
                    "驗收時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "採購單號": his_po,
                    "品號": his_code,
                    "藥名": row["標準藥名"],
                    "廠商": his_supplier,
                    "收貨單號": delivery_no,
                    "批號": lot_no,
                    "有效日期": expiry,
                    "驗收數量": int(receive_qty),
                    "驗收人": inspector,
                    "驗收日期": str(inspect_date),
                    "採購數量": int(row["採購數量"]),
                    "累計驗收數量": after,
                    "狀態": status,
                    "備註": note,
                    "收貨單OCR原文": st.session_state.delivery_ocr_text
                })
                st.success("已寫入驗收紀錄")
                st.rerun()

with tab5:
    st.subheader("⑤ 匯出 HIS 格式")

    if not st.session_state.records:
        st.info("目前尚無驗收紀錄")
    else:
        result_df = pd.DataFrame(st.session_state.records)
        st.dataframe(result_df, use_container_width=True)

        his_cols = [
            "採購單號", "品號", "批號", "有效日期", "驗收數量",
            "廠商", "收貨單號", "驗收人", "驗收日期"
        ]

        his_df = result_df[his_cols]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            his_df.to_excel(writer, index=False, sheet_name="HIS驗收匯入格式")
            result_df.to_excel(writer, index=False, sheet_name="完整驗收紀錄")

        st.download_button(
            "📥 下載 HIS 驗收 Excel",
            data=output.getvalue(),
            file_name=f"HIS驗收匯入_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
