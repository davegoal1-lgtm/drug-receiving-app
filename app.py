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


def build_his_confirm_df(ocr_result, selected_po_pool):
    rows = []

    drug_name = str(ocr_result.get("drug_name", "")).strip()
    qty = ocr_result.get("quantity", 0)
    lot = ocr_result.get("lot", "")
    expiry = ocr_result.get("expiry", "")
    vendor = ocr_result.get("vendor", "")

    matched_drug, score = match_ai_drug_name_to_master(drug_name)

    matched_code = ""
    if matched_drug is not None and score >= 0.35:
        matched_code = str(matched_drug.get("品號", ""))

    for _, po_row in selected_po_pool.iterrows():
        po_code = str(po_row.get("品號", "")).strip()
        is_match = matched_code != "" and po_code == matched_code

        base_row = {
            "辨識狀態": "已辨識完成" if is_match else "尚未辨識完成",
            "採購單號": po_row.get("採購單號", ""),
            "品號": po_row.get("品號", ""),
            "藥名": po_row.get("藥名", po_row.get("標準藥名", "")),
            "標準藥名": po_row.get("標準藥名", ""),
            "採購數量": po_row.get("採購數量", ""),
            "驗收數量": qty if is_match else "",
            "批號": lot if is_match else "",
            "有效日期": expiry if is_match else "",
            "廠商": vendor if is_match and vendor else po_row.get("廠商", "")
        }

        rows.append(base_row)

    return pd.DataFrame(rows)


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
        po_file = st.file_uploader("上傳採購單 Excel", type=["xlsx", "xls"], key="po_excel_upload_2")

        if po_file:
            raw = pd.read_excel(po_file)
            st.dataframe(raw.head(20), use_container_width=True)
            cols = raw.columns.tolist()

            # 自動抓固定欄位
            po_col = "採購單號"
            po_date_col = "採購日期"
            item_col = "品號"
            drug_col = "品名/規格"
            qty_col = "採購數量"
            supplier_col = "廠商名稱"

            # 檢查欄位是否存在
            
   
            required_cols = [po_col, po_date_col, supplier_col]
            missing = [c for c in required_cols if c not in cols]

            if missing:
                st.error(f"Excel 缺少欄位: {missing}")
                st.stop()
   
                

            if st.button("載入 BB 採購單"):
                df = raw.copy()
                df = df[df[po_col].astype(str).str.startswith("BB")].copy()

                if df.empty:
                    st.error("找不到 BB 開頭採購單")
                else:
                    df = df.rename(columns={
                        po_col: "採購單號",
                        po_date_col: "採購日期",
                        item_col: "品號",
                        drug_col: "藥名",
                        qty_col: "採購數量"
                    })

                df["品號"] = df["品號"].astype(str).str.strip()
                df["採購數量"] = pd.to_numeric(df["採購數量"], errors="coerce").fillna(0).astype(int)

                if supplier_col != "無" and supplier_col in df.columns:
                    df = df.rename(columns={supplier_col: "廠商"})
                else:
                    df["廠商"] = ""

                if st.session_state.master_df is not None:
                    df = df.merge(st.session_state.master_df, on="品號", how="left")

                if "標準藥名" not in df.columns:
                    df["標準藥名"] = df["藥名"]
                for col in ["學名", "別名1", "別名2"]:
                    if col not in df.columns:
                        df[col] = ""
                    else:
                        df[col] = df[col].fillna("")

                df["已驗收數量"] = 0
                df["狀態"] = "待驗收"

                st.session_state.po_df = df[[
                    "採購日期", "採購單號", "品號", "藥名", "標準藥名", "學名",
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
    st.subheader("② 選擇採購日期")

    if st.session_state.po_df is None:
        st.warning("請先匯入採購單")
    else:
        df = st.session_state.po_df.copy()
        df["採購日期"] = pd.to_datetime(df["採購日期"], errors="coerce").dt.date

        selected_date = st.date_input("選擇採購日期", value=date.today())

        date_df = df[df["採購日期"] == selected_date].copy()

        if date_df.empty:
            st.warning("這一天沒有採購單")
        else:
            st.markdown("### 當日採購單")

            date_df.insert(0, "選取", False)

            select_all = st.checkbox("全部選取")

            if select_all:
                date_df["選取"] = True

            edited_df = st.data_editor(
                date_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "選取": st.column_config.CheckboxColumn("選取")
                }
            )

            selected_po_pool = edited_df[
                edited_df["選取"] == True
            ].copy()

            st.info(f"已選取 {len(selected_po_pool)} 筆採購明細")

            if st.button("確認選取採購單"):
                if selected_po_pool.empty:
                    st.warning("請至少選一筆")
                else:
                    st.session_state["selected_po_pool"] = selected_po_pool
                    st.success("已儲存選取採購單，可進行收貨單 OCR")

            rows = []

            for _, row in selected_po_pool.iterrows():
                received = get_received_total(row["採購單號"], row["品號"])
                row["已驗收數量"] = received
                row["狀態"] = make_status(int(row["採購數量"]), received)
                rows.append(row)

            if rows:
                st.dataframe(pd.DataFrame(rows), use_container_width=True)
with tab3:
    st.subheader("③ 收貨單 AI OCR")

    ocr_file = st.file_uploader(
        "上傳收貨單圖片",
        type=["png", "jpg", "jpeg"],
        key="delivery_img"
    )

    if ocr_file:
        image = Image.open(ocr_file)
        st.image(image, caption="收貨單", use_container_width=True)

        if st.button("開始 AI 辨識收貨單"):
            try:
                ai_result = ai_parse_delivery_image(ocr_file)

                st.session_state.delivery_ai_result = ai_result
                st.session_state.delivery_ocr_text = json.dumps(
                    ai_result,
                    ensure_ascii=False,
                    indent=2
                )

                st.success("AI 收貨單辨識完成")

                selected_po_pool = st.session_state.get("selected_po_pool")

                if selected_po_pool is not None and not selected_po_pool.empty:
                    his_confirm_df = build_his_confirm_df(
                        ai_result,
                        selected_po_pool
                    )

                    st.session_state["his_confirm_df"] = his_confirm_df
                    st.success("已產生 HIS 驗收確認資料，請到第④步確認")

                else:
                    st.warning("請先到第②步選擇採購單")

            except Exception as e:
                st.error("AI 收貨單辨識失敗")
                st.code(str(e))

    if "delivery_ai_result" in st.session_state:
        result = st.session_state.delivery_ai_result

        st.markdown("### AI 辨識結果")
        st.json(result)

        drug_name = result.get("drug_name", "")
        qty = result.get("quantity", 0)
        lot = result.get("lot", "")
        expiry = result.get("expiry", "")
        delivery_no = result.get("delivery_no", "")
        vendor = result.get("vendor", "")

        matched_drug, score = match_ai_drug_name_to_master(drug_name)

        if matched_drug is not None and score >= 0.35:
            st.success(
                f"對應主檔：{matched_drug['品號']}｜相似度 {score:.2f}"
            )
        else:
            st.warning("尚未明確對應主檔")
with tab4:
    st.subheader("④ HIS 驗收確認")

    his_confirm_df = st.session_state.get("his_confirm_df")

    if his_confirm_df is None or his_confirm_df.empty:
        st.warning("尚未有辨識結果，請先完成收貨單 OCR")

    else:
        if "選取" not in his_confirm_df.columns:
            his_confirm_df.insert(0, "選取", False)

        required_cols = [
            "採購單號",
            "品號",
            "驗收數量",
            "批號",
            "有效日期",
            "廠商"
        ]

        def check_status(row):
            for col in required_cols:
                value = row.get(col)

                if pd.isna(value) or str(value).strip() == "":
                    return "尚未辨識完成"

            return "已辨識完成"

        his_confirm_df["辨識狀態"] = his_confirm_df.apply(
            check_status,
            axis=1
        )

        edited_df = st.data_editor(
            his_confirm_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic"
        )

        edited_df["辨識狀態"] = edited_df.apply(
            check_status,
            axis=1
        )

        complete_df = edited_df[
            edited_df["辨識狀態"] == "已辨識完成"
        ]

        incomplete_df = edited_df[
            edited_df["辨識狀態"] == "尚未辨識完成"
        ]

        st.success(f"已辨識完成：{len(complete_df)} 筆")
        st.warning(f"尚未辨識完成：{len(incomplete_df)} 筆")

        if st.button("確認寫入驗收紀錄"):
            checked_df = edited_df[
                edited_df["選取"] == True
            ].copy()

            if checked_df.empty:
                st.warning("請先勾選要寫入的資料")

            else:
                checked_df["辨識狀態"] = checked_df.apply(
                    check_status,
                    axis=1
                )

                incomplete_checked_df = checked_df[
                    checked_df["辨識狀態"] == "尚未辨識完成"
                ]

                if not incomplete_checked_df.empty:
                    st.error("勾選資料仍有未完成欄位")

                else:
                    for _, row in checked_df.iterrows():
                        st.session_state.records.append({
                            "驗收時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "採購單號": row.get("採購單號", ""),
                            "品號": row.get("品號", ""),
                            "藥名": row.get("藥名", ""),
                            "標準藥名": row.get("標準藥名", ""),
                            "採購數量": row.get("採購數量", ""),
                            "驗收數量": int(row.get("驗收數量", 0)),
                            "批號": row.get("批號", ""),
                            "有效日期": row.get("有效日期", ""),
                            "廠商": row.get("廠商", ""),
                            "驗收日期": str(date.today()),
                            "狀態": "已完成",
                        })

                    st.success("已寫入驗收紀錄")
                    st.rerun()
        st.markdown("### 已驗收完成紀錄")

        if st.session_state.records:
            done_df = pd.DataFrame(st.session_state.records)
            st.dataframe(done_df, use_container_width=True)
        else:
            st.info("目前尚無已寫入的驗收紀錄")
with tab5:
    st.subheader("⑤ 匯出 HIS 格式")

    if not st.session_state.records:
        st.info("目前尚無驗收紀錄")
    else:
        result_df = pd.DataFrame(st.session_state.records)
        st.dataframe(result_df, use_container_width=True)

        his_cols = [
            "採購單號", "品號", "批號", "有效日期", "驗收數量",
            "廠商", "驗收人", "驗收日期"
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
