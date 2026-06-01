import streamlit as st
import pandas as pd
import datetime
import json
import gspread
from google.oauth2.service_account import Credentials
import os
import base64
import re

st.set_page_config(page_title="발주서 및 거래명세서 작성", page_icon="📝", layout="wide")

# 1. 구글 시트 연결
@st.cache_resource
def init_connection():
    creds_info = json.loads(st.secrets["google_credentials"])
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def fetch_mother_data():
    try:
        client = init_connection()
        doc = client.open("석미_마더데이터")
        try:
            sheet = doc.worksheet("발주내역")
        except:
            sheet = doc.sheet1
            
        all_data = sheet.get_all_values()
        if not all_data: return pd.DataFrame()
        return pd.DataFrame(all_data[1:], columns=all_data[0])
    except:
        return pd.DataFrame()

# 3개의 탭(품목, 매입, 매출)을 불러오는 함수
@st.cache_data(ttl=60)
def load_google_sheet_data():
    df_items, df_purch, df_sales = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    try:
        client = init_connection()
        doc = client.open("석미_마더데이터")
        for worksheet in doc.worksheets():
            title = worksheet.title
            data = worksheet.get_all_values()
            if not data or len(data) < 2: continue
            
            df = pd.DataFrame(data[1:], columns=data[0])
            if '품목' in title: df_items = df
            elif '매입' in title: df_purch = df
            elif '매출' in title: df_sales = df
    except Exception as e:
        pass
    return df_items, df_purch, df_sales

st.title("📝 발주서 및 거래명세서 작성")

today = datetime.date.today()
month_list = [(pd.to_datetime(today) + pd.DateOffset(months=i)).strftime("%Y-%m") for i in range(-12, 61)]

# 세션 초기화
keys = [
    ('f_order_no', ''), ('f_close_month', month_list[12]), ('f_date', today), ('f_due_date', today),
    ('f_due_time', ''), ('f_sales_v', ''), ('f_site', ''), ('f_manager', ''),
    ('f_phone', ''), ('f_purch_v', ''), ('f_address', ''),
    ('f_ship_cost', ''), ('f_ship_car', ''), ('f_ship_driver', ''), ('f_ship_phone', ''),
    ('f_depot', ''), ('f_sender', ''), ('f_sender_phone', ''), ('f_receiver', ''), ('f_po_note', '')
]
for k, v in keys:
    if k not in st.session_state: st.session_state[k] = v

df_history = fetch_mother_data()
df_items, df_purch, df_sales = load_google_sheet_data()

# --- 0. 과거 내역 불러오기 ---
st.subheader("0. 과거 발주내역 불러오기")
if not df_history.empty and '발주일' in df_history.columns:
    df_temp = df_history.sort_values(by='발주일', ascending=False).copy()
    if '주문번호' in df_temp.columns:
        df_temp['OrderKey'] = df_temp['주문번호'].astype(str) + " | " + df_temp['납품처'].astype(str) + " | " + df_temp['현장명'].astype(str)
    else:
        df_temp['OrderKey'] = df_temp['발주일'].astype(str) + " | " + df_temp['납품처'].astype(str) + " | " + df_temp['현장명'].astype(str)
    
    unique_orders = df_temp['OrderKey'].drop_duplicates().tolist()
    
    with st.expander("🔄 기존 내역 검색 및 양식 채우기"):
        c_sel, c_btn = st.columns([4, 1])
        with c_sel:
            sel_order = st.selectbox("불러올 내역 (주문번호/발주일 | 납품처 | 현장명)", ["선택안함"] + unique_orders, label_visibility="collapsed")
        with c_btn:
            if st.button("📥 불러오기", use_container_width=True) and sel_order != "선택안함":
                target_df = df_temp[df_temp['OrderKey'] == sel_order]
                if not target_df.empty:
                    first_row = target_df.iloc[0]
                    st.session_state.f_close_month = first_row.get('마감월', st.session_state.f_close_month)
                    try: st.session_state.f_date = pd.to_datetime(first_row.get('발주일')).date()
                    except: pass
                    try: st.session_state.f_due_date = pd.to_datetime(first_row.get('납기일')).date()
                    except: pass
                    for k, col in [('f_due_time', '납기시간'), ('f_sales_v', '납품처'), ('f_site', '현장명'), 
                                 ('f_manager', '담당(수령인)'), ('f_phone', '수령인전화'), ('f_purch_v', '매입업체'), 
                                 ('f_address', '도착지주소'), ('f_ship_cost', '운임'), ('f_ship_car', '차량번호'), 
                                 ('f_ship_driver', '기사명'), ('f_ship_phone', '기사연락처'), ('f_depot', '출고지'), 
                                 ('f_sender', '출고자'), ('f_sender_phone', '출고자전화'), ('f_receiver', '인수자'), ('f_po_note', '특이사항')]:
                        st.session_state[k] = str(first_row.get(col, ''))
                        
                    items_cols = [c for c in ['품목', '규격', '수량', '단위', '색상', '가공', 'KS', '비고', '매입단가', '매출단가', '매입업체'] if c in target_df.columns]
                    new_items = target_df[items_cols].copy()
                    if '매입업체' not in new_items.columns: new_items['매입업체'] = st.session_state.f_purch_v
                    
                    new_items['수량'] = pd.to_numeric(new_items.get('수량', 1), errors='coerce').fillna(1)
                    new_items['매입단가'] = pd.to_numeric(new_items.get('매입단가', 0), errors='coerce').fillna(0)
                    new_items['매출단가'] = pd.to_numeric(new_items.get('매출단가', 0), errors='coerce').fillna(0)
                    
                    for c in ["품목", "규격", "수량", "단위", "색상", "가공", "KS", "비고", "매입단가", "매출단가", "매입업체"]:
                        if c not in new_items.columns: new_items[c] = ""
                        
                    st.session_state.order_items = new_items[["품목", "규격", "수량", "단위", "색상", "가공", "KS", "비고", "매입단가", "매출단가", "매입업체"]]
                    st.rerun()

# --- 2. 기본 정보 입력창 ---
st.subheader("1. 기본 정보")

today_str = today.strftime("%y%m%d")
seq = 1
if not df_history.empty and '주문번호' in df_history.columns:
    today_orders = df_history[df_history['주문번호'].astype(str).str.startswith(today_str)]
    if not today_orders.empty:
        try:
            seqs = today_orders['주문번호'].str.split('-').str[-1].astype(int)
            seq = seqs.max() + 1
        except: pass
default_order_no = f"{today_str}-{seq:02d}"

f_order_no = st.text_input("주문번호", value=st.session_state.f_order_no if st.session_state.f_order_no else default_order_no)

r1c1, r1c2, r1c3, r1c4 = st.columns(4)
with r1c1: f_close_month = st.selectbox("마감월", month_list, index=month_list.index(st.session_state.f_close_month) if st.session_state.f_close_month in month_list else 12)
with r1c2: f_date = st.date_input("발주일", st.session_state.f_date)
with r1c3: f_due_date = st.date_input("납기일", st.session_state.f_due_date)
with r1c4: f_due_time = st.text_input("납기시간", value=st.session_state.f_due_time)

r2c1, r2c2, r2c3, r2c4 = st.columns(4)
with r2c1: f_sales_v = st.text_input("납품처(매출업체)", value=st.session_state.f_sales_v)
with r2c2: f_site = st.text_input("현장명", value=st.session_state.f_site)
with r2c3: f_manager = st.text_input("담당(수령인)", value=st.session_state.f_manager)
with r2c4: f_phone = st.text_input("수령인전화", value=st.session_state.f_phone)

r3c1, r3c2 = st.columns([3, 1])
with r3c1: f_address = st.text_input("도착지주소", value=st.session_state.f_address)
with r3c2: f_purch_v = st.text_input("기본 매입업체", value=st.session_state.f_purch_v)


# --- 3. 품목 상세 입력창 ---
st.subheader("2. 품목 상세")

# 면적 계산 함수 (안전망 단위 변환용)
def extract_area(spec_str):
    nums = [float(x) for x in re.findall(r'(\d+(?:\.\d+)?)', str(spec_str))]
    if len(nums) >= 2: return nums[0] * nums[1]
    elif len(nums) == 1: return nums[0]
    return 1.0

# 품목 옵션 추출
item_options, spec_options, unit_options = [], [], []
if not df_items.empty and '품목' in df_items.columns:
    item_options = sorted([str(x) for x in df_items['품목'].dropna().unique() if str(x).strip()])
if not item_options: item_options = ["안전망2cm(방염)", "안전망1cm", "멀티망"]

if 'order_items' not in st.session_state:
    st.session_state.order_items = pd.DataFrame(columns=["품목", "규격", "수량", "단위", "색상", "가공", "KS", "비고", "매입단가", "매출단가", "매입업체"])

st.markdown("#### 🔹 품목 추가 (품목과 규격을 선택하면 관련 정보와 단가가 자동 입력됩니다)")
with st.container(border=True):
    c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
    with c1: 
        sel_item = st.selectbox("1️⃣ 품목 선택", [""] + item_options + ["[직접 입력]"])
        in_item = st.text_input("품목 직접입력", label_visibility="collapsed") if sel_item == "[직접 입력]" else sel_item
    
    # 선택된 품목에 맞는 규격만 필터링
    spec_options = []
    if sel_item and sel_item != "[직접 입력]" and not df_items.empty and '규격' in df_items.columns:
        spec_options = sorted([str(x) for x in df_items[df_items['품목'] == sel_item]['규격'].dropna().unique() if str(x).strip()])
        
    with c2:
        sel_spec = st.selectbox("2️⃣ 규격 선택", [""] + spec_options + ["[직접 입력]"])
        in_spec = st.text_input("규격 직접입력", label_visibility="collapsed") if sel_spec == "[직접 입력]" else sel_spec

    # VLOOKUP 로직 시작
    auto_unit, auto_color, auto_proc, auto_ks = "롤", "", "", ""
    auto_p_price, auto_s_price = 0, 0
    
    if in_item:
        # 1. 품목정보 탭 매칭
        if not df_items.empty:
            m_item = df_items[df_items['품목'] == in_item]
            if in_spec: m_item = m_item[m_item['규격'] == in_spec]
            if not m_item.empty:
                r = m_item.iloc[0]
                auto_unit = str(r.get('단위', auto_unit))
                auto_color = str(r.get('색상', ''))
                auto_proc = str(r.get('가공', ''))
                auto_ks = str(r.get('KS', ''))
        
        # 안전망 면적 배수 설정 (해배 -> 롤 단가 변환용)
        multiplier = 1.0
        if '안전망' in in_item or '멀티망' in in_item:
            multiplier = extract_area(in_spec)
        
        # 2. 매입단가 탭 매칭 (기본 매입업체 기준)
        if not df_purch.empty and f_purch_v in df_purch.columns:
            p_item_col = next((c for c in df_purch.columns if '품목' == c.strip()), None)
            p_spec_col = next((c for c in df_purch.columns if '규격' == c.strip()), None)
            if p_item_col:
                p_match = df_purch[df_purch[p_item_col] == in_item]
                if p_spec_col and in_spec: p_match = p_match[p_match[p_spec_col] == in_spec]
                if not p_match.empty:
                    val = p_match.iloc[0][f_purch_v]
                    try: 
                        auto_p_price = int(float(str(val).replace(',', '')) * multiplier) # 롤 단위 변환
                    except: pass
        
        # 3. 매출업체 탭 매칭 (납품처 기준)
        if not df_sales.empty and f_sales_v in df_sales.columns:
            s_item_col = next((c for c in df_sales.columns if '품목' == c.strip()), None)
            s_spec_col = next((c for c in df_sales.columns if '규격' == c.strip()), None)
            if s_item_col:
                s_match = df_sales[df_sales[s_item_col] == in_item]
                if s_spec_col and in_spec: s_match = s_match[s_match[s_spec_col] == in_spec]
                if not s_match.empty:
                    val = s_match.iloc[0][f_sales_v]
                    try: 
                        auto_s_price = int(float(str(val).replace(',', '')) * multiplier) # 롤 단위 변환
                    except: pass

    with c3: in_qty = st.number_input("수량", min_value=1, value=1)
    with c4: in_unit = st.text_input("단위", value=auto_unit if auto_unit != "nan" else "롤")
    
    c5, c6, c7, c8 = st.columns(4)
    with c5: in_color = st.text_input("색상", value=auto_color if auto_color != "nan" else "")
    with c6: in_proc = st.text_input("가공", value=auto_proc if auto_proc != "nan" else "")
    with c7: in_ks = st.text_input("KS", value=auto_ks if auto_ks != "nan" else "")
    with c8: in_note = st.text_input("비고", value="")
    
    c9, c10, c11, c12 = st.columns([1.5, 1.5, 1.5, 1.5])
    with c9: in_p_price = st.number_input("매입단가 (롤당)", min_value=0, value=auto_p_price, step=100)
    with c10: in_s_price = st.number_input("매출단가 (롤당)", min_value=0, value=auto_s_price, step=100)
    with c11: in_vendor = st.text_input("매입업체 지정", value=f_purch_v)
    with c12:
        st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
        if st.button("➕ 표에 추가", use_container_width=True, type="primary"):
            if in_item.strip():
                new_row = pd.DataFrame([{
                    "품목": in_item, "규격": in_spec, "수량": in_qty, "단위": in_unit,
                    "색상": in_color, "가공": in_proc, "KS": in_ks, "비고": in_note,
                    "매입단가": in_p_price, "매출단가": in_s_price, "매입업체": in_vendor
                }])
                st.session_state.order_items = pd.concat([st.session_state.order_items, new_row], ignore_index=True)
                st.rerun()
            else:
                st.warning("품목을 선택해주세요.")

# --- 메인 데이터 표 ---
edited_df = st.data_editor(
    st.session_state.order_items,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_order=["품목", "규격", "수량", "단위", "색상", "가공", "KS", "비고", "매입단가", "매출단가", "매입업체"]
)
st.session_state.order_items = edited_df.copy()

st.markdown("**[발주서] 특이사항**")
f_po_note = st.text_area("특이사항", value=st.session_state.f_po_note, label_visibility="collapsed")

st.divider()

# --- 4. 운송 정보 ---
st.subheader("3. 운송정보")
r4c1, r4c2, r4c3, r4c4 = st.columns(4)
with r4c1: f_ship_cost = st.text_input("운임", value=st.session_state.f_ship_cost)
with r4c2: f_ship_car = st.text_input("차량번호", value=st.session_state.f_ship_car)
with r4c3: f_ship_driver = st.text_input("기사명", value=st.session_state.f_ship_driver)
with r4c4: f_ship_phone = st.text_input("기사연락처", value=st.session_state.f_ship_phone)

r5c1, r5c2, r5c3, r5c4 = st.columns(4)
with r5c1: f_depot = st.text_input("출고지", value=st.session_state.f_depot)
with r5c2: f_sender = st.text_input("출고자", value=st.session_state.f_sender)
with r5c3: f_sender_phone = st.text_input("출고자 전화", value=st.session_state.f_sender_phone)
with r5c4: f_receiver = st.text_input("인수자", value=st.session_state.f_receiver)

SUPPLIER_INFO = {"company": "석미세이프", "biznum": "524-38-00469", "address": "경기도 남양주시 수동면 남가로 1771-1"}
st.divider()

# --- 5. 저장 및 PDF 발행 ---
st.subheader("4. 장부 저장 및 PDF 통합 발행")

if 'is_saved' not in st.session_state: st.session_state.is_saved = False

if st.button("💾 장부 저장 및 PDF 다운로드", type="primary", use_container_width=True):
    valid_df = edited_df[edited_df['품목'].notna() & (edited_df['품목'].astype(str).str.strip() != "")].copy()
    if valid_df.empty:
        st.error("⚠️ 품목을 하나 이상 입력해주세요.")
    else:
        try:
            with st.spinner("구글 시트에 저장 중입니다..."):
                client = init_connection()
                doc = client.open("석미_마더데이터")
                try: sheet = doc.worksheet("발주내역")
                except: sheet = doc.add_worksheet(title="발주내역", rows="1000", cols="30")
                
                expected_headers = [
                    '주문번호', '마감월', '발주일', '납기일', '납기시간', '납품처', '현장명', '담당(수령인)', '수령인전화',
                    '도착지주소', '매입업체',
                    '품목', '규격', '수량', '단위', '색상', '가공', 'KS', '비고', 
                    '매입단가', '매출단가',
                    '운임', '차량번호', '기사명', '기사연락처', '출고지', '출고자', '출고자전화', '인수자', '특이사항'
                ]
                
                existing_data = sheet.get_all_values()
                if not existing_data or existing_data[0] != expected_headers:
                    if not existing_data: sheet.append_row(expected_headers)
                    else: sheet.insert_row(expected_headers, index=1)
                
                rows_to_append = []
                for _, row in valid_df.iterrows():
                    row_vendor = row.get('매입업체') if pd.notna(row.get('매입업체')) and str(row.get('매입업체')).strip() else f_purch_v
                    rows_to_append.append([
                        f_order_no, f_close_month, f_date.strftime("%Y-%m-%d"), f_due_date.strftime("%Y-%m-%d"),
                        f_due_time, f_sales_v, f_site, f_manager, f_phone, f_address, row_vendor,
                        row['품목'], row['규격'], row['수량'], row['단위'], row['색상'], row['가공'], row['KS'], row['비고'], 
                        row['매입단가'], row['매출단가'],
                        f_ship_cost, f_ship_car, f_ship_driver, f_ship_phone, f_depot, f_sender, f_sender_phone, f_receiver, f_po_note
                    ])
                    
                sheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
                st.success("✅ 저장 완료! 곧 PDF가 다운로드됩니다.")
                st.session_state.is_saved = True
        except Exception as e:
            st.error(f"저장 중 오류 발생: {e}")

# (이하 PDF 출력용 HTML/JS 코드는 기존과 동일하게 작동합니다)
TOTAL_ROWS = 14
tbody_html = ""
valid_rows = edited_df[edited_df['품목'].notna() & (edited_df['품목'].astype(str).str.strip() != "")]
for i, row in valid_rows.iterrows():
    tbody_html += f"""
    <tr>
        <td style='text-align:center; padding:0 4px; border:1px solid #000; height: 23px;'>{i+1}</td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'>{row.get('품목', '')}</td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'>{row.get('규격', '')}</td>
        <td style='text-align:center; padding:0 4px; border:1px solid #000; height: 23px;'>{row.get('수량', '')}</td>
        <td style='text-align:center; padding:0 4px; border:1px solid #000; height: 23px;'>{row.get('단위', '')}</td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'>{row.get('색상', '')} {row.get('가공', '')} {row.get('KS', '')}</td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'>{row.get('비고', '')}</td>
    </tr>
    """
for _ in range(max(0, TOTAL_ROWS - len(valid_rows))):
    tbody_html += "<tr>" + "<td style='border:1px solid #000; height: 23px;'></td>"*7 + "</tr>"

# (거래명세서 템플릿 생략 없이 전체 병합)
def create_ts_block(receiver_name):
    return f"""
    <div style="width: 48%; padding: 10px; box-sizing: border-box; font-family: 'Malgun Gothic', sans-serif;">
        <h1 style="text-align: center; letter-spacing: 15px; border-bottom: 2px solid #000;">거 래 명 세 서</h1>
        <table style="width: 100%; font-size: 11px;">
            <tr><td style="font-weight: bold;">납품처: {receiver_name}</td></tr>
            <tr><td style="font-weight: bold;">현장명: {f_site}</td></tr>
        </table>
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000; font-size: 11px; text-align: center;">
            <tr style="background-color: #f0f0f0;">
                <th style="border: 1px solid #000;">No</th><th style="border: 1px solid #000;">품목</th>
                <th style="border: 1px solid #000;">규격</th><th style="border: 1px solid #000;">수량</th>
                <th style="border: 1px solid #000;">단위</th><th style="border: 1px solid #000;">상세</th><th style="border: 1px solid #000;">비고</th>
            </tr>
            {tbody_html}
        </table>
    </div>
    """

ts_block = create_ts_block(f_sales_v)

auto_download_js = ""
if st.session_state.is_saved:
    auto_download_js = "window.onload = function() { setTimeout(downloadPDF, 500); };"
    st.session_state.is_saved = False

html_template = f"""
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
<button onclick="downloadPDF()">📥 PDF 수동 다운로드</button>
<div id="capture-area" style="width: 1050px; display: flex; background: #fff;">
    {ts_block} {ts_block}
</div>
<script>
    function downloadPDF() {{
        html2pdf().set({{
            filename: '거래명세서.pdf', image: {{ type: 'jpeg', quality: 1.0 }},
            html2canvas: {{ scale: 2 }}, jsPDF: {{ format: 'a4', orientation: 'landscape' }}
        }}).from(document.getElementById('capture-area')).save();
    }}
    {auto_download_js}
</script>
"""
st.components.v1.html(html_template, height=800)
