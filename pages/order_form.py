import streamlit as st
import pandas as pd
import datetime
import json
import gspread
from google.oauth2.service_account import Credentials
import re

st.set_page_config(page_title="발주서 및 거래명세서 작성", page_icon="📝", layout="wide")

# 1. 구글 시트 연결 및 데이터 로드
@st.cache_resource
def init_connection():
    creds_info = json.loads(st.secrets["google_credentials"])
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_google_sheet_data():
    df_history, df_purch, df_sales = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    try:
        client = init_connection()
        doc = client.open("석미_마더데이터")
        for ws in doc.worksheets():
            data = ws.get_all_values()
            if not data or len(data) < 2: continue
            df = pd.DataFrame(data[1:], columns=data[0])
            if '발주내역' in ws.title: df_history = df
            elif '매입' in ws.title: df_purch = df
            elif '매출' in ws.title: df_sales = df
    except: pass
    return df_history, df_purch, df_sales

df_history, df_purch, df_sales = load_google_sheet_data()

st.title("📝 발주서 및 거래명세서 작성")

# 세션 초기화
today = datetime.date.today()
month_list = [(pd.to_datetime(today) + pd.DateOffset(months=i)).strftime("%Y-%m") for i in range(-12, 61)]
keys = [
    ('f_order_no', ''), ('f_close_month', month_list[12]), ('f_date', today), ('f_due_date', today),
    ('f_due_time', ''), ('f_sales_v', ''), ('f_site', ''), ('f_manager', ''),
    ('f_phone', ''), ('f_purch_v', ''), ('f_address', ''),
    ('f_ship_cost', ''), ('f_ship_car', ''), ('f_ship_driver', ''), ('f_ship_phone', ''),
    ('f_depot', ''), ('f_sender', ''), ('f_sender_phone', ''), ('f_receiver', ''), ('f_po_note', '')
]
for k, v in keys:
    if k not in st.session_state: st.session_state[k] = v

# --- 0. 과거 내역 불러오기 ---
st.subheader("0. 과거 발주내역 불러오기")

MAX_ROWS = 20
if 'row_cnt' not in st.session_state: st.session_state.row_cnt = 5

# 폼(Form)용 위젯 세션 초기화
for i in range(MAX_ROWS):
    if f"f_item_{i}" not in st.session_state: st.session_state[f"f_item_{i}"] = ""
    if f"f_spec_{i}" not in st.session_state: st.session_state[f"f_spec_{i}"] = ""
    if f"f_qty_{i}" not in st.session_state: st.session_state[f"f_qty_{i}"] = 1.0
    if f"f_unit_{i}" not in st.session_state: st.session_state[f"f_unit_{i}"] = "롤"
    if f"f_color_{i}" not in st.session_state: st.session_state[f"f_color_{i}"] = ""
    if f"f_proc_{i}" not in st.session_state: st.session_state[f"f_proc_{i}"] = ""
    if f"f_ks_{i}" not in st.session_state: st.session_state[f"f_ks_{i}"] = ""
    if f"f_note_{i}" not in st.session_state: st.session_state[f"f_note_{i}"] = ""
    if f"f_pprice_{i}" not in st.session_state: st.session_state[f"f_pprice_{i}"] = 0
    if f"f_sprice_{i}" not in st.session_state: st.session_state[f"f_sprice_{i}"] = 0
    if f"f_vendor_{i}" not in st.session_state: st.session_state[f"f_vendor_{i}"] = ""

if not df_history.empty and '발주일' in df_history.columns:
    df_temp = df_history.sort_values(by='발주일', ascending=False).copy()
    df_temp['OrderKey'] = df_temp['주문번호'].astype(str) + " | " + df_temp['납품처'].astype(str) + " | " + df_temp['현장명'].astype(str)
    
    with st.expander("🔄 기존 내역 검색 및 양식 채우기"):
        c_sel, c_btn = st.columns([4, 1])
        with c_sel:
            sel_order = st.selectbox("불러올 내역 선택", ["선택안함"] + df_temp['OrderKey'].drop_duplicates().tolist(), label_visibility="collapsed")
        with c_btn:
            if st.button("📥 일괄입력폼에 불러오기", use_container_width=True) and sel_order != "선택안함":
                target_df = df_temp[df_temp['OrderKey'] == sel_order]
                if not target_df.empty:
                    r = target_df.iloc[0]
                    st.session_state.f_close_month = r.get('마감월', st.session_state.f_close_month)
                    try: st.session_state.f_date = pd.to_datetime(r.get('발주일')).date()
                    except: pass
                    try: st.session_state.f_due_date = pd.to_datetime(r.get('납기일')).date()
                    except: pass
                    
                    for k, col in [('f_due_time', '납기시간'), ('f_sales_v', '납품처'), ('f_site', '현장명'), 
                                   ('f_manager', '담당(수령인)'), ('f_phone', '수령인전화'), ('f_purch_v', '매입업체'), 
                                   ('f_address', '도착지주소'), ('f_ship_cost', '운임'), ('f_ship_car', '차량번호'), 
                                   ('f_ship_driver', '기사명'), ('f_ship_phone', '기사연락처'), ('f_depot', '출고지'), 
                                   ('f_sender', '출고자'), ('f_sender_phone', '출고자전화'), ('f_receiver', '인수자'), ('f_po_note', '특이사항')]:
                        st.session_state[k] = str(r.get(col, ''))
                    
                    # 폼에 과거 내역 주입
                    st.session_state.row_cnt = min(MAX_ROWS, max(5, len(target_df) + 2))
                    for i in range(MAX_ROWS):
                        if i < len(target_df):
                            row_data = target_df.iloc[i]
                            st.session_state[f"f_item_{i}"] = str(row_data.get('품목', ''))
                            st.session_state[f"f_spec_{i}"] = str(row_data.get('규격', ''))
                            try: st.session_state[f"f_qty_{i}"] = float(row_data.get('수량', 1.0))
                            except: st.session_state[f"f_qty_{i}"] = 1.0
                            st.session_state[f"f_unit_{i}"] = str(row_data.get('단위', '롤'))
                            st.session_state[f"f_color_{i}"] = str(row_data.get('색상', ''))
                            st.session_state[f"f_proc_{i}"] = str(row_data.get('가공', ''))
                            st.session_state[f"f_ks_{i}"] = str(row_data.get('KS', ''))
                            st.session_state[f"f_note_{i}"] = str(row_data.get('비고', ''))
                            try: st.session_state[f"f_pprice_{i}"] = int(float(row_data.get('매입단가', 0)))
                            except: st.session_state[f"f_pprice_{i}"] = 0
                            try: st.session_state[f"f_sprice_{i}"] = int(float(row_data.get('매출단가', 0)))
                            except: st.session_state[f"f_sprice_{i}"] = 0
                            st.session_state[f"f_vendor_{i}"] = str(row_data.get('매입업체', ''))
                        else:
                            st.session_state[f"f_item_{i}"] = ""
                            st.session_state[f"f_spec_{i}"] = ""
                            st.session_state[f"f_qty_{i}"] = 1.0
                            st.session_state[f"f_unit_{i}"] = "롤"
                            st.session_state[f"f_color_{i}"] = ""
                            st.session_state[f"f_proc_{i}"] = ""
                            st.session_state[f"f_ks_{i}"] = ""
                            st.session_state[f"f_note_{i}"] = ""
                            st.session_state[f"f_pprice_{i}"] = 0
                            st.session_state[f"f_sprice_{i}"] = 0
                            st.session_state[f"f_vendor_{i}"] = ""
                    st.rerun()

# --- 1. 기본 정보 ---
st.subheader("1. 기본 정보")
today_str = today.strftime("%y%m%d")
seq = 1
if not df_history.empty and '주문번호' in df_history.columns:
    today_orders = df_history[df_history['주문번호'].astype(str).str.startswith(today_str)]
    if not today_orders.empty:
        try: seq = today_orders['주문번호'].str.split('-').str[-1].astype(int).max() + 1
        except: pass

f_order_no = st.text_input("주문번호", value=st.session_state.f_order_no if st.session_state.f_order_no else f"{today_str}-{seq:02d}")

r1c1, r1c2, r1c3, r1c4 = st.columns(4)
with r1c1: f_close_month = st.selectbox("마감월", month_list, index=month_list.index(st.session_state.f_close_month) if st.session_state.f_close_month in month_list else 12)
with r1c2: f_date = st.date_input("발주일", st.session_state.f_date)
with r1c3: f_due_date = st.date_input("납기일", st.session_state.f_due_date)
with r1c4: f_due_time = st.text_input("납기시간", value=st.session_state.f_due_time)

p_vendors = sorted([c for c in df_purch.columns if c not in ['품목', '규격', '단위', '가공', '색상', 'KS', '비고']]) if not df_purch.empty else []
s_vendors = sorted([c for c in df_sales.columns if c not in ['품목', '규격', '단위', '가공', '색상', 'KS', '비고']]) if not df_sales.empty else []

r2c1, r2c2, r2c3, r2c4 = st.columns(4)
with r2c1: 
    sel_s_v = st.selectbox("납품처(매출업체)", [""] + s_vendors + ["[직접 입력]"], index=s_vendors.index(st.session_state.f_sales_v)+1 if st.session_state.f_sales_v in s_vendors else 0)
    f_sales_v = st.text_input("매출업체 직접입력", value=st.session_state.f_sales_v, label_visibility="collapsed") if sel_s_v == "[직접 입력]" else sel_s_v
with r2c2: f_site = st.text_input("현장명", value=st.session_state.f_site)
with r2c3: f_manager = st.text_input("담당(수령인)", value=st.session_state.f_manager)
with r2c4: f_phone = st.text_input("수령인전화", value=st.session_state.f_phone)

r3c1, r3c2 = st.columns([3, 1])
with r3c1: f_address = st.text_input("도착지주소", value=st.session_state.f_address)
with r3c2: 
    sel_p_v = st.selectbox("기본 매입업체", [""] + p_vendors + ["[직접 입력]"], index=p_vendors.index(st.session_state.f_purch_v)+1 if st.session_state.f_purch_v in p_vendors else 0)
    f_purch_v = st.text_input("매입업체 직접입력", value=st.session_state.f_purch_v, label_visibility="collapsed") if sel_p_v == "[직접 입력]" else sel_p_v

st.session_state.f_sales_v = f_sales_v
st.session_state.f_purch_v = f_purch_v

# --- 2. 품목 상세 (일괄 입력 폼) ---
st.subheader("2. 품목 상세 (모바일 최적화 일괄 폼)")

def get_opts(df, col): return [str(x) for x in df[col].dropna().unique() if str(x).strip()] if not df.empty and col in df.columns else []

def build_opts(base_list, state_prefix):
    opts = set(base_list)
    for i in range(MAX_ROWS):
        v = st.session_state.get(f"{state_prefix}_{i}", "")
        if v: opts.add(v)
    return [""] + sorted(list(opts))

target_order = ["안전망1cm", "안전망2cm", "PP로프", "와이어로프", "와이어클립", "럿셀망", "멀티망", "케이블타이", "PE로프", "웨빙띠"]
raw_items = list(set(get_opts(df_purch, '품목') + get_opts(df_sales, '품목')))
item_opts = [""] + sorted(raw_items, key=lambda x: target_order.index(x) if x in target_order else 999)
for i in range(MAX_ROWS):
    if st.session_state[f"f_item_{i}"] and st.session_state[f"f_item_{i}"] not in item_opts:
        item_opts.append(st.session_state[f"f_item_{i}"])

spec_opts = build_opts(get_opts(df_purch, '규격') + get_opts(df_sales, '규격'), "f_spec")
unit_opts = build_opts(["롤", "m2", "R/L", "M", "EA", "봉", "장", "박스", "kg", "포"], "f_unit")
color_opts = build_opts(["녹색", "청색", "청+노", "백색", "흑색", "주황"], "f_color")
proc_opts = build_opts(["미가공", "6mm가공", "8mm가공", "10mm가공", "가공품"], "f_proc")
ks_opts = build_opts(["KS", "일반"], "f_ks")
vendor_opts = [""] + p_vendors

def extract_area(item, spec):
    nums = [float(x) for x in re.findall(r'(\d+(?:\.\d+)?)', str(spec))]
    if '안전망' in str(item) or '멀티망' in str(item):
        if len(nums) >= 2: return nums[0] * nums[1]
        elif len(nums) == 1: return nums[0]
    elif '와이어로프' in str(item):
        if len(nums) >= 2: return nums[-1]
        elif len(nums) == 1: return nums[0]
    elif '와이어클립' in str(item):
        if nums: return nums[0]
    return 1.0

# 폼 생성
with st.form("bulk_input_form", border=True):
    st.markdown("👇 **필요한 데이터를 입력/수정한 후, 맨 아래 파란색 [단가 자동계산 및 내역 확정] 버튼을 꼭 누르세요!**")
    
    for i in range(st.session_state.row_cnt):
        st.markdown(f"**🔹 [{i+1}번 품목]**")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.selectbox("품목", item_opts, key=f"f_item_{i}")
        with c2: st.selectbox("규격", spec_opts, key=f"f_spec_{i}")
        with c3: st.number_input("수량", min_value=0.0, step=1.0, key=f"f_qty_{i}")
        with c4: st.selectbox("단위", unit_opts, key=f"f_unit_{i}")
        
        c5, c6, c7, c8 = st.columns(4)
        with c5: st.selectbox("가공", proc_opts, key=f"f_proc_{i}")
        with c6: st.selectbox("색상", color_opts, key=f"f_color_{i}")
        with c7: st.selectbox("KS", ks_opts, key=f"f_ks_{i}")
        with c8: st.text_input("비고", key=f"f_note_{i}", placeholder="비고 입력")
        
        c9, c10, c11 = st.columns([1,1,2])
        with c9: st.number_input("매입단가 (자동)", key=f"f_pprice_{i}", step=100)
        with c10: st.number_input("매출단가 (자동)", key=f"f_sprice_{i}", step=100)
        with c11: st.selectbox("개별 매입업체 지정", vendor_opts, key=f"f_vendor_{i}", index=vendor_opts.index(st.session_state[f"f_vendor_{i}"]) if st.session_state[f"f_vendor_{i}"] in vendor_opts else 0)
        
        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
    
    c_btn1, c_btn2 = st.columns([1, 4])
    with c_btn1:
        if st.form_submit_button("➕ 입력칸 3개 늘리기"):
            st.session_state.row_cnt = min(MAX_ROWS, st.session_state.row_cnt + 3)
            st.rerun()
            
    with c_btn2:
        submit_btn = st.form_submit_button("✨ 단가 자동계산 및 내역 확정 (저장/PDF용)", type="primary", use_container_width=True)

    if submit_btn:
        for i in range(st.session_state.row_cnt):
            item = st.session_state[f"f_item_{i}"]
            spec = st.session_state[f"f_spec_{i}"]
            color = st.session_state[f"f_color_{i}"]
            proc = st.session_state[f"f_proc_{i}"]
            row_vendor = st.session_state[f"f_vendor_{i}"]
            vendor = row_vendor if row_vendor else f_purch_v
            
            if not item: continue
            multiplier = extract_area(item, spec)
            
            # 매입단가 계산
            if not df_purch.empty and vendor in df_purch.columns:
                m = df_purch[df_purch['품목'] == item]
                if spec and '규격' in m.columns: m = m[m['규격'] == spec]
                if not m.empty:
                    try:
                        u_price = float(str(m.iloc[0][vendor]).replace(',', ''))
                        if '안전망' in item and proc == '미가공' and color not in ['', '녹색', '청색', '청+노', '녹', '청', '청노']:
                            u_price += 100
                        st.session_state[f"f_pprice_{i}"] = int(u_price * multiplier)
                    except: pass

            # 매출단가 계산
            if not df_sales.empty and f_sales_v in df_sales.columns:
                m = df_sales[df_sales['품목'] == item]
                if spec and '규격' in m.columns: m = m[m['규격'] == spec]
                if not m.empty:
                    try: st.session_state[f"f_sprice_{i}"] = int(float(str(m.iloc[0][f_sales_v]).replace(',', '')) * multiplier)
                    except: pass
        st.rerun()

# 폼 데이터 DataFrame 변환
valid_rows = []
for i in range(st.session_state.row_cnt):
    item = st.session_state[f"f_item_{i}"]
    if str(item).strip():
        valid_rows.append({
            "품목": item, "규격": st.session_state[f"f_spec_{i}"], 
            "수량": st.session_state[f"f_qty_{i}"], "단위": st.session_state[f"f_unit_{i}"],
            "색상": st.session_state[f"f_color_{i}"], "가공": st.session_state[f"f_proc_{i}"], 
            "KS": st.session_state[f"f_ks_{i}"], "비고": st.session_state[f"f_note_{i}"], 
            "매입단가": st.session_state[f"f_pprice_{i}"], "매출단가": st.session_state[f"f_sprice_{i}"], 
            "매입업체": st.session_state[f"f_vendor_{i}"]
        })
valid_df = pd.DataFrame(valid_rows)

st.markdown("**[발주서] 특이사항**")
f_po_note = st.text_area("특이사항", value=st.session_state.f_po_note, label_visibility="collapsed")

st.divider()

# --- 3. 운송 정보 ---
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

st.divider()

# --- 4. 저장 및 PDF 발행 ---
st.subheader("4. 장부 저장 및 PDF 통합 발행")

if 'is_saved' not in st.session_state: st.session_state.is_saved = False

if st.button("💾 장부 저장 및 PDF 다운로드", type="primary", use_container_width=True):
    if valid_df.empty:
        st.error("⚠️ 폼에 품목을 하나 이상 추가하고 확정해주세요.")
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

TOTAL_ROWS = 14
tbody_html = ""
for i, row in valid_df.iterrows():
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
for _ in range(max(0, TOTAL_ROWS - len(valid_df))):
    tbody_html += "<tr>" + "<td style='border:1px solid #000; height: 23px;'></td>"*7 + "</tr>"

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
