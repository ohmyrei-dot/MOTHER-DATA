import streamlit as st
import pandas as pd
import datetime
import json
import gspread
from google.oauth2.service_account import Credentials
import os
import base64

st.set_page_config(page_title="발주서 및 거래명세서 작성", page_icon="📝", layout="wide")

# 1. 구글 시트 연결
@st.cache_resource
def init_connection():
    creds_info = json.loads(st.secrets["google_credentials"])
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

st.title("📝 발주서 및 거래명세서 작성")

# 2. 기본 정보 입력창
st.subheader("1. 기본 정보")

today = datetime.date.today()
month_list = [(pd.to_datetime(today) + pd.DateOffset(months=i)).strftime("%Y-%m") for i in range(-12, 61)]

# 1번째 줄
r1c1, r1c2, r1c3, r1c4 = st.columns(4)
with r1c1: f_close_month = st.selectbox("마감월 (저장용)", month_list, index=12)
with r1c2: f_date = st.date_input("발주일", today)
with r1c3: f_due_date = st.date_input("납기일", today)
with r1c4: f_due_time = st.text_input("납기시간", placeholder="예: 오전 10시")

# 2번째 줄
r2c1, r2c2, r2c3, r2c4 = st.columns(4)
with r2c1: f_sales_v = st.text_input("납품처")
with r2c2: f_site = st.text_input("현장명")
with r2c3: f_manager = st.text_input("담당(수령인)")
with r2c4: f_phone = st.text_input("수령인전화")

# 3번째 줄
r3c1, r3c2 = st.columns([1, 3])
with r3c1: f_purch_v = st.text_input("매입업체")
with r3c2: f_address = st.text_input("도착지주소 (상세 입력)")

# 3. 품목 상세 입력창
st.subheader("2. 품목 상세")

# --- 드롭다운(검색)용 데이터 불러오기 ---
file_path = 'price_list.xlsx'
item_options, spec_options = [], []
if os.path.exists(file_path):
    try:
        temp_df = pd.read_excel(file_path, sheet_name='Sales_매출단가')
        item_options = [str(x) for x in temp_df['품목'].dropna().unique() if str(x).strip()]
        if '규격' in temp_df.columns:
            spec_options = [str(x) for x in temp_df['규격'].dropna().unique() if str(x).strip()]
    except: pass

if not item_options:
    item_options = ["안전망1cm", "안전망2cm", "멀티망", "럿셀망", "PP로프", "와이어로프", "와이어클립", "케이블타이"]
if not spec_options:
    spec_options = ["미가공", "6mm가공", "8mm가공", "10mm가공", "1200D", "1.2", "1.5", "1.8"]
unit_options = ["롤", "m2", "R/L", "M", "EA", "봉", "장", "박스", "kg", "포"]

if 'order_items' not in st.session_state:
    st.session_state.order_items = pd.DataFrame([
        {"품목": "", "규격": "", "수량": 1, "단위": "롤", "색상": "", "가공": "", "KS": "", "비고": "", "매입단가": 0, "매출단가": 0}
    ])

# --- 수기 입력창 (표 위쪽에 배치) ---
with st.expander("➕ 수기 입력 (드롭다운 목록에 없는 항목 강제 추가)"):
    c_m1, c_m2, c_m3, c_m4, c_m5 = st.columns(5)
    with c_m1: m_item = st.text_input("품목 (직접입력)", key="m_item")
    with c_m2: m_spec = st.text_input("규격 (직접입력)", key="m_spec")
    with c_m3: m_qty = st.number_input("수량", min_value=1, value=1, key="m_qty")
    with c_m4: m_unit = st.text_input("단위", value="EA", key="m_unit")
    with c_m5:
        st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
        if st.button("표에 추가", use_container_width=True):
            if m_item.strip():
                new_row = {"품목": m_item, "규격": m_spec, "수량": m_qty, "단위": m_unit, "색상": "", "가공": "", "KS": "", "비고": "", "매입단가": 0, "매출단가": 0}
                st.session_state.order_items = pd.concat([st.session_state.order_items, pd.DataFrame([new_row])], ignore_index=True)
                st.rerun()

# 현재 테이블에 있는 값도 옵션에 포함 (오류 방지: 모두 문자열로 강제 변환)
current_items = [str(x) for x in st.session_state.order_items['품목'].unique() if str(x).strip()]
current_specs = [str(x) for x in st.session_state.order_items['규격'].unique() if str(x).strip()]
current_units = [str(x) for x in st.session_state.order_items['단위'].unique() if str(x).strip()]

final_item_opts = sorted(list(set(item_options + current_items + [""])), key=str)
final_spec_opts = sorted(list(set(spec_options + current_specs + [""])), key=str)
final_unit_opts = sorted(list(set(unit_options + current_units + [""])), key=str)

# --- 메인 데이터 표 (드롭다운 적용) ---
edited_df = st.data_editor(
    st.session_state.order_items,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "품목": st.column_config.SelectboxColumn("품목 (선택)", options=final_item_opts, width="medium"),
        "규격": st.column_config.SelectboxColumn("규격 (선택)", options=final_spec_opts, width="medium"),
        "단위": st.column_config.SelectboxColumn("단위 (선택)", options=final_unit_opts, width="small"),
        "수량": st.column_config.NumberColumn("수량", min_value=0.01, step=1, width="small")
    }
)

# 수정된 데이터를 세션에 동기화
st.session_state.order_items = edited_df.copy()

st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
st.markdown("**[발주서] 특이사항**")
f_po_note = st.text_area("특이사항", placeholder="발주서 하단에 출력될 특이사항을 기재하세요.", height=80, label_visibility="collapsed")

st.divider()

# 4. 하단 출력용 추가 정보 (운송정보)
st.subheader("3. 운송정보")
r4c1, r4c2, r4c3, r4c4 = st.columns(4)
with r4c1: f_ship_cost = st.text_input("운임")
with r4c2: f_ship_car = st.text_input("차량번호")
with r4c3: f_ship_driver = st.text_input("기사명")
with r4c4: f_ship_phone = st.text_input("기사연락처")

r5c1, r5c2, r5c3, r5c4 = st.columns(4)
with r5c1: f_depot = st.text_input("출고지")
with r5c2: f_sender = st.text_input("출고자")
with r5c3: f_sender_phone = st.text_input("출고자 전화")
with r5c4: f_receiver = st.text_input("인수자")

# 공급자 정보 (고정)
SUPPLIER_INFO = {
    "company": "석미세이프",
    "biznum": "524-38-00469",
    "address": "경기도 남양주시 수동면 남가로<br>1771-1"
}

st.divider()

# 5. 저장 및 PDF 발행 통합 로직
st.subheader("4. 장부 저장 및 PDF 통합 발행")
st.info("💡 아래 버튼을 누르면 구글 시트에 자동 저장된 후, 거래명세서(양면)와 발주서(단면)가 포함된 PDF가 다운로드됩니다.")

if 'is_saved' not in st.session_state:
    st.session_state.is_saved = False

if st.button("💾 장부 저장 및 PDF 다운로드", type="primary", use_container_width=True):
    valid_df = edited_df[edited_df['품목'].astype(str).str.strip() != ""].copy()
    
    if valid_df.empty:
        st.error("⚠️ 품목을 하나 이상 입력해주세요.")
        st.session_state.is_saved = False
    else:
        try:
            with st.spinner("구글 시트에 저장 중입니다..."):
                client = init_connection()
                sheet = client.open("석미_마더데이터").sheet1 
                
                expected_headers = [
                    '마감월', '발주일', '납기일', '납기시간', '납품처', '현장명', '담당(수령인)', '수령인전화',
                    '도착지주소', '매입업체',
                    '품목', '규격', '수량', '단위', '색상', '가공', 'KS', '비고', 
                    '매입단가', '매출단가',
                    '운임', '차량번호', '기사명', '기사연락처', '출고지', '출고자', '출고자전화', '인수자', '특이사항'
                ]
                
                existing_data = sheet.get_all_values()
                if not existing_data or existing_data[0] != expected_headers:
                    if not existing_data:
                        sheet.append_row(expected_headers)
                    else:
                        sheet.insert_row(expected_headers, index=1)
                
                rows_to_append = []
                for _, row in valid_df.iterrows():
                    rows_to_append.append([
                        f_close_month,
                        f_date.strftime("%Y-%m-%d"), 
                        f_due_date.strftime("%Y-%m-%d"),
                        f_due_time, f_sales_v, f_site, f_manager, f_phone, 
                        f_address, f_purch_v,
                        row['품목'], row['규격'], row['수량'], row['단위'], 
                        row['색상'], row['가공'], row['KS'], row['비고'], 
                        row['매입단가'], row['매출단가'],
                        f_ship_cost, f_ship_car, f_ship_driver, f_ship_phone,
                        f_depot, f_sender, f_sender_phone, f_receiver, f_po_note
                    ])
                    
                sheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
                st.success("✅ 구글 시트 저장 완료! PDF 다운로드를 시작합니다.")
                st.session_state.is_saved = True
                
        except Exception as e:
            st.error(f"저장 중 오류 발생: {e}")
            st.session_state.is_saved = False

# 6. PDF 출력용 템플릿 준비
TOTAL_ROWS = 14  # 페이지 넘침 방지하면서 14줄로 확장
tbody_html = ""
valid_rows = edited_df[edited_df['품목'].astype(str).str.strip() != ""]
for i, row in valid_rows.iterrows():
    tbody_html += f"""
    <tr>
        <td style='text-align:center; padding:0 4px; border:1px solid #000; height: 23px;'><div style='height:23px; line-height:23px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;'>{i+1}</div></td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'><div style='height:23px; line-height:23px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;'>{row.get('품목', '')}</div></td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'><div style='height:23px; line-height:23px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;'>{row.get('규격', '')}</div></td>
        <td style='text-align:center; padding:0 4px; border:1px solid #000; height: 23px;'><div style='height:23px; line-height:23px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;'>{row.get('수량', '')}</div></td>
        <td style='text-align:center; padding:0 4px; border:1px solid #000; height: 23px;'><div style='height:23px; line-height:23px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;'>{row.get('단위', '')}</div></td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'><div style='height:23px; line-height:23px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;'>{row.get('색상', '')} {row.get('가공', '')} {row.get('KS', '')}</div></td>
        <td style='padding:0 4px; border:1px solid #000; height: 23px;'><div style='height:23px; line-height:23px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;'>{row.get('비고', '')}</div></td>
    </tr>
    """

# 빈 칸 채우기
empty_rows_count = max(0, TOTAL_ROWS - len(valid_rows))
for _ in range(empty_rows_count):
    tbody_html += f"""
    <tr>
        <td style='border:1px solid #000; height: 23px;'><div style='height:23px;'></div></td>
        <td style='border:1px solid #000; height: 23px;'><div style='height:23px;'></div></td>
        <td style='border:1px solid #000; height: 23px;'><div style='height:23px;'></div></td>
        <td style='border:1px solid #000; height: 23px;'><div style='height:23px;'></div></td>
        <td style='border:1px solid #000; height: 23px;'><div style='height:23px;'></div></td>
        <td style='border:1px solid #000; height: 23px;'><div style='height:23px;'></div></td>
        <td style='border:1px solid #000; height: 23px;'><div style='height:23px;'></div></td>
    </tr>
    """

# 거래명세서 템플릿 생성기
def create_ts_block(receiver_name):
    # 운임란 간격 벌리기
    display_cost = f'<div style="display: flex; justify-content: space-between; padding: 0 5px;"><span>₩</span><span style="text-align: center; flex-grow: 1;">{f_ship_cost}</span><span>원</span></div>'
    
    driver_phone_arr = []
    if f_ship_driver.strip(): driver_phone_arr.append(f_ship_driver.strip())
    if f_ship_phone.strip(): driver_phone_arr.append(f_ship_phone.strip())
    driver_phone_display = " / ".join(driver_phone_arr)
    
    sender_arr = []
    if f_depot.strip(): sender_arr.append(f_depot.strip())
    if f_sender.strip(): sender_arr.append(f_sender.strip())
    f_sender_display = " / ".join(sender_arr)
    
    return f"""
    <div style="width: 48%; padding: 10px; box-sizing: border-box; font-family: 'Malgun Gothic', sans-serif;">
        <div style="position: relative; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 10px;">
            <h1 style="text-align: center; letter-spacing: 15px; font-size: 22px; margin: 0;">거 래 명 세 서</h1>
            <div style="position: absolute; right: 0; bottom: 5px; font-size: 12px; font-weight: bold;">{f_date.strftime('%Y - %m - %d')}</div>
        </div>
        
        <div style="display: flex; justify-content: space-between; margin-bottom: 5px; font-size: 12px; align-items: stretch;">
            <!-- 좌측: 수신처 정보 (높이 및 글꼴 동일하게 맞춤) -->
            <div style="width: 49%;">
                <table style="width: 100%; height: 100%; border-collapse: collapse; text-align: left; line-height: 1.5; font-size: 12px;">
                    <tr><td style="width: 55px; font-weight: bold;">납기일</td><td style="width: 10px;">:</td><td><span style="color: #d32f2f; font-weight: bold;">{f_due_date.strftime('%Y-%m-%d')} {f_due_time}</span></td></tr>
                    <tr><td style="font-weight: bold;">납품처</td><td>:</td><td>{receiver_name}</td></tr>
                    <tr><td style="font-weight: bold;">현장명</td><td>:</td><td>{f_site}</td></tr>
                    <tr><td style="font-weight: bold;">담 당</td><td>:</td><td>{f_manager}</td></tr>
                    <tr><td style="font-weight: bold;">전 화</td><td>:</td><td>{f_phone}</td></tr>
                </table>
            </div>
            
            <!-- 우측: 공급자 정보 (높이 100% 꽉 채우기) -->
            <div style="width: 49%;">
                <table style="width: 100%; height: 100%; border-collapse: collapse; border: 2px solid #000; text-align: left; font-size: 11px;">
                    <tr>
                        <td rowspan="4" style="width: 18px; text-align: center; border-right: 1px solid #000; font-weight: bold; padding: 2px;">공<br>급<br>자</td>
                        <td style="padding: 3px 5px; border-right: 1px solid #000; border-bottom: 1px solid #000; width: 55px; white-space: nowrap;">등록번호</td>
                        <td style="padding: 3px 5px; border-bottom: 1px solid #000; white-space: nowrap; font-weight: bold;">{SUPPLIER_INFO['biznum']}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 5px; border-right: 1px solid #000; border-bottom: 1px solid #000; white-space: nowrap;">상호</td>
                        <td style="padding: 3px 5px; border-bottom: 1px solid #000; font-weight: bold;">{SUPPLIER_INFO['company']}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 5px; border-right: 1px solid #000; border-bottom: 1px solid #000; white-space: nowrap;">주소</td>
                        <td style="padding: 3px 5px; border-bottom: 1px solid #000; word-break: keep-all; line-height: 1.3;">{SUPPLIER_INFO['address']}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 5px; border-right: 1px solid #000; white-space: nowrap;">업태/종목</td>
                        <td style="padding: 3px 5px; word-break: keep-all;">제조,도소매 / 안전용품</td>
                    </tr>
                </table>
            </div>
        </div>
        
        <!-- 도착지 주소 (박스 밑 가로 전체) -->
        <table style="width: 100%; border-collapse: collapse; text-align: left; line-height: 1.5; font-size: 12px; margin-bottom: 5px;">
            <tr><td style="width: 65px; font-weight: bold;">도착지주소</td><td style="width: 10px;">:</td><td style="word-break: keep-all;">{f_address}</td></tr>
        </table>
        
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000; font-size: 11px; text-align: center;">
            <tr style="background-color: #f0f0f0;">
                <th style="padding: 4px; border: 1px solid #000; width: 30px;">No</th>
                <th style="padding: 4px; border: 1px solid #000;">품목</th>
                <th style="padding: 4px; border: 1px solid #000;">규격</th>
                <th style="padding: 4px; border: 1px solid #000; width: 35px;">수량</th>
                <th style="padding: 4px; border: 1px solid #000; width: 35px;">단위</th>
                <th style="padding: 4px; border: 1px solid #000;">상세(색상/가공/KS)</th>
                <th style="padding: 4px; border: 1px solid #000;">비고</th>
            </tr>
            {tbody_html}
        </table>
        
        <!-- 하단 운송 정보 표 -->
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000; font-size: 11px; text-align: left; margin-top: 5px;">
            <tr>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9; width: 18%;">운임 ( 후불 )</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center; width: 32%;">{display_cost}</td>
                <td rowspan="4" style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9; width: 8%;">출<br>고<br>지</td>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9; width: 14%;">출고자</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center; width: 28%;">{f_sender_display}</td>
            </tr>
            <tr>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9;">차량번호</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center;">{f_ship_car}</td>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9;">전 화</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center;">{f_sender_phone}</td>
            </tr>
            <tr>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9;">기사명 / 전화</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center;">{driver_phone_display}</td>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9;">FAX</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center;">02-495-4856</td>
            </tr>
            <tr>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9;">인수자</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center;">{f_receiver}</td>
                <td style="border: 1px solid #000; font-weight: bold; padding: 4px; text-align: center; background-color: #f9f9f9;">Mobile</td>
                <td style="border: 1px solid #000; padding: 4px; text-align: center;">010-8645-4854</td>
            </tr>
        </table>
    </div>
    """

# 발주서 분리 템플릿 생성기
def create_po_block(receiver_name):
    # 특이사항 줄바꿈 처리
    po_note_html = str(f_po_note).replace('\n', '<br>')
    
    return f"""
    <div style="width: 48%; padding: 10px; box-sizing: border-box; font-family: 'Malgun Gothic', sans-serif;">
        <div style="position: relative; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 15px;">
            <h1 style="text-align: center; letter-spacing: 15px; font-size: 22px; margin: 0;">발 주 서</h1>
        </div>
        
        <div style="display: flex; justify-content: space-between; margin-bottom: 10px; font-size: 12px; align-items: stretch;">
            <!-- 좌측: 수신처 정보 (글꼴 동일 적용) -->
            <div style="width: 49%;">
                <table style="width: 100%; height: 100%; border-collapse: collapse; text-align: left; line-height: 1.8; font-size: 12px;">
                    <tr><td style="width: 55px; font-weight: bold;">수 신</td><td style="width: 10px;">:</td><td>{receiver_name}</td></tr>
                    <tr><td style="font-weight: bold;">발주일</td><td>:</td><td>{f_date.strftime('%Y-%m-%d')}</td></tr>
                    <tr><td style="font-weight: bold;">납기일</td><td>:</td><td><span style="color: #d32f2f; font-weight: bold;">{f_due_date.strftime('%Y-%m-%d')} {f_due_time}</span></td></tr>
                </table>
            </div>
            
            <!-- 우측: 발신처 정보 (높이 채우기) -->
            <div style="width: 49%;">
                <table style="width: 100%; height: 100%; border-collapse: collapse; text-align: left; line-height: 1.5; font-size: 11px;">
                    <tr><td style="width: 40px; font-weight: bold;">발 신</td><td style="width: 10px;">:</td><td style="font-weight: bold;">석미세이프</td></tr>
                    <tr><td style="font-weight: bold; vertical-align: top;">주 소</td><td style="vertical-align: top;">:</td><td style="word-break: keep-all;">경기도 남양주시 수동면 남가로<br>1771-1</td></tr>
                    <tr><td style="font-weight: bold;">전 화</td><td>:</td><td>031-559-4854</td></tr>
                    <tr><td style="font-weight: bold;">팩 스</td><td>:</td><td>02-6008-4854</td></tr>
                    <tr><td style="font-weight: bold;">E-mail</td><td>:</td><td>sm_safe@naver.com</td></tr>
                </table>
            </div>
        </div>
        
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000; font-size: 11px; text-align: center;">
            <tr style="background-color: #f0f0f0;">
                <th style="padding: 4px; border: 1px solid #000; width: 30px;">No</th>
                <th style="padding: 4px; border: 1px solid #000;">품목</th>
                <th style="padding: 4px; border: 1px solid #000;">규격</th>
                <th style="padding: 4px; border: 1px solid #000; width: 35px;">수량</th>
                <th style="padding: 4px; border: 1px solid #000; width: 35px;">단위</th>
                <th style="padding: 4px; border: 1px solid #000;">상세(색상/가공/KS)</th>
                <th style="padding: 4px; border: 1px solid #000;">비고</th>
            </tr>
            {tbody_html}
        </table>
        
        <!-- 하단 특이사항 박스 -->
        <div style="margin-top: 5px; border: 2px solid #000; padding: 6px 10px; font-size: 12px; text-align: left; height: 85px; box-sizing: border-box; overflow: hidden;">
            <div style="font-weight: bold; margin-bottom: 5px; text-decoration: underline;">[ 특이사항 ]</div>
            <div style="line-height: 1.4;">{po_note_html}</div>
        </div>
    </div>
    """

# 블록 생성
ts_block = create_ts_block(f_sales_v)
po_block = create_po_block(f_purch_v)

auto_download_js = ""
if st.session_state.is_saved:
    auto_download_js = """
    window.onload = function() {
        setTimeout(function() { downloadPDF(); }, 500);
    };
    """
    st.session_state.is_saved = False

html_template = f"""
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>

<div style="text-align: right; max-width: 1050px; margin: 0 auto 10px auto;">
    <button onclick="downloadPDF()" style="padding: 10px 20px; background-color: #ff4b4b; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; font-weight: bold;">
        📥 PDF 즉시 수동 다운로드 (A4 가로)
    </button>
</div>

<!-- 캡처 오류(좌측 잘림) 원천 차단: 무조건 좌측 정렬(margin: 0) 및 가로 스크롤 허용 -->
<div style="width: 100%; overflow-x: auto; background-color: #f0f2f6; padding: 20px 0; text-align: left;">
    
    <!-- 실제 캡처 영역 (너비 1050px 고정) -->
    <div id="capture-area" style="width: 1050px; margin: 0; background: #fff; padding: 20px; box-sizing: border-box; color: #000; font-family: 'Malgun Gothic', sans-serif;">
        
        <!-- 1페이지: 거래명세서 -->
        <div style="display: flex; justify-content: space-between; width: 100%; position: relative;">
            <!-- 중앙 절취선 -->
            <div style="position: absolute; left: 50%; top: 0; bottom: 0; border-left: 1px dashed #666; transform: translateX(-50%);"></div>
            {ts_block}
            {ts_block}
        </div>
        
        <!-- 강제 페이지 넘김 딱 1번만 적용 -->
        <div class="html2pdf__page-break"></div>
        
        <!-- 2페이지: 발주서 -->
        <div style="display: flex; justify-content: space-between; width: 100%; padding-top: 30px;">
            {po_block}
            <div style="width: 48%;"></div>
        </div>

    </div>
</div>

<script>
    function downloadPDF() {{
        var element = document.getElementById('capture-area');
        
        var opt = {{
            margin:       [5, 9, 5, 4], // ★ 좌측 여백 9mm, 우측 여백 4mm
            filename:     '거래명세서_및_발주서_{f_sales_v}.pdf',
            image:        {{ type: 'jpeg', quality: 1.0 }},
            html2canvas:  {{ scale: 2, useCORS: true }},
            jsPDF:        {{ unit: 'mm', format: 'a4', orientation: 'landscape' }},
            pagebreak:    {{ mode: 'legacy' }} // class="html2pdf__page-break" 전용 모드
        }};
        
        html2pdf().set(opt).from(element).save();
    }}
    
    {auto_download_js}
</script>
"""

# 스크롤 방해 없이 깔끔하게 렌더링
st.components.v1.html(html_template, height=1400, scrolling=True)
