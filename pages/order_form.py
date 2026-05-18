import streamlit as st
import pandas as pd
import datetime
import json
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="발주서 작성", page_icon="📝", layout="wide")

# 1. 구글 시트 연결
@st.cache_resource
def init_connection():
    creds_info = json.loads(st.secrets["google_credentials"])
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

st.title("📝 발주서 작성 (마더데이터 저장 테스트)")

# 2. 기본 정보 입력창 (누락 항목 모두 추가)
st.subheader("1. 기본 정보")
c1, c2, c3, c4 = st.columns(4)
with c1: 
    f_date = st.date_input("발주일", datetime.date.today())
    f_sales_v = st.text_input("매출업체")
    f_supplier = st.text_input("공급자정보")
with c2: 
    f_due_date = st.text_input("납기일(시간포함)", placeholder="예: 5/20 오전 10시")
    f_purch_v = st.text_input("매입업체")
with c3: 
    f_site = st.text_input("현장명")
    f_manager = st.text_input("담당(수령인)")
with c4: 
    f_address = st.text_input("도착지주소")
    f_phone = st.text_input("담당전화번호")

# 3. 품목 상세 입력창 (누락 항목 모두 추가)
st.subheader("2. 품목 상세")
if 'order_items' not in st.session_state:
    st.session_state.order_items = pd.DataFrame([
        {"품목": "", "규격": "", "수량": 1, "단위": "롤", "색상": "", "가공": "", "KS": "", "비고": "", "매입단가": 0, "매출단가": 0}
    ])

edited_df = st.data_editor(
    st.session_state.order_items,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True
)

st.divider()

# 4. 저장 버튼 및 헤더 자동 생성 로직
if st.button("💾 마더데이터에 저장", type="primary"):
    # 품목이 비어있지 않은 줄만 필터링
    valid_df = edited_df[edited_df['품목'].astype(str).str.strip() != ""].copy()
    
    if valid_df.empty:
        st.error("⚠️ 품목을 하나 이상 입력해주세요.")
    else:
        try:
            client = init_connection()
            sheet = client.open("석미_마더데이터").sheet1 
            
            # 구글 시트에 들어갈 전체 헤더(제목) 정의
            expected_headers = [
                '발주일', '납기일', '매출업체', '매입업체', '현장명', '도착지주소', 
                '담당(수령인)', '담당전화번호', '공급자정보',
                '품목', '규격', '수량', '단위', '색상', '가공', 'KS', '비고', 
                '매입단가', '매출단가'
            ]
            
            # 시트의 기존 데이터 확인
            existing_data = sheet.get_all_values()
            
            # 헤더가 없거나 다르면 1행에 헤더를 자동으로 추가
            if not existing_data or existing_data[0] != expected_headers:
                if not existing_data:
                    sheet.append_row(expected_headers)
                else:
                    sheet.insert_row(expected_headers, index=1)
            
            # 데이터 추가 준비
            rows_to_append = []
            for _, row in valid_df.iterrows():
                rows_to_append.append([
                    f_date.strftime("%Y-%m-%d"), 
                    f_due_date, f_sales_v, f_purch_v, f_site, f_address, 
                    f_manager, f_phone, f_supplier,
                    row['품목'], row['규격'], row['수량'], row['단위'], 
                    row['색상'], row['가공'], row['KS'], row['비고'], 
                    row['매입단가'], row['매출단가']
                ])
                
            # 시트에 데이터 저장
            sheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
            st.success("✅ 구글 시트에 1행 헤더와 함께 성공적으로 저장되었습니다!")
            
        except Exception as e:
            st.error(f"저장 중 오류 발생: {e}")
