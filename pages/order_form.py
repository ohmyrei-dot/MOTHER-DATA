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
c1, c2, c3, c4 = st.columns(4)

# 마감월 선택용 리스트 (과거 1년 ~ 미래 5년)
today = datetime.date.today()
month_list = [(pd.to_datetime(today) + pd.DateOffset(months=i)).strftime("%Y-%m") for i in range(-12, 61)]

with c1: 
    f_close_month = st.selectbox("마감월 (시트저장용)", month_list, index=12)
    f_sales_v = st.text_input("납품처")
    f_address = st.text_input("도착지주소")
with c2: 
    f_date = st.date_input("발주일", today)
    f_site = st.text_input("현장명")
    f_purch_v = st.text_input("매입업체")
with c3: 
    f_due_date = st.date_input("납기일", today)
    f_due_time = st.time_input("납기시간", datetime.time(10, 0))
    f_manager = st.text_input("담당(수령인)")
with c4: 
    st.markdown("<div style='margin-top: 73px;'></div>", unsafe_allow_html=True) # 줄맞춤용 공백
    f_phone = st.text_input("수령인전화")

# 공급자 정보 (고정)
SUPPLIER_INFO = {
    "company": "석미세이프",
    "biznum": "524-38-00469",
    "address": "경기도 남양주시 수동면 남가로 1771-1"
}

# 3. 품목 상세 입력창
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
    valid_df = edited_df[edited_df['품목'].astype(str).str.strip() != ""].copy()
    
    if valid_df.empty:
        st.error("⚠️ 품목을 하나 이상 입력해주세요.")
    else:
        try:
            client = init_connection()
            sheet = client.open("석미_마더데이터").sheet1 
            
            # 구글 시트 헤더 (나중에 이 리스트 순서만 바꾸면 시트 순서도 변경됨)
            expected_headers = [
                '마감월', '발주일', '납기일', '납기시간', '납품처', '현장명', '담당(수령인)', '수령인전화',
                '도착지주소', '매입업체',
                '품목', '규격', '수량', '단위', '색상', '가공', 'KS', '비고', 
                '매입단가', '매출단가'
            ]
            
            existing_data = sheet.get_all_values()
            
            if not existing_data or existing_data[0] != expected_headers:
                if not existing_data:
                    sheet.append_row(expected_headers)
                else:
                    sheet.insert_row(expected_headers, index=1)
            
            rows_to_append = []
            f_due_time_str = f_due_time.strftime("%p %I:%M").replace("AM", "오전").replace("PM", "오후")
            for _, row in valid_df.iterrows():
                rows_to_append.append([
                    f_close_month, # 마감월 (문자열 그대로 저장)
                    f_date.strftime("%Y-%m-%d"), 
                    f_due_date.strftime("%Y-%m-%d"),
                    f_due_time_str, f_sales_v, f_site, f_manager, f_phone, 
                    f_address, f_purch_v,
                    row['품목'], row['규격'], row['수량'], row['단위'], 
                    row['색상'], row['가공'], row['KS'], row['비고'], 
                    row['매입단가'], row['매출단가']
                ])
                
            sheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
            st.success("✅ 구글 시트에 성공적으로 저장되었습니다!")
            
        except Exception as e:
            st.error(f"저장 중 오류 발생: {e}")

st.divider()

# 5. PDF 출력 미리보기 (마감월 제외, 고정 공급자 정보 사용)
st.subheader("3. 거래명세서/발주서 미리보기 및 PDF 다운로드")

tbody_html = ""
valid_rows = edited_df[edited_df['품목'].astype(str).str.strip() != ""]
for i, row in valid_rows.iterrows():
    tbody_html += f"""
    <tr>
        <td style='text-align:center; padding:5px; border:1px solid #000;'>{i+1}</td>
        <td style='padding:5px; border:1px solid #000;'>{row.get('품목', '')}</td>
        <td style='padding:5px; border:1px solid #000;'>{row.get('규격', '')}</td>
        <td style='text-align:center; padding:5px; border:1px solid #000;'>{row.get('수량', '')}</td>
        <td style='text-align:center; padding:5px; border:1px solid #000;'>{row.get('단위', '')}</td>
        <td style='padding:5px; border:1px solid #000;'>{row.get('색상', '')} {row.get('가공', '')} {row.get('KS', '')}</td>
        <td style='padding:5px; border:1px solid #000;'>{row.get('비고', '')}</td>
    </tr>
    """

html_template = f"""
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>

<div style="text-align: right; max-width: 840px; margin: 0 auto 10px auto;">
    <button onclick="downloadPDF()" style="padding: 10px 20px; background-color: #ff4b4b; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; font-weight: bold;">
        📥 PDF 다운로드
    </button>
</div>

<div id="capture-area" style="max-width: 820px; margin: 0 auto; padding: 20px; background: #fff; color: #000; font-family: 'Malgun Gothic', sans-serif;">
    <h1 style="text-align: center; letter-spacing: 10px; border-bottom: 2px solid #000; padding-bottom: 10px;">거 래 명 세 서</h1>
    
    <div style="display: flex; justify-content: space-between; margin-top: 20px; font-size: 14px;">
        <div style="width: 48%;">
            <table style="width: 100%; border-collapse: collapse;">
                <tr><td style="padding: 5px; width: 80px; font-weight: bold;">납품처</td><td>: {f_sales_v}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold;">현장명</td><td>: {f_site}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold;">도착지주소</td><td>: {f_address}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold;">수령인/연락처</td><td>: {f_manager} / {f_phone}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold;">납기일시</td><td>: {f_due_date.strftime('%Y-%m-%d')} {f_due_time_str}</td></tr>
            </table>
        </div>
        
        <div style="width: 48%;">
            <table style="width: 100%; border-collapse: collapse; border: 2px solid #000;">
                <tr>
                    <td rowspan="3" style="width: 25px; text-align: center; border-right: 1px solid #000; font-weight: bold;">공<br>급<br>자</td>
                    <td style="padding: 5px; border-right: 1px solid #000; border-bottom: 1px solid #000; width: 70px;">등록번호</td>
                    <td style="padding: 5px; border-bottom: 1px solid #000;">{SUPPLIER_INFO['biznum']}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; border-right: 1px solid #000; border-bottom: 1px solid #000;">상호</td>
                    <td style="padding: 5px; border-bottom: 1px solid #000; font-weight: bold;">{SUPPLIER_INFO['company']}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; border-right: 1px solid #000;">주소</td>
                    <td style="padding: 5px;">{SUPPLIER_INFO['address']}</td>
                </tr>
            </table>
        </div>
    </div>
    
    <div style="margin-top: 20px;">
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000; font-size: 13px; text-align: center;">
            <tr style="background-color: #f0f0f0;">
                <th style="padding: 8px; border: 1px solid #000; width: 50px;">No</th>
                <th style="padding: 8px; border: 1px solid #000;">품목</th>
                <th style="padding: 8px; border: 1px solid #000;">규격</th>
                <th style="padding: 8px; border: 1px solid #000; width: 60px;">수량</th>
                <th style="padding: 8px; border: 1px solid #000; width: 60px;">단위</th>
                <th style="padding: 8px; border: 1px solid #000;">상세(색상/가공/KS)</th>
                <th style="padding: 8px; border: 1px solid #000;">비고</th>
            </tr>
            {tbody_html}
        </table>
    </div>
</div>

<script>
    function downloadPDF() {{
        var element = document.getElementById('capture-area');
        var opt = {{
            margin:       10,
            filename:     '거래명세서_{f_sales_v}.pdf',
            image:        {{ type: 'jpeg', quality: 0.98 }},
            html2canvas:  {{ scale: 2 }},
            jsPDF:        {{ unit: 'mm', format: 'a4', orientation: 'portrait' }}
        }};
        html2pdf().set(opt).from(element).save();
    }}
</script>
"""

st.components.v1.html(html_template, height=800, scrolling=True)
