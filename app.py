import streamlit as st
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import json

# 앱 제목
st.title("🚀 석미 발주시스템 연결 확인")

# 1. 스트림릿 Secrets에서 권한 정보 가져오기
try:
    # Secrets에 저장한 JSON 문자열을 딕셔너리로 변환
    creds_info = json.loads(st.secrets["google_credentials"])
    
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    client = gspread.authorize(creds)

    # 2. 구글 시트 열기 (공유한 파일 이름과 정확히 같아야 함)
    spreadsheet_name = "석미_마더데이터"
    sheet = client.open(spreadsheet_name).sheet1
    
    # 3. 데이터 읽어오기
    data = sheet.get_all_records()
    df = pd.DataFrame(data)

    st.success("✅ 구글 시트 연결에 성공했습니다!")
    
    st.write("### 📊 현재 마더데이터 내용")
    if not df.empty:
        st.dataframe(df)
    else:
        st.info("시트에 데이터가 비어 있습니다. 구글 시트에 내용을 입력해 보세요.")

except Exception as e:
    st.error(f"❌ 연결 중 오류가 발생했습니다: {e}")
    st.info("구글 시트의 [공유] 설정에서 'bot-876@seokmi-order-2.iam.gserviceaccount.com' 이 편집자로 추가되었는지 다시 확인해 주세요.")
