import streamlit as st
import pandas as pd
import datetime
import json
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="발주서 작성 (마더데이터 연동)", page_icon="📝", layout="wide")

# -----------------------------------------------------------------------------
# 1. 구글 시트 연결 (MOTHER-DATA)
# -----------------------------------------------------------------------------
@st.cache_resource
def init_connection():
    creds_info = json.loads(st.secrets["google_credentials"])
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(creds)

# -----------------------------------------------------------------------------
# 2. UI - 기본 정보 입력
# -----------------------------------------------------------------------------
st.title("📝 발주서 작성 및 장부 자동저장")
st.markdown("발주 내용을 입력하면 **PDF(거래명세서+발주서 2장) 출력**과 동시에 **구글 마더데이터에 자동 저장**됩니다.")

st.subheader("1. 기본 정보")
c1, c2, c3, c4 = st.columns(4)
with c1: f_date = st.date_input("발주일", datetime.date.today())
with c2: f_sales_vendor = st.text_input("매출업체 (고객사)", placeholder="예: 네오이앤씨")
with c3: f_site = st.text_input("현장명", placeholder="예: 분당느티마을")
with c4: f_purch_vendor = st.text_input("매입업체 (발주처)", placeholder="예: 삼원이앤씨")

# -----------------------------------------------------------------------------
# 3. UI - 품목 리스트 입력
# -----------------------------------------------------------------------------
st.subheader("2. 품목 상세 입력")
st.caption("💡 빈 행을 클릭하여 품목을 추가하세요. 금액과 부가세는 장부 저장 시 자동 계산됩니다.")

if 'order_items' not in st.session_state:
    st.session_state.order_items = pd.DataFrame([{
        "품목": "", "규격": "", "색상(비고1)": "", "수량": 1, "단위": "롤", "매입단가": 0, "매출단가": 0
    }])

edited_df = st.data_editor(
    st.session_state.order_items,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "수량": st.column_config.NumberColumn("수량", min_value=1, step=1),
        "매입단가": st.column_config.NumberColumn("매입단가(원)", min_value=0, step=100, format="%d"),
        "매출단가": st.column_config.NumberColumn("매출단가(원)", min_value=0, step=100, format="%d"),
    }
)

# -----------------------------------------------------------------------------
# 4. 데이터 저장 및 렌더링 로직
# -----------------------------------------------------------------------------
st.divider()
col_btn, col_msg = st.columns([2, 8])

if col_btn.button("💾 장부 저장 및 PDF 다운로드", type="primary"):
    valid_df = edited_df[edited_df['품목'].str.strip() != ""].copy()
    
    if valid_df.empty or not f_sales_vendor or not f_purch_vendor:
        st.error("⚠️ 매출업체, 매입업체, 그리고 최소 1개의 품목을 입력해주세요.")
    else:
        try:
            client = init_connection()
            sheet = client.open("석미_마더데이터").sheet1
            
            rows_to_append = []
            total_amt_sum = 0
            
            for _, row in valid_df.iterrows():
                qty = int(row['수량'])
                p_price = int(row['매입단가'])
                s_price = int(row['매출단가'])
                
                p_amt = p_price * qty
                p_tax = int(p_amt * 0.1)
                p_total = p_amt + p_tax
                
                s_amt = s_price * qty
                s_tax = int(s_amt * 0.1)
                s_total = s_amt + s_tax
                total_amt_sum += s_total
                
                append_row = [
                    f_date.strftime("%Y-%m-%d"), "", f_sales_vendor, f_site, f_purch_vendor,
                    row['품목'], row['규격'], row['색상(비고1)'], qty, row['단위'],
                    s_price, s_amt, s_tax, s_total,
                    p_price, p_amt, p_tax, p_total,
                    "", ""
                ]
                rows_to_append.append(append_row)
                
            sheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
            st.success("✅ 구글 마더데이터 시트에 성공적으로 저장되었습니다!")
            
            # --- HTML PDF 생성 ---
            # 거래명세서 표 생성 (최소 10줄 맞춤)
            trs = ""
            for idx, r in valid_df.reset_index(drop=True).iterrows():
                s_price = int(r['매출단가']); s_amt = s_price * int(r['수량'])
                trs += f"""
                <tr>
                    <td style='border:1px solid #333; padding:5px; text-align:center;'>{idx+1}</td>
                    <td style='border:1px solid #333; padding:5px;'>{r['품목']}</td>
                    <td style='border:1px solid #333; padding:5px;'>{r['규격']} {r['색상(비고1)']}</td>
                    <td style='border:1px solid #333; padding:5px; text-align:center;'>{r['수량']}</td>
                    <td style='border:1px solid #333; padding:5px; text-align:center;'>{r['단위']}</td>
                    <td style='border:1px solid #333; padding:5px; text-align:right;'>{s_price:,}</td>
                    <td style='border:1px solid #333; padding:5px; text-align:right;'>{s_amt:,}</td>
                </tr>
                """
            # 남은 빈 줄 채우기
            for i in range(len(valid_df), 10):
                trs += f"""
                <tr>
                    <td style='border:1px solid #333; height:26px;'></td><td style='border:1px solid #333;'></td>
                    <td style='border:1px solid #333;'></td><td style='border:1px solid #333;'></td>
                    <td style='border:1px solid #333;'></td><td style='border:1px solid #333;'></td>
                    <td style='border:1px solid #333;'></td>
                </tr>
                """

            html_pdf = f"""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
            <div id="pdf-wrap" style="background:#fff; color:#000; font-family:'Malgun Gothic',sans-serif; font-size:12px; width: 297mm; padding: 0;">
                
                <!-- 1페이지: 거래명세서 (좌우 대칭) -->
                <div style="page-break-after: always; display: flex; justify-content: space-between; padding: 20px 10px; width: 100%; box-sizing: border-box;">
                    
                    <!-- 좌측 (공급받는자용) -->
                    <div style="width: 48%; border: 2px solid #000; padding: 15px; box-sizing: border-box;">
                        <h2 style="text-align:center; letter-spacing: 15px; margin-top: 0; margin-bottom: 20px; text-decoration: underline;">거래명세서</h2>
                        <div style="text-align:center; margin-bottom: 10px; font-weight: bold;">(공급받는자 보관용)</div>
                        
                        <div style="display:flex; justify-content:space-between; margin-bottom:15px; font-size:14px;">
                            <div style="border-bottom: 1px solid #000; padding-bottom: 3px;"><b>{f_sales_vendor}</b> 귀하</div>
                            <div style="border-bottom: 1px solid #000; padding-bottom: 3px;">날짜: {f_date.strftime('%Y-%m-%d')}</div>
                        </div>
                        
                        <div style="margin-bottom:10px; font-size: 14px; color: #333;">
                            <b>합계금액: ₩ {total_amt_sum:,}</b> (VAT 포함)
                        </div>

                        <table style="width:100%; border-collapse:collapse; text-align:center;">
                            <colgroup>
                                <col width="8%"><col width="22%"><col width="22%"><col width="10%"><col width="10%"><col width="13%"><col width="15%">
                            </colgroup>
                            <tr style="background:#f0f0f0;">
                                <th style="border:1px solid #333; padding:5px;">No</th>
                                <th style="border:1px solid #333; padding:5px;">품목</th>
                                <th style="border:1px solid #333; padding:5px;">규격/색상</th>
                                <th style="border:1px solid #333; padding:5px;">수량</th>
                                <th style="border:1px solid #333; padding:5px;">단위</th>
                                <th style="border:1px solid #333; padding:5px;">단가</th>
                                <th style="border:1px solid #333; padding:5px;">금액</th>
                            </tr>
                            {trs}
                        </table>
                        
                        <div style="margin-top:15px; display:flex; justify-content:space-between; font-size: 14px;">
                            <div><b>현장명:</b> {f_site}</div>
                            <div><b>공급자:</b> 석미세이프 (인)</div>
                        </div>
                    </div>
                    
                    <!-- 우측 (공급자보관용) -->
                    <div style="width: 48%; border: 2px solid #000; padding: 15px; box-sizing: border-box;">
                        <h2 style="text-align:center; letter-spacing: 15px; margin-top: 0; margin-bottom: 20px; text-decoration: underline;">거래명세서</h2>
                        <div style="text-align:center; margin-bottom: 10px; font-weight: bold;">(공급자 보관용)</div>
                        
                        <div style="display:flex; justify-content:space-between; margin-bottom:15px; font-size:14px;">
                            <div style="border-bottom: 1px solid #000; padding-bottom: 3px;"><b>{f_sales_vendor}</b> 귀하</div>
                            <div style="border-bottom: 1px solid #000; padding-bottom: 3px;">날짜: {f_date.strftime('%Y-%m-%d')}</div>
                        </div>

                        <div style="margin-bottom:10px; font-size: 14px; color: #333;">
                            <b>합계금액: ₩ {total_amt_sum:,}</b> (VAT 포함)
                        </div>

                        <table style="width:100%; border-collapse:collapse; text-align:center;">
                            <colgroup>
                                <col width="8%"><col width="22%"><col width="22%"><col width="10%"><col width="10%"><col width="13%"><col width="15%">
                            </colgroup>
                            <tr style="background:#f0f0f0;">
                                <th style="border:1px solid #333; padding:5px;">No</th>
                                <th style="border:1px solid #333; padding:5px;">품목</th>
                                <th style="border:1px solid #333; padding:5px;">규격/색상</th>
                                <th style="border:1px solid #333; padding:5px;">수량</th>
                                <th style="border:1px solid #333; padding:5px;">단위</th>
                                <th style="border:1px solid #333; padding:5px;">단가</th>
                                <th style="border:1px solid #333; padding:5px;">금액</th>
                            </tr>
                            {trs}
                        </table>
                        
                        <div style="margin-top:15px; display:flex; justify-content:space-between; font-size: 14px;">
                            <div><b>현장명:</b> {f_site}</div>
                            <div><b>인수자:</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(인)</div>
                        </div>
                    </div>
                </div>

                <!-- 2페이지: 발주서(운송장) -->
                <div style="padding: 40px; box-sizing: border-box; width: 100%;">
                    <h1 style="text-align:center; letter-spacing: 20px; margin-bottom: 40px; font-size: 32px; text-decoration: underline;">발 주 서</h1>
                    
                    <table style="width:100%; margin-bottom: 30px; font-size:16px;">
                        <tr><td style="width:120px; padding: 5px 0;"><b>발주처:</b></td><td style="font-size: 18px; font-weight: bold;">{f_purch_vendor}</td></tr>
                        <tr><td style="padding: 5px 0;"><b>발주일:</b></td><td>{f_date.strftime('%Y-%m-%d')}</td></tr>
                        <tr><td style="padding: 5px 0;"><b>도착지(현장):</b></td><td>{f_site} <span style="color: #555;">({f_sales_vendor})</span></td></tr>
                        <tr><td style="padding: 5px 0;"><b>발송자:</b></td><td>석미세이프</td></tr>
                    </table>
                    
                    <table style="width:100%; border-collapse:collapse; font-size:15px; text-align:center;">
                        <colgroup>
                            <col width="10%"><col width="30%"><col width="30%"><col width="15%"><col width="15%">
                        </colgroup>
                        <tr style="background:#eaeaea;">
                            <th style="border:2px solid #333; padding:12px;">연번</th>
                            <th style="border:2px solid #333; padding:12px;">품목</th>
                            <th style="border:2px solid #333; padding:12px;">규격 / 색상</th>
                            <th style="border:2px solid #333; padding:12px;">수량</th>
                            <th style="border:2px solid #333; padding:12px;">단위</th>
                        </tr>
            """
            
            for idx, r in valid_df.reset_index(drop=True).iterrows():
                html_pdf += f"""
                <tr>
                    <td style='border:1px solid #333; padding:12px;'>{idx+1}</td>
                    <td style='border:1px solid #333; padding:12px;'>{r['품목']}</td>
                    <td style='border:1px solid #333; padding:12px;'>{r['규격']} {r['색상(비고1)']}</td>
                    <td style='border:1px solid #333; padding:12px; font-weight:bold; font-size:16px;'>{r['수량']}</td>
                    <td style='border:1px solid #333; padding:12px;'>{r['단위']}</td>
                </tr>
                """
                
            html_pdf += """
                    </table>
                </div>
            </div>
            
            <script>
                var opt = {
                    margin:       0,
                    filename:     '발주서_및_거래명세서.pdf',
                    image:        { type: 'jpeg', quality: 0.98 },
                    html2canvas:  { scale: 2 },
                    jsPDF:        { unit: 'mm', format: 'a4', orientation: 'landscape' }
                };
                html2pdf().set(opt).from(document.getElementById('pdf-wrap')).save();
            </script>
            """
            st.components.v1.html(html_pdf, height=0)
            
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
