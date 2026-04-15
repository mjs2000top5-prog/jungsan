import streamlit as st
import pandas as pd
import gspread
import msoffcrypto
import io
import xlsxwriter
from datetime import datetime
from google.oauth2.service_account import Credentials

# --- [1. 구글 스프레드시트 설정 (Secrets 기반)] ---
SPREADSHEET_ID = '16oZFGDacad4ewfy_tQTz3OXkgiqPW2-IwuklU-An8Yk'

def get_gspread_client():
    try:
        creds_info = st.secrets["gcp_service_account"]
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"⚠️ 구글 인증 정보를 찾을 수 없습니다: {e}")
        return None

# --- [2. 엑셀 파일 복호화 및 데이터 정제] ---
def process_excel(uploaded_file, columns, input_password):
    file_extension = uploaded_file.name.split('.')[-1].lower()
    file_bytes = uploaded_file.getvalue()
    try:
        decrypted_workbook = io.BytesIO()
        if file_extension == 'xlsx':
            office_file = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
            if office_file.is_encrypted():
                if not input_password:
                    st.warning("🔑 암호가 걸린 파일입니다.")
                    return None
                office_file.load_key(password=input_password)
                office_file.decrypt(decrypted_workbook)
            else:
                decrypted_workbook = io.BytesIO(file_bytes)
            df = pd.read_excel(decrypted_workbook, usecols=columns, dtype=str, index_col=None, engine='openpyxl').fillna('')
        elif file_extension == 'xls':
            df = pd.read_excel(io.BytesIO(file_bytes), usecols=columns, engine='xlrd', dtype=str, index_col=None).fillna('')
        
        cols = pd.Series(df.columns)
        for i, col in enumerate(cols):
            if cols.duplicated()[i]: cols[i] = f"{col}_{i}"
        df.columns = cols
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        return df.reset_index(drop=True)
    except Exception as e:
        st.error(f"❌ 파일 처리 오류: {e}")
        return None

# --- [3. UI 구성] ---
st.set_page_config(page_title="위멤버스 통합 정산 시스템", layout="wide")
st.title("💰 위멤버스 통합 정산 및 데이터 관리")

menu = st.sidebar.selectbox("📂 작업 선택", ["가입자 시트 업로드", "사용자 시트 업로드", "정산 데이터 생성", "데이터 초기화"])

def run_upload_ui(title, columns, sheet_name):
    st.subheader(f"📑 {title}")
    uploaded_file = st.file_uploader("엑셀 파일 선택", type=["xlsx", "xls"], key=f"file_{sheet_name}")
    if uploaded_file:
        pw = st.text_input("비밀번호 입력", type="password", key=f"pw_{sheet_name}")
        if st.button("🔓 데이터 불러오기", key=f"btn_{sheet_name}"):
            df = process_excel(uploaded_file, columns, pw)
            if df is not None: st.session_state[f'df_{sheet_name}'] = df
        
        if f'df_{sheet_name}' in st.session_state:
            df = st.session_state[f'df_{sheet_name}']
            st.dataframe(df.head())
            if st.button(f"🚀 {sheet_name} 시트로 최종 업로드"):
                with st.spinner('업로드 중...'):
                    client = get_gspread_client()
                    if client:
                        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
                        header, data_values = df.columns.values.tolist(), df.values.tolist()
                        sheet.clear()
                        sheet.update('A1', [header] + data_values, value_input_option='USER_ENTERED')
                        st.success("✅ 업로드 완료!")
                        del st.session_state[f'df_{sheet_name}']

# --- [4. 정산 데이터 생성] ---
if menu == "정산 데이터 생성":
    st.subheader("📅 월별 정산 실행")
    current_year = datetime.now().year
    target_month = st.selectbox("정산 대상월 선택", [f"{current_year}-{m:02d}" for m in range(1, 13)])
    
    if st.button("📊 정산 실행"):
        try:
            with st.spinner('정산 계산 중...'):
                client = get_gspread_client()
                if client:
                    gaib_sheet = client.open_by_key(SPREADSHEET_ID).worksheet("위멤버스 가입자")
                    user_sheet = client.open_by_key(SPREADSHEET_ID).worksheet("위멤버스 사용자")
                    special_sheet = client.open_by_key(SPREADSHEET_ID).worksheet("별도 요금제")
                    
                    gaib_raw = gaib_sheet.get_all_values()
                    user_raw = user_sheet.get_all_values()
                    special_raw = special_sheet.get_all_values()

                    if len(gaib_raw) < 2:
                        st.error("가입자 데이터가 부족합니다.")
                    else:
                        df_gaib = pd.DataFrame(gaib_raw[1:], columns=gaib_raw[0])
                        # 인덱스 초기화 및 컬럼 정제
                        temp_cols = pd.Series(df_gaib.columns)
                        for i, col in enumerate(temp_cols):
                            if temp_cols.duplicated()[i]: temp_cols[i] = f"{col}_{i}"
                        df_gaib.columns = temp_cols
                        df_gaib = df_gaib.loc[:, ~df_gaib.columns.str.contains('^Unnamed|^$')].reset_index(drop=True)

                        # 별도 요금제 매핑
                        special_map = {}
                        if len(special_raw) >= 2:
                            df_special = pd.DataFrame(special_raw[1:], columns=special_raw[0])
                            special_map = df_special.set_index(df_special.columns[0])[df_special.columns[4]].to_dict()

                        # 날짜 처리
                        target_dt = datetime.strptime(target_month, "%Y-%m")
                        payment_date = target_dt.strftime("%Y%m") + "11"
                        prev_dt = datetime(target_dt.year - 1, 12, 1) if target_dt.month == 1 else datetime(target_dt.year, target_dt.month - 1, 1)
                        prev_month_str = prev_dt.strftime("%Y%m")

                        # 사용자 수 매핑
                        if len(user_raw) >= 2:
                            df_user = pd.DataFrame(user_raw[1:], columns=user_raw[0])
                            df_user_filtered = df_user[~df_user.iloc[:, 0].apply(lambda x: str(x).replace("-", "")[:6] == prev_month_str)]
                            user_counts = df_user_filtered.iloc[:, 4].str.strip().value_counts().to_dict()
                            df_gaib['사용자수'] = df_gaib.iloc[:, 10].str.strip().map(user_counts).fillna(0).astype(int)
                        else:
                            df_gaib['사용자수'] = 0

                        # [핵심 로직] 비대면 바우처 및 기본 필터링
                        # H열=인덱스 7, I열=인덱스 8
                        def filter_rows(row):
                            # 1. 기본 필터링 (TEST 업체, 휴폐업, 베이직 제외)
                            if str(row.iloc[10]) == 'TEST' or str(row.iloc[7]) == '휴폐업' or str(row.iloc[2]) == '위멤버스 베이직':
                                return False
                            # 2. 전월 가입자 제외
                            if str(row.iloc[3]).replace("-", "")[:6] == prev_month_str:
                                return False
                            # 3. 비대면_바우처 제외 로직 (H열 확인)
                            if str(row.iloc[7]) == '비대면_바우처':
                                try:
                                    # I열의 면제 종료월 추출 (YYYY-MM 형식 가정)
                                    myeonje_end = str(row.iloc[8]).replace("-", "")[:6]
                                    target_str = target_month.replace("-", "")
                                    # 면제 종료월 >= 정산 대상월 이면 제외 (정산 안 함)
                                    if myeonje_end >= target_str:
                                        return False
                                except: pass
                            return True

                        df_gaib = df_gaib[df_gaib.apply(filter_rows, axis=1)]

                        # 금액 계산
                        def calculate_final(row):
                            gaib_no = str(row.iloc[0]).strip()
                            if gaib_no in special_map:
                                try:
                                    p = int(str(special_map[gaib_no]).replace(",", "")); return pd.Series([p, int(p * 0.1)])
                                except: pass
                            try:
                                service, user_cnt = str(row.iloc[2]), int(row['사용자수'])
                                join_dt, base_dt = pd.to_datetime(row.iloc[3]), pd.to_datetime('2025-01-01')
                                base_price = 0
                                if '스탠다드' in service: base_price = 30000 if join_dt < base_dt else 36000
                                elif '프리미엄' in service: base_price = 50000 if join_dt < base_dt else 60000
                                extra = 0
                                if '스탠다드' in service and user_cnt > 2: extra = (user_cnt - 2) * 10000
                                elif '프리미엄' in service and user_cnt > 5: extra = (user_cnt - 5) * 10000
                                f = base_price + extra; return pd.Series([f, int(f * 0.1)])
                            except: return pd.Series([0, 0])

                        df_gaib[['최종정산금액', '부가세']] = df_gaib.apply(calculate_final, axis=1)
                        df_gaib['입금일자'] = payment_date
                        df_gaib['결제코드'] = df_gaib.iloc[:, 11].apply(lambda x: 'A' if '자동이체' in str(x) else ('C' if '신용카드' in str(x) else x))

                        result_df = df_gaib[[df_gaib.columns[0], df_gaib.columns[1], '입금일자', df_gaib.columns[2], '결제코드', '사용자수', '최종정산금액', '부가세']]
                        st.session_state['result_df'] = result_df.reset_index(drop=True)
                        st.success(f"✅ {target_month} 정산 완료!")

        except Exception as e:
            st.error(f"정산 실패: {e}")

    if 'result_df' in st.session_state:
        res = st.session_state['result_df']
        st.dataframe(res)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res.to_excel(writer, index=False, sheet_name='정산내역')
        st.download_button("📥 엑셀 다운로드", output.getvalue(), f"정산_{target_month}.xlsx", "application/vnd.ms-excel")

elif menu == "가입자 시트 업로드":
    run_upload_ui("가입자 데이터", [0, 2, 4, 6, 16, 17, 18, 22, 23, 24, 25, 68, 74, 80, 83], "위멤버스 가입자")
elif menu == "사용자 시트 업로드":
    run_upload_ui("사용자 데이터", [0, 2, 3, 9, 10], "위멤버스 사용자")
elif menu == "데이터 초기화":
    if st.button("🔥 시트 초기화"):
        client = get_gspread_client()
        if client: client.open_by_key(SPREADSHEET_ID).worksheet("위멤버스 가입자").clear(); st.success("초기화됨")