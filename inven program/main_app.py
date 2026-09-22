import streamlit as st
import pandas as pd
import re
from datetime import datetime, timezone, timedelta
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

# 자동 번역으로 인한 글자 깨짐 방지
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

# ==========================================
# 🇰🇷 한국 표준시(KST) 구하기 함수 (새로 추가)
# ==========================================
def get_kst_now():
    kst = timezone(timedelta(hours=9))
    return datetime.now(kst).strftime("%Y-%m-%d %H:%M")

# ==========================================
# 0. 페이지 기본 설정 및 구글 API 연동
# ==========================================
st.set_page_config(
    page_title="러셀대치학원 시설보수 및 재고관리 시스템", 
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    
    /* 자동 번역 방지 및 기본 폰트 적용 */
    html, body, [class*="css"], .stApp {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
    }

    /* 사이드바 배경 및 전체 글자색 고정 */
    [data-testid="stSidebar"] {
        background-color: #f1f5f9 !important;
    }
    [data-testid="stSidebar"] * {
        color: #0f172a !important;
    }

    /* 🚨 메뉴(Radio Button) 글자색 및 선택 항목 강제 고정 */
    [data-testid="stRadio"] label p {
        color: #0f172a !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
    }

    /* 라디오 버튼 선택된 항목 배경 및 글자 강조 */
    [data-testid="stRadio"] div[role="radiogroup"] > label {
        padding: 8px 12px;
        border-radius: 8px;
        margin-bottom: 4px;
        transition: background-color 0.2s;
    }
    [data-testid="stRadio"] div[role="radiogroup"] > label:hover {
        background-color: #e2e8f0 !important;
    }

    /* 일반 입력창/라벨 글자색 고정 */
    label, p, span, h1, h2, h3, h4, h5, h6 {
        color: #0f172a !important;
    }

    /* Input 입력 박스 배경 및 글자색 보장 */
    .stTextInput input, .stSelectbox div[data-baseweb="select"], .stNumberInput input, .stTextArea textarea {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
    }

    /* 버튼 스타일 */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        background-color: #1e3a8a !important;
        color: #ffffff !important;
        border: none;
    }
    .stButton>button:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

# 구글 API 연결 함수
@st.cache_resource
def get_google_services():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    if "private_key" in creds_dict:
        pk = str(creds_dict["private_key"])
        pk = pk.replace("\\n", "\n")
        creds_dict["private_key"] = pk
        
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(credentials)
    drive_service = build('drive', 'v3', credentials=credentials)
    return gc, drive_service

try:
    gc, drive_service = get_google_services()
    sh = gc.open("러셀대치_통합DB")
    ws_users = sh.worksheet("users")
    ws_facility = sh.worksheet("facility")
    ws_inventory = sh.worksheet("inventory")
except Exception as e:
    st.error(f"⚠️ 구글 시트 및 드라이브 연동에 실패했습니다: {e}")
    st.stop()

# 구글 드라이브 사진 업로드 함수
def upload_photo_to_drive(uploaded_file):
    if uploaded_file is None:
        return ""
    try:
        file_metadata = {'name': f"facility_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uploaded_file.name}"}
        media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype=uploaded_file.type)
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        
        # 누구나 열람 가능하도록 권한 변경
        drive_service.permissions().create(
            fileId=file.get('id'),
            body={'role': 'reader', 'type': 'anyone'}
        ).execute()
        return file.get('webViewLink', '')
    except Exception as e:
        st.warning(f"사진 업로드 실패: {e}")
        return ""

def clean_phone(phone_str):
    if pd.isna(phone_str): return ""
    return re.sub(r"[^0-9]", "", str(phone_str))

CONSUMABLES = [
    "A4용지", "A3용지", "B4용지", "미색A4용지", "미색A3용지",
    "분필(백)", "분필(청)", "분필(빨)", "분필(노)",
    "점보롤", "핸드타월", "물티슈", "각티슈", "물비누",
    "종이컵", "세모금컵", "생수", "AAA건전지", "AA건전지", "마이크 커버"
]
EQUIPMENT = [
    "노트북", "출결리더기", "보조배터리", "캠코더", 
    "삼각대 및 플레이트", "SD카드", "빔포인터", "빔리모콘", "에어컨리모콘"
]
ALL_ITEMS = CONSUMABLES + EQUIPMENT

SPACES = {
    "강의실": ["대형강의실", "201호", "202호", "203호", "204호", "601호", "602호", "701호", "702호", "703호", "704호"],
    "자습관": ["3-1관", "3-2관", "3-3관", "3-4관", "3-5관", "4-1관", "4-2관", "4-3관", "4-4관", "4-5관", "5-1관", "5-2관", "5-3관", "6-1관", "6-2관", "7-1관"],
    "화장실": ["2층 남자화장실", "2층 여자화장실", "3층 남자화장실", "3층 여자화장실", "5층 남자화장실", "5층 여자화장실", "6층 남자화장실", "6층 여자화장실", "7층 남자화장실", "7층 여자화장실"],
    "기타공간": ["기타공간 (복도/엘리베이터/로비/사무실 등)"]
}

CATEGORY_BY_SPACE = {
    "강의실": ["빔프로젝터/음향", "냉난방/환기", "조명/전기", "책상/의자/칠판", "문/창문/열쇠", "신규 물품 구매 요청", "기타 시설"],
    "자습관": ["자습관 책상/시디즈 의자", "스탠드/전기/콘센트", "냉난방/공기청정기", "공용 스탠딩 책상", "문/창문/소음 문제", "신규 물품 구매 요청", "기타 시설"],
    "화장실": ["대변기 (막힘/고장/부속)", "소변기 (막힘/자동센서/누수)", "세면대/수도꼭지", "휴지걸이/비누디스펜서", "조명/환풍기", "바닥 배수구/타일", "신규 물품 구매 요청", "기타 시설"],
    "기타공간": ["엘리베이터", "복도/계단/난간", "정수기/음료대", "자동문/출입문", "조명/전기", "신규 물품 구매 요청", "기타 시설"]
}

# 세션 초기화
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
if "user_name" not in st.session_state: st.session_state["user_name"] = ""
if "user_phone" not in st.session_state: st.session_state["user_phone"] = ""
if "user_role" not in st.session_state: st.session_state["user_role"] = ""

# ==========================================
# 1. 로그인 & 회원가입
# ==========================================
if not st.session_state["logged_in"]:
    st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>🏢 러셀대치학원 시설보수 및 재고관리 시스템</h2>", unsafe_allow_html=True)
    tab_login, tab_signup = st.tabs(["🔑 로그인", "📝 회원가입"])

    with tab_login:
        st.subheader("로그인")
        login_name = st.text_input("이름", key="l_name")
        login_phone = st.text_input("전화번호", type="password", key="l_phone")
        
        if st.button("로그인하기", use_container_width=True):
            c_name = login_name.strip()
            c_phone = clean_phone(login_phone)
            
            df_u = pd.DataFrame(ws_users.get_all_records())
            matched = None
            if not df_u.empty:
                for idx, row in df_u.iterrows():
                    if str(row.get("이름","")).strip() == c_name and clean_phone(str(row.get("전화번호",""))) == c_phone:
                        matched = row
                        break
            
            if matched is not None:
                st.session_state["logged_in"] = True
                st.session_state["user_name"] = c_name
                st.session_state["user_phone"] = c_phone
                st.session_state["user_role"] = str(matched.get("분류","직원")).strip()
                st.success(f"🎉 {c_name}님 환영합니다!")
                st.rerun()
            else:
                st.error("이름 또는 전화번호가 일치하지 않습니다.")

    with tab_signup:
        st.subheader("신규 회원가입")
        signup_role = st.selectbox("구분", ["교무팀", "조교", "직원", "강사"])
        signup_name = st.text_input("이름", key="s_name")
        signup_phone = st.text_input("전화번호", key="s_phone")

        if st.button("회원가입 완료", use_container_width=True):
            c_name = signup_name.strip()
            c_phone = clean_phone(signup_phone)
            if c_name and len(c_phone) >= 8:
                now_str = get_kst_now()
                ws_users.append_row([c_name, f"'{c_phone}", signup_role, now_str])
                st.success("✅ 회원가입 완료! 로그인 탭에서 로그인해 주세요.")
            else:
                st.error("올바른 정보를 입력해 주세요.")
    st.stop()

# ==========================================
# 2. 메인 화면 및 메뉴
# ==========================================
st.sidebar.markdown("### 🏢 러셀대치학원 시스템")
st.sidebar.info(f"👤 **{st.session_state['user_name']}** ({st.session_state['user_role']})")

if st.sidebar.button("🚪 로그아웃", use_container_width=True):
    st.session_state["logged_in"] = False
    st.rerun()

user_role = st.session_state["user_role"]
menu_options = ["🛠️ 시설 보수 및 물품 구매 요청", "📦 물품/비품 재고 관리"] if user_role in ["교무팀", "조교"] else ["🛠️ 시설 보수 및 물품 구매 요청"]
menu = st.sidebar.radio("메뉴 이동", menu_options)

# ==========================================
# 3. 요청 등록 (사진 첨부 기능 포함)
# ==========================================
if menu == "🛠️ 시설 보수 및 물품 구매 요청":
    st.title("🛠️ 시설 보수 및 물품 구매 요청")
    tab1, tab2 = st.tabs(["📝 요청 등록", "📋 현황 확인"])

    with tab1:
        req_type = st.radio("요청 유형", ["🔧 시설 보수 요청", "🛒 신규 물품/비품 구매 요청"], horizontal=True)
        space_cat = st.selectbox("공간 구분", list(SPACES.keys()))
        det_space = st.selectbox("세부 위치", SPACES[space_cat])

        if req_type == "🔧 시설 보수 요청":
            cat = st.selectbox("보수 분류", [c for c in CATEGORY_BY_SPACE[space_cat] if c != "신규 물품 구매 요청"])
            details = st.text_area("상세 고장 내용")
            photo_file = st.file_uploader("📷 현장 사진 첨부 (선택사항)", type=["png", "jpg", "jpeg"])

            if st.button("보수 요청 제출", use_container_width=True):
                if details.strip():
                    with st.spinner("사진 업로드 및 요청 저장 중..."):
                        photo_url = upload_photo_to_drive(photo_file)
                        now_str = get_kst_now()
                        ws_facility.append_row([now_str, space_cat, det_space, f"[보수] {cat}", details, photo_url, "접수완료", st.session_state['user_name']])
                    st.success("✅ 구글 시트에 요청과 사진 정보가 저장되었습니다.")
                else:
                    st.error("상세 고장 내용을 입력해 주세요.")
        else:
            item_n = st.text_input("구매 물품명")
            item_q = st.number_input("수량", min_value=1, value=1)
            item_r = st.text_area("구매 사유")
            photo_file = st.file_uploader("📷 참고 사진/참고자료 첨부 (선택사항)", type=["png", "jpg", "jpeg"])

            if st.button("구매 요청 제출", use_container_width=True):
                if item_n.strip() and item_r.strip():
                    with st.spinner("요청 저장 중..."):
                        photo_url = upload_photo_to_drive(photo_file)
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                        full_d = f"[구매물품] {item_n} ({item_q}개)\n[사유] {item_r}"
                        ws_facility.append_row([now_str, space_cat, det_space, "[구매요청] 물품구매", full_d, photo_url, "접수완료", st.session_state['user_name']])
                    st.success("✅ 구글 시트에 구매 요청이 저장되었습니다.")
                else:
                    st.error("물품명과 사유를 모두 입력해 주세요.")

    with tab2:
        df_f = pd.DataFrame(ws_facility.get_all_records())
        st.dataframe(df_f, use_container_width=True)

# ==========================================
# 4. 재고 관리 (실사/입고/출고 분리 보완판)
# ==========================================
elif menu == "📦 물품/비품 재고 관리":
    st.title("📦 재고 관리 System")
    tab1, tab2, tab3 = st.tabs(["📝 마감 재고 실사", "📥 입고 및 출고(사용) 등록", "📊 재고 현황 & 이력"])

    # 안전하게 구글 시트 데이터 불러오기
    def get_safe_inventory_df():
        try:
            records = ws_inventory.get_all_records()
            return pd.DataFrame(records)
        except Exception:
            all_cols = ["일시", "구분", "작성자"] + ALL_ITEMS + ["비고"]
            return pd.DataFrame(columns=all_cols)

    # 탭 1: 마감 재고 실사 (기준 재고 등록)
    with tab1:
        st.subheader("📝 일일 마감 실사 재고 입력")
        st.caption("※ 실제 창고에 남아있는 실사 수량을 입력합니다. (이후 입고/출고 계산의 기준점이 됩니다)")
        
        with st.form("inv_form"):
            now_str = get_kst_now()
            st.write(f"등록일시: **{now_str}** | 작성자: **{st.session_state['user_name']}**")
            
            st.markdown("#### 📄 소모품 (BOX)")
            c_vals = {item: st.number_input(f"{item}", min_value=0, value=0, key=f"c_{item}") for item in CONSUMABLES}
            
            st.markdown("#### 💻 비품/기기 (개)")
            e_vals = {item: st.number_input(f"{item}", min_value=0, value=0, key=f"e_{item}") for item in EQUIPMENT}

            if st.form_submit_button("마감 실사 저장", use_container_width=True):
                all_v = {**c_vals, **e_vals}
                # 구분: '실사'
                row = [now_str, "실사", st.session_state['user_name']] + [all_v[i] for i in ALL_ITEMS] + ["정기 마감 실사"]
                ws_inventory.append_row(row)
                st.success("✅ 마감 실사 데이터가 구글 시트에 독립적으로 저장되었습니다.")
                st.rerun()

    # 탭 2: 입고 및 출고 등록 (독립 내역 기록)
    with tab2:
        st.subheader("📥 입고 / 📤 출고(사용) 등록")
        st.caption("※ 수량 변동 내역(입고/사용)만 별도로 등록합니다.")
        
        inout_type = st.radio("작업 구분", ["📥 입고 (수량 추가)", "📤 출고 (사용/차감)"], horizontal=True)
        target_item = st.selectbox("물품 선택", ALL_ITEMS)
        qty = st.number_input("수량", min_value=1, value=1)
        memo = st.text_input("비고/메모 (예: OO문구 구매분, 3층 자습관 교체용 등)", placeholder="사유 입력")

        if st.button("내역 등록하기", use_container_width=True):
            now_str = get_kst_now()
            is_in = "입고" in inout_type
            action_label = "입고" if is_in else "출고"
            
            # 입고는 양수(+), 출고는 음수(-) 처리하여 단일 물품 수량 기록
            record_qty = qty if is_in else -qty
            
            row_data = [now_str, action_label, st.session_state['user_name']]
            for item in ALL_ITEMS:
                if item == target_item:
                    row_data.append(record_qty)
                else:
                    row_data.append(0) # 해당 없는 품목은 0으로 기록
            row_data.append(memo)

            ws_inventory.append_row(row_data)
            st.success(f"🎉 [{target_item}] {qty}개 {action_label} 등록 완료! (비고: {memo})")
            st.rerun()

    # 탭 3: 현황 및 누적 계산
    with tab3:
        df_i = get_safe_inventory_df()
        
        if df_i.empty:
            st.info("등록된 재고 데이터가 없습니다.")
        else:
            st.subheader("📊 현재 계산 재고")
            
            # 가장 최근 '실사' 행 위치 찾기
            df_silsa = df_i[df_i["구분"] == "실사"]
            
            if not df_silsa.empty:
                last_silsa_idx = df_silsa.index[-1]
                last_silsa_date = df_i.loc[last_silsa_idx, "일시"]
                st.caption(f"💡 최근 실사일({last_silsa_date}) 이후의 입출고 내역을 자동으로 합산한 현재 재고입니다.")
                
                # 최근 실사 이후의 모든 데이터 누적 합산 (실사값 + 입고값 - 출고값)
                df_calc = df_i.loc[last_silsa_idx:].copy()
                
                current_stock = {}
                for item in ALL_ITEMS:
                    if item in df_calc.columns:
                        # 숫자형으로 변환 후 합산
                        current_stock[item] = pd.to_numeric(df_calc[item], errors='coerce').fillna(0).sum()
                    else:
                        current_stock[item] = 0
                
                # 요약 대시보드 표시
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("##### 📄 주요 소모품 재고")
                    df_c_stock = pd.DataFrame([{"품목": k, "현재재고(BOX)": int(v)} for k, v in current_stock.items() if k in CONSUMABLES])
                    st.dataframe(df_c_stock, use_container_width=True, hide_index=True)
                with col2:
                    st.markdown("##### 💻 비품/기기 재고")
                    df_e_stock = pd.DataFrame([{"품목": k, "현재재고(개)": int(v)} for k, v in current_stock.items() if k in EQUIPMENT])
                    st.dataframe(df_e_stock, use_container_width=True, hide_index=True)
            else:
                st.warning("⚠️ 등록된 '마감 실사' 데이터가 없습니다. 먼저 1번째 탭에서 마감 실사를 진행해 주세요.")

            st.markdown("---")
            st.subheader("📋 전체 이력 히스토리 (실사/입고/출고)")
            st.dataframe(df_i, use_container_width=True)