import streamlit as st
import pandas as pd
import os
from datetime import datetime

# 데이터 저장 파일 이름
DATA_FILE = "facility_requests.csv"

# 저장 파일이 없으면 자동 생성
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=["접수시간", "구분", "세부위치", "보수분류", "상세내용", "처리상태"])
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

# 공간 데이터 정의
SPACES = {
    "강의실": [
        "대형강의실", "201호", "202호", "203호", "204호", 
        "601호", "602호", "701호", "702호", "703호", "704호"
    ],
    "자습관": [
        "3-1관", "3-2관", "3-3관", "3-4관", "3-5관",
        "4-1관", "4-2관", "4-3관", "4-4관", "4-5관",
        "5-1관", "5-2관", "5-3관", "6-1관", "6-2관", "7-1관"
    ],
    "화장실": [
        "2층 남자화장실", "2층 여자화장실",
        "3층 남자화장실", "3층 여자화장실",
        "5층 남자화장실", "5층 여자화장실",
        "6층 남자화장실", "6층 여자화장실",
        "7층 남자화장실", "7층 여자화장실"
    ],
    "기타공간": ["기타공간 (복도/엘리베이터/로비 등)"]
}

# 웹페이지 기본 설정
st.set_page_config(page_title="러셀대치 시설 보수 요청 시스템", layout="centered")

st.title("🛠️ 러셀대치학원 시설 보수 요청")
st.caption("강의실, 자습관, 화장실 등 시설물 보수 및 점검 요청을 등록해 주세요.")

# 탭 나누기 (1. 요청 등록 / 2. 담당자 관리)
tab1, tab2 = st.tabs(["📝 보수 요청 등록", "📋 담당자 보수 현황 관리"])

# [탭 1] 강사 및 임직원 요청 등록
with tab1:
    st.subheader("📍 위치 및 보수 내용 입력")
    
    # 1. 대분류 선택
    space_category = st.selectbox("1. 공간 구분 선택", list(SPACES.keys()))
    
    # 2. 대분류에 따른 세부위치 선택
    detailed_space = st.selectbox("2. 세부 위치 선택", SPACES[space_category])
    
    # 3. 보수 분류 선택
    category = st.selectbox(
        "3. 보수 분류", 
        ["빔프로젝터/음향", "냉난방/환기", "조명/전기", "책상/의자/책장", "수도/화장실 설비", "문/창문/열쇠", "기타 시설"]
    )
    
    # 4. 상세 내용 작성
    details = st.text_area("4. 상세 요청 내용", placeholder="예: 3층 남자화장실 세면대 수도꼭지 누수, 201호 프로젝터 화면 수평 미조정 등")
    
    if st.button("🚨 보수 요청 제출하기", use_container_width=True):
        if not details.strip():
            st.error("상세 요청 내용을 입력해 주세요!")
        else:
            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            new_data = pd.DataFrame(
                [[now, space_category, detailed_space, category, details, "접수완료"]], 
                columns=["접수시간", "구분", "세부위치", "보수분류", "상세내용", "처리상태"]
            )
            
            # 기존 데이터에 추가 저장
            df_existing = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
            df_updated = pd.concat([df_existing, new_data], ignore_index=True)
            df_updated.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
            
            st.success(f"✅ [{space_category} > {detailed_space}] 보수 요청이 정상 접수되었습니다!")
            st.balloons() # 접수 완료 시 축하 효과

# [탭 2] 시설담당관 현황 조회 및 상태 변경
with tab2:
    st.subheader("📊 실시간 보수 요청 접수 내역")
    
    df_current = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
    
    if len(df_current) == 0:
        st.info("현재 접수된 보수 요청이 없습니다.")
    else:
        # 필터 기능 (구분별 보기)
        filter_cat = st.selectbox("구분별 필터링", ["전체보기"] + list(SPACES.keys()))
        if filter_cat != "전체보기":
            display_df = df_current[df_current["구분"] == filter_cat]
        else:
            display_df = df_current
            
        st.dataframe(display_df, use_container_width=True)
        
        st.divider()
        st.write("🔧 **처리 상태 업데이트**")
        
        # 선택한 항목 처리 상태 변경
        col1, col2 = st.columns(2)
        with col1:
            selected_index = st.number_input("상태 변경할 행 번호 (Index)", min_value=0, max_value=len(df_current)-1, step=1)
        with col2:
            new_status = st.selectbox("변경할 상태", ["접수완료", "진행중(업체예약)", "보수완료", "보류"])
        
        if st.button("상태 변경 저장"):
            df_current.loc[selected_index, "처리상태"] = new_status
            df_current.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
            st.success(f"Index {selected_index}번 항목이 [{new_status}] 상태로 변경되었습니다.")
            st.rerun()