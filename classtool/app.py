"""
메가스터디 러셀대치학원 단과 강좌 [RMS 수업목록(.xls/.xlsx) ↔ 강의계획서 PDF] 종합 검수 시스템
1. 메인 상단 2열 드롭존 & 컨트롤 바:
   - [📊 1. RMS 강좌정보 엑셀 업로드] & [📄 2. 강의계획서 PDF 업로드]
   - [☑️ 판매중지 강좌 제외 (기본값 ON)] 원클릭 토글
   - [📥 통합 결과지 엑셀(.xlsx) 다운로드] & [📋 구글 시트 붙여넣기용 복사] 고정 배치
2. 상단 고정 헤드라인 배너 & 5대 KPI 카드:
   - 유효 강좌수, 판매중지 제외 건수, 총 강좌수, 총 수강생, 결제완료, 미결제 가등록, 결제율(%)
3. 직관적인 4대 핵심 탭 구성:
   - Tab 1: 📈 강사별 수강생 & 월별 누적 통계 (당월 순위 + 월별 누적 피벗 테이블 + 전월 대비 증감 MoM + 드릴다운)
   - Tab 2: 🔍 엑셀 vs PDF 강의계획서 정밀 검수 (7대 항목 정밀 대조 + RMS 판매중지/폐강 의심 경고)
   - Tab 3: ✍️ 맞춤법 교정 & ⚡ 초단축 문장·강의명 제안 (Gemini 3.8 초단축 3종 스타일 + 긴문장 축약)
   - Tab 4: 💾 누적 DB 백업/관리 & 제외 강좌 확인 (판매중지 강좌 리스트 + 누적 DB 백업/복원 + 강사 매핑)
"""

from __future__ import annotations

import importlib
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import core.cumulative_db as _cum_db_mod
import core.matcher as _matcher_mod
import core.text_improver as _text_improver_mod
import core.title_suggester as _title_suggester_mod
import parsers.pdf_parser as _pdf_parser_mod
import parsers.rms_parser as _rms_parser_mod

# 핫리로드 시 하위 모듈 변경사항 즉시 반영
importlib.reload(_rms_parser_mod)
importlib.reload(_pdf_parser_mod)
importlib.reload(_text_improver_mod)
importlib.reload(_title_suggester_mod)
importlib.reload(_cum_db_mod)
importlib.reload(_matcher_mod)

from core.cumulative_db import (
    build_pivot_table,
    delete_month_data,
    export_enrollment_backup_excel,
    get_available_months,
    get_monthly_summary_table,
    import_enrollment_backup,
    load_enrollment_db,
    upsert_month_data,
)
from core.matcher import (
    EXPORT_DISPLAY_COLUMNS,
    InstructorMapper,
    export_crosscheck_to_csv_bytes,
    export_crosscheck_to_excel,
    export_crosscheck_to_tsv_string,
    export_results_to_google_sheets,
    generate_pdf_fix_request_messages,
    run_rms_pdf_crosscheck,
)
from core.title_suggester import (
    DIRECTION_PROMPTS,
    batch_suggest_course_titles,
    get_gemini_api_key,
    suggest_course_titles_with_gemini,
)
from parsers.pdf_parser import parse_syllabus_pdf
from parsers.rms_parser import (
    detect_dominant_year_month,
    extract_rms_summary,
    filter_rms_active_courses,
    generate_instructor_student_summary,
    parse_rms_xls,
)

# -------------------------------------------------------------------------
# 기본 페이지 설정 및 커스텀 CSS
# -------------------------------------------------------------------------
st.set_page_config(
    page_title="러셀대치 RMS ↔ 강의계획서 PDF 종합 검수 시스템",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #0f2b46 0%, #1e3a8a 50%, #2563eb 100%);
        color: white;
        padding: 1.15rem 1.5rem;
        border-radius: 10px;
        margin-bottom: 0.8rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-header h2 {
        color: white !important;
        margin: 0;
        font-size: 1.4rem;
        font-weight: 700;
    }
    .main-header p {
        color: #dbeafe;
        margin: 0.35rem 0 0 0;
        font-size: 0.92rem;
    }
    .rms-headline-banner {
        background: #f0fdf4;
        border: 2px solid #22c55e;
        border-radius: 8px;
        padding: 0.75rem 1.25rem;
        margin-bottom: 0.8rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .rms-headline-title {
        font-size: 1.02rem;
        font-weight: 700;
        color: #15803d;
    }
    .proposal-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.9rem;
        margin-bottom: 0.8rem;
    }
    .proposal-title {
        font-weight: 700;
        font-size: 0.95rem;
        color: #1e293b;
        margin-bottom: 0.25rem;
    }
    .export-bar-box {
        background: #f8fafc;
        border: 1px solid #cbd5e1;
        padding: 0.75rem 1rem;
        border-radius: 8px;
        margin-bottom: 0.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

CONFIG_PATH = PROJECT_ROOT / "config" / "instructor_mapping.json"
USER_SETTINGS_PATH = PROJECT_ROOT / "config" / "user_settings.json"

DEFAULT_USER_SETTINGS: Dict[str, Any] = {
    "target_ym": "2027-01",
    "title_sim_threshold": 55,
    "max_char_threshold": 55,
    "gemini_api_key": "",
    "export_sheet_url": "",
    "exclude_stopped": True,
}


def load_user_settings() -> Dict[str, Any]:
    settings = dict(DEFAULT_USER_SETTINGS)
    if USER_SETTINGS_PATH.exists():
        try:
            loaded = json.loads(USER_SETTINGS_PATH.read_text(encoding="utf-8-sig"))
            if isinstance(loaded, dict):
                settings.update(loaded)
        except Exception:
            pass
    return settings


def save_user_settings(settings: Dict[str, Any]) -> None:
    try:
        USER_SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
        USER_SETTINGS_PATH.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


@st.cache_resource
def get_or_init_mapper() -> InstructorMapper:
    return InstructorMapper(CONFIG_PATH)


mapper = get_or_init_mapper()
user_settings = load_user_settings()

# 세션 상태 초기화
if "rms_df_raw" not in st.session_state:
    st.session_state["rms_df_raw"] = None
if "syllabus_pdf_df" not in st.session_state:
    st.session_state["syllabus_pdf_df"] = None
if "crosscheck_result" not in st.session_state:
    st.session_state["crosscheck_result"] = None
if "ai_titles_df" not in st.session_state:
    st.session_state["ai_titles_df"] = None
if "ai_single_suggestions" not in st.session_state:
    st.session_state["ai_single_suggestions"] = {}

# -------------------------------------------------------------------------
# 상단 타이틀 배너
# -------------------------------------------------------------------------
st.markdown(
    """
    <div class="main-header">
        <h2>메가스터디 러셀 대치학원 — RMS ↔ 강의계획서 PDF 종합 검수 시스템</h2>
        <p>RMS 강좌목록 엑셀(.xls/.xlsx)과 강의계획서 PDF(.pdf) 7대 항목 정밀 검수 · 수강생/결제자수 누적 통계 · 15자 이내 초단축 AI 강의명 제안</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------------
# 사이드바 설정 영역
# -------------------------------------------------------------------------
with st.sidebar:
    st.subheader("⚙️ 시스템 및 검수 설정")

    # Gemini API 키 설정
    existing_key = user_settings.get("gemini_api_key", "") or get_gemini_api_key() or ""
    gemini_key_input = st.text_input(
        "🔑 Gemini API Key (Gemini 3.8 초단축 AI 카피라이팅)",
        value=existing_key,
        type="password",
        help="Google AI Studio에서 발급받은 Gemini API 키를 입력하면 최신 Gemini 3.8 초단축 리브랜딩이 활성화됩니다.",
    )
    if gemini_key_input != user_settings.get("gemini_api_key", ""):
        user_settings["gemini_api_key"] = gemini_key_input
        save_user_settings(user_settings)

    if gemini_key_input:
        st.caption("🟢 **Gemini 3.8 (High-Impact Copywriting Engine) 연결 준비 완료**")
    else:
        st.caption("🟡 **내장 대치동 초단축 규칙 엔진(Fallback) 가동 중**")

    st.markdown("---")
    st.subheader("✍️ 문장 검수 옵션")
    max_char_threshold = st.slider(
        "긴 문장 감지 기준 (글자 수)",
        min_value=40,
        max_value=100,
        value=int(user_settings.get("max_char_threshold", 55)),
        step=5,
        help="PDF 내 강의특징 문장이 이 글자 수를 초과하면 간결한 문장형/개조식 축약안을 제시합니다.",
    )
    if max_char_threshold != user_settings.get("max_char_threshold", 55):
        user_settings["max_char_threshold"] = max_char_threshold
        save_user_settings(user_settings)

    target_ym_input = st.text_input(
        "📅 검수 기준 연월 (YYYY-MM)",
        value=user_settings.get("target_ym", "2027-01"),
        help="RMS 개강일 및 PDF 진도표 첫날의 연월 기준입니다.",
    )
    if target_ym_input != user_settings.get("target_ym", "2027-01"):
        user_settings["target_ym"] = target_ym_input
        save_user_settings(user_settings)

# -------------------------------------------------------------------------
# 원스톱 드롭존: 최상단 2열 파일 업로더
# -------------------------------------------------------------------------
st.markdown("#### 📂 1. 데이터 파일 업로드 (RMS 강좌정보 엑셀 & 강의계획서 PDF)")
col_upload_rms, col_upload_pdf = st.columns(2)

with col_upload_rms:
    uploaded_rms_files = st.file_uploader(
        "📋 **1. RMS 강좌정보 엑셀 업로드** (`.xls`, `.xlsx` 다중 지원)",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        help="메가스터디 RMS에서 다운로드한 강좌목록 엑셀 파일입니다. 여러 달의 파일을 한 번에 올리면 월별 누적 DB에 즉시 반영됩니다.",
        key="uploader_rms_files",
    )

with col_upload_pdf:
    uploaded_pdf_files = st.file_uploader(
        "📄 **2. 강의계획서 PDF 업로드** (`.pdf` 다중 지원)",
        type=["pdf"],
        accept_multiple_files=True,
        help="브로슈어 제작용 단일 통합 PDF 또는 강사별 개별 PDF 파일들을 동시에 업로드할 수 있습니다.",
        key="uploader_pdf_files",
    )

# -------------------------------------------------------------------------
# 상단 컨트롤 바: [☑️ 판매중지 강좌 제외] & [다운로드/복사 버튼]
# -------------------------------------------------------------------------
col_toggle, col_dn_excel, col_btn_copy = st.columns([2.5, 1.5, 1.5])

with col_toggle:
    exclude_stopped = st.checkbox(
        "☑️ **판매중지 강좌 제외** (판매중인 강좌만 집계·검수)",
        value=True,
        help="체크 시 RMS 엑셀의 '판매' 컬럼이 '판매'인 유효 강좌만 집계하고, '판매중지' 강좌는 제외합니다. (사전 검수 시에는 체크를 해제하세요)",
        key="exclude_stopped_checkbox",
    )

# -------------------------------------------------------------------------
# RMS 파일 파싱 및 처리 (세션 캐싱 적용)
# -------------------------------------------------------------------------
raw_rms_df = None
selected_rms_ym = user_settings.get("target_ym", "2027-01")

if uploaded_rms_files:
    rms_sig = tuple((getattr(f, "name", ""), getattr(f, "size", 0)) for f in uploaded_rms_files)
    if st.session_state.get("cached_rms_sig") == rms_sig and st.session_state.get("rms_df_raw") is not None:
        raw_rms_df = st.session_state["rms_df_raw"]
        selected_rms_ym = st.session_state.get("selected_rms_ym", selected_rms_ym)
    else:
        # 파일이 1개 또는 여러 개인 경우 처리
        if len(uploaded_rms_files) == 1:
            f = uploaded_rms_files[0]
            try:
                raw_rms_df = parse_rms_xls(f)
                detected_ym = detect_dominant_year_month(raw_rms_df, default_ym=selected_rms_ym)
                selected_rms_ym = detected_ym
                st.session_state["rms_df_raw"] = raw_rms_df
                st.session_state["selected_rms_ym"] = selected_rms_ym
                st.session_state["cached_rms_sig"] = rms_sig
            except Exception as e:
                st.error(f"❌ RMS 엑셀 파싱 중 오류가 발생했습니다: {e}")
        else:
            # 여러 개 파일이 업로드된 경우
            file_map: Dict[str, pd.DataFrame] = {}
            for f in uploaded_rms_files:
                try:
                    parsed = parse_rms_xls(f)
                    ym_found = detect_dominant_year_month(parsed, default_ym="2027-01")
                    file_map[f"{f.name} ({ym_found})"] = parsed
                    # 누적 DB 자동 Upsert
                    upsert_month_data(parsed, ym_found, filter_stopped=exclude_stopped)
                except Exception as e:
                    st.warning(f"⚠️ '{f.name}' 파싱 실패: {e}")

            if file_map:
                st.success(f"🎉 **총 {len(file_map)}개 월의 RMS 파일이 파싱되어 누적 DB에 자동 반영되었습니다.**")
                selected_name = st.selectbox(
                    "현재 검수 및 대시보드 기준으로 볼 RMS 파일을 선택하세요:",
                    options=list(file_map.keys()),
                    index=0,
                )
                raw_rms_df = file_map[selected_name]
                st.session_state["rms_df_raw"] = raw_rms_df
                selected_rms_ym = detect_dominant_year_month(raw_rms_df, default_ym=selected_rms_ym)
                st.session_state["selected_rms_ym"] = selected_rms_ym
                st.session_state["cached_rms_sig"] = rms_sig
elif st.session_state.get("rms_df_raw") is not None:
    raw_rms_df = st.session_state["rms_df_raw"]
    selected_rms_ym = st.session_state.get("selected_rms_ym", selected_rms_ym)

# -------------------------------------------------------------------------
# PDF 파일 파싱 및 처리 (세션 캐싱으로 0.00초 즉시 반환)
# -------------------------------------------------------------------------
syllabus_pdf_df = None
if uploaded_pdf_files:
    pdf_sig = tuple((getattr(f, "name", ""), getattr(f, "size", 0)) for f in uploaded_pdf_files)
    if st.session_state.get("cached_pdf_sig") == pdf_sig and st.session_state.get("syllabus_pdf_df") is not None:
        syllabus_pdf_df = st.session_state["syllabus_pdf_df"]
    else:
        pdf_source = uploaded_pdf_files[0] if len(uploaded_pdf_files) == 1 else uploaded_pdf_files
        try:
            with st.spinner("📄 PDF 강의계획서 초고속 정밀 파싱 중..."):
                syllabus_pdf_df = parse_syllabus_pdf(pdf_source)
                st.session_state["syllabus_pdf_df"] = syllabus_pdf_df
                st.session_state["cached_pdf_sig"] = pdf_sig
        except Exception as e:
            st.error(f"❌ PDF 파싱 중 오류가 발생했습니다: {e}")
elif st.session_state.get("syllabus_pdf_df") is not None:
    syllabus_pdf_df = st.session_state["syllabus_pdf_df"]

# -------------------------------------------------------------------------
# 판매중지 필터 적용 및 요약 배너 생성
# -------------------------------------------------------------------------
active_rms_df = None
stopped_rms_df = pd.DataFrame()
rms_summary = extract_rms_summary(None)

if raw_rms_df is not None:
    active_rms_df, stopped_rms_df = filter_rms_active_courses(raw_rms_df, exclude_stopped=exclude_stopped)
    rms_summary = extract_rms_summary(
        active_rms_df,
        total_raw_count=len(raw_rms_df),
        stopped_count=len(stopped_rms_df) if exclude_stopped else 0,
    )

    # 상단 요약 배너 표시
    st.markdown(
        f"""
        <div class="rms-headline-banner">
            <span class="rms-headline-title">{rms_summary["headline_banner"]}</span>
            <span style="font-size: 0.85rem; color: #166534; font-weight: 600;">
                기준 연월: {selected_rms_ym} | 필터: {'판매중지 강좌 제외' if exclude_stopped else '전체 강좌 포함'}
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 5대 핵심 KPI 메트릭 카드
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    with kpi_col1:
        st.metric(
            label="유효 강좌수",
            value=f"{rms_summary['total_courses']}건",
            delta=f"판매중지 {len(stopped_rms_df)}건 제외" if (exclude_stopped and len(stopped_rms_df) > 0) else None,
            delta_color="inverse" if (exclude_stopped and len(stopped_rms_df) > 0) else "normal",
        )
    with kpi_col2:
        st.metric(
            label="총 수강생수",
            value=f"{rms_summary['total_students']:,}명",
            delta="가등록 포함",
        )
    with kpi_col3:
        st.metric(
            label="결제완료 인원",
            value=f"{rms_summary['paid_students']:,}명",
            delta=f"완납·분납 {rms_summary.get('installment_students', 0):,}명",
        )
    with kpi_col4:
        st.metric(
            label="미결제 가등록",
            value=f"{rms_summary['unpaid_students']:,}명",
            delta="결제 대기",
            delta_color="inverse",
        )
    with kpi_col5:
        st.metric(
            label="수강생 결제율",
            value=f"{rms_summary['payment_rate']}%",
            delta="결제완료/가등록",
        )

# -------------------------------------------------------------------------
# 교차 검수 실행
# -------------------------------------------------------------------------
crosscheck_result = None
if active_rms_df is not None:
    crosscheck_result = run_rms_pdf_crosscheck(
        rms_df=active_rms_df,
        syllabus_pdf_df=syllabus_pdf_df,
        mapper=mapper,
        target_ym=selected_rms_ym,
        title_similarity_threshold=float(user_settings.get("title_sim_threshold", 55)),
        max_char_threshold=int(max_char_threshold),
        gemini_api_key=gemini_key_input or None,
        stopped_rms_df=stopped_rms_df,
    )
    # AI 강의명 추천 데이터프레임이 세션에 있으면 병합
    if st.session_state.get("ai_titles_df") is not None:
        crosscheck_result["ai_titles_df"] = st.session_state["ai_titles_df"]
    st.session_state["crosscheck_result"] = crosscheck_result

# 상단 컨트롤 바의 [다운로드 버튼] 및 [복사 버튼] 활성화
if crosscheck_result is not None:
    with col_dn_excel:
        excel_bytes = export_crosscheck_to_excel(crosscheck_result, target_ym=selected_rms_ym)
        st.download_button(
            label="📥 **통합 결과지 엑셀(.xlsx)**",
            data=excel_bytes,
            file_name=f"러셀대치_강의계획서_정밀검수결과지_{selected_rms_ym}_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_btn_copy:
        tsv_str = export_crosscheck_to_tsv_string(crosscheck_result.get("comparison_df", pd.DataFrame()))
        with st.popover("📋 **구글시트 복사**", use_container_width=True):
            st.caption("아래 텍스트를 복사(Ctrl+C)하여 구글 스프레드시트에 붙여넣기(Ctrl+V)하세요:")
            st.text_area("TSV 데이터", value=tsv_str, height=140)
            st.link_button("🌐 sheets.new (새 시트 열기)", "https://sheets.new")

st.markdown("---")

# -------------------------------------------------------------------------
# 4대 핵심 탭 인터페이스 구성
# -------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📈 Tab 1: 강사별 수강생 & 월별 누적 통계",
        "🔍 Tab 2: 엑셀 vs PDF 강의계획서 정밀 검수",
        "✍️ Tab 3: 맞춤법 교정 & ⚡ 초단축 문장·강의명 제안",
        "💾 Tab 4: 누적 DB 백업/관리 & 제외 강좌 확인",
    ]
)

# =========================================================================
# Tab 1: 📈 강사별 수강생 & 월별 누적 통계
# =========================================================================
with tab1:
    if active_rms_df is None:
        st.info("💡 상단 업로드 영역에서 **RMS 수업목록 엑셀 파일(.xls, .xlsx)**을 업로드하면 수강생 현황이 즉시 표시됩니다.")
    else:
        st.markdown(f"### 📊 1. 당월({selected_rms_ym}) 강사별 수강생 순위 집계표")
        inst_summary_df = generate_instructor_student_summary(active_rms_df)

        col_search_inst, col_save_db = st.columns([3, 1])
        with col_search_inst:
            kw_inst = st.text_input("🔍 강사명 또는 과목 검색:", placeholder="강사명(예: 이지원, 지인해) 또는 과목 입력", key="kw_inst_search")
        with col_save_db:
            st.write("")
            st.write("")
            if st.button("💾 **현재 월 누적 DB에 저장**", use_container_width=True):
                ok, msg = upsert_month_data(raw_rms_df, selected_rms_ym, filter_stopped=exclude_stopped)
                if ok:
                    st.toast(msg, icon="✅")
                    st.success(msg)
                else:
                    st.error(msg)

        filtered_inst_df = inst_summary_df.copy()
        if kw_inst.strip():
            filtered_inst_df = filtered_inst_df[
                filtered_inst_df["강사명"].str.contains(kw_inst.strip(), case=False, na=False)
                | filtered_inst_df["과목(영역)"].str.contains(kw_inst.strip(), case=False, na=False)
            ]

        st.dataframe(
            filtered_inst_df,
            use_container_width=True,
            column_config={
                "순위": st.column_config.NumberColumn("순위", format="%d위"),
                "총 수강생수(가등록포함)": st.column_config.NumberColumn("총 수강생", format="%d명"),
                "결제완료 인원": st.column_config.NumberColumn("결제완료", format="%d명"),
                "완납&분납 인원": st.column_config.NumberColumn("완납·분납", format="%d명"),
                "가등록(미결제) 인원": st.column_config.NumberColumn("미결제", format="%d명"),
                "결제율(%)": st.column_config.ProgressColumn("결제율", min_value=0, max_value=100, format="%.1f%%"),
                "강좌당 평균 수강생": st.column_config.NumberColumn("강좌당 평균", format="%.1f명"),
            },
            hide_index=True,
        )

        st.markdown("---")
        # -------------------------------------------------------------
        # 월별 강사 결제자수 누적 추이 피벗 테이블
        # -------------------------------------------------------------
        st.markdown("### 📅 2. 월별 강사 누적 통계 피벗 테이블 (전월 대비 증감 MoM)")
        col_metric, col_dummy = st.columns([2.5, 2.5])
        with col_metric:
            selected_metric = st.radio(
                "조회할 누적 통계 지표 선택:",
                options=["결제완료 인원", "총 수강생(가등록포함)", "유효 강좌수", "결제 매출액"],
                horizontal=True,
                key="pivot_metric_radio",
            )

        pivot_df = build_pivot_table(metric_label=selected_metric)
        if pivot_df.empty:
            st.info("💡 아직 저장된 월별 누적 데이터가 없습니다. 상단의 `[💾 현재 월 누적 DB에 저장]` 버튼을 누르거나, 여러 달의 RMS 엑셀 파일을 업로드하세요.")
        else:
            st.dataframe(
                pivot_df,
                use_container_width=True,
                column_config={
                    "순위": st.column_config.NumberColumn("순위", format="%d위"),
                    "누적 합계": st.column_config.NumberColumn("누적 합계", format="%d"),
                    "월평균": st.column_config.NumberColumn("월평균", format="%.1f"),
                },
                hide_index=True,
            )

        st.markdown("---")
        # -------------------------------------------------------------
        # 강사별 개별 강좌 상세 드릴다운
        # -------------------------------------------------------------
        st.markdown("### 🔎 3. 강사별 개설 강좌 상세 드릴다운")
        all_instructors = sorted(list(active_rms_df["강사"].dropna().unique()))
        selected_drill_inst = st.selectbox("상세 조회할 강사를 선택하세요:", options=all_instructors)
        if selected_drill_inst:
            drill_df = active_rms_df[active_rms_df["강사"] == selected_drill_inst].copy()
            show_cols = [
                "강좌코드",
                "과목",
                "학년",
                "강좌명",
                "간결수업시간",
                "개강일_짧은표시",
                "종강일_짧은표시",
                "회차",
                "수강료",
                "가등록포함수강생수",
                "결제인원",
                "가등록(미결제)",
                "결제율(%)",
            ]
            valid_show_cols = [c for c in show_cols if c in drill_df.columns]
            st.dataframe(
                drill_df[valid_show_cols].reset_index(drop=True),
                use_container_width=True,
                column_config={
                    "수강료": st.column_config.NumberColumn("수강료", format="%d원"),
                    "가등록포함수강생수": st.column_config.NumberColumn("가등록", format="%d명"),
                    "결제인원": st.column_config.NumberColumn("결제완료", format="%d명"),
                    "가등록(미결제)": st.column_config.NumberColumn("미결제", format="%d명"),
                },
                hide_index=True,
            )

# =========================================================================
# Tab 2: 🔍 엑셀 vs PDF 강의계획서 정밀 검수
# =========================================================================
with tab2:
    if crosscheck_result is None or active_rms_df is None:
        st.info("💡 상단에서 **RMS 엑셀 파일**과 **강의계획서 PDF 파일**을 업로드하면 7대 항목 정밀 대조가 실행됩니다.")
    else:
        # 1. 사전 검수 알림 및 폐강 경고
        pdf_stopped_df = crosscheck_result.get("pdf_stopped_courses_df", pd.DataFrame())
        if not pdf_stopped_df.empty:
            st.warning(
                f"⚠️ **현재 PDF에 수록된 강좌 중 {len(pdf_stopped_df)}건이 RMS상 '판매중지' 상태입니다.**\n\n"
                f"• **접수 오픈 전 사전 검수**를 진행하시는 경우라면 상단의 **`[☑️ 판매중지 강좌 제외]`** 체크를 해제하세요.\n"
                f"• 만약 **실제 폐강된 강좌**라면 배포용 강의계획서 PDF에서 해당 강좌의 삭제가 필요합니다."
            )
            with st.expander(f"📋 RMS 판매중지(폐강 의심) 강좌 리스트 ({len(pdf_stopped_df)}건 보기)", expanded=True):
                st.dataframe(
                    pdf_stopped_df[
                        [
                            "PDF_페이지",
                            "과목",
                            "학년",
                            "강사명",
                            "PDF 강좌명",
                            "RMS 강좌명",
                            "강좌코드",
                            "수업시간",
                            "수강료",
                            "확인 및 조치사항",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

        st.markdown("### 🔍 7대 항목 정밀 대조 검수표")
        kpi = crosscheck_result.get("kpi", {})
        k_col1, k_col2, k_col3, k_col4 = st.columns(4)
        k_col1.metric("총 RMS 유효 강좌", f"{kpi.get('total_rms', 0)}건")
        k_col2.metric("정상 일치 강좌", f"{kpi.get('normal_count', 0)}건")
        k_col3.metric("오류·확인필요 강좌", f"{kpi.get('error_count', 0)}건", delta_color="inverse")
        k_col4.metric("PDF 미수록 강좌", f"{kpi.get('missing_pdf_count', 0)}건", delta_color="inverse")

        # 필터링 라디오
        filter_status = st.radio(
            "표시할 강좌 필터:",
            options=["전체 강좌", "🔴 오류·확인필요 강좌만", "✅ 정상 일치 강좌만", "⚪ PDF 미수록 강좌만"],
            horizontal=True,
        )

        comp_df = crosscheck_result.get("comparison_df", pd.DataFrame())
        if not comp_df.empty:
            if filter_status == "🔴 오류·확인필요 강좌만":
                disp_df = comp_df[comp_df["검수상태"].str.startswith(("🔴", "🟡", "⚠️"))].copy()
            elif filter_status == "✅ 정상 일치 강좌만":
                disp_df = comp_df[comp_df["검수상태"].str.startswith("✅")].copy()
            elif filter_status == "⚪ PDF 미수록 강좌만":
                disp_df = comp_df[comp_df["검수상태"].str.startswith("⚪")].copy()
            else:
                disp_df = comp_df.copy()

            view_cols = [
                "검수상태",
                "No",
                "강좌코드",
                "과목",
                "학년",
                "강사명",
                "RMS 강좌명",
                "PDF 강좌명",
                "PDF 페이지",
                "RMS 개강~종강(회차)",
                "PDF 개요 개강~종강(회차)",
                "PDF 1월진도표(첫날~끝날/회차)",
                "RMS 수강료",
                "PDF 수강료",
                "수업시간 비교",
                "바자관시수 비교",
                "오류 상세 내역",
                "자동 생성된 수정 요청 사항 요약",
            ]
            valid_view_cols = [c for c in view_cols if c in disp_df.columns]
            st.dataframe(disp_df[valid_view_cols], use_container_width=True, hide_index=True)

            # 디자인팀 및 강사용 수정 요청 안내문 자동 생성
            st.markdown("---")
            with st.expander("📢 디자인팀 / 강사용 [수정 요청 메시지] 자동 생성 결과"):
                err_only_df = crosscheck_result.get("error_df", pd.DataFrame())
                msg_dict = generate_pdf_fix_request_messages(err_only_df, target_ym=selected_rms_ym)
                if msg_dict:
                    st.text_area("전체 통합 수정 요청문", value=msg_dict.get("__ALL__", ""), height=220)
                else:
                    st.info("수정이 필요한 오류 강좌가 없습니다.")

# =========================================================================
# Tab 3: ✍️ 맞춤법 교정 & ⚡ 초단축 문장·강의명 제안
# =========================================================================
with tab3:
    subtab_text, subtab_ai = st.tabs(
        [
            "📝 서브탭 A: 국어 맞춤법 교정 & 긴 문장 축약",
            "⚡ 서브탭 B: 15자 이내 초단축 임팩트 강의명 제안",
        ]
    )

    # -------------------------------------------------------------
    # 서브탭 A: 국어 맞춤법 및 긴 문장 축약
    # -------------------------------------------------------------
    with subtab_text:
        if crosscheck_result is None or syllabus_pdf_df is None:
            st.info("💡 강의계획서 PDF를 업로드하면 맞춤법 및 긴 문장 축약 검수가 자동으로 실행됩니다.")
        else:
            spelling_df = crosscheck_result.get("spelling_df", pd.DataFrame())
            long_sentence_df = crosscheck_result.get("long_sentence_df", pd.DataFrame())

            st.markdown(f"#### 🔍 1. 국어 맞춤법 및 오탈자 검수 ({len(spelling_df)}건 발견)")
            if not spelling_df.empty:
                st.dataframe(spelling_df, use_container_width=True, hide_index=True)
            else:
                st.success("✅ 감지된 맞춤법 및 오탈자 오류가 없습니다.")

            st.markdown("---")
            st.markdown(f"#### ✍️ 2. 긴 문장({max_char_threshold}자 초과) 축약 및 개조식 제안 ({len(long_sentence_df)}건)")
            if not long_sentence_df.empty:
                st.dataframe(
                    long_sentence_df[
                        [
                            "강좌명",
                            "강사명",
                            "섹션",
                            "글자수",
                            "원문문장",
                            "제안1_서술형",
                            "제안2_개조식",
                            "최종문장",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.success(f"✅ {max_char_threshold}자를 초과하는 긴 문장이 없습니다.")

    # -------------------------------------------------------------
    # 서브탭 B: 15자 이내 초단축 임팩트 강의명 제안
    # -------------------------------------------------------------
    with subtab_ai:
        st.markdown("### ⚡ 15자 이내 초단축·임팩트형 AI 강의명 리브랜딩 (Gemini 3.8 Engine)")
        st.markdown(
            "시간표 한 칸과 브로슈어에 0.5초 만에 각인될 수 있도록, "
            "서술형 수식어를 배제하고 **공백 포함 10~15자 내외의 명사형 초단축 강의명**을 제안합니다."
        )

        direction_choice = st.radio(
            "카피라이팅 방향성 선택:",
            options=[
                "대치동 트렌드 종합",
                "내신 1등급/학교별 기출 강조",
                "수능 평가원/실전 킬러 강조",
                "개념 완성/중위권 도약 강조",
            ],
            horizontal=True,
            key="ai_direction_radio",
        )

        if active_rms_df is not None:
            courses_for_ai = active_rms_df.to_dict(orient="records")
            col_ai_btn, col_ai_cnt = st.columns([2, 3])
            with col_ai_btn:
                btn_run_ai = st.button("⚡ **15자 이내 초단축 강의명 일괄 생성**", type="primary", use_container_width=True)

            if btn_run_ai:
                progress_bar = st.progress(0.0)
                status_text = st.empty()

                def _ai_progress(cur: int, tot: int, name: str) -> None:
                    progress_bar.progress(cur / tot)
                    status_text.caption(f"🤖 AI 초단축 카피 생성 중 ({cur}/{tot}): **{name}**")

                ai_results = batch_suggest_course_titles(
                    course_list=courses_for_ai,
                    api_key=gemini_key_input or None,
                    direction=direction_choice,
                    progress_callback=_ai_progress,
                )
                progress_bar.empty()
                status_text.empty()

                st.session_state["ai_titles_df"] = pd.DataFrame(ai_results)
                if crosscheck_result is not None:
                    crosscheck_result["ai_titles_df"] = st.session_state["ai_titles_df"]
                st.success(f"🎉 **총 {len(ai_results)}개 강좌의 15자 이내 초단축 강의명 생성이 완료되었습니다.**")

            # 결과 데이터프레임 표시
            ai_df = st.session_state.get("ai_titles_df")
            if ai_df is not None and not ai_df.empty:
                st.dataframe(
                    ai_df[
                        [
                            "순번",
                            "과목",
                            "학년",
                            "강사명",
                            "현재강좌명",
                            "제안1_초단축직관형",
                            "제안2_초단축시그니처",
                            "제안3_초단축킬러형",
                            "임팩트점수",
                            "진단비평",
                        ]
                    ],
                    use_container_width=True,
                    column_config={
                        "임팩트점수": st.column_config.NumberColumn("임팩트점수", format="%d점"),
                    },
                    hide_index=True,
                )

                st.markdown("---")
                st.markdown("#### 🎴 개별 강좌 초단축 스타일 카드 및 원클릭 복사")
                sel_c_idx = st.selectbox(
                    "상세 카드 및 재추천을 확인할 강좌를 선택하세요:",
                    options=range(len(ai_df)),
                    format_func=lambda i: f"[{ai_df.iloc[i]['과목']} {ai_df.iloc[i]['강사명']}T] {ai_df.iloc[i]['현재강좌명']}",
                )
                c_data = ai_df.iloc[sel_c_idx].to_dict()

                col_c1, col_c2, col_c3 = st.columns(3)
                with col_c1:
                    st.markdown(
                        f"""
                        <div class="proposal-card">
                            <div class="proposal-title">⚡ 스타일 1: 핵심직관형 (12자 내외)</div>
                            <h4 style="color: #1e3a8a; margin: 0.4rem 0;">{c_data.get('제안1_초단축직관형', '')}</h4>
                            <p style="font-size: 0.85rem; color: #475569;">글자수: {len(c_data.get('제안1_초단축직관형', ''))}자 | 직관적 목표 각인</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with col_c2:
                    st.markdown(
                        f"""
                        <div class="proposal-card">
                            <div class="proposal-title">💎 스타일 2: 시그니처형 (14자 내외)</div>
                            <h4 style="color: #065f46; margin: 0.4rem 0;">{c_data.get('제안2_초단축시그니처', '')}</h4>
                            <p style="font-size: 0.85rem; color: #475569;">글자수: {len(c_data.get('제안2_초단축시그니처', ''))}자 | 강사 브랜드 시그니처</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with col_c3:
                    st.markdown(
                        f"""
                        <div class="proposal-card">
                            <div class="proposal-title">🔥 스타일 3: 킬러임팩트형 (12자 내외)</div>
                            <h4 style="color: #991b1b; margin: 0.4rem 0;">{c_data.get('제안3_초단축킬러형', '')}</h4>
                            <p style="font-size: 0.85rem; color: #475569;">글자수: {len(c_data.get('제안3_초단축킬러형', ''))}자 | 1등급 킬러 공략</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                st.info(f"💡 **진단 비평:** {c_data.get('진단비평', '')}")

                if st.button(f"🔄 [{c_data['강사명']}T {c_data['현재강좌명']}] 이 강좌만 다시 추천받기"):
                    with st.spinner("AI가 새로운 초단축 카피를 생성 중입니다..."):
                        new_sug = suggest_course_titles_with_gemini(
                            course_info=courses_for_ai[sel_c_idx],
                            api_key=gemini_key_input or None,
                            direction=direction_choice,
                        )
                        ai_df.at[sel_c_idx, "제안1_초단축직관형"] = new_sug["style_1_short"]
                        ai_df.at[sel_c_idx, "제안2_초단축시그니처"] = new_sug["style_2_brand"]
                        ai_df.at[sel_c_idx, "제안3_초단축킬러형"] = new_sug["style_3_punch"]
                        ai_df.at[sel_c_idx, "진단비평"] = new_sug["critique"]
                        st.session_state["ai_titles_df"] = ai_df
                        st.rerun()
                        st.rerun()

# =========================================================================
# Tab 4: 💾 누적 DB 백업/관리 & 제외 강좌 확인
# =========================================================================
with tab4:
    # -------------------------------------------------------------
    # 섹션 1: 판매중지로 제외된 강좌 전체 목록
    # -------------------------------------------------------------
    st.markdown(f"### 🚫 1. 판매중지로 제외된 강좌 목록 ({len(stopped_rms_df)}건)")
    if stopped_rms_df.empty:
        st.success("✅ 현재 RMS 엑셀 파일에 '판매중지'로 등록된 강좌가 없습니다.")
    else:
        st.caption("아래 강좌들은 RMS 상에서 '판매중지' 상태여서 검수 및 통계 집계에서 제외되었습니다:")
        st.dataframe(
            stopped_rms_df[
                [
                    "No",
                    "강좌코드",
                    "영역",
                    "과목",
                    "학년",
                    "강사",
                    "강좌명",
                    "간결수업시간",
                    "개강일_짧은표시",
                    "종강일_짧은표시",
                    "수강료",
                    "판매",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("---")
    # -------------------------------------------------------------
    # 섹션 2: 누적 DB 백업 및 복원 관리
    # -------------------------------------------------------------
    st.markdown("### 💾 2. 월별 누적 수강생 DB 백업 및 복원")
    col_bk_dn, col_bk_up = st.columns(2)

    with col_bk_dn:
        st.markdown("#### 📥 누적 DB 백업 파일 다운로드")
        st.caption("클라우드 재시작 시에도 데이터를 안전하게 보존하기 위해 누적 DB를 엑셀 백업본으로 저장하세요.")
        backup_excel = export_enrollment_backup_excel()
        st.download_button(
            label="📥 **누적 DB 백업 엑셀(.xlsx) 다운로드**",
            data=backup_excel,
            file_name=f"러셀대치_월별수강생누적DB_백업_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_bk_up:
        st.markdown("#### 📂 백업 파일로부터 복원하기")
        uploaded_backup = st.file_uploader(
            "백업 파일(.xlsx 또는 .json) 업로드:",
            type=["xlsx", "json"],
            key="uploader_backup_file",
        )
        if uploaded_backup:
            ok, msg, restored_cnt = import_enrollment_backup(uploaded_backup)
            if ok:
                st.success(f"🎉 {msg}")
            else:
                st.error(f"❌ {msg}")

    # 현재 DB에 저장된 월 목록 조회 및 삭제 기능
    st.markdown("---")
    st.markdown("#### 🗄️ 현재 등록된 월별 데이터 목록")
    db_months = get_available_months()
    if db_months:
        st.write(f"현재 등록된 연월: **{', '.join(db_months)}** (총 {len(db_months)}개 월)")
        col_del_m, col_del_btn = st.columns([3, 1])
        with col_del_m:
            del_m_target = st.selectbox("삭제할 월을 선택하세요:", options=db_months)
        with col_del_btn:
            st.write("")
            st.write("")
            if st.button("🗑️ 해당 월 데이터 삭제", use_container_width=True):
                if delete_month_data(del_m_target):
                    st.success(f"'{del_m_target}' 데이터가 삭제되었습니다.")
                    st.rerun()
    else:
        st.info("등록된 월별 누적 데이터가 없습니다.")

    st.markdown("---")
    # -------------------------------------------------------------
    # 섹션 3: 강사명 정규화 매핑 사전 설정
    # -------------------------------------------------------------
    st.markdown("### 🏷️ 3. 강사명 정규화 및 매핑 사전 관리")
    st.caption("PDF 상의 표기(예: '박소현')와 RMS 상의 등록명(예: '박소현(리아)')이 다른 경우 매핑 규칙을 추가하세요.")
    mappings = mapper.mappings
    st.dataframe(pd.DataFrame(mappings), use_container_width=True, hide_index=True)
