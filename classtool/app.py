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
import io
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
os.chdir(str(PROJECT_ROOT))

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import core.cumulative_db as _cum_db_mod
import core.exporter as _exporter_mod
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
importlib.reload(_exporter_mod)

from core.cumulative_db import (
    build_pivot_table,
    build_snapshot_pivot_table,
    compare_two_snapshots,
    delete_month_data,
    delete_snapshot_data,
    detect_snapshot_date_from_filename,
    export_enrollment_backup_excel,
    format_short_snapshot_header,
    format_snapshot_label,
    get_available_months,
    get_available_snapshots,
    get_monthly_summary_table,
    import_enrollment_backup,
    load_enrollment_db,
    normalize_snapshot_key,
    upsert_month_data,
    upsert_snapshot_data,
)
from core.exporter import (
    APPS_SCRIPT_SAMPLE_CODE,
    build_all_in_one_master_df,
    build_multi_section_report_tsv,
    export_all_in_one_csv_bytes,
    export_all_sheets_zip_bytes,
    render_google_sheets_copy_button,
    send_all_tabs_to_apps_script,
)
from core.matcher import (
    EXPORT_DISPLAY_COLUMNS,
    InstructorMapper,
    export_ai_titles_only_excel,
    export_crosscheck_to_csv_bytes,
    export_crosscheck_to_excel,
    export_crosscheck_to_tsv_string,
    export_results_to_google_sheets,
    generate_pdf_fix_request_messages,
    run_rms_pdf_crosscheck,
)
from core.text_improver import (
    analyze_course_text_content,
    batch_run_external_spelling_parallel,
)
from core.title_suggester import (
    DIRECTION_PROMPTS,
    batch_suggest_course_titles,
    get_gemini_api_key,
    suggest_course_titles_with_gemini,
)
from parsers.pdf_parser import compute_pdf_hash, parse_syllabus_pdf
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


@st.cache_data(show_spinner=False, max_entries=20)
def parse_pdf_cached(
    _file_bytes: bytes,
    file_hash: str,
    file_name: str,
    target_ym: str = "2027-01",
) -> pd.DataFrame:
    """
    대용량 PDF 파일 바이트(_file_bytes)의 Streamlit 자체 해싱 부하를 100% 제거(인자명 앞 '_')하고
    경량 문자열 file_hash만 해시 키로 활용하여 0.00초 캐시 조회를 보장합니다.
    """
    return parse_syllabus_pdf(_file_bytes, target_ym=target_ym)


@st.cache_data(show_spinner=False, max_entries=20)
def parse_all_pdfs_cached(
    _files_data: List[Tuple[str, bytes]],
    files_hash: str,
    target_ym: str = "2027-01",
) -> pd.DataFrame:
    """
    다중/단일 PDF 파일 바이트 리스트(_files_data)에 대해 해싱 오버헤드를 완전 제거하고
    초고속 pypdf 엔진으로 병합된 DataFrame을 캐싱하여 반환합니다.
    """
    parsed_dfs = []
    for fname, fbytes in _files_data:
        df_part = parse_syllabus_pdf(fbytes, target_ym=target_ym)
        if df_part is not None and not df_part.empty:
            parsed_dfs.append(df_part)
    if parsed_dfs:
        return pd.concat(parsed_dfs, ignore_index=True)
    return pd.DataFrame()


@st.cache_data(show_spinner=False, max_entries=20)
def parse_rms_cached(
    _file_bytes: bytes,
    file_hash: str,
    file_name: str,
) -> pd.DataFrame:
    """
    RMS 엑셀 파일 바이트(_file_bytes)의 Streamlit 자체 해싱을 방지하고
    경량 file_hash만 키로 활용하여 즉시 캐시 반환합니다.
    """
    return parse_rms_xls(io.BytesIO(_file_bytes))


@st.cache_data(show_spinner=False, max_entries=20)
def run_crosscheck_cached(
    _rms_df: pd.DataFrame,
    _syllabus_pdf_df: Optional[pd.DataFrame],
    _mapper: InstructorMapper,
    _stopped_rms_df: Optional[pd.DataFrame],
    rms_hash: str,
    pdf_hash: str,
    target_ym: str,
    title_sim_threshold: float,
    max_char_threshold: int,
) -> Dict[str, Any]:
    """
    RMS ↔ PDF 대조 검수를 캐싱하여, 탭 이동이나 필터 조작 시 0.00초 만에 즉시 반환합니다.
    대용량 DataFrame 인자들은 '_'로 시작하여 Streamlit의 불필요한 해싱을 생략합니다.
    (PDF 업로드 시 외부 네트워크 호출 100% 차단, 순수 로컬 규칙만 실행)
    """
    return run_rms_pdf_crosscheck(
        rms_df=_rms_df,
        syllabus_pdf_df=_syllabus_pdf_df,
        mapper=_mapper,
        target_ym=target_ym,
        title_similarity_threshold=title_sim_threshold,
        max_char_threshold=max_char_threshold,
        gemini_api_key=None,
        stopped_rms_df=_stopped_rms_df,
    )


@st.cache_data(show_spinner=False, max_entries=10)
def export_crosscheck_to_excel_cached(
    _crosscheck_result: Dict[str, Any],
    cache_key: str,
    target_ym: str,
    snap_a: Optional[str] = None,
    snap_b: Optional[str] = None,
) -> bytes:
    """
    10개 서식 시트를 포함하는 무거운 엑셀 생성을 캐싱하여
    사용자가 화면을 조작하거나 리런될 때 openpyxl 연산 지연(0.8~1.5초)을 0.00초로 완전 제거합니다.
    """
    return export_crosscheck_to_excel(
        _crosscheck_result,
        target_ym=target_ym,
        snapshot_a_key=snap_a,
        snapshot_b_key=snap_b,
    )


@st.cache_data(show_spinner=False, max_entries=10)
def build_all_in_one_master_df_cached(
    _crosscheck_result: Dict[str, Any],
    cache_key: str,
    target_ym: str,
) -> pd.DataFrame:
    return build_all_in_one_master_df(_crosscheck_result, target_ym=target_ym)


@st.cache_data(show_spinner=False, max_entries=10)
def build_multi_section_report_tsv_cached(
    _crosscheck_result: Dict[str, Any],
    cache_key: str,
    target_ym: str,
    snap_a: Optional[str] = None,
    snap_b: Optional[str] = None,
) -> str:
    return build_multi_section_report_tsv(
        _crosscheck_result,
        target_ym=target_ym,
        snapshot_a_key=snap_a,
        snapshot_b_key=snap_b,
    )


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
        <p>RMS 강좌목록 엑셀(.xls/.xlsx)과 강의계획서 PDF(.pdf) 7대 항목 정밀 검수 · 수강생/결제자수 누적 통계 · 15자 이내 초단축 AI 강의명 제안 <span style="background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 4px; font-size: 0.8rem; font-weight: 600; margin-left: 6px;">v2.2 최신배포본</span></p>
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
# 상단 컨트롤 바: [📅 스냅샷 기준일 설정] & [☑️ 판매중지 강좌 제외]
# -------------------------------------------------------------------------
col_snap_cfg, col_toggle = st.columns([1.5, 2.0])

# RMS 엑셀 파일명에서 기준 스냅샷 날짜 선감지
detected_default_snap = "2027-01-20"
if uploaded_rms_files:
    first_f = uploaded_rms_files[0]
    detected_default_snap = detect_snapshot_date_from_filename(first_f.name, fallback_ym=user_settings.get("target_ym", "2027-01"))

with col_snap_cfg:
    snapshot_date_input = st.text_input(
        "📅 **통계 기준일(스냅샷)**",
        value=st.session_state.get("custom_snapshot_date", detected_default_snap),
        help="RMS 수강생 누적 관리 기준일입니다. 기본값: 파일명 기반 매월 20일(YYYY-MM-20).",
        key="input_snapshot_date",
    )
    current_snapshot_key = normalize_snapshot_key(snapshot_date_input)
    st.session_state["custom_snapshot_date"] = current_snapshot_key
    st.caption(f"📌 **{format_snapshot_label(current_snapshot_key)}**")

with col_toggle:
    exclude_stopped = st.checkbox(
        "☑️ **판매중지 강좌 제외** (판매중인 강좌만 집계·검수)",
        value=True,
        help="체크 시 RMS 엑셀의 '판매' 컬럼이 '판매'인 유효 강좌만 집계하고, '판매중지' 강좌는 제외합니다. (사전 검수 시에는 체크를 해제하세요)",
        key="exclude_stopped_checkbox",
    )

# -------------------------------------------------------------------------
# RMS 파일 파싱 및 처리 (@st.cache_data 파일 바이트 캐싱 적용)
# -------------------------------------------------------------------------
raw_rms_df = None
selected_rms_ym = user_settings.get("target_ym", "2027-01")

if uploaded_rms_files:
    if len(uploaded_rms_files) == 1:
        f = uploaded_rms_files[0]
        try:
            f_bytes = f.getvalue()
            f_hash = f"{f.name}_{len(f_bytes)}"
            raw_rms_df = parse_rms_cached(_file_bytes=f_bytes, file_hash=f_hash, file_name=f.name)
            detected_ym = detect_dominant_year_month(raw_rms_df, default_ym=selected_rms_ym)
            selected_rms_ym = detected_ym
            st.session_state["rms_df_raw"] = raw_rms_df
            st.session_state["selected_rms_ym"] = selected_rms_ym
            st.session_state["rms_file_hash"] = f_hash
            # 매월 20일 기준 스냅샷 DB에 자동 등록/업데이트
            upsert_snapshot_data(raw_rms_df, current_snapshot_key, filter_stopped=exclude_stopped)
            st.session_state["active_snapshot_key"] = current_snapshot_key
        except Exception as e:
            st.error(f"❌ RMS 엑셀 파싱 중 오류가 발생했습니다: {e}")
    else:
        file_map: Dict[str, pd.DataFrame] = {}
        combined_hash_parts = []
        for f in uploaded_rms_files:
            try:
                f_bytes = f.getvalue()
                f_hash = f"{f.name}_{len(f_bytes)}"
                combined_hash_parts.append(f_hash)
                parsed = parse_rms_cached(_file_bytes=f_bytes, file_hash=f_hash, file_name=f.name)
                ym_found = detect_dominant_year_month(parsed, default_ym="2027-01")
                f_snap = detect_snapshot_date_from_filename(f.name, fallback_ym=ym_found)
                file_map[f"{f.name} ({format_snapshot_label(f_snap)})"] = (parsed, f_snap, ym_found)
                # 각 파일별 매월 20일 스냅샷 DB에 자동 등록
                upsert_snapshot_data(parsed, f_snap, filter_stopped=exclude_stopped)
            except Exception as e:
                st.warning(f"⚠️ '{f.name}' 파싱 실패: {e}")

        if file_map:
            st.success(f"🎉 **총 {len(file_map)}개 시점의 RMS 파일이 파싱되어 매월 20일 스냅샷 DB에 자동 분류·저장되었습니다.**")
            selected_name = st.selectbox(
                "현재 검수 및 대시보드 기준으로 볼 RMS 파일을 선택하세요:",
                options=list(file_map.keys()),
                index=0,
            )
            raw_rms_df, cur_snap, cur_ym = file_map[selected_name]
            st.session_state["rms_df_raw"] = raw_rms_df
            selected_rms_ym = cur_ym
            st.session_state["selected_rms_ym"] = selected_rms_ym
            st.session_state["active_snapshot_key"] = cur_snap
            st.session_state["rms_file_hash"] = "_".join(combined_hash_parts)
elif st.session_state.get("rms_df_raw") is not None:
    raw_rms_df = st.session_state["rms_df_raw"]
    selected_rms_ym = st.session_state.get("selected_rms_ym", selected_rms_ym)

# -------------------------------------------------------------------------
# PDF 파일 파싱 및 처리 (@st.cache_data 파일 바이트 캐싱 적용 - 0.00초 즉시 반환)
# -------------------------------------------------------------------------
syllabus_pdf_df = None
pdf_files_hash = "none"
if uploaded_pdf_files:
    try:
        pdf_files_data = [(f.name, f.getvalue()) for f in uploaded_pdf_files]
        pdf_files_hash = compute_pdf_hash(pdf_files_data, extra_tag=selected_rms_ym)
        syllabus_pdf_df = parse_all_pdfs_cached(_files_data=pdf_files_data, files_hash=pdf_files_hash, target_ym=selected_rms_ym)
        st.session_state["syllabus_pdf_df"] = syllabus_pdf_df
        st.session_state["pdf_files_hash"] = pdf_files_hash
    except Exception as e:
        st.error(f"❌ PDF 파싱 중 오류가 발생했습니다: {e}")
elif st.session_state.get("syllabus_pdf_df") is not None:
    syllabus_pdf_df = st.session_state["syllabus_pdf_df"]
    pdf_files_hash = st.session_state.get("pdf_files_hash", f"session_{len(syllabus_pdf_df)}_{selected_rms_ym}")

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
# 교차 검수 실행 (캐싱 적용: 탭 이동/필터 조작 시 0.00초 즉시 반환)
# -------------------------------------------------------------------------
crosscheck_result = None
if active_rms_df is not None:
    rms_hash_key = f"{st.session_state.get('rms_file_hash', 'raw')}_{exclude_stopped}_{len(active_rms_df)}_{len(stopped_rms_df)}"
    pdf_hash_key = f"{pdf_files_hash}_{selected_rms_ym}_{len(syllabus_pdf_df) if syllabus_pdf_df is not None else 0}"

    raw_cross = run_crosscheck_cached(
        _rms_df=active_rms_df,
        _syllabus_pdf_df=syllabus_pdf_df,
        _mapper=mapper,
        _stopped_rms_df=stopped_rms_df,
        rms_hash=rms_hash_key,
        pdf_hash=pdf_hash_key,
        target_ym=selected_rms_ym,
        title_sim_threshold=float(user_settings.get("title_sim_threshold", 55)),
        max_char_threshold=int(max_char_threshold),
    )
    crosscheck_result = dict(raw_cross)
    # AI 강의명 추천 데이터프레임이 세션에 있으면 병합
    if st.session_state.get("ai_titles_df") is not None:
        crosscheck_result["ai_titles_df"] = st.session_state["ai_titles_df"]
    # 세션에 커스텀 맞춤법 검사 결과가 있으면 병합
    if st.session_state.get("custom_spelling_df") is not None:
        crosscheck_result["spelling_df"] = st.session_state["custom_spelling_df"]
    st.session_state["crosscheck_result"] = crosscheck_result

# =========================================================================
# 상단 통합 결과지 다운로드 & 구글 시트(Google Sheets) 원스톱 내보내기 액션 바
# =========================================================================
if crosscheck_result is not None:
    st.markdown("#### 📥 2. 통합 결과지 다운로드 & 🌐 구글 시트(Google Sheets) 원스톱 내보내기")
    st.caption("기존 엑셀(.xlsx) 파일 다운로드뿐만 아니라, 모든 통계·검수·맞춤법·강의명 정보를 구글 시트로 즉시 전송하거나 복사할 수 있습니다.")

    snap_a = st.session_state.get("selected_snapshot_a", None)
    snap_b = st.session_state.get("selected_snapshot_b", None)
    ai_df_state = st.session_state.get("ai_titles_df")
    if ai_df_state is not None and not ai_df_state.empty:
        ai_hash = hashlib.md5(f"{len(ai_df_state)}_{ai_df_state.iloc[0].to_json()}".encode()).hexdigest()[:8]
    else:
        ai_hash = "none"

    sp_df_state = st.session_state.get("custom_spelling_df")
    if sp_df_state is not None and not sp_df_state.empty:
        sp_hash = hashlib.md5(f"{len(sp_df_state)}_{sp_df_state.iloc[0].to_json()}".encode()).hexdigest()[:8]
    else:
        sp_hash = "none"

    export_cache_key = f"{rms_hash_key}_{pdf_hash_key}_{selected_rms_ym}_{snap_a}_{snap_b}_{ai_hash}_{sp_hash}"

    # 1. 올인원 마스터 DataFrame 및 전체 TSV 준비 (캐싱 적용: 0.00초)
    all_in_one_master_df = build_all_in_one_master_df_cached(
        _crosscheck_result=crosscheck_result,
        cache_key=export_cache_key,
        target_ym=selected_rms_ym,
    )
    multi_sec_tsv = build_multi_section_report_tsv_cached(
        _crosscheck_result=crosscheck_result,
        cache_key=export_cache_key,
        target_ym=selected_rms_ym,
        snap_a=snap_a,
        snap_b=snap_b,
    )

    # 복사 대상 선택 라디오 (버튼 2와 연동)
    copy_target_opt = st.radio(
        "구글 시트 복사 대상 선택:",
        options=[
            "📋 전체 리포트 한 번에 복사 (Multi-Section TSV)",
            "📊 올인원 마스터 표만 복사 (강좌별 가로 완전 병합)",
            "📈 강사별 20일 비교 통계만 복사",
            "✍️ 맞춤법·문장축약 상세내역만 복사",
        ],
        horizontal=True,
        key="radio_copy_target_opt",
    )

    # 선택된 복사 텍스트 결정
    if copy_target_opt.startswith("📊"):
        selected_copy_tsv = all_in_one_master_df.to_csv(index=False, sep="\t")
    elif copy_target_opt.startswith("📈"):
        comp_snap_res = compare_two_snapshots(snap_a, snap_b)
        selected_copy_tsv = comp_snap_res.get("comparison_df", pd.DataFrame()).to_csv(index=False, sep="\t")
    elif copy_target_opt.startswith("✍️"):
        sp_df = crosscheck_result.get("spelling_df", pd.DataFrame())
        selected_copy_tsv = sp_df.to_csv(index=False, sep="\t")
    else:
        selected_copy_tsv = multi_sec_tsv

    # 4열 액션 바
    col_btn_excel, col_btn_js_copy, col_btn_csv, col_btn_tsv = st.columns([1.35, 1.85, 1.35, 1.05])

    # 버튼 1: 기존 엑셀 다운로드 (지연 연산 캐싱 적용으로 0.00초 즉시 제공)
    with col_btn_excel:
        excel_bytes = export_crosscheck_to_excel_cached(
            _crosscheck_result=crosscheck_result,
            cache_key=export_cache_key,
            target_ym=selected_rms_ym,
            snap_a=snap_a,
            snap_b=snap_b,
        )
        st.download_button(
            label="📥 **통합 결과지 엑셀 (.xlsx)**",
            data=excel_bytes,
            file_name=f"러셀대치_강의계획서_정밀검수결과지_{selected_rms_ym}_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Sheet 0(올인원)부터 Sheet 9까지 모든 서식이 포함된 엑셀입니다. 구글 시트에 그대로 업로드해도 완벽 호환됩니다.",
        )

    # 버튼 2: JavaScript 브라우저 클립보드 복사 & sheets.new 자동 오픈
    with col_btn_js_copy:
        render_google_sheets_copy_button(
            tsv_text=selected_copy_tsv,
            button_label="🚀 모든 정보 복사 & 새 구글 시트 열기",
            height=52,
        )

    # 버튼 3: 구글 시트용 CSV / ZIP 패키지 다운로드
    with col_btn_csv:
        with st.popover("📊 **구글시트 CSV/ZIP 다운로드**", use_container_width=True):
            st.markdown("##### 📊 구글 시트 전용 CSV 다운로드 (UTF-8-SIG)")
            st.caption("구글 시트 '파일 ➔ 가져오기'에서 한글 깨짐 없이 바로 열립니다.")

            csv_master_bytes = export_all_in_one_csv_bytes(all_in_one_master_df)
            st.download_button(
                label="📄 **올인원 마스터 CSV (.csv)**",
                data=csv_master_bytes,
                file_name=f"러셀대치_올인원_전체통합마스터_{selected_rms_ym}.csv",
                mime="text/csv",
                use_container_width=True,
            )

            zip_all_bytes = export_all_sheets_zip_bytes(
                crosscheck_result,
                target_ym=selected_rms_ym,
                snapshot_a_key=snap_a,
                snapshot_b_key=snap_b,
            )
            st.download_button(
                label="🗂️ **전체 시트 CSV 압축파일 (.zip)**",
                data=zip_all_bytes,
                file_name=f"러셀대치_전체시트_CSV_패키지_{selected_rms_ym}.zip",
                mime="application/zip",
                use_container_width=True,
            )

    # 버튼 4: 기존 TSV 수동 복사 팝오버 (호환성 유지)
    with col_btn_tsv:
        with st.popover("📋 **TSV 텍스트**", use_container_width=True):
            st.caption("텍스트를 직접 복사할 수 있습니다:")
            st.text_area("Multi-Section TSV", value=selected_copy_tsv, height=180)
            st.link_button("🌐 sheets.new (새 시트 열기)", "https://sheets.new")

    # 웹앱 연동: 내 구글 시트 URL에 모든 탭 직접 꽂아넣기 (Apps Script 웹앱 연동)
    with st.expander("☁️ 내 구글 시트 URL에 모든 탭 직접 꽂아넣기 (JSON 키 불필요)", expanded=False):
        st.markdown(
            """
            **💡 구글 시트 웹앱(Apps Script) 연동 안내:**
            1. 데이터를 전송받을 구글 스프레드시트를 엽니다.
            2. 상단 메뉴 **[확장 프로그램] ➔ [Apps Script]**를 클릭합니다.
            3. 아래 코드를 붙여넣고 상단 **[배포] ➔ [새 배포] ➔ 유형: [웹 앱]**을 선택합니다.
               - *액세스 권한이 있는 사용자:* **모든 사용자(Anyone)**로 설정 후 배포합니다.
            4. 생성된 **웹 앱 URL**을 아래 입력란에 넣고 **[🚀 내 구글 시트에 전송]**을 누르면 끝!
            """
        )
        st.code(APPS_SCRIPT_SAMPLE_CODE, language="javascript")

        col_gs_url, col_gs_btn = st.columns([3.5, 1.5])
        with col_gs_url:
            user_webhook_url = st.text_input(
                "구글 Apps Script 웹앱 URL (https://script.google.com/macros/s/...):",
                value=user_settings.get("export_sheet_url", ""),
                placeholder="https://script.google.com/macros/s/AKfycb.../exec",
                key="input_gs_webhook_url",
            )
        with col_gs_btn:
            st.write("")
            st.write("")
            btn_push_gs = st.button("🚀 **내 구글 시트에 전송**", type="primary", use_container_width=True)

        if btn_push_gs:
            if not user_webhook_url.strip():
                st.warning("⚠️ 구글 Apps Script 웹앱 URL을 먼저 입력해 주세요.")
            else:
                user_settings["export_sheet_url"] = user_webhook_url.strip()
                save_user_settings(user_settings)
                with st.spinner("☁️ 구글 시트로 올인원 마스터 및 전 탭 데이터를 전송 중입니다..."):
                    ok_push, msg_push = send_all_tabs_to_apps_script(
                        webhook_url=user_webhook_url.strip(),
                        crosscheck_result=crosscheck_result,
                        target_ym=selected_rms_ym,
                        snapshot_a_key=snap_a,
                        snapshot_b_key=snap_b,
                    )
                    if ok_push:
                        st.success(f"🎉 {msg_push}")
                    else:
                        st.error(f"❌ {msg_push}")

st.markdown("---")

# -------------------------------------------------------------------------
# 4대 핵심 탭 인터페이스 구성
# -------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📈 Tab 1: 강사별 수강생 & 매월 20일 기준 비교 통계",
        "🔍 Tab 2: 엑셀 vs PDF 강의계획서 정밀 검수",
        "✍️ Tab 3: 맞춤법 교정 & ⚡ 초단축 문장·강의명 제안",
        "💾 Tab 4: 누적 DB 백업/관리 & 제외 강좌 확인",
    ]
)

# =========================================================================
# Tab 1: 📈 강사별 수강생 & 매월 20일 기준 비교 통계
# =========================================================================
with tab1:
    avail_snaps = get_available_snapshots()

    if not avail_snaps:
        st.info("💡 상단 업로드 영역에서 **RMS 수업목록 엑셀 파일(.xls, .xlsx)**을 업로드하면 매월 20일 기준 스냅샷이 자동 등록되어 비교 분석할 수 있습니다.")
    else:
        st.markdown("### 📅 1. 스냅샷 비교 시점 선택 (Snapshot Selector)")
        st.caption("매월 20일 시점의 수강생 및 결제자수 데이터를 1:1로 대조하여 정확한 증감 통계를 분석합니다.")

        # 최근 2개 스냅샷 자동 선택 (2개 이상이면 이전 달과 당월)
        if len(avail_snaps) >= 2:
            def_idx_a = len(avail_snaps) - 2
            def_idx_b = len(avail_snaps) - 1
        else:
            def_idx_a = 0
            def_idx_b = 0

        col_snap_a, col_snap_b, col_btn_sync = st.columns([2, 2, 1.2])
        with col_snap_a:
            sel_snap_a = st.selectbox(
                "📅 **기준 시점 A (전월 20일)**",
                options=avail_snaps,
                index=def_idx_a,
                format_func=format_snapshot_label,
                key="tab1_sel_snap_a",
            )
        with col_snap_b:
            sel_snap_b = st.selectbox(
                "📅 **비교 시점 B (당월 20일)**",
                options=avail_snaps,
                index=def_idx_b,
                format_func=format_snapshot_label,
                key="tab1_sel_snap_b",
            )
        with col_btn_sync:
            st.write("")
            st.write("")
            if raw_rms_df is not None:
                if st.button("💾 **현재 RMS 스냅샷 갱신**", use_container_width=True):
                    cur_key = st.session_state.get("active_snapshot_key", normalize_snapshot_key(selected_rms_ym))
                    ok, msg = upsert_snapshot_data(raw_rms_df, cur_key, filter_stopped=exclude_stopped)
                    if ok:
                        st.toast(msg, icon="✅")
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

        st.session_state["selected_snapshot_a"] = sel_snap_a
        st.session_state["selected_snapshot_b"] = sel_snap_b

        # 1:1 비교 데이터 산출
        comp_res = compare_two_snapshots(sel_snap_a, sel_snap_b)
        kpi = comp_res.get("kpi", {})
        comp_df = comp_res.get("comparison_df", pd.DataFrame())
        short_a = comp_res.get("short_a", "전월")
        short_b = comp_res.get("short_b", "당월")

        # 4대 핵심 KPI 카드
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            crs_diff = kpi.get("courses_diff", 0)
            st.metric(
                label=f"유효 강좌수 변화 ({short_b})",
                value=f"{kpi.get('courses_b', 0)}건",
                delta=f"{'+' if crs_diff > 0 else ''}{crs_diff}건 ({short_a} {kpi.get('courses_a', 0)}건)",
            )
        with c2:
            stu_diff = kpi.get("students_diff", 0)
            stu_pct = kpi.get("students_pct", 0.0)
            st.metric(
                label=f"총 수강생수 변화 ({short_b})",
                value=f"{kpi.get('students_b', 0):,}명",
                delta=f"{'+' if stu_diff > 0 else ''}{stu_diff:,}명 ({'+' if stu_pct > 0 else ''}{stu_pct}%)",
            )
        with c3:
            paid_diff = kpi.get("paid_diff", 0)
            paid_pct = kpi.get("paid_pct", 0.0)
            st.metric(
                label=f"결제완료 인원 변화 ({short_b})",
                value=f"{kpi.get('paid_b', 0):,}명",
                delta=f"{'+' if paid_diff > 0 else ''}{paid_diff:,}명 ({'+' if paid_pct > 0 else ''}{paid_pct}%)",
            )
        with c4:
            rate_diff = kpi.get("rate_diff", 0.0)
            st.metric(
                label=f"평균 결제율 변화 ({short_b})",
                value=f"{kpi.get('rate_b', 0.0)}%",
                delta=f"{'+' if rate_diff > 0 else ''}{rate_diff}%p ({short_a} {kpi.get('rate_a', 0.0)}%)",
            )

        st.markdown("---")
        # -------------------------------------------------------------
        # 2. 강사별 [A vs B] 1:1 정밀 비교 분석표 (Full Outer Join)
        # -------------------------------------------------------------
        st.markdown(f"### 👥 2. 강사별 [{short_a} vs {short_b}] 1:1 정밀 비교 분석표 (Full Outer Join)")
        st.caption("• 신규개설(`🆕`) 강좌 및 종강(`⚪`) 강좌까지 누락 없이 자동 분류·대조됩니다.")

        if not comp_df.empty:
            # 4열 필터 컨트롤 바
            f_col1, f_col2, f_col3, f_col4 = st.columns([1.5, 1.5, 2.0, 2.0])
            all_subjects = sorted(list(comp_df["과목"].dropna().unique()))
            with f_col1:
                subj_choice = st.selectbox("과목(영역) 필터:", options=["전체"] + all_subjects, key="tab1_filter_subj")
            with f_col2:
                status_choice = st.selectbox(
                    "변동 상태 필터:",
                    options=["전체", "🔺 증가", "🔻 감소", "➖ 유지", "🆕 당월 신규개설", "⚪ 당월 미개설(종강)"],
                    key="tab1_filter_status",
                )
            with f_col3:
                search_kw = st.text_input("🔍 강사명 검색:", placeholder="강사명(예: 이지원, 지인해)", key="tab1_search_inst")
            with f_col4:
                sort_choice = st.selectbox(
                    "정렬 기준:",
                    options=[
                        f"[B] {short_b} 결제자수 많은 순",
                        "결제자 증감(명) 많은 순",
                        f"[B] {short_b} 총수강생 많은 순",
                        "강좌수 많은 순",
                        "강사명 가나다순",
                    ],
                    key="tab1_sort_choice",
                )

            filtered_df = comp_df.copy()
            if subj_choice != "전체":
                filtered_df = filtered_df[filtered_df["과목"] == subj_choice]
            if status_choice != "전체":
                filtered_df = filtered_df[filtered_df["변동 상태"] == status_choice]
            if search_kw.strip():
                filtered_df = filtered_df[
                    filtered_df["강사명"].str.contains(search_kw.strip(), case=False, na=False)
                    | filtered_df["과목(영역)"].str.contains(search_kw.strip(), case=False, na=False)
                ]

            if sort_choice == f"[B] {short_b} 결제자수 많은 순":
                filtered_df = filtered_df.sort_values(by=f"[B] {short_b} 결제자수", ascending=False)
            elif sort_choice == "결제자 증감(명) 많은 순":
                filtered_df = filtered_df.sort_values(by="결제자_증감_정수", ascending=False)
            elif sort_choice == f"[B] {short_b} 총수강생 많은 순":
                filtered_df = filtered_df.sort_values(by=f"[B] {short_b} 총수강생", ascending=False)
            elif sort_choice == "강좌수 많은 순":
                filtered_df = filtered_df.sort_values(by=f"[B] {short_b} 강좌수", ascending=False)
            elif sort_choice == "강사명 가나다순":
                filtered_df = filtered_df.sort_values(by="강사명", ascending=True)

            # 순위 재부여
            filtered_df = filtered_df.reset_index(drop=True)
            filtered_df["순위"] = filtered_df.index + 1

            disp_cols = [
                "순위",
                "과목(영역)",
                "강사명",
                f"[A] {short_a} 강좌수",
                f"[B] {short_b} 강좌수",
                "강좌수 증감",
                f"[A] {short_a} 결제자수",
                f"[B] {short_b} 결제자수",
                "결제자 증감(명)",
                "결제자 증감률(%)",
                f"[A] {short_a} 총수강생",
                f"[B] {short_b} 총수강생",
                "총수강생 증감(명)",
                "[A] 결제율(%)",
                "[B] 결제율(%)",
                "결제율 증감(%p)",
                "변동 상태",
            ]
            valid_disp_cols = [c for c in disp_cols if c in filtered_df.columns]

            st.dataframe(
                filtered_df[valid_disp_cols],
                use_container_width=True,
                column_config={
                    "순위": st.column_config.NumberColumn("순위", format="%d위"),
                    f"[A] {short_a} 결제자수": st.column_config.NumberColumn(f"[A] {short_a} 결제", format="%d명"),
                    f"[B] {short_b} 결제자수": st.column_config.NumberColumn(f"[B] {short_b} 결제", format="%d명"),
                    f"[A] {short_a} 총수강생": st.column_config.NumberColumn(f"[A] {short_a} 수강생", format="%d명"),
                    f"[B] {short_b} 총수강생": st.column_config.NumberColumn(f"[B] {short_b} 수강생", format="%d명"),
                },
                hide_index=True,
            )
        else:
            st.info("비교할 강사 데이터가 없습니다.")

        st.markdown("---")
        # -------------------------------------------------------------
        # 3. 매월 20일 기준 연간 시계열 누적 피벗 테이블
        # -------------------------------------------------------------
        st.markdown("### 📊 3. 매월 20일 기준 연간 시계열 누적 피벗 테이블")
        col_metric, _ = st.columns([3, 2])
        with col_metric:
            selected_metric = st.radio(
                "조회할 누적 통계 지표 선택:",
                options=["결제완료 인원", "총 수강생(가등록포함)", "유효 강좌수", "결제 매출액"],
                horizontal=True,
                key="snapshot_pivot_metric_radio",
            )

        pivot_df = build_snapshot_pivot_table(metric_label=selected_metric)
        if pivot_df.empty:
            st.info("💡 아직 저장된 매월 20일 기준 누적 데이터가 없습니다.")
        else:
            st.dataframe(
                pivot_df,
                use_container_width=True,
                column_config={
                    "순위": st.column_config.NumberColumn("순위", format="%d위"),
                    "기간 누적 합계": st.column_config.NumberColumn("누적 합계", format="%d"),
                    "매월 20일 평균": st.column_config.NumberColumn("월평균", format="%.1f"),
                },
                hide_index=True,
            )

        st.markdown("---")
        # -------------------------------------------------------------
        # 4. 강사별 개설 강좌 상세 드릴다운
        # -------------------------------------------------------------
        st.markdown("### 🔎 4. 강사별 개설 강좌 상세 드릴다운")
        if active_rms_df is not None:
            all_instructors = sorted(list(active_rms_df["강사"].dropna().unique()))
            selected_drill_inst = st.selectbox("상세 조회할 강사를 선택하세요:", options=all_instructors, key="tab1_drill_inst")
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
            # 세션에 커스텀 맞춤법 검사 결과가 있으면 우선 사용
            if "custom_spelling_df" in st.session_state:
                spelling_df = st.session_state["custom_spelling_df"]
            else:
                spelling_df = crosscheck_result.get("spelling_df", pd.DataFrame())

            long_sentence_df = crosscheck_result.get("long_sentence_df", pd.DataFrame())

            col_sp_head, col_sp_btn = st.columns([3, 2])
            with col_sp_head:
                st.markdown(f"#### 🔍 1. 국어 맞춤법 및 오탈자 검수 ({len(spelling_df)}건 발견)")
            with col_sp_btn:
                btn_run_ext_spell = st.button("⚡ **외부 정밀 맞춤법 추가 스캔 (10개 병렬 스레드)**", type="primary", use_container_width=True)

            if btn_run_ext_spell:
                with st.spinner("⚡ 10개 병렬 스레드로 외부 정밀 맞춤법을 고속 스캔 중입니다..."):
                    courses_for_ext = syllabus_pdf_df.to_dict(orient="records")
                    ext_hits = batch_run_external_spelling_parallel(
                        pdf_courses=courses_for_ext,
                        gemini_api_key=gemini_key_input or None,
                        max_workers=10,
                    )
                    if ext_hits:
                        ext_df = pd.DataFrame(ext_hits)
                        merged_spell = pd.concat([spelling_df, ext_df], ignore_index=True)
                        if "원문문장" in merged_spell.columns and "대상어구" in merged_spell.columns:
                            merged_spell = merged_spell.drop_duplicates(subset=["강좌명", "원문문장", "대상어구"])
                        st.session_state["custom_spelling_df"] = merged_spell
                        if crosscheck_result is not None:
                            crosscheck_result["spelling_df"] = merged_spell
                        st.success(f"🎉 총 {len(ext_hits)}건의 외부 정밀 맞춤법 교정안이 10개 병렬 스레드로 신속하게 스캔 완료되었습니다.")
                        st.rerun()
                    else:
                        st.info("ℹ️ 외부 정밀 맞춤법 스캔 결과 추가 오류가 발견되지 않았습니다. (내장 규칙 120종 결과 유지)")

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
                st.success(f"🎉 **총 {len(ai_results)}개 강좌의 임팩트 분석 및 초단축 강의명 생성이 완료되었습니다.**")
                st.rerun()

            # 결과 데이터프레임 표시
            ai_df = st.session_state.get("ai_titles_df")
            if ai_df is not None and not ai_df.empty:
                # 1. 상단 요약 배지 및 필터 제공
                rev_mask = ai_df["needs_revision"] == True if "needs_revision" in ai_df.columns else ai_df["판정결과"].str.contains("수정")
                rev_cnt = int(rev_mask.sum())
                keep_cnt = len(ai_df) - rev_cnt

                st.markdown(
                    f"""
                    <div style="background: #f0fdf4; border: 1.5px solid #22c55e; border-radius: 8px; padding: 0.75rem 1.25rem; margin-bottom: 0.8rem; display: flex; align-items: center; justify-content: space-between;">
                        <span style="font-size: 1.05rem; font-weight: 700; color: #15803d;">
                            📊 전체 {len(ai_df)}개 강좌 분석 완료 — 🔧 수정 추천 : {rev_cnt}건 &nbsp;|&nbsp; ✅ 기존 유지 권장 : {keep_cnt}건
                        </span>
                        <span style="font-size: 0.88rem; color: #166534; font-weight: 600;">
                            Gemini 3.8 고유 시그니처 자동 보존 로직 가동
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # AI 추천 강의명 전용 다운로드 액션 바
                col_ai_down1, col_ai_down2, col_ai_down3 = st.columns([1.5, 1.8, 1.2])
                with col_ai_down1:
                    ai_only_bytes = export_ai_titles_only_excel(ai_df)
                    st.download_button(
                        label="📥 **AI 추천 강의명 전용 엑셀 (.xlsx)**",
                        data=ai_only_bytes,
                        file_name=f"러셀대치_AI추천강의명_리브랜딩_{selected_rms_ym}_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                        key="btn_down_ai_only_excel",
                        help="임팩트 평가 및 15자 이내 초단축 강의명 제안 결과만 깔끔하게 서식화된 전용 엑셀입니다.",
                    )
                with col_ai_down2:
                    st.download_button(
                        label="📥 **전 항목 통합 결과지 엑셀 (.xlsx)**",
                        data=excel_bytes,
                        file_name=f"러셀대치_강의계획서_종합검수결과_{selected_rms_ym}_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_down_full_excel_tab3",
                        help="Sheet 0(올인원)부터 Sheet 7(AI추천강의명)까지 모든 시트가 수록된 통합 마스터 엑셀입니다.",
                    )
                with col_ai_down3:
                    ai_tsv_str = ai_df.to_csv(index=False, sep="\t")
                    render_google_sheets_copy_button(
                        tsv_text=ai_tsv_str,
                        button_label="📋 표 전체 복사 (시트 붙여넣기)",
                        height=38,
                    )

                # 라디오 필터: [🔧 수정 추천 강좌만 보기 (기본값)] vs [전체 강좌 보기 (기존 유지 포함)]
                filter_ai_view = st.radio(
                    "표시 필터:",
                    options=["🔧 수정 추천 강좌만 보기 (기본값)", "전체 강좌 보기 (기존 유지 포함)"],
                    horizontal=True,
                    key="radio_filter_ai_view",
                )

                if filter_ai_view.startswith("🔧"):
                    disp_ai_df = ai_df[rev_mask].copy()
                    st.caption(f"💡 평이하거나 서술적이어서 15자 이내 초단축 리브랜딩이 권장되는 **{len(disp_ai_df)}건의 강좌**만 표시 중입니다.")
                else:
                    disp_ai_df = ai_df.copy()
                    st.caption(f"💡 기존 유지 권장 강좌를 포함한 **전체 {len(disp_ai_df)}건의 강좌**를 표시 중입니다.")

                table_cols = [
                    "순번",
                    "판정결과",
                    "과목",
                    "학년",
                    "강사명",
                    "현재강의명",
                    "임팩트점수",
                    "판정사유",
                    "⚡ 추천 강의명 1 (직관·시그니처형)",
                    "🔥 추천 강의명 2 (킬러·임팩트형)",
                    "생성엔진",
                ]
                valid_tbl_cols = [c for c in table_cols if c in disp_ai_df.columns]

                st.dataframe(
                    disp_ai_df[valid_tbl_cols],
                    use_container_width=True,
                    column_config={
                        "임팩트점수": st.column_config.NumberColumn("임팩트점수", format="%d점"),
                        "순번": st.column_config.NumberColumn("순번", format="%d"),
                    },
                    hide_index=True,
                )

                st.markdown("---")
                st.markdown("#### 🎴 개별 강좌 상세 카드 및 실시간 재추천")
                sel_c_idx = st.selectbox(
                    "상세 카드 및 재추천을 확인할 강좌를 선택하세요:",
                    options=range(len(ai_df)),
                    format_func=lambda i: f"[{ai_df.iloc[i].get('판정결과', '')}] [{ai_df.iloc[i]['과목']} {ai_df.iloc[i]['강사명']}T] {ai_df.iloc[i].get('현재강의명', ai_df.iloc[i].get('현재강좌명', ''))}",
                )
                c_data = ai_df.iloc[sel_c_idx].to_dict()
                is_keep = not bool(c_data.get("needs_revision", False)) or ("기존 유지" in str(c_data.get("판정결과", "")))

                if is_keep:
                    st.success(
                        f"✅ **[기존 강의명 유지 권장]** 현재 강의명이 충분히 임팩트 있거나 고유 시그니처 브랜드가 확립되어 있어 수정을 추천하지 않습니다.\n\n"
                        f"• **현재 강의명:** `{c_data.get('현재강의명', c_data.get('현재강좌명', ''))}`\n"
                        f"• **판정 사유:** {c_data.get('판정사유', c_data.get('진단비평', ''))}"
                    )
                else:
                    col_c1, col_c2 = st.columns(2)
                    r1_title = c_data.get("⚡ 추천 강의명 1 (직관·시그니처형)") or c_data.get("제안1_초단축직관형", "")
                    r2_title = c_data.get("🔥 추천 강의명 2 (킬러·임팩트형)") or c_data.get("제안2_초단축시그니처", "")

                    with col_c1:
                        st.markdown(
                            f"""
                            <div class="proposal-card">
                                <div class="proposal-title">⚡ 추천 1: 직관·시그니처형 (12자 내외)</div>
                                <h4 style="color: #1e3a8a; margin: 0.4rem 0;">{r1_title}</h4>
                                <p style="font-size: 0.85rem; color: #475569;">글자수: {len(r1_title)}자 | 목표와 시그니처 0.5초 직관 각인</p>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                    with col_c2:
                        st.markdown(
                            f"""
                            <div class="proposal-card">
                                <div class="proposal-title">🔥 추천 2: 킬러·임팩트형 (12자 내외)</div>
                                <h4 style="color: #991b1b; margin: 0.4rem 0;">{r2_title}</h4>
                                <p style="font-size: 0.85rem; color: #475569;">글자수: {len(r2_title)}자 | 1등급 킬러 정복·실전 무기 공략</p>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

                    st.info(f"💡 **판정 사유:** {c_data.get('판정사유', c_data.get('진단비평', ''))}")

                if st.button(f"🔄 [{c_data['강사명']}T {c_data.get('현재강의명', c_data.get('현재강좌명', ''))}] 이 강좌만 다시 분석/추천받기"):
                    with st.spinner("AI가 강좌명을 다시 분석 중입니다..."):
                        new_sug = suggest_course_titles_with_gemini(
                            course_info=courses_for_ai[sel_c_idx],
                            api_key=gemini_key_input or None,
                            direction=direction_choice,
                        )
                        ai_df.at[sel_c_idx, "판정결과"] = new_sug["status_badge"]
                        ai_df.at[sel_c_idx, "needs_revision"] = new_sug["needs_revision"]
                        ai_df.at[sel_c_idx, "임팩트점수"] = new_sug["impact_score"]
                        ai_df.at[sel_c_idx, "판정사유"] = new_sug["evaluation_reason"]
                        ai_df.at[sel_c_idx, "진단비평"] = new_sug["evaluation_reason"]
                        ai_df.at[sel_c_idx, "⚡ 추천 강의명 1 (직관·시그니처형)"] = new_sug["recommended_title_1"]
                        ai_df.at[sel_c_idx, "🔥 추천 강의명 2 (킬러·임팩트형)"] = new_sug["recommended_title_2"]
                        ai_df.at[sel_c_idx, "제안1_초단축직관형"] = new_sug["recommended_title_1"]
                        ai_df.at[sel_c_idx, "제안2_초단축시그니처"] = new_sug["recommended_title_2"]
                        st.session_state["ai_titles_df"] = ai_df
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

    # 현재 DB에 저장된 매월 20일 스냅샷 목록 조회 및 삭제 기능
    st.markdown("---")
    st.markdown("#### 🗄️ 현재 등록된 매월 20일 기준 스냅샷 목록")
    db_snaps = get_available_snapshots()
    if db_snaps:
        disp_snap_labels = [f"{format_snapshot_label(k)} ({k})" for k in db_snaps]
        st.write(f"현재 등록된 스냅샷: **{', '.join(disp_snap_labels)}** (총 {len(db_snaps)}개 시점)")
        col_del_m, col_del_btn = st.columns([3, 1])
        with col_del_m:
            del_target_idx = st.selectbox(
                "삭제할 스냅샷을 선택하세요:",
                options=range(len(db_snaps)),
                format_func=lambda i: f"{format_snapshot_label(db_snaps[i])} [{db_snaps[i]}]",
                key="tab4_del_snap_target",
            )
        with col_del_btn:
            st.write("")
            st.write("")
            if st.button("🗑️ 해당 스냅샷 데이터 삭제", use_container_width=True):
                target_to_del = db_snaps[del_target_idx]
                if delete_snapshot_data(target_to_del):
                    st.success(f"'{format_snapshot_label(target_to_del)}' 스냅샷 데이터가 삭제되었습니다.")
                    st.rerun()
    else:
        st.info("등록된 매월 20일 기준 스냅샷 데이터가 없습니다.")

    st.markdown("---")
    # -------------------------------------------------------------
    # 섹션 3: 강사명 정규화 매핑 사전 설정
    # -------------------------------------------------------------
    st.markdown("### 🏷️ 3. 강사명 정규화 및 매핑 사전 관리")
    st.caption("PDF 상의 표기(예: '박소현')와 RMS 상의 등록명(예: '박소현(리아)')이 다른 경우 매핑 규칙을 추가하세요.")
    mappings = mapper.mappings
    st.dataframe(pd.DataFrame(mappings), use_container_width=True, hide_index=True)
