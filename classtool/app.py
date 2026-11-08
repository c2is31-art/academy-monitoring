"""
메가스터디 러셀대치학원 단과 강좌 [RMS 수업목록(.xls/.xlsx) ↔ 강의계획서 PDF] 종합 검수 시스템
1. RMS 강의정보 업로드 시 '총 강좌수 / 총 수강생(결제완료) / 강사별 수강생 현황' 즉시 확인 대시보드
2. RMS ↔ PDF 7대 항목 정밀 대조 (수강료, 회차, 개강일/종강일, 진도표 요일오타/혼입, 수업시간, 바자관시수, 강좌명)
3. 국어 맞춤법·오탈자 검수 및 긴 문장(기본 55자) 간략화(서술형/개조식) 제안 모듈
4. Gemini 2.5 Flash 기반 대치동 단과 '강의명 & 슬로건 임팩트 수정 제안(AI 카피라이팅)' 모듈
5. 7개 시트 서식 포함 엑셀(.xlsx), 구글 시트 호환 CSV/TSV, sheets.new 원클릭 연동 및 결과지 다운로드
"""

from __future__ import annotations

import importlib
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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
importlib.reload(_matcher_mod)

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
        font-size: 1.45rem;
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
        margin-bottom: 0.9rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .rms-headline-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #15803d;
    }
    .error-brief-card {
        background-color: #fff5f5;
        border-left: 5px solid #e11d48;
        border-top: 1px solid #fecdd3;
        border-right: 1px solid #fecdd3;
        border-bottom: 1px solid #fecdd3;
        padding: 0.9rem 1.1rem;
        border-radius: 8px;
        margin-bottom: 0.7rem;
    }
    .error-brief-title {
        font-weight: 700;
        font-size: 1.02rem;
        color: #881337;
        margin-bottom: 0.35rem;
    }
    .error-brief-body {
        font-size: 0.9rem;
        color: #1f2937;
        line-height: 1.5;
    }
    .error-fix-box {
        background-color: #fef2f2;
        border: 1px dashed #f87171;
        padding: 0.45rem 0.75rem;
        border-radius: 6px;
        margin-top: 0.4rem;
        font-weight: 600;
        color: #b91c1c;
        font-size: 0.9rem;
    }
    .proposal-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 0.85rem;
    }
    .proposal-title {
        font-weight: 700;
        font-size: 0.98rem;
        color: #1e293b;
        margin-bottom: 0.3rem;
    }
    .export-banner {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        padding: 0.85rem 1rem;
        border-radius: 8px;
        margin-bottom: 0.85rem;
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


user_settings = load_user_settings()
initial_key = user_settings.get("gemini_api_key") or get_gemini_api_key() or ""

if "target_ym_val" not in st.session_state:
    st.session_state["target_ym_val"] = str(user_settings.get("target_ym", "2027-01"))
if "title_sim_threshold_val" not in st.session_state:
    st.session_state["title_sim_threshold_val"] = int(user_settings.get("title_sim_threshold", 55))
if "max_char_threshold_val" not in st.session_state:
    st.session_state["max_char_threshold_val"] = int(user_settings.get("max_char_threshold", 55))
if "gemini_api_key_val" not in st.session_state:
    st.session_state["gemini_api_key_val"] = str(initial_key)
if "export_sheet_url_val" not in st.session_state:
    st.session_state["export_sheet_url_val"] = str(
        user_settings.get("export_sheet_url", user_settings.get("sheet_url", ""))
    )
if "ai_titles_df" not in st.session_state:
    st.session_state["ai_titles_df"] = pd.DataFrame()

mapper = InstructorMapper(config_path=CONFIG_PATH)


def render_clipboard_and_sheets_new_widget(
    tsv_text: str, button_id: str, label_text: str = "📋 구글 시트 붙여넣기용 원클릭 복사"
) -> None:
    """클릭 한 번으로 전체 검수 결과 표(TSV)를 클립보드에 복사하고 sheets.new 바로가기 링크를 제공합니다."""
    escaped_tsv = json.dumps(tsv_text, ensure_ascii=False)
    html_code = f"""
    <div style="display:flex; gap:8px; align-items:center; font-family:'Malgun Gothic',sans-serif; margin:0; padding:0;">
        <button id="copy_btn_{button_id}" onclick="copyTsv_{button_id}()"
            style="flex:1; background:#2563eb; color:#ffffff; border:none; border-radius:6px; padding:9px 12px;
                   font-size:13.5px; font-weight:600; cursor:pointer; transition:background 0.15s;">
            {label_text}
        </button>
        <a href="https://sheets.new" target="_blank" rel="noopener noreferrer"
            style="background:#16a34a; color:#ffffff; text-decoration:none; border-radius:6px; padding:9px 13px;
                   font-size:13.5px; font-weight:600; white-space:nowrap; display:inline-block;">
            🚀 새 구글 시트 열기 (sheets.new)
        </a>
    </div>
    <div id="copy_msg_{button_id}" style="font-family:'Malgun Gothic',sans-serif; font-size:12px; color:#15803d; font-weight:600; margin-top:4px; display:none;">
        ✅ 클립보드에 복사되었습니다! 새 시트 A1 셀에서 <b>Ctrl+V</b>를 누르세요.
    </div>
    <script>
    function copyTsv_{button_id}() {{
        const text = {escaped_tsv};
        const btn = document.getElementById("copy_btn_{button_id}");
        const msg = document.getElementById("copy_msg_{button_id}");
        function onSuccess() {{
            btn.style.background = "#1d4ed8";
            btn.innerText = "✅ 클립보드 복사 완료! (Ctrl+V로 붙여넣기)";
            msg.style.display = "block";
            setTimeout(() => {{
                btn.style.background = "#2563eb";
                btn.innerText = "{label_text}";
            }}, 4000);
        }}
        if (navigator.clipboard && navigator.clipboard.writeText) {{
            navigator.clipboard.writeText(text).then(onSuccess).catch(() => fallbackCopy(text));
        }} else {{
            fallbackCopy(text);
        }}
        function fallbackCopy(str) {{
            const ta = document.createElement("textarea");
            ta.value = str;
            ta.style.position = "fixed";
            ta.style.left = "-9999px";
            document.body.appendChild(ta);
            ta.select();
            try {{
                document.execCommand("copy");
                onSuccess();
            }} catch (e) {{
                alert("클립보드 복사 실패: " + e);
            }}
            document.body.removeChild(ta);
        }}
    }}
    </script>
    """
    components.html(html_code, height=64)


def render_export_toolbar(
    crosscheck: Dict[str, Any],
    excel_bytes: bytes,
    csv_bytes: bytes,
    tsv_str: str,
    target_ym: str,
    section_key: str,
) -> None:
    """결과지 내보내기 툴바"""
    comparison_df: pd.DataFrame = crosscheck.get("comparison_df", pd.DataFrame())
    if comparison_df.empty:
        return

    now_tag = datetime.now().strftime("%Y%m%d_%H%M")
    ym_clean = target_ym.replace("-", "")[-4:] if target_ym else "2701"

    st.markdown('<div class="export-banner">', unsafe_allow_html=True)
    st.markdown("##### 📥 검수 결과지 즉시 다운로드 & 구글 시트 내보내기")

    col_ex1, col_ex2, col_ex3 = st.columns([1.15, 1.0, 1.65])
    with col_ex1:
        st.download_button(
            label="📥 엑셀 결과지 다운로드 (.xlsx - 서식·7개 시트)",
            data=excel_bytes,
            file_name=f"러셀대치_강의계획서_종합검수결과_{target_ym}_{now_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
            key=f"dl_xlsx_{section_key}",
            help="전체검수결과, 강사별_수강생현황, 데이터오류_수정요청, 맞춤법_및_긴문장축약, AI_강의명_수정제안, PDF누락강좌, PDF진도표상세검수 시트 포함",
        )
    with col_ex2:
        st.download_button(
            label="📊 구글 시트 호환 CSV 다운로드 (UTF-8-SIG)",
            data=csv_bytes,
            file_name=f"러셀대치_강의계획서_종합검수결과_{target_ym}_{now_tag}.csv",
            mime="text/csv",
            width="stretch",
            key=f"dl_csv_{section_key}",
            help="구글 드라이브 업로드 또는 엑셀에서 바로 열었을 때 한글이 깨지지 않는 UTF-8-SIG CSV 파일",
        )
    with col_ex3:
        render_clipboard_and_sheets_new_widget(
            tsv_text=tsv_str,
            button_id=f"top_{section_key}",
            label_text="📋 구글 시트 붙여넣기용 원클릭 복사",
        )

    with st.expander("📤 구글 시트 URL로 직접 내보내기 (선택 옵션)"):
        g_col1, g_col2, g_col3 = st.columns([2.4, 1.0, 1.0])
        with g_col1:
            input_gs_url = st.text_input(
                "결과를 업로드할 Google Sheets URL",
                value=st.session_state.get("export_sheet_url_val", ""),
                placeholder="https://docs.google.com/spreadsheets/d/.../edit",
                key=f"gs_url_input_{section_key}",
            )
        with g_col2:
            default_tab_name = f"검수결과_{ym_clean}"
            input_tab_name = st.text_input(
                "업로드할 시트 탭 이름",
                value=default_tab_name,
                key=f"gs_tab_input_{section_key}",
            )
        with g_col3:
            st.write("")
            if st.button("📤 구글 시트로 업로드", width="stretch", key=f"gs_up_btn_{section_key}"):
                if not input_gs_url.strip():
                    st.warning("업로드할 구글 시트 URL을 입력해 주세요.")
                else:
                    st.session_state["export_sheet_url_val"] = input_gs_url.strip()
                    user_settings["export_sheet_url"] = input_gs_url.strip()
                    save_user_settings(user_settings)
                    with st.spinner(f"구글 시트에 [{input_tab_name}] 탭으로 검수 결과를 업로드하는 중..."):
                        try:
                            up_res = export_results_to_google_sheets(
                                sheet_url=input_gs_url.strip(),
                                crosscheck_result=crosscheck,
                                target_ym=target_ym,
                                tab_name=input_tab_name.strip(),
                            )
                            st.success(
                                f"🎉 구글 시트({up_res['spreadsheet_title']}) 업로드 완료! "
                                + " / ".join(up_res["logs"])
                            )
                        except Exception as e:
                            st.error(f"구글 시트 직접 내보내기 오류: {e}")
    st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------------------------
# 상단 메인 헤더
# -------------------------------------------------------------------------
st.markdown(
    """
    <div class="main-header">
        <h2>📋 메가스터디 러셀대치 단과 강의계획서 종합 검수 시스템</h2>
        <p>RMS 수업목록 엑셀 ↔ 강의계획서 PDF 7대 정밀 대조 · 수강생/결제완료 실시간 집계 · 국어 맞춤법/긴문장 교정 · Gemini AI 강의명 리브랜딩</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------------
# 사이드바 (Sidebar): 4단계 업로더 & 설정
# -------------------------------------------------------------------------
with st.sidebar:
    st.subheader("Step 1. 📋 RMS 수업목록 엑셀 업로드")
    rms_uploaded_file = st.file_uploader(
        "RMS 수업목록 (.xls, .xlsx) 파일 선택",
        type=["xls", "xlsx", "xlsm", "html", "htm"],
        help="확장자가 .xls인 EUC-KR HTML 테이블 파일과 일반 .xlsx 파일을 모두 지원합니다.",
    )

    workspace_xls_files = sorted(
        list(PROJECT_ROOT.glob("*.xls")) + list(PROJECT_ROOT.glob("*.xlsx"))
    )
    rms_df = pd.DataFrame()
    rms_file_label = ""

    try:
        if rms_uploaded_file is not None:
            rms_df = parse_rms_xls(rms_uploaded_file)
            rms_file_label = rms_uploaded_file.name
        elif workspace_xls_files:
            rms_df = parse_rms_xls(workspace_xls_files[0])
            rms_file_label = f"{workspace_xls_files[0].name} (기본 로드)"

        if not rms_df.empty:
            detected_ym = detect_dominant_year_month(rms_df, default_ym=st.session_state["target_ym_val"])
            if st.session_state.get("last_rms_label") != rms_file_label:
                st.session_state["target_ym_val"] = detected_ym
                st.session_state["last_rms_label"] = rms_file_label
                user_settings["target_ym"] = detected_ym
                save_user_settings(user_settings)

            st.success(
                f"✅ **{rms_file_label}** 로드 완료\n\n"
                f"• 전체 개설 강좌: **{len(rms_df)}강좌**\n"
                f"• 기준 개강 연월: **`{st.session_state['target_ym_val']}`**"
            )
        else:
            st.info("RMS 수업목록(`.xls` 또는 `.xlsx`) 파일을 업로드해 주세요.")
    except Exception as e:
        st.error(f"RMS 파일 파싱 오류: {e}")

    st.divider()

    st.subheader("Step 2. 📄 강의계획서 PDF 업로드")
    pdf_uploaded_files = st.file_uploader(
        "강의계획서 (.pdf) 파일 선택 (다중 업로드 지원)",
        type=["pdf"],
        accept_multiple_files=True,
        help="통합 PDF 1개 또는 여러 개의 개별 PDF 파일을 동시에 드래그 앤 드롭할 수 있습니다.",
    )

    workspace_pdf_files = sorted(
        PROJECT_ROOT.glob("*.pdf"),
        key=lambda p: (0 if "강의계획서" in p.name else 1, p.name),
    )
    syllabus_pdf_df = pd.DataFrame()
    pdf_file_label = ""

    try:
        target_ym_for_pdf = st.session_state.get("target_ym_val", "2027-01")
        if pdf_uploaded_files:
            cache_parts = [f"{uf.name}:{uf.size}" for uf in pdf_uploaded_files]
            pdf_cache_key = f"upload::{'|'.join(cache_parts)}::{target_ym_for_pdf}"
            pdf_file_label = (
                pdf_uploaded_files[0].name
                if len(pdf_uploaded_files) == 1
                else f"{pdf_uploaded_files[0].name} 외 {len(pdf_uploaded_files) - 1}개 파일"
            )
            if st.session_state.get("pdf_cache_key") != pdf_cache_key:
                st.session_state["syllabus_pdf_df"] = parse_syllabus_pdf(
                    pdf_uploaded_files,
                    target_ym=target_ym_for_pdf,
                )
                st.session_state["pdf_cache_key"] = pdf_cache_key
            syllabus_pdf_df = st.session_state.get("syllabus_pdf_df", pd.DataFrame())
        elif workspace_pdf_files:
            default_pdf_path = workspace_pdf_files[0]
            pdf_cache_key = f"local::{default_pdf_path.name}::{default_pdf_path.stat().st_mtime_ns}::{target_ym_for_pdf}"
            pdf_file_label = f"{default_pdf_path.name} (기본 로드)"
            if st.session_state.get("pdf_cache_key") != pdf_cache_key:
                st.session_state["syllabus_pdf_df"] = parse_syllabus_pdf(
                    default_pdf_path,
                    target_ym=target_ym_for_pdf,
                )
                st.session_state["pdf_cache_key"] = pdf_cache_key
            syllabus_pdf_df = st.session_state.get("syllabus_pdf_df", pd.DataFrame())

        if not syllabus_pdf_df.empty:
            st.success(
                f"📄 **{pdf_file_label}**\n\n"
                f"• 추출된 강좌 블록: **{len(syllabus_pdf_df)}강좌**"
            )
        else:
            st.info("강의계획서 `.pdf` 파일을 업로드해 주세요.")
    except Exception as e:
        st.error(f"강의계획서 PDF 파싱 오류: {e}")

    st.divider()

    st.subheader("Step 3. ✍️ 문장 검수 옵션 설정")
    new_char_threshold = st.slider(
        "긴 문장 감지 기준 글자 수",
        min_value=40,
        max_value=100,
        value=int(st.session_state.get("max_char_threshold_val", 55)),
        step=5,
        help="설정한 글자 수 이상의 문장을 자동으로 감지하여 서술형/개조식 축약안을 제시합니다. (기본 55자)",
    )
    if new_char_threshold != st.session_state["max_char_threshold_val"]:
        st.session_state["max_char_threshold_val"] = new_char_threshold
        user_settings["max_char_threshold"] = new_char_threshold
        save_user_settings(user_settings)

    st.divider()

    st.subheader("Step 4. 🔑 Gemini API Key 설정")
    gemini_key_input = st.text_input(
        "Gemini API Key",
        value=st.session_state.get("gemini_api_key_val", ""),
        type="password",
        placeholder="AI 카피라이팅 & 강의명 추천 활성화",
        help="Google AI Studio에서 발급받은 Gemini API 키를 입력하시면 대치동 트렌드 AI 카피라이팅이 100% 활성화됩니다. 미입력 시에도 내장 규칙 엔진이 작동합니다.",
    )
    if gemini_key_input != st.session_state.get("gemini_api_key_val", ""):
        st.session_state["gemini_api_key_val"] = gemini_key_input.strip()
        user_settings["gemini_api_key"] = gemini_key_input.strip()
        save_user_settings(user_settings)

    if st.session_state.get("gemini_api_key_val"):
        st.caption("🟢 **Gemini AI 엔진 활성화됨** (`gemini-2.5-flash`)")
    else:
        st.caption("⚪ **내장 규칙 기반 엔진 작동 중** (API 키 입력 시 AI 업그레이드)")

# -------------------------------------------------------------------------
# RMS 요약 통계 추출 및 상단 고정 헤드라인 배너 표시 (파트 1)
# -------------------------------------------------------------------------
rms_summary = extract_rms_summary(rms_df)

if not rms_df.empty:
    st.markdown(
        f"""
        <div class="rms-headline-banner">
            <span class="rms-headline-title">{rms_summary['headline_banner']}</span>
            <span style="font-size:0.92rem; font-weight:600; color:#166534; background:#dcfce7; padding:4px 10px; border-radius:12px;">
                결제율: {rms_summary['payment_rate']}% (완납&분납 {rms_summary['installment_students']:,}명)
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

# -------------------------------------------------------------------------
# 코어 7대 검수 및 문장 개선 실행
# -------------------------------------------------------------------------
target_ym = st.session_state["target_ym_val"]
title_sim_threshold = st.session_state["title_sim_threshold_val"]
max_char_threshold = st.session_state["max_char_threshold_val"]
gemini_api_key = st.session_state.get("gemini_api_key_val", "").strip() or None

crosscheck = run_rms_pdf_crosscheck(
    rms_df=rms_df,
    syllabus_pdf_df=syllabus_pdf_df,
    mapper=mapper,
    target_ym=target_ym,
    title_similarity_threshold=float(title_sim_threshold),
    max_char_threshold=int(max_char_threshold),
    gemini_api_key=gemini_api_key,
)

comparison_df: pd.DataFrame = crosscheck["comparison_df"]
error_df: pd.DataFrame = crosscheck["error_df"]
missing_pdf_df: pd.DataFrame = crosscheck["missing_pdf_df"]
unmatched_pdf_df: pd.DataFrame = crosscheck["unmatched_pdf_df"]
pdf_schedule_df: pd.DataFrame = crosscheck["pdf_schedule_df"]
spelling_df: pd.DataFrame = crosscheck.get("spelling_df", pd.DataFrame())
long_sentence_df: pd.DataFrame = crosscheck.get("long_sentence_df", pd.DataFrame())
instructor_summary_df: pd.DataFrame = crosscheck.get("instructor_summary_df", pd.DataFrame())
kpi: Dict[str, int] = crosscheck["kpi"]

# AI 강의명 추천 데이터프레임이 세션에 있으면 crosscheck 결과에 결합
if not st.session_state["ai_titles_df"].empty:
    crosscheck["ai_titles_df"] = st.session_state["ai_titles_df"]

excel_bytes = export_crosscheck_to_excel(crosscheck, target_ym=target_ym) if not comparison_df.empty else b""
csv_bytes = export_crosscheck_to_csv_bytes(comparison_df) if not comparison_df.empty else b""
tsv_str = export_crosscheck_to_tsv_string(comparison_df) if not comparison_df.empty else ""

# 사이드바 하단 퀵 다운로드
with st.sidebar:
    if not comparison_df.empty:
        st.divider()
        st.subheader("📥 결과지 빠른 다운로드")
        now_short = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            label="📥 엑셀 결과지 다운로드 (.xlsx)",
            data=excel_bytes,
            file_name=f"러셀대치_강의계획서_종합검수결과_{target_ym}_{now_short}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
            key="dl_xlsx_sidebar",
        )
        st.download_button(
            label="📊 구글 시트용 CSV 다운로드",
            data=csv_bytes,
            file_name=f"러셀대치_강의계획서_종합검수결과_{target_ym}_{now_short}.csv",
            mime="text/csv",
            width="stretch",
            key="dl_csv_sidebar",
        )

# -------------------------------------------------------------------------
# 상단 요약 영역 내보내기 툴바
# -------------------------------------------------------------------------
if not comparison_df.empty:
    render_export_toolbar(
        crosscheck=crosscheck,
        excel_bytes=excel_bytes,
        csv_bytes=csv_bytes,
        tsv_str=tsv_str,
        target_ym=target_ym,
        section_key="header",
    )

# -------------------------------------------------------------------------
# 6개 메인 탭 구성
# -------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "📊 Tab 1: RMS 수강생 & 강사별 현황",
        "🔍 Tab 2: RMS ↔ PDF 정밀 대조표",
        "📅 Tab 3: 주차별 진도표 역산 검증",
        "✍️ Tab 4: 맞춤법 교정 & 긴 문장 간략화",
        "✨ Tab 5: AI 강의명 & 슬로건 제안",
        "📥 Tab 6: 결과지 다운로드 & 안내문 생성",
    ]
)

# =========================================================================
# Tab 1: 📊 RMS 수강생 & 강사별 현황 (파트 1 신규 전용 대시보드)
# =========================================================================
with tab1:
    if rms_df.empty:
        st.warning("좌측 사이드바 **Step 1**에서 RMS 수업목록(`.xls`/`.xlsx`) 파일을 업로드해 주세요.")
    else:
        st.subheader("📊 RMS 수강생 모집 및 강사별 실시간 현황")
        st.caption("RMS 강의정보 엑셀에서 추출한 총 강좌수, 가등록 포함 총 수강생수, 결제완료 인원 및 강사별 상세 지표입니다.")

        s_col1, s_col2, s_col3, s_col4, s_col5 = st.columns(5)
        with s_col1:
            st.metric(
                label="📋 전체 개설 강좌수",
                value=f"{rms_summary['total_courses']}건",
                delta=f"기준 {target_ym}",
            )
        with s_col2:
            st.metric(
                label="👥 총 수강생 (가등록포함)",
                value=f"{rms_summary['total_students']:,}명",
            )
        with s_col3:
            st.metric(
                label="💳 결제완료 인원",
                value=f"{rms_summary['paid_students']:,}명",
                delta=f"결제율 {rms_summary['payment_rate']}%",
            )
        with s_col4:
            st.metric(
                label="⏳ 가등록 (미결제 인원)",
                value=f"{rms_summary['unpaid_students']:,}명",
                delta_color="inverse",
            )
        with s_col5:
            st.metric(
                label="🧾 완납 & 분납 인원",
                value=f"{rms_summary['installment_students']:,}명",
            )

        st.divider()

        # 강사별 수강생 순위 집계표
        st.markdown("#### 🏆 강사별 수강생 순위 및 결제 현황 집계표")
        st.caption("총 수강생수 내림차순 정렬 (동률 시 결제완료 인원 -> 개설 강좌수 순)")

        if not instructor_summary_df.empty:
            st.dataframe(
                instructor_summary_df,
                width="stretch",
                hide_index=True,
            )
        else:
            st.info("강사별 집계 데이터가 없습니다.")

        st.divider()

        # 과목별 / 학년별 수강생 합계 요약표
        col_st_sub, col_st_grd = st.columns(2)
        with col_st_sub:
            st.markdown("##### 📚 과목(영역)별 수강생 합계")
            if "과목" in rms_df.columns:
                sub_group = (
                    rms_df.groupby("과목", as_index=False)
                    .agg(
                        개설강좌수=("강좌코드", "count"),
                        총수강생=("가등록포함수강생수", "sum"),
                        결제완료=("결제인원", "sum"),
                        미결제가등록=("가등록(미결제)", "sum"),
                    )
                    .sort_values(by="총수강생", ascending=False)
                )
                sub_group["결제율(%)"] = [
                    round(p / t * 100.0, 1) if t > 0 else 0.0
                    for p, t in zip(sub_group["결제완료"], sub_group["총수강생"])
                ]
                st.dataframe(sub_group, width="stretch", hide_index=True)

        with col_st_grd:
            st.markdown("##### 🎓 학년별 수강생 합계")
            if "학년" in rms_df.columns:
                grd_group = (
                    rms_df.groupby("학년", as_index=False)
                    .agg(
                        개설강좌수=("강좌코드", "count"),
                        총수강생=("가등록포함수강생수", "sum"),
                        결제완료=("결제인원", "sum"),
                        미결제가등록=("가등록(미결제)", "sum"),
                    )
                    .sort_values(by="총수강생", ascending=False)
                )
                grd_group["결제율(%)"] = [
                    round(p / t * 100.0, 1) if t > 0 else 0.0
                    for p, t in zip(grd_group["결제완료"], grd_group["총수강생"])
                ]
                st.dataframe(grd_group, width="stretch", hide_index=True)

        st.divider()

        # 강사별 상세 개별 강좌 펼침 보기
        st.markdown("#### 🔍 강사별 개별 강좌 수강생 상세 드릴다운")
        inst_list = sorted(rms_df["강사"].dropna().unique().tolist())
        selected_inst_for_drill = st.selectbox("확인할 강사 선택", options=inst_list)

        if selected_inst_for_drill:
            inst_courses = rms_df[rms_df["강사"] == selected_inst_for_drill].copy()
            st.markdown(
                f"**[{selected_inst_for_drill}T]** 총 {len(inst_courses)}개 강좌 개설 | "
                f"총 수강생: {inst_courses['가등록포함수강생수'].sum():,}명 | "
                f"결제완료: {inst_courses['결제인원'].sum():,}명"
            )

            detail_disp_cols = [
                "No",
                "과목",
                "학년",
                "강좌명",
                "수업시간",
                "개강일",
                "가등록포함수강생수",
                "결제인원",
                "가등록(미결제)",
                "수강료",
            ]
            st.dataframe(
                inst_courses[[c for c in detail_disp_cols if c in inst_courses.columns]],
                width="stretch",
                hide_index=True,
            )


# =========================================================================
# Tab 2: 🔍 RMS ↔ PDF 정밀 대조표
# =========================================================================
with tab2:
    if comparison_df.empty:
        st.warning("RMS 수업목록 파일을 먼저 업로드해 주세요.")
    else:
        render_export_toolbar(
            crosscheck=crosscheck,
            excel_bytes=excel_bytes,
            csv_bytes=csv_bytes,
            tsv_str=tsv_str,
            target_ym=target_ym,
            section_key="tab2",
        )

        f_col1, f_col2, f_col3, f_col4 = st.columns([1.5, 1.1, 1.1, 1.5])
        with f_col1:
            status_filter = st.radio(
                "검수 상태 필터",
                options=[
                    "전체 보기",
                    "📄 PDF 수록 강좌만 (정상+오류)",
                    "🔴 오류(수정필요)만",
                    "✅ 정상(완벽일치)만",
                    "⚪ PDF 미수록만",
                ],
                horizontal=False,
            )
        with f_col2:
            subj_options = ["전체"] + sorted(comparison_df["과목"].dropna().unique().tolist())
            selected_subj = st.selectbox("과목 필터", options=subj_options)
        with f_col3:
            grade_options = ["전체"] + sorted(comparison_df["학년"].dropna().unique().tolist())
            selected_grade = st.selectbox("학년 필터", options=grade_options)
        with f_col4:
            keyword_q = st.text_input(
                "검색 (강사명 / 강좌명 / 강좌코드)",
                placeholder="예: 심수현, 박성민, 301647",
            )

        filtered_df = comparison_df.copy()
        if status_filter == "📄 PDF 수록 강좌만 (정상+오류)":
            filtered_df = filtered_df[filtered_df["검수상태"] != "⚪ PDF 미수록"]
        elif status_filter == "🔴 오류(수정필요)만":
            filtered_df = filtered_df[filtered_df["검수상태"].str.startswith(("🔴", "🟡"))]
        elif status_filter == "✅ 정상(완벽일치)만":
            filtered_df = filtered_df[filtered_df["검수상태"].str.startswith("✅")]
        elif status_filter == "⚪ PDF 미수록만":
            filtered_df = filtered_df[filtered_df["검수상태"] == "⚪ PDF 미수록"]

        if selected_subj != "전체":
            filtered_df = filtered_df[filtered_df["과목"] == selected_subj]
        if selected_grade != "전체":
            filtered_df = filtered_df[filtered_df["학년"] == selected_grade]
        if keyword_q.strip():
            kq = keyword_q.strip().lower()
            filtered_df = filtered_df[
                filtered_df["강사명"].str.lower().str.contains(kq, na=False)
                | filtered_df["RMS 강좌명"].str.lower().str.contains(kq, na=False)
                | filtered_df["PDF 강좌명"].str.lower().str.contains(kq, na=False)
                | filtered_df["강좌코드"].str.lower().str.contains(kq, na=False)
            ]

        st.caption(f"조회 결과: 총 **{len(filtered_df)}건** (전체 {len(comparison_df)}건 중)")

        display_cols = [
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
            "총수강생(가등록포함)",
            "결제완료인원",
            "결제율(%)",
            "수업시간 비교",
            "바자관시수 비교",
            "오류 상세 내역",
            "자동 생성된 수정 요청 사항 요약",
        ]

        def highlight_crosscheck_row(row):
            status = str(row.get("검수상태", ""))
            if status.startswith(("🔴", "🟡")) or "오류" in status:
                return ["background-color: #fef2f2; color: #7f1d1d;"] * len(row)
            elif status.startswith("✅"):
                return ["background-color: #f0fdf4; color: #14532d;"] * len(row)
            elif status.startswith("⚪"):
                return ["background-color: #f8fafc; color: #64748b;"] * len(row)
            return [""] * len(row)

        sub_display_df = filtered_df[[c for c in display_cols if c in filtered_df.columns]].copy()
        styled_table = sub_display_df.style.apply(highlight_crosscheck_row, axis=1)
        st.dataframe(styled_table, width="stretch", height=520, hide_index=True)


# =========================================================================
# Tab 3: 📅 주차별 진도표 역산 검증
# =========================================================================
with tab3:
    if pdf_schedule_df.empty:
        st.info("PDF 강좌 정보가 로드되지 않았습니다.")
    else:
        st.subheader("📅 주차별 진도표 날짜 & 달력 역산 자가모순 검증")
        st.caption("PDF 1월 과정 날짜 리스트를 2027년 실제 달력 요일과 datetime으로 대조하여 비정상 요일 섞임 및 오타를 100% 검출합니다.")

        sched_cols = [
            "PDF_순번",
            "PDF_페이지",
            "검수상태",
            "과목",
            "학년",
            "강사명",
            "PDF 강좌명",
            "수업시간",
            "PDF 개요 개강~종강(회차)",
            "PDF 1월진도표(첫날~끝날/회차)",
            "1월 진도표 추출 날짜 리스트",
            "2월표 내 1월 날짜",
            "진도표 요일 정합성",
            "오류 상세 내역",
        ]

        st.dataframe(
            pdf_schedule_df[[c for c in sched_cols if c in pdf_schedule_df.columns]],
            width="stretch",
            hide_index=True,
        )

        st.divider()
        st.markdown("#### 🔍 특정 강좌 진도표 상세 내역 드릴다운")
        course_names = pdf_schedule_df["PDF 강좌명"].tolist()
        sel_cname = st.selectbox("확인할 강좌 선택", options=course_names)
        sel_row = pdf_schedule_df[pdf_schedule_df["PDF 강좌명"] == sel_cname].iloc[0]

        d_col1, d_col2 = st.columns(2)
        with d_col1:
            st.markdown(f"**📌 [{sel_row.get('강사명')}] 1월 과정 진도표**")
            jan_list = sel_row.get("주차별진도계획_1월리스트", []) or []
            if jan_list:
                jan_tbl_rows = []
                for idx_j, jit in enumerate(jan_list, start=1):
                    wk_ok = jit.get("요일일치여부", True)
                    wk_badge = "✅ 정상" if wk_ok else f"🔴 요일오타 (실제: {jit.get('실제요일')}요일)"
                    jan_tbl_rows.append(
                        {
                            "회차": f"{idx_j}회",
                            "PDF 표기 날짜": jit.get("날짜_원본", ""),
                            "달력 실제 요일": jit.get("실제요일", ""),
                            "요일 검증": wk_badge,
                            "단원 및 수업내용": jit.get("단원_및_수업내용", ""),
                        }
                    )
                st.dataframe(pd.DataFrame(jan_tbl_rows), width="stretch", hide_index=True)
            else:
                st.info("추출된 1월 진도표 항목이 없습니다.")

        with d_col2:
            st.markdown(f"**📌 [{sel_row.get('강사명')}] 2월 과정 진도표 (1월 날짜 혼입 점검)**")
            feb_list = sel_row.get("주차별진도계획_2월리스트", []) or []
            if feb_list:
                feb_tbl_rows = []
                for idx_f, fit in enumerate(feb_list, start=1):
                    is_jan_spill = str(fit.get("날짜", "")).startswith(str(target_ym).strip())
                    spill_badge = "🟡 1월 날짜가 2월표에 혼입됨!" if is_jan_spill else "✅ 2월"
                    feb_tbl_rows.append(
                        {
                            "순번": f"{idx_f}",
                            "PDF 표기 날짜": fit.get("날짜_원본", ""),
                            "실제 요일": fit.get("실제요일", ""),
                            "구분": spill_badge,
                            "단원 및 수업내용": fit.get("단원_및_수업내용", ""),
                        }
                    )
                st.dataframe(pd.DataFrame(feb_tbl_rows), width="stretch", hide_index=True)
            else:
                st.info("2월 과정 표가 없는 단일 1월 과정 강좌입니다.")


# =========================================================================
# Tab 4: ✍️ 맞춤법 교정 & 긴 문장 간략화 제안
# =========================================================================
with tab4:
    st.subheader("✍️ 국어 맞춤법 검수 및 긴 문장 간략화(축약/개조식) 제안")
    st.caption("PDF 본문 텍스트를 분석하여 띄어쓰기·오탈자·조사 오류를 적발하고, 긴 문장(기본 55자 이상)에 대해 [간결한 서술형]과 [브로슈어 개조식 요약] 2가지 수정안을 제공합니다.")

    sub_t1, sub_t2 = st.tabs([
        f"🔴 국어 맞춤법 & 오탈자 교정 ({len(spelling_df)}건)",
        f"✂️ 긴 문장 간략화 제안 ({len(long_sentence_df)}건 - 기준: {max_char_threshold}자)",
    ])

    with sub_t1:
        if spelling_df.empty:
            st.success("🎉 PDF 강의계획서 본문에서 발견된 국어 맞춤법 오류나 오탈자가 없습니다.")
        else:
            col_sp_f1, col_sp_f2 = st.columns([1.5, 2.5])
            with col_sp_f1:
                sp_type_opts = ["전체"] + sorted(spelling_df["오류유형"].unique().tolist())
                sel_sp_type = st.selectbox("오류 유형 필터", options=sp_type_opts, key="sp_type_filter")
            with col_sp_f2:
                sp_inst_opts = ["전체"] + sorted(spelling_df["강사명"].unique().tolist())
                sel_sp_inst = st.selectbox("강사 필터", options=sp_inst_opts, key="sp_inst_filter")

            flt_sp_df = spelling_df.copy()
            if sel_sp_type != "전체":
                flt_sp_df = flt_sp_df[flt_sp_df["오류유형"] == sel_sp_type]
            if sel_sp_inst != "전체":
                flt_sp_df = flt_sp_df[flt_sp_df["강사명"] == sel_sp_inst]

            st.caption(f"검출된 맞춤법/오탈자 내역: 총 **{len(flt_sp_df)}건**")

            for s_idx, (_, s_row) in enumerate(flt_sp_df.iterrows(), start=1):
                orig_sent = s_row["원문문장"]
                target_w = s_row["대상어구"]
                rep_w = s_row["교정제안"]
                hl_orig = orig_sent.replace(target_w, f"<b style='color:#e11d48; text-decoration:line-through;'>{target_w}</b>")
                hl_sug = orig_sent.replace(target_w, f"<b style='color:#16a34a; text-decoration:underline;'>{rep_w}</b>")

                st.markdown(
                    f"""
                    <div class="proposal-card">
                        <div class="proposal-title">
                            #{s_idx}. <span style="color:#b91c1c;">[{s_row['오류유형']}]</span> &nbsp;
                            <b>{s_row['강사명']}T</b> &nbsp;|&nbsp; {s_row['강좌명']}
                            <span style="font-size:0.85rem; color:#64748b;">({s_row['섹션']})</span>
                        </div>
                        <div style="font-size:0.92rem; line-height:1.6; margin-top:0.4rem;">
                            • <b>오류 어구</b> : <code>{target_w}</code> ➔ <b>교정 제안</b> : <code>{rep_w}</code> &nbsp;
                            <span style="color:#475569; font-size:0.88rem;">({s_row['사유']})</span><br>
                            • <b>원문 문장</b> : {hl_orig}<br>
                            • <b>교정 문장</b> : {hl_sug}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("#### 📋 맞춤법/오탈자 전체 표")
            st.dataframe(
                flt_sp_df[["강좌명", "강사명", "섹션", "오류유형", "대상어구", "교정제안", "사유", "원문문장", "교정문장"]],
                width="stretch",
                hide_index=True,
            )

    with sub_t2:
        if long_sentence_df.empty:
            st.success(f"🎉 기준 글자수({max_char_threshold}자)를 초과하는 긴 문장이 없습니다.")
        else:
            col_lg_f1, col_lg_f2 = st.columns([1.5, 2.5])
            with col_lg_f1:
                lg_inst_opts = ["전체"] + sorted(long_sentence_df["강사명"].unique().tolist())
                sel_lg_inst = st.selectbox("강사 필터", options=lg_inst_opts, key="lg_inst_filter")
            with col_lg_f2:
                lg_sec_opts = ["전체"] + sorted(long_sentence_df["섹션"].unique().tolist())
                sel_lg_sec = st.selectbox("섹션 필터", options=lg_sec_opts, key="lg_sec_filter")

            flt_lg_df = long_sentence_df.copy()
            if sel_lg_inst != "전체":
                flt_lg_df = flt_lg_df[flt_lg_df["강사명"] == sel_lg_inst]
            if sel_lg_sec != "전체":
                flt_lg_df = flt_lg_df[flt_lg_df["섹션"] == sel_lg_sec]

            st.caption(f"긴 문장 감지 내역: 총 **{len(flt_lg_df)}건** (기준: {max_char_threshold}자 이상)")

            for l_idx, (orig_index, l_row) in enumerate(flt_lg_df.head(30).iterrows(), start=1):
                c_len = l_row["글자수"]
                p1 = l_row["제안1_서술형"]
                p2 = l_row["제안2_개조식"]
                eng_type = l_row.get("생성엔진", "내장 규칙")

                state_key_choice = f"choice_lg_{orig_index}"
                state_key_text = f"text_lg_{orig_index}"

                if state_key_choice not in st.session_state:
                    st.session_state[state_key_choice] = "제안 1 (간결한 서술형)"
                if state_key_text not in st.session_state:
                    st.session_state[state_key_text] = p1

                with st.container():
                    st.markdown(
                        f"""
                        <div class="proposal-card">
                            <div class="proposal-title">
                                #{l_idx}. <b>[{l_row['강사명']}T]</b> {l_row['강좌명']} &nbsp;|&nbsp;
                                <span style="color:#d97706; font-weight:700;">{c_len}자 초과</span> &nbsp;
                                <span style="font-size:0.83rem; background:#e0f2fe; color:#0369a1; padding:2px 8px; border-radius:12px;">{eng_type}</span>
                            </div>
                            <div style="font-size:0.9rem; color:#334155; margin:0.4rem 0;">
                                <b>[원문 문장]</b> ({c_len}자)<br>
                                <span style="color:#64748b;">{l_row['원문문장']}</span>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    c_opt1, c_opt2 = st.columns(2)
                    with c_opt1:
                        st.info(f"**💡 [제안 1] 간결한 서술형** ({len(p1)}자)\n\n{p1}")
                    with c_opt2:
                        st.success(f"**📌 [제안 2] 브로슈어 개조식 요약** ({len(p2)}자)\n\n{p2}")

                    sel_opt = st.radio(
                        "반영할 문장 스타일 선택",
                        options=["제안 1 (간결한 서술형)", "제안 2 (브로슈어 개조식)", "직접 수정"],
                        horizontal=True,
                        key=state_key_choice,
                    )

                    if sel_opt == "제안 1 (간결한 서술형)":
                        final_text = p1
                    elif sel_opt == "제안 2 (브로슈어 개조식)":
                        final_text = p2
                    else:
                        final_text = st.text_input(
                            "사용자 직접 수정 문장",
                            value=st.session_state.get(state_key_text, p1),
                            key=state_key_text,
                        )

                    long_sentence_df.loc[orig_index, "최종문장"] = final_text
                    st.markdown("<hr style='margin:0.6rem 0;'>", unsafe_allow_html=True)


# =========================================================================
# Tab 5: ✨ AI 강의명 & 슬로건 제안 (파트 2 신규 기능)
# =========================================================================
with tab5:
    st.subheader("✨ Gemini AI 기반 대치동 단과 '강의명 & 슬로건 임팩트 수정 제안'")
    st.caption("대치동 학원가 최신 트렌드를 반영하여 평이한 강의명을 1등급 쟁취 및 학부모 수강 전환율을 극대화하는 임팩트 강의명과 한 줄 슬로건으로 리브랜딩합니다.")

    # 1. 상단 옵션 바
    ai_col1, ai_col2 = st.columns([1.6, 1.2])
    with ai_col1:
        direction_option = st.selectbox(
            "🎯 카피라이팅 방향성 선택",
            options=list(DIRECTION_PROMPTS.keys()),
            help="선택한 방향성에 맞춰 Gemini 2.5 Flash가 타깃 맞춤형 카피라이팅을 생성합니다.",
        )
    with ai_col2:
        st.write("")
        st.caption(f"선택된 전략: **{DIRECTION_PROMPTS[direction_option][:45]}...**")

    # 강좌 소스 선택 (PDF 수록 강좌 우선, 없으면 RMS 강좌)
    target_courses_source = []
    if syllabus_pdf_df is not None and not syllabus_pdf_df.empty:
        target_courses_source = syllabus_pdf_df.to_dict(orient="records")
        source_label = f"PDF 추출 강좌 ({len(target_courses_source)}개)"
    elif not rms_df.empty:
        target_courses_source = rms_df.to_dict(orient="records")
        source_label = f"RMS 개설 강좌 ({len(target_courses_source)}개)"

    btn_col1, btn_col2 = st.columns([1.5, 2.5])
    with btn_col1:
        run_batch_btn = st.button(
            f"✨ [{source_label}] 전체 강좌 임팩트 강의명 일괄 생성",
            type="primary",
            use_container_width=True,
            disabled=len(target_courses_source) == 0,
        )

    # 일괄 생성 로직 실행
    if run_batch_btn:
        progress_bar = st.progress(0)
        status_text = st.empty()

        def update_prog(cur, tot, cname):
            pct = int(cur / tot * 100)
            progress_bar.progress(pct)
            status_text.text(f"AI 카피라이팅 생성 중... ({cur}/{tot}) - {cname}")

        with st.spinner("Gemini 2.5 Flash가 대치동 스타 강사 카피라이팅을 생성 중입니다..."):
            generated_list = batch_suggest_course_titles(
                course_list=target_courses_source,
                api_key=gemini_api_key,
                direction=direction_option,
                progress_callback=update_prog,
            )
            st.session_state["ai_titles_df"] = pd.DataFrame(generated_list)
            progress_bar.empty()
            status_text.empty()
            st.success(f"🎉 총 {len(generated_list)}개 강좌의 임팩트 강의명 & 슬로건 생성이 완료되었습니다!")

    ai_df = st.session_state.get("ai_titles_df", pd.DataFrame())

    if ai_df.empty:
        st.info("💡 위의 **[✨ 전체 강좌 임팩트 강의명 일괄 생성]** 버튼을 누르시면 전체 강좌에 대한 3종 스타일 추천 강좌명이 즉시 생성됩니다.")
    else:
        st.divider()
        st.markdown("#### 📋 3종 스타일 강의명 & 슬로건 전체 비교표")

        table_cols = [
            "순번",
            "과목",
            "학년",
            "강사명",
            "현재강좌명",
            "임팩트점수",
            "추천슬로건",
            "제안1_직관형",
            "제안2_시그니처형",
            "제안3_킬러임팩트형",
            "진단비평",
            "생성엔진",
        ]
        st.dataframe(
            ai_df[[c for c in table_cols if c in ai_df.columns]],
            width="stretch",
            hide_index=True,
        )

        st.divider()
        st.markdown("#### 🗂️ 개별 강좌별 상세 카드 & 실시간 재수정")

        # 강좌별 카드 뷰
        for c_idx, (_, r_data) in enumerate(ai_df.iterrows(), start=1):
            with st.container():
                st.markdown(
                    f"""
                    <div class="proposal-card">
                        <div class="proposal-title">
                            #{c_idx}. [{r_data['과목']} {r_data['학년']}] <b>{r_data['강사명']}T</b> &nbsp;|&nbsp;
                            현재 강좌명: <span style="color:#475569;">{r_data['현재강좌명']}</span> &nbsp;
                            <span style="background:#fee2e2; color:#b91c1c; padding:2px 8px; border-radius:10px; font-size:0.83rem;">
                                기존 임팩트: {r_data['임팩트점수']}점
                            </span>
                        </div>
                        <div style="font-size:0.88rem; color:#64748b; margin:0.3rem 0;">
                            💡 <b>진단 및 비평:</b> {r_data['진단비평']}
                        </div>
                        <div style="font-size:0.92rem; font-weight:600; color:#1e40af; margin-bottom:0.6rem;">
                            ✨ <b>추천 슬로건:</b> <code>{r_data['추천슬로건']}</code>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                st_c1, st_c2, st_c3 = st.columns(3)
                with st_c1:
                    st.info(f"**🎯 [제안 1] 타깃·성적향상 직관형**\n\n{r_data['제안1_직관형']}")
                with st_c2:
                    st.success(f"**💎 [제안 2] 시그니처·브랜딩형**\n\n{r_data['제안2_시그니처형']}")
                with st_c3:
                    st.warning(f"**🔥 [제안 3] 최상위권·실전압도형**\n\n{r_data['제안3_킬러임팩트형']}")

                # 개별 다시 생성 버튼
                btn_re_key = f"regen_btn_{c_idx}"
                if st.button(f"🔄 [{r_data['강사명']}T] 이 강좌만 다른 스타일로 다시 생성", key=btn_re_key):
                    match_course = next((c for c in target_courses_source if (c.get('PDF 강좌명') or c.get('강좌명')) == r_data['현재강좌명']), None)
                    if match_course:
                        with st.spinner("해당 강좌를 다시 카피라이팅하는 중..."):
                            re_sug = suggest_course_titles_with_gemini(match_course, api_key=gemini_api_key, direction=direction_option)
                            ai_df.loc[ai_df["현재강좌명"] == r_data["현재강좌명"], "추천슬로건"] = re_sug["slogan_suggestion"]
                            ai_df.loc[ai_df["현재강좌명"] == r_data["현재강좌명"], "제안1_직관형"] = re_sug["style_1_intuitive"]
                            ai_df.loc[ai_df["현재강좌명"] == r_data["현재강좌명"], "제안2_시그니처형"] = re_sug["style_2_signature"]
                            ai_df.loc[ai_df["현재강좌명"] == r_data["현재강좌명"], "제안3_킬러임팩트형"] = re_sug["style_3_killer"]
                            st.session_state["ai_titles_df"] = ai_df
                            st.rerun()

                st.markdown("<hr style='margin:0.8rem 0;'>", unsafe_allow_html=True)


# =========================================================================
# Tab 6: 📥 결과지 다운로드 & 수정 안내문 생성
# =========================================================================
with tab6:
    if comparison_df.empty:
        st.warning("RMS 수업목록 파일을 먼저 업로드해 주세요.")
    else:
        st.subheader("📥 검수 결과지 다운로드 & 구글 시트 내보내기")
        render_export_toolbar(
            crosscheck=crosscheck,
            excel_bytes=excel_bytes,
            csv_bytes=csv_bytes,
            tsv_str=tsv_str,
            target_ym=target_ym,
            section_key="tab6",
        )

        st.divider()
        st.subheader("📝 디자인팀 / 강사 전달용 [PDF 수정 요청 텍스트] 자동 생성")

        if error_df.empty:
            st.success("오류가 발견된 강좌가 없어 수정 요청할 내역이 없습니다.")
        else:
            fix_messages = generate_pdf_fix_request_messages(error_df, target_ym=target_ym)
            all_fix_text = fix_messages.get("__ALL__", "")

            msg_col1, msg_col2 = st.columns([1.3, 1.0])
            with msg_col1:
                st.markdown("##### 📋 [전체 통합본] 디자인팀/교무입시팀 전달용 PDF 수정 요청문")
                render_clipboard_and_sheets_new_widget(
                    tsv_text=all_fix_text,
                    button_id="fix_all_msg",
                    label_text="📋 전체 수정 요청 텍스트 원클릭 복사",
                )
                st.text_area(
                    "전체 PDF 수정 요청 텍스트 (Ctrl+A → Ctrl+C 가능)",
                    value=all_fix_text,
                    height=340,
                )

            with msg_col2:
                st.markdown("##### 👨‍🏫 [강사별 개별 안내문] 확인 및 복사")
                inst_keys = [k for k in fix_messages.keys() if k != "__ALL__"]
                if inst_keys:
                    sel_inst = st.selectbox("강사 선택", options=inst_keys)
                    inst_msg_text = fix_messages.get(sel_inst, "")
                    render_clipboard_and_sheets_new_widget(
                        tsv_text=inst_msg_text,
                        button_id="fix_inst_msg",
                        label_text=f"📋 {sel_inst}T 안내문 원클릭 복사",
                    )
                    st.text_area(
                        f"{sel_inst} 선생님 개별 안내문",
                        value=inst_msg_text,
                        height=290,
                    )

        st.divider()
        with st.expander("⚙️ 강사명 매핑 사전 관리 (RMS 강사명 ↔ PDF 강사명 별칭 설정)"):
            st.caption("동명이인 구분(예: 박소현(리아) ↔ 리아, 박소현 ↔ 박소현(셀리나)) 및 필명/본명 매핑을 관리합니다.")
            mapping_df = pd.DataFrame(mapper.mappings)
            edited_mapping_df = st.data_editor(
                mapping_df,
                num_rows="dynamic",
                width="stretch",
                key="instructor_mapping_editor",
            )
            if st.button("💾 강사명 매핑 설정 저장"):
                mapper.save(edited_mapping_df.to_dict(orient="records"))
                st.success("✅ 강사명 매핑 사전이 저장되었습니다.")
                st.rerun()
