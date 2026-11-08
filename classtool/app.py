"""
메가스터디 러셀대치학원 단과 강좌 [RMS 수업목록(.xls/.xlsx) ↔ 강의계획서 PDF] 종합 검수 시스템
- 7대 핵심 항목 정밀 대조 (수강료/단가, 회차 3중 대조, 개강일/종강일 3중 대조 및 요일 모순, 진도표 요일오타/혼입, 수업시간, 바자관시수, 강좌명)
- 국어 맞춤법·오탈자 검수 및 긴 문장(기본 55자) 간략화 제안 모듈
- 검수 결과지 내보내기: 5개 시트 서식 포함 엑셀(.xlsx), 구글 시트 호환 CSV/TSV, 원클릭 클립보드 복사 + sheets.new 연동, 구글 시트 URL 직접 내보내기
- 고정 접속 포트: 8504 (http://localhost:8504)
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
import parsers.pdf_parser as _pdf_parser_mod
import parsers.rms_parser as _rms_parser_mod

# 핫리로드 시 하위 모듈 변경사항 즉시 반영
importlib.reload(_rms_parser_mod)
importlib.reload(_pdf_parser_mod)
importlib.reload(_text_improver_mod)
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
from parsers.pdf_parser import parse_syllabus_pdf
from parsers.rms_parser import detect_dominant_year_month, parse_rms_xls

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
        margin-bottom: 1rem;
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

PROJECT_ROOT = Path(__file__).resolve().parent
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
if "target_ym_val" not in st.session_state:
    st.session_state["target_ym_val"] = str(user_settings.get("target_ym", "2027-01"))
if "title_sim_threshold_val" not in st.session_state:
    st.session_state["title_sim_threshold_val"] = int(user_settings.get("title_sim_threshold", 55))
if "max_char_threshold_val" not in st.session_state:
    st.session_state["max_char_threshold_val"] = int(user_settings.get("max_char_threshold", 55))
if "gemini_api_key_val" not in st.session_state:
    st.session_state["gemini_api_key_val"] = str(user_settings.get("gemini_api_key", os.environ.get("GEMINI_API_KEY", "")))
if "export_sheet_url_val" not in st.session_state:
    st.session_state["export_sheet_url_val"] = str(
        user_settings.get("export_sheet_url", user_settings.get("sheet_url", ""))
    )

mapper = InstructorMapper(config_path=CONFIG_PATH)


def render_clipboard_and_sheets_new_widget(
    tsv_text: str, button_id: str, label_text: str = "📋 구글 시트 붙여넣기용 원클릭 복사"
) -> None:
    """
    클릭 한 번으로 전체 검수 결과 표(TSV)를 클립보드에 복사하고,
    바로 옆에 [🚀 새 구글 시트 열기 (sheets.new)] 링크 버튼을 렌더링합니다.
    """
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
    """상단 요약 영역과 대조표 탭 상단에 배치되는 결과지 내보내기 툴바"""
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
            label="📥 엑셀 결과지 다운로드 (.xlsx - 서식·5개 시트)",
            data=excel_bytes,
            file_name=f"러셀대치_강의계획서_종합검수결과_{target_ym}_{now_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
            key=f"dl_xlsx_{section_key}",
            help="전체검수결과, 데이터오류_수정요청, 맞춤법_오탈자교정, 긴문장_간략화제안, PDF누락강좌 시트가 포함된 서식 적용 엑셀 파일",
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
        <p>[RMS 수업목록(.xls/.xlsx)] ↔ [강의계획서 PDF] 7대 정밀 대조 · 국어 맞춤법/오탈자 검수 · 긴 문장 축약/개조식 제안 · 5개 시트 결과지 다운로드</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------------
# 사이드바 (Sidebar): 3단계 업로더 & 검수 옵션
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
                f"✅ **{rms_file_label}**\n\n"
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

    gemini_key_input = st.text_input(
        "Gemini API Key (선택사항)",
        value=st.session_state.get("gemini_api_key_val", ""),
        type="password",
        placeholder="AI 카피라이팅 활성화 시 입력",
        help="입력 시 Gemini AI 기반 최상위권 대치동 브로슈어 스타일 카피라이팅이 활성화됩니다. 미입력 시에도 내장 규칙 기반 축약 엔진이 100% 작동합니다.",
    )
    if gemini_key_input != st.session_state.get("gemini_api_key_val", ""):
        st.session_state["gemini_api_key_val"] = gemini_key_input.strip()
        user_settings["gemini_api_key"] = gemini_key_input.strip()
        save_user_settings(user_settings)

    st.divider()
    with st.expander("⚙️ 고급 검수 기준 설정"):
        target_ym_input = st.text_input(
            "검수 대상 연월 (YYYY-MM)",
            value=st.session_state["target_ym_val"],
            help="1월 진도표 날짜 요일 역산 및 개강일 검수 기준 연월",
        )
        if target_ym_input != st.session_state["target_ym_val"]:
            st.session_state["target_ym_val"] = target_ym_input.strip()
            user_settings["target_ym"] = target_ym_input.strip()
            save_user_settings(user_settings)

        title_sim_input = st.slider(
            "강좌명 유사도 판정 임계치 (%)",
            min_value=40,
            max_value=90,
            value=int(st.session_state["title_sim_threshold_val"]),
            step=5,
        )
        if title_sim_input != st.session_state["title_sim_threshold_val"]:
            st.session_state["title_sim_threshold_val"] = title_sim_input
            user_settings["title_sim_threshold"] = title_sim_input
            save_user_settings(user_settings)

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
kpi: Dict[str, int] = crosscheck["kpi"]

excel_bytes = export_crosscheck_to_excel(crosscheck, target_ym=target_ym) if not comparison_df.empty else b""
csv_bytes = export_crosscheck_to_csv_bytes(comparison_df) if not comparison_df.empty else b""
tsv_str = export_crosscheck_to_tsv_string(comparison_df) if not comparison_df.empty else ""

# 사이드바 하단 퀵 다운로드 버튼
with st.sidebar:
    if not comparison_df.empty:
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
# 상단 요약 영역 결과지 내보내기 툴바
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
# 5개 메인 탭 구성
# -------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [
        "📊 Tab 1: 검수 요약 & 오류 브리핑",
        "🔍 Tab 2: RMS ↔ PDF 정밀 대조표",
        "📅 Tab 3: 주차별 진도표 역산 검증",
        "✍️ Tab 4: 맞춤법 교정 & 긴 문장 간략화 제안",
        "📥 Tab 5: 결과지 다운로드 & 수정 요청문 생성",
    ]
)

# =========================================================================
# Tab 1: 📊 검수 요약 & 오류 브리핑
# =========================================================================
with tab1:
    if comparison_df.empty:
        st.warning("좌측 사이드바 **Step 1**에서 RMS 수업목록(`.xls`/`.xlsx`) 파일을 업로드해 주세요.")
    else:
        total_rms = kpi["total_rms"]
        pdf_matched = kpi["pdf_matched_count"]
        normal_cnt = kpi["normal_count"]
        error_cnt = kpi["error_count"]
        missing_cnt = kpi["missing_pdf_count"]
        spelling_cnt = kpi.get("spelling_error_count", len(spelling_df))
        long_cnt = kpi.get("long_sentence_count", len(long_sentence_df))

        st.subheader(f"📌 {target_ym} 단과 강의계획서 종합 검수 현황 요약")

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        with m1:
            st.metric(
                label="📋 전체 RMS 강좌",
                value=f"{total_rms}건",
                delta=f"기준 {target_ym}",
            )
        with m2:
            pdf_rate = (pdf_matched / total_rms * 100.0) if total_rms > 0 else 0.0
            st.metric(
                label="📄 PDF 수록 강좌",
                value=f"{pdf_matched}건",
                delta=f"수록률 {pdf_rate:.1f}%",
            )
        with m3:
            ok_rate = (normal_cnt / pdf_matched * 100.0) if pdf_matched > 0 else 0.0
            st.metric(
                label="✅ 데이터 완벽 일치",
                value=f"{normal_cnt}건",
                delta=f"일치율 {ok_rate:.1f}%",
            )
        with m4:
            st.metric(
                label="🚨 데이터 오류 (수정필요)",
                value=f"{error_cnt}건",
                delta=(
                    f"수강료 {kpi['fee_error_count']} · 회차 {kpi['session_error_count']} · 날짜 {kpi['date_error_count']}"
                    if error_cnt > 0
                    else "오류 없음"
                ),
                delta_color="inverse" if error_cnt > 0 else "normal",
            )
        with m5:
            st.metric(
                label="🔴 맞춤법·오탈자 검출",
                value=f"{spelling_cnt}건",
                delta="Tab 4 확인" if spelling_cnt > 0 else "오탈자 없음",
                delta_color="inverse" if spelling_cnt > 0 else "normal",
            )
        with m6:
            st.metric(
                label="✂️ 긴 문장 축약 대상",
                value=f"{long_cnt}건",
                delta=f"{max_char_threshold}자 이상",
                delta_color="inverse" if long_cnt > 0 else "normal",
            )

        st.divider()

        # 🚨 수정 필요 강좌 핵심 요약 브리핑 카드
        st.markdown("### 🚨 수정 필요 강좌 핵심 요약 브리핑")
        if error_df.empty:
            st.success("🎉 PDF에 수록된 모든 강좌의 수강료, 회차, 개강일, 종강일, 1월 진도표 날짜/요일, 수업시간, 바자관 시수가 RMS와 100% 일치합니다!")
        else:
            st.error(
                f"총 **{len(error_df)}건**의 강좌에서 **수강료 불일치, 회차 불일치, 개강일 요일 모순, 종강일 불일치, 또는 진도표 날짜/요일 오타**가 검출되었습니다."
            )

            for idx, (_, erow) in enumerate(error_df.iterrows(), start=1):
                st.markdown(
                    f"""
                    <div class="error-brief-card">
                        <div class="error-brief-title">
                            {idx}. {erow['검수상태']} &nbsp;|&nbsp; [{erow['과목']} · {erow['학년']}] {erow['강사명']}T — {erow['PDF 강좌명']}
                            <span style="font-size:0.85rem; font-weight:500; color:#6b7280;">(PDF {erow['PDF 페이지']} / RMS No.{erow['No']} · 강좌코드 {erow['강좌코드']})</span>
                        </div>
                        <div class="error-brief-body">
                            • <b>개강~종강(회차) 3중 대조</b> :
                              <b>[RMS]</b> {erow['RMS 개강~종강(회차)']} &nbsp;↔&nbsp;
                              <b>[PDF 개요]</b> {erow['PDF 개요 개강~종강(회차)']} &nbsp;↔&nbsp;
                              <b>[PDF 1월 진도표]</b> {erow['PDF 1월진도표(첫날~끝날/회차)']}<br>
                            • <b>수강료 대조</b> : <b>[RMS]</b> {erow['RMS 수강료']} &nbsp;↔&nbsp; <b>[PDF]</b> {erow['PDF 수강료']} &nbsp;|&nbsp;
                              <b>수업시간</b> : {erow['수업시간 비교']} &nbsp;|&nbsp; <b>바자관</b> : {erow['바자관시수 비교']}<br>
                            • <b>검출된 오류 상세</b> : <span style="color:#be123c; font-weight:600;">{erow['오류 상세 내역']}</span>
                        </div>
                        <div class="error-fix-box">
                            🛠️ 자동 생성된 수정 요청 사항 : {erow['자동 생성된 수정 요청 사항 요약']}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.divider()

        # 과목별 / 학년별 통계
        col_stat1, col_stat2 = st.columns(2)
        with col_stat1:
            st.markdown("#### 📊 과목별 PDF 수록 및 검수 현황")
            subj_rows = []
            for subj_name, grp in comparison_df.groupby("과목", sort=False):
                tot = len(grp)
                n_ok = int(grp["검수상태"].str.startswith("✅").sum())
                n_err = int(grp["검수상태"].str.startswith(("🔴", "🟡")).sum())
                n_miss = int((grp["검수상태"] == "⚪ PDF 미수록").sum())
                subj_rows.append(
                    {
                        "과목": subj_name,
                        "RMS 전체": tot,
                        "PDF 수록": n_ok + n_err,
                        "✅ 일치": n_ok,
                        "🚨 오류": n_err,
                        "⚪ 미수록": n_miss,
                    }
                )
            st.dataframe(pd.DataFrame(subj_rows), width="stretch", hide_index=True)

        with col_stat2:
            st.markdown("#### 🎓 학년별 PDF 수록 및 검수 현황")
            grade_rows = []
            for grade_name, grp in comparison_df.groupby("학년", sort=False):
                tot = len(grp)
                n_ok = int(grp["검수상태"].str.startswith("✅").sum())
                n_err = int(grp["검수상태"].str.startswith(("🔴", "🟡")).sum())
                n_miss = int((grp["검수상태"] == "⚪ PDF 미수록").sum())
                grade_rows.append(
                    {
                        "학년": grade_name,
                        "RMS 전체": tot,
                        "PDF 수록": n_ok + n_err,
                        "✅ 일치": n_ok,
                        "🚨 오류": n_err,
                        "⚪ 미수록": n_miss,
                    }
                )
            st.dataframe(pd.DataFrame(grade_rows), width="stretch", hide_index=True)


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
                placeholder="예: 심수현, 박성민, 10103",
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
# Tab 4: ✍️ 국어 맞춤법 & 긴 문장 간략화 제안 (인라인 에디터 & 승인/반영)
# =========================================================================
with tab4:
    st.subheader("✍️ 국어 맞춤법 검수 및 긴 문장 간략화(축약/개조식) 제안")
    st.caption("PDF 본문 텍스트를 분석하여 띄어쓰기·오탈자·조사 오류를 적발하고, 긴 문장(기본 55자 이상)에 대해 [간결한 서술형]과 [브로슈어 개조식 요약] 2가지 수정안을 제공합니다.")

    sub_t1, sub_t2 = st.tabs([
        f"🔴 국어 맞춤법 & 오탈자 교정 ({len(spelling_df)}건)",
        f"✂️ 긴 문장 간략화 제안 ({len(long_sentence_df)}건 - 기준: {max_char_threshold}자)",
    ])

    # ---------------------------------------------------------------------
    # 서브탭 1: 맞춤법 및 오탈자 교정
    # ---------------------------------------------------------------------
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

            # 카드 형태로 세부 교정안 표시
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

    # ---------------------------------------------------------------------
    # 서브탭 2: 긴 문장 간략화 제안
    # ---------------------------------------------------------------------
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

            # 인터랙티브 수정 카드
            for l_idx, (orig_index, l_row) in enumerate(flt_lg_df.head(40).iterrows(), start=1):
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

                    # 선택 라디오 및 인라인 수정기
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

                    # DataFrame 및 crosscheck 결과에 실시간 반영
                    long_sentence_df.loc[orig_index, "최종문장"] = final_text
                    st.markdown("<hr style='margin:0.6rem 0;'>", unsafe_allow_html=True)

            if len(flt_lg_df) > 40:
                st.info(f"💡 전체 {len(flt_lg_df)}건 중 상위 40건의 에디터를 표시 중입니다. 전체 목록은 아래 표 및 Tab 5 엑셀 결과지에서 모두 확인하실 수 있습니다.")

            st.markdown("#### 📋 긴 문장 간략화 제안 전체 표")
            st.dataframe(
                long_sentence_df[["강좌명", "강사명", "섹션", "글자수", "원문문장", "제안1_서술형", "제안2_개조식", "최종문장"]],
                width="stretch",
                hide_index=True,
            )


# =========================================================================
# Tab 5: 📥 결과지 다운로드 & 수정 요청문 생성
# =========================================================================
with tab5:
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
            section_key="tab5",
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
