"""
Google Sheets 및 외부 연동 통합 익스포터 모듈 (core/exporter.py)
1. [올인원(All-in-One) 마스터 통합 시트]: 강좌별 기본정보, 수강생 통계, PDF 대조 검수, 맞춤법/문장축약, AI 추천강의명을 가로로 완전 병합
2. [전체 리포트 일괄 시트 (Multi-Section TSV)]: A1 셀에 한 번 붙여넣기로 모든 섹션(요약/통계/검수/교정)을 일괄 열람 가능한 TSV
3. 브라우저 JavaScript 기반 클립보드 원클릭 복사 & https://sheets.new 자동 오픈 컴포넌트
4. 구글 시트 전용 UTF-8-SIG CSV 및 전체 시트 ZIP 패키지 다운로드 지원
5. Google Apps Script Web App 연동: JSON 키 없이 구글 시트 URL/웹앱으로 전 탭 자동 전송
"""

from __future__ import annotations

import io
import json
import re
import zipfile
from datetime import datetime
from typing import Any, Dict, List, Optional, Set, Tuple, Union

import pandas as pd
import requests
import streamlit.components.v1 as components

from core.cumulative_db import (
    build_snapshot_pivot_table,
    compare_two_snapshots,
    format_snapshot_label,
    get_available_snapshots,
    normalize_snapshot_key,
)


# -------------------------------------------------------------------------
# 1. 올인원(All-in-One) 마스터 통합 DataFrame 생성
# -------------------------------------------------------------------------
def build_all_in_one_master_df(
    crosscheck_result: Dict[str, Any],
    target_ym: str = "2027-01",
) -> pd.DataFrame:
    """
    RMS 강좌 목록을 기준으로 PDF 검수 결과, 수강생 수, 맞춤법/문장 축약, AI 추천 강의명까지
    한 행(Row)에 모든 정보를 가로로 완전 결합한 올인원 마스터 DataFrame을 생성합니다.
    """
    comparison_df = crosscheck_result.get("comparison_df")
    if comparison_df is None:
        comparison_df = pd.DataFrame()
    stopped_df = crosscheck_result.get("stopped_rms_df")
    if stopped_df is None:
        stopped_df = pd.DataFrame()
    spelling_df = crosscheck_result.get("spelling_df")
    if spelling_df is None:
        spelling_df = pd.DataFrame()
    bullet_issues_df = crosscheck_result.get("bullet_issues_df")
    if bullet_issues_df is None:
        bullet_issues_df = pd.DataFrame()
    disallowed_punct_df = crosscheck_result.get("disallowed_punct_df")
    if disallowed_punct_df is None:
        disallowed_punct_df = pd.DataFrame()
    long_sentence_df = crosscheck_result.get("long_sentence_df")
    if long_sentence_df is None:
        long_sentence_df = pd.DataFrame()
    ai_titles_df = crosscheck_result.get("ai_titles_df")

    # 만약 세션에 AI 강의명이 아직 없다면 엑셀 통합 시 기본 내장 규칙으로 0.01초 만에 자동 생성
    if ai_titles_df is None or (isinstance(ai_titles_df, pd.DataFrame) and ai_titles_df.empty):
        if comparison_df is not None and not comparison_df.empty:
            from core.title_suggester import batch_suggest_course_titles
            courses_list = comparison_df.to_dict(orient="records")
            ai_results = batch_suggest_course_titles(
                course_list=courses_list,
                api_key=None,
                direction="대치동 트렌드 종합",
            )
            ai_titles_df = pd.DataFrame(ai_results)
            crosscheck_result["ai_titles_df"] = ai_titles_df
        else:
            ai_titles_df = pd.DataFrame()

    # 1. 맞춤법 요약 사전 구축 (강좌명 및 강사명 매핑)
    spell_map: Dict[Tuple[str, str], List[str]] = {}
    if not spelling_df.empty:
        for _, r in spelling_df.iterrows():
            c_title = str(r.get("강좌명", "")).strip()
            inst = str(r.get("강사명", "")).strip()
            err_w = str(r.get("오류어구", r.get("대상어구", ""))).strip()
            fix_w = str(r.get("교정안", r.get("교정제안", ""))).strip()
            item_str = f"{err_w} ➔ {fix_w}" if err_w and fix_w else (err_w or fix_w)
            if item_str:
                spell_map.setdefault((inst, c_title), []).append(item_str)
                spell_map.setdefault(("", c_title), []).append(item_str)

    # 1-2. 문장부호(■) 미통일 요약 사전 구축
    bullet_map: Dict[Tuple[str, str], List[str]] = {}
    if not bullet_issues_df.empty:
        for _, r in bullet_issues_df.iterrows():
            c_title = str(r.get("강좌명", "")).strip()
            inst = str(r.get("강사명", "")).strip()
            item_str = str(r.get("발견항목", f"{r.get('대상어구', '')} → ■")).strip()
            if item_str:
                bullet_map.setdefault((inst, c_title), []).append(item_str)
                bullet_map.setdefault(("", c_title), []).append(item_str)

    # 1-3. 비허용 문장부호(■, -, [ ] 외) 요약 사전 구축
    disallowed_map: Dict[Tuple[str, str], List[str]] = {}
    if not disallowed_punct_df.empty:
        for _, r in disallowed_punct_df.iterrows():
            c_title = str(r.get("강좌명", "")).strip()
            inst = str(r.get("강사명", "")).strip()
            item_str = str(r.get("발견항목", r.get("대상어구", ""))).strip()
            if item_str:
                disallowed_map.setdefault((inst, c_title), []).append(item_str)
                disallowed_map.setdefault(("", c_title), []).append(item_str)

    # 2. 긴 문장 축약 요약 사전 구축
    long_map: Dict[Tuple[str, str], List[str]] = {}
    if not long_sentence_df.empty:
        for _, r in long_sentence_df.iterrows():
            c_title = str(r.get("강좌명", "")).strip()
            inst = str(r.get("강사명", "")).strip()
            s1 = str(r.get("제안1_서술형", "")).strip()
            s2 = str(r.get("제안2_개조식", "")).strip()
            summary_part = []
            if s1:
                summary_part.append(f"[서술형] {s1}")
            if s2:
                summary_part.append(f"[개조식] {s2}")
            if summary_part:
                long_map.setdefault((inst, c_title), []).append(" | ".join(summary_part))
                long_map.setdefault(("", c_title), []).append(" | ".join(summary_part))

    # 3. AI 추천 강의명 사전 구축
    ai_map: Dict[Tuple[str, str], Dict[str, Any]] = {}
    if ai_titles_df is not None and not ai_titles_df.empty:
        for _, r in ai_titles_df.iterrows():
            c_title = str(r.get("현재강의명", r.get("현재강좌명", ""))).strip()
            inst = str(r.get("강사명", "")).strip()
            ai_info = {
                "판정결과": str(r.get("판정결과", "🔧 수정 추천" if r.get("needs_revision") else "✅ 기존 유지")),
                "임팩트점수": r.get("임팩트점수", 70),
                "판정사유": str(r.get("판정사유", r.get("진단비평", ""))),
                "추천1": str(r.get("⚡ 추천 강의명 1 (직관·시그니처형)", r.get("제안1_초단축직관형", "-"))),
                "추천2": str(r.get("🔥 추천 강의명 2 (킬러·임팩트형)", r.get("제안2_초단축시그니처", "-"))),
            }
            ai_map[(inst, c_title)] = ai_info
            ai_map[("", c_title)] = ai_info
            from parsers.rms_parser import strip_course_brackets
            pure_t, _ = strip_course_brackets(c_title)
            if pure_t:
                ai_map[(inst, pure_t)] = ai_info
                ai_map[("", pure_t)] = ai_info

    master_rows: List[Dict[str, Any]] = []

    # comparison_df 순회
    if not comparison_df.empty:
        for _, row in comparison_df.iterrows():
            inst = str(row.get("강사명", "")).strip()
            rms_inst = str(row.get("RMS_강사명", "") or row.get("강사명", "")).strip()
            norm_inst = str(row.get("정규화강사명", "")).strip()
            rms_title = str(row.get("RMS 강좌명", "")).strip()
            pdf_title = str(row.get("PDF 강좌명", "")).strip()

            # 매핑 데이터 찾기
            sp_list = spell_map.get((inst, pdf_title)) or spell_map.get((inst, rms_title)) or spell_map.get(("", pdf_title)) or []
            bl_list = bullet_map.get((inst, pdf_title)) or bullet_map.get((inst, rms_title)) or bullet_map.get(("", pdf_title)) or []
            dp_list = disallowed_map.get((inst, pdf_title)) or disallowed_map.get((inst, rms_title)) or disallowed_map.get(("", pdf_title)) or []
            lg_list = long_map.get((inst, pdf_title)) or long_map.get((inst, rms_title)) or long_map.get(("", pdf_title)) or []
            ai_info = (
                ai_map.get((inst, pdf_title))
                or ai_map.get((inst, rms_title))
                or ai_map.get((rms_inst, rms_title))
                or ai_map.get((norm_inst, rms_title))
                or ai_map.get((rms_inst, pdf_title))
                or ai_map.get(("", pdf_title))
                or ai_map.get(("", rms_title))
                or {}
            )

            # 수강료 단가 계산
            rms_fee = row.get("RMS 수강료", 0)
            pdf_fee = row.get("PDF 수강료", 0)

            # 결제 관련
            tot_stu = row.get("총수강생(가등록포함)", 0)
            paid_stu = row.get("결제완료인원", 0)
            unpaid_stu = int(tot_stu) - int(paid_stu) if str(tot_stu).isdigit() and str(paid_stu).isdigit() else 0

            status_val = str(row.get("검수상태", ""))

            # 강의상태 결정
            raw_sale = str(row.get("강의상태", "") or row.get("판매", "")).strip()
            if not raw_sale or ("판매" in raw_sale and "중지" not in raw_sale):
                sale_status_str = "판매 중"
            elif "중지" in raw_sale:
                sale_status_str = "[판매중지]"
            else:
                sale_status_str = raw_sale

            master_rows.append({
                # 1. 기본 정보 및 강의명 좌우 인접 대조
                "No": row.get("No", len(master_rows) + 1),
                "강의상태": sale_status_str,
                "과목": row.get("과목", ""),
                "학년": row.get("학년", ""),
                "강사명": inst,
                "RMS 강좌명": rms_title,
                "PDF 강좌명": pdf_title or "-",
                "최종검수상태": status_val,
                "오류상세내역": row.get("오류 상세 내역", ""),
                "자동생성_수정요청요약": row.get("자동 생성된 수정 요청 사항 요약", ""),
                "PDF 페이지": row.get("PDF 페이지", "-"),
                "총수강생(가등록포함)": tot_stu,
                "결제완료인원": paid_stu,
                "미결제(가등록)": unpaid_stu if unpaid_stu >= 0 else 0,
                "결제율(%)": row.get("결제율(%)", 0.0),
                "판매상태": sale_status_str,
                "RMS 개강~종강(회차)": row.get("RMS 개강~종강(회차)", ""),
                "PDF 개요 개강~종강(회차)": row.get("PDF 개요 개강~종강(회차)", ""),
                "PDF 1월진도표(첫날~끝날/실제회차)": row.get("PDF 1월진도표(첫날~끝날/회차)", ""),
                "RMS 수강료": rms_fee,
                "PDF 수강료": pdf_fee,
                "RMS 수업시간": row.get("수업시간 비교", "").split(" / ")[0].replace("RMS: ", "") if " / " in str(row.get("수업시간 비교", "")) else row.get("수업시간 비교", ""),
                "PDF 수업시간": row.get("수업시간 비교", "").split(" / ")[1].replace("PDF: ", "") if " / " in str(row.get("수업시간 비교", "")) else "",
                "바자관시수 비교": row.get("바자관시수 비교", ""),
                # 3. 문장/맞춤법/기호/AI 네이밍 정보
                "맞춤법오류건수": len(sp_list),
                "맞춤법교정요약(원문->수정)": " | ".join(sp_list) if sp_list else "오류 없음",
                "기호미통일건수": len(bl_list),
                "기호통일요약(기호->■)": " | ".join(bl_list) if bl_list else "■ 통일 완료",
                "비허용부호건수(■·-·[]외)": len(dp_list),
                "비허용부호요약": " | ".join(dp_list) if dp_list else "규정 부호 준수",
                "긴문장축약건수": len(lg_list),
                "긴문장축약요약(간결형/개조식)": " // ".join(lg_list) if lg_list else "축약 대상 없음",
                "강의명판정(유지/수정추천)": ai_info.get("판정결과", "미분석"),
                "AI_임팩트점수": ai_info.get("임팩트점수", "-"),
                "AI_판정사유": ai_info.get("판정사유", "-"),
                "AI추천강의명1(직관·시그니처)": ai_info.get("추천1", "-"),
                "AI추천강의명2(킬러·임팩트)": ai_info.get("추천2", "-"),
            })

    # 판매중지 강좌 중 comparison_df에 이미 포함되지 않은 강좌만 추가 (토글로 제외된 경우)
    existing_nos = {r.get("No") for r in master_rows if r.get("No") is not None}
    existing_titles = {(str(r.get("강사명")).strip(), str(r.get("RMS 강좌명")).strip()) for r in master_rows}

    if not stopped_df.empty:
        for _, s_row in stopped_df.iterrows():
            s_no = s_row.get("No")
            s_inst = str(s_row.get("강사", "")).strip()
            s_title = str(s_row.get("강좌명", "")).strip()
            # 이미 comparison_df에 포함되어 검수된 강좌라면 중복 추가 방지 (정밀 대조 결과 보존)
            if (s_no is not None and s_no in existing_nos) or (s_inst, s_title) in existing_titles:
                continue

            master_rows.append({
                "No": s_no or (len(master_rows) + 1),
                "강의상태": "[판매중지]",
                "과목": s_row.get("과목", ""),
                "학년": s_row.get("학년", ""),
                "강사명": s_inst,
                "RMS 강좌명": s_title,
                "PDF 강좌명": "-",
                "최종검수상태": "⚪ RMS 판매중지 (검수 제외 강좌)",
                "오류상세내역": "RMS 상에서 판매중지로 등록되어 검수에서 제외되었습니다.",
                "자동생성_수정요청요약": "사전 검수 시 상단 필터 해제 필요",
                "PDF 페이지": "-",
                "총수강생(가등록포함)": s_row.get("가등록포함수강생수", 0),
                "결제완료인원": s_row.get("결제인원", 0),
                "미결제(가등록)": s_row.get("가등록(미결제)", 0),
                "결제율(%)": s_row.get("결제율(%)", 0.0),
                "판매상태": "[판매중지]",
                "RMS 개강~종강(회차)": f"{s_row.get('개강일_짧은표시', '')}~{s_row.get('종강일_짧은표시', '')} ({s_row.get('회차', 0)}회)",
                "PDF 개요 개강~종강(회차)": "-",
                "PDF 1월진도표(첫날~끝날/실제회차)": "-",
                "RMS 수강료": s_row.get("수강료", 0),
                "PDF 수강료": "-",
                "RMS 수업시간": s_row.get("간결수업시간", ""),
                "PDF 수업시간": "-",
                "바자관시수 비교": "-",
                "맞춤법오류건수": 0,
                "맞춤법교정요약(원문->수정)": "-",
                "기호미통일건수": 0,
                "기호통일요약(기호->■)": "-",
                "비허용부호건수(■·-·[]외)": 0,
                "비허용부호요약": "-",
                "긴문장축약건수": 0,
                "긴문장축약요약(간결형/개조식)": "-",
                "강의명판정(유지/수정추천)": "-",
                "AI_임팩트점수": "-",
                "AI_판정사유": "-",
                "AI추천강의명1(직관·시그니처)": "-",
                "AI추천강의명2(킬러·임팩트)": "-",
            })

    master_df = pd.DataFrame(master_rows)
    return master_df


# -------------------------------------------------------------------------
# 2. 전체 리포트 일괄 시트 (Multi-Section TSV) 생성
# -------------------------------------------------------------------------
def build_multi_section_report_tsv(
    crosscheck_result: Dict[str, Any],
    target_ym: str = "2027-01",
    snapshot_a_key: Optional[str] = None,
    snapshot_b_key: Optional[str] = None,
) -> str:
    """
    구글 스프레드시트 A1 셀에 'Ctrl+V' 한 번으로 모든 정보를 위에서 아래로
    섹션별 구분선과 함께 일괄 붙여넣을 수 있는 Multi-Section TSV 문자열을 생성합니다.
    """
    rms_sum = crosscheck_result.get("rms_summary", {}) or {}
    all_in_one_df = build_all_in_one_master_df(crosscheck_result, target_ym=target_ym)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 스냅샷 비교 데이터
    avail_snaps = get_available_snapshots()
    norm_target = normalize_snapshot_key(target_ym)
    snap_b = snapshot_b_key or (norm_target if norm_target in avail_snaps else (avail_snaps[-1] if avail_snaps else norm_target))
    snap_a = snapshot_a_key
    if not snap_a:
        if snap_b in avail_snaps:
            b_idx = avail_snaps.index(snap_b)
            snap_a = avail_snaps[b_idx - 1] if b_idx > 0 else snap_b
        elif len(avail_snaps) >= 2:
            snap_a = avail_snaps[-2]
        else:
            snap_a = snap_b

    comp_res = compare_two_snapshots(snap_a, snap_b)
    kpi = comp_res.get("kpi", {})
    comp_snap_df = comp_res.get("comparison_df", pd.DataFrame())

    tsv_lines: List[str] = []

    def _df_to_tsv(df: pd.DataFrame) -> None:
        if df is None or df.empty:
            tsv_lines.append("(해당 내역 없음)")
            return
        # 헤더
        cols = list(df.columns)
        tsv_lines.append("\t".join(cols))
        for _, r in df.iterrows():
            row_vals = []
            for c in cols:
                val = r.get(c, "")
                clean_v = re.sub(r"[\t\r\n]+", " ", str(val) if pd.notna(val) else "").strip()
                row_vals.append(clean_v)
            tsv_lines.append("\t".join(row_vals))

    # =========================================================================
    # [섹션 1] 전체 요약 헤드라인
    # =========================================================================
    tsv_lines.append("=" * 70)
    tsv_lines.append("📋 [러셀대치 RMS ↔ 강의계획서 PDF 종합 검수 및 수강생 통계 전체 리포트]")
    tsv_lines.append(f"생성일시: {now_str}\t검수 기준 연월: {target_ym}\t스냅샷 비교: {kpi.get('label_a', snap_a)} vs {kpi.get('label_b', snap_b)}")
    tsv_lines.append(
        f"총 강좌수: {rms_sum.get('total_courses', 0)}건\t"
        f"총 수강생: {rms_sum.get('total_students', 0):,}명\t"
        f"결제완료 인원: {rms_sum.get('paid_students', 0):,}명\t"
        f"미결제 가등록: {rms_sum.get('unpaid_students', 0):,}명\t"
        f"수강생 결제율: {rms_sum.get('payment_rate', 0.0)}%"
    )
    tsv_lines.append("=" * 70)
    tsv_lines.append("")

    # =========================================================================
    # [섹션 2] 강사별 매월 20일 기준 비교 통계표
    # =========================================================================
    tsv_lines.append("■ [섹션 2] 강사별 매월 20일 기준 1:1 정밀 비교 통계 (Full Outer Join)")
    tsv_lines.append(
        f"비교 요약:\t강좌수 증감: {'+' if kpi.get('courses_diff', 0) > 0 else ''}{kpi.get('courses_diff', 0)}건\t"
        f"결제완료 증감: {'+' if kpi.get('paid_diff', 0) > 0 else ''}{kpi.get('paid_diff', 0):,}명 ({'+' if kpi.get('paid_pct', 0) > 0 else ''}{kpi.get('paid_pct', 0)}%)\t"
        f"총수강생 증감: {'+' if kpi.get('students_diff', 0) > 0 else ''}{kpi.get('students_diff', 0):,}명\t"
        f"결제율 변화: {'+' if kpi.get('rate_diff', 0) > 0 else ''}{kpi.get('rate_diff', 0)}%p"
    )
    _df_to_tsv(comp_snap_df)
    tsv_lines.append("")

    # =========================================================================
    # [섹션 3] 전 강좌 올인원(All-in-One) 정밀 검수 및 AI 추천 강의명 통합표
    # =========================================================================
    tsv_lines.append("■ [섹션 3] 전 강좌 올인원(All-in-One) 정밀 검수 및 AI 추천 강의명 통합표")
    _df_to_tsv(all_in_one_df)
    tsv_lines.append("")

    # =========================================================================
    # [섹션 4] 강좌별 국어 맞춤법 교정 및 긴 문장 축약 상세 내역
    # =========================================================================
    tsv_lines.append("■ [섹션 4-1] 맞춤법·기호(■)통일 & 문장축약 통합 검수 내역")
    comb_proof_df = crosscheck_result.get("combined_proofread_df", pd.DataFrame())
    if comb_proof_df is None or comb_proof_df.empty:
        comb_proof_df = crosscheck_result.get("spelling_df", pd.DataFrame())
    _df_to_tsv(comb_proof_df)
    tsv_lines.append("")

    tsv_lines.append("■ [섹션 4-2] 긴 문장 간략화 (서술형/개조식 축약) 상세 내역")
    long_sentence_df = crosscheck_result.get("long_sentence_df", pd.DataFrame())
    _df_to_tsv(long_sentence_df)
    tsv_lines.append("")

    return "\n".join(tsv_lines)


# -------------------------------------------------------------------------
# 3. 브라우저 JavaScript 기반 클립보드 원클릭 복사 & sheets.new 자동 오픈 컴포넌트
# -------------------------------------------------------------------------
def render_google_sheets_copy_button(
    tsv_text: str,
    button_label: str = "🚀 모든 정보 복사 & 새 구글 시트 열기",
    height: int = 58,
) -> None:
    """
    Streamlit Cloud(리눅스 컨테이너) 환경에서도 100% 작동하도록
    브라우저 JavaScript(navigator.clipboard.writeText)를 통해 클립보드에 복사하고
    새 탭으로 https://sheets.new 를 띄워 즉시 Ctrl+V 할 수 있게 합니다.
    """
    safe_tsv = json.dumps(tsv_text)
    html_code = f"""
    <div style="display:flex; width:100%; align-items:center;">
      <button id="gs-copy-btn" style="
          background: linear-gradient(90deg, #0F9D58 0%, #16A34A 100%);
          color: white;
          border: none;
          padding: 11px 20px;
          border-radius: 8px;
          font-weight: 700;
          font-size: 14px;
          cursor: pointer;
          width: 100%;
          box-shadow: 0 2px 4px rgba(0,0,0,0.12);
          transition: all 0.2s ease-in-out;
      ">
        {button_label}
      </button>
    </div>
    <script>
      const btn = document.getElementById('gs-copy-btn');
      const data = {safe_tsv};
      btn.addEventListener('click', async () => {{
        try {{
          await navigator.clipboard.writeText(data);
          btn.innerText = '✅ 모든 정보 복사 완료! 새 구글 시트 A1 셀에서 Ctrl+V 하세요';
          btn.style.background = '#1a73e8';
          window.open('https://sheets.new', '_blank');
          setTimeout(() => {{
            btn.innerText = '{button_label}';
            btn.style.background = 'linear-gradient(90deg, #0F9D58 0%, #16A34A 100%)';
          }}, 5000);
        }} catch (err) {{
          prompt('클립보드 자동 복사에 실패했습니다. 아래 텍스트를 Ctrl+C 후 구글시트에 붙여넣으세요:', data.substring(0, 500));
          window.open('https://sheets.new', '_blank');
        }}
      }});
      btn.addEventListener('mouseover', () => {{
        btn.style.opacity = '0.92';
      }});
      btn.addEventListener('mouseout', () => {{
        btn.style.opacity = '1.0';
      }});
    </script>
    """
    components.html(html_code, height=height)


# -------------------------------------------------------------------------
# 4. 구글 시트용 CSV 및 전체 시트 ZIP 압축 패키지 생성
# -------------------------------------------------------------------------
def export_all_in_one_csv_bytes(all_in_one_df: pd.DataFrame) -> bytes:
    """구글 시트에서 한글 깨짐 없이 바로 열리는 UTF-8-SIG CSV 바이트를 생성합니다."""
    if all_in_one_df is None or all_in_one_df.empty:
        return b"\xef\xbb\xbf\xeb\x8d\xb0\xec\x9d\xb4\xed\x84\xb0\xea\xb0\x80 \xec\x97\x86\xec\x8a\xb5\xeb\x8b\x88\xeb\x8b\xa4."
    return all_in_one_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def export_all_sheets_zip_bytes(
    crosscheck_result: Dict[str, Any],
    target_ym: str = "2027-01",
    snapshot_a_key: Optional[str] = None,
    snapshot_b_key: Optional[str] = None,
) -> bytes:
    """
    모든 개별 시트(올인원, 비교통계, 연간누적추이, 검수결과, 오류목록, 맞춤법, 강의명 등)를
    각각 UTF-8-SIG CSV 파일로 담은 ZIP 압축 패키지 바이트를 생성합니다.
    """
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        # 1. 올인원 마스터
        master_df = build_all_in_one_master_df(crosscheck_result, target_ym=target_ym)
        zf.writestr("01_올인원_전체통합마스터.csv", master_df.to_csv(index=False, encoding="utf-8-sig"))

        # 2. 매월 20일 기준 비교
        avail_snaps = get_available_snapshots()
        norm_target = normalize_snapshot_key(target_ym)
        snap_b = snapshot_b_key or (norm_target if norm_target in avail_snaps else (avail_snaps[-1] if avail_snaps else norm_target))
        snap_a = snapshot_a_key or (avail_snaps[-2] if len(avail_snaps) >= 2 else snap_b)
        comp_res = compare_two_snapshots(snap_a, snap_b)
        comp_snap_df = comp_res.get("comparison_df", pd.DataFrame())
        zf.writestr("02_매월20일_강사별비교통계.csv", comp_snap_df.to_csv(index=False, encoding="utf-8-sig"))

        # 3. 연간 누적 추이
        pivot_df = build_snapshot_pivot_table("결제완료 인원")
        zf.writestr("03_매월20일_연간누적추이.csv", pivot_df.to_csv(index=False, encoding="utf-8-sig"))

        # 4. 당월 상세 순위표
        inst_df = crosscheck_result.get("instructor_summary_df", pd.DataFrame())
        zf.writestr("04_당월20일_강사별상세순위.csv", inst_df.to_csv(index=False, encoding="utf-8-sig"))

        # 5. 전체 검수 결과 유효강좌
        comp_df = crosscheck_result.get("comparison_df", pd.DataFrame())
        zf.writestr("05_전체검수결과_유효강좌.csv", comp_df.to_csv(index=False, encoding="utf-8-sig"))

        # 6. 오류 및 수정 필요 강좌
        err_df = crosscheck_result.get("error_df", pd.DataFrame())
        zf.writestr("06_오류및수정필요강좌.csv", err_df.to_csv(index=False, encoding="utf-8-sig"))

        # 7. 맞춤법 및 기호통일, 긴문장 축약
        comb_df = crosscheck_result.get("combined_proofread_df", pd.DataFrame())
        if comb_df is not None and not comb_df.empty:
            zf.writestr("07_맞춤법_기호통일_문장축약.csv", comb_df.to_csv(index=False, encoding="utf-8-sig"))
        else:
            spell_df = crosscheck_result.get("spelling_df", pd.DataFrame())
            zf.writestr("07_맞춤법교정내역.csv", spell_df.to_csv(index=False, encoding="utf-8-sig"))

        long_df = crosscheck_result.get("long_sentence_df", pd.DataFrame())
        zf.writestr("08_긴문장축약내역.csv", long_df.to_csv(index=False, encoding="utf-8-sig"))

        # 8. AI 강의명 추천
        ai_df = crosscheck_result.get("ai_titles_df", pd.DataFrame())
        zf.writestr("09_AI추천강의명.csv", ai_df.to_csv(index=False, encoding="utf-8-sig"))

    return bio.getvalue()


# -------------------------------------------------------------------------
# 5. Google Apps Script Web App 자동 전송 (JSON 키 불필요)
# -------------------------------------------------------------------------
APPS_SCRIPT_SAMPLE_CODE = """
// [구글 스프레드시트 -> 확장 프로그램 -> Apps Script에 붙여넣기]
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // tabs: { "시트명": [ [헤더1, 헤더2], [값1, 값2], ... ] }
    var tabs = data.tabs || {};
    for (var tabName in tabs) {
      var rows = tabs[tabName];
      if (!rows || rows.length === 0) continue;
      
      var sheet = ss.getSheetByName(tabName);
      if (!sheet) {
        sheet = ss.insertSheet(tabName);
      }
      sheet.clear();
      
      var numRows = rows.length;
      var numCols = rows[0].length;
      sheet.getRange(1, 1, numRows, numCols).setValues(rows);
      sheet.getRange(1, 1, 1, numCols).setBackground("#1F4E78").setFontColor("#FFFFFF").setFontWeight("bold");
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", updated_sheets: Object.keys(tabs) }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
""".strip()


def send_all_tabs_to_apps_script(
    webhook_url: str,
    crosscheck_result: Dict[str, Any],
    target_ym: str = "2027-01",
    snapshot_a_key: Optional[str] = None,
    snapshot_b_key: Optional[str] = None,
    timeout_sec: int = 20,
) -> Tuple[bool, str]:
    """
    Google Apps Script 웹앱 URL로 올인원 마스터, 비교통계, 검수결과 등 전 탭 데이터를
    JSON 포맷으로 직접 전송하여 구글 시트 탭들을 자동 생성/덮어쓰기합니다.
    """
    url = str(webhook_url).strip()
    if not url.startswith("http"):
        return False, "유효한 Apps Script 웹앱 URL(https://script.google.com/macros/s/...)을 입력하세요."

    try:
        master_df = build_all_in_one_master_df(crosscheck_result, target_ym=target_ym)

        avail_snaps = get_available_snapshots()
        norm_target = normalize_snapshot_key(target_ym)
        snap_b = snapshot_b_key or (norm_target if norm_target in avail_snaps else (avail_snaps[-1] if avail_snaps else norm_target))
        snap_a = snapshot_a_key or (avail_snaps[-2] if len(avail_snaps) >= 2 else snap_b)
        comp_res = compare_two_snapshots(snap_a, snap_b)
        comp_snap_df = comp_res.get("comparison_df", pd.DataFrame())

        comparison_df = crosscheck_result.get("comparison_df", pd.DataFrame())
        err_df = crosscheck_result.get("error_df", pd.DataFrame())
        ai_df = crosscheck_result.get("ai_titles_df", pd.DataFrame())

        def _df_to_2d_list(df: pd.DataFrame) -> List[List[Any]]:
            if df is None or df.empty:
                return []
            headers = [str(c) for c in df.columns]
            rows = [headers]
            for _, r in df.iterrows():
                row_vals = []
                for c in df.columns:
                    val = r.get(c, "")
                    row_vals.append("" if pd.isna(val) else str(val))
                rows.append(row_vals)
            return rows

        tabs_data = {
            "올인원_전체통합마스터": _df_to_2d_list(master_df),
            "매월20일_강사별비교": _df_to_2d_list(comp_snap_df),
            "전체검수결과": _df_to_2d_list(comparison_df),
            "오류및수정요청": _df_to_2d_list(err_df),
            "AI_추천강의명": _df_to_2d_list(ai_df),
        }

        payload = {
            "target_ym": target_ym,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "tabs": tabs_data,
        }

        resp = requests.post(url, json=payload, timeout=timeout_sec)
        if resp.status_code == 200:
            res_json = resp.json()
            if res_json.get("status") == "success":
                up_sheets = ", ".join(res_json.get("updated_sheets", []))
                return True, f"구글 시트에 성공적으로 전송되었습니다! 갱신된 시트: [{up_sheets}]"
            return False, f"Apps Script 처리 오류: {res_json.get('message')}"
        return False, f"HTTP 응답 오류 (상태 코드: {resp.status_code})"
    except Exception as e:
        return False, f"구글 시트 전송 중 오류 발생: {e}"

