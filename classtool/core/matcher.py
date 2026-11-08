"""
핵심 매칭 및 RMS 엑셀 ↔ 강의계획서 PDF 7대 정밀 검수 모듈 (core/matcher.py)
1. 강사명 정규화 및 매핑 사전 관리 (config/instructor_mapping.json)
2. 강좌 자동 매칭 알고리즘 (정규화 강사명 + 요일/시간 + 강좌명/슬로건 유사도 + 대상학년 조합)
3. 7대 핵심 검수 규칙:
   ① 수강료 & 회차당 단가 검증
   ② 회차 3중 검증 (RMS 회차 vs PDF 개요 회차 vs PDF 1월 진도표 날짜 수)
   ③ 개강일 & 종강일 3중 검증 (RMS vs PDF 개요 vs PDF 진도표 첫날/끝날 + 개강일 요일 모순)
   ④ 진도표 내 비정상 요일 섞임 검증 (수업요일 vs 실제 달력 요일 datetime 대조)
   ⑤ 수업 요일/시간 및 바자관 시수(6T, 3T) 검증
   ⑥ 과목 및 강좌명(슬로건 포함) 비교
   ⑦ PDF 미수록(누락) 강좌 및 RMS 미개설 강좌 분리
4. 4개 시트 구성 서식 포함 엑셀(.xlsx) 생성, UTF-8-SIG CSV/TSV 생성, 디자인팀/강사용 수정 안내문 자동 생성
"""

from __future__ import annotations

import io
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

import pandas as pd
from rapidfuzz import fuzz

from parsers.rms_parser import (
    extract_rms_summary,
    format_short_md_with_weekday,
    generate_instructor_student_summary,
    get_korean_weekday,
    normalize_baza_hours,
    strip_course_brackets,
)
from core.text_improver import analyze_course_text_content

DEFAULT_MAPPING_LIST: List[Dict[str, str]] = [
    {
        "rms_name": "리아",
        "sheet_name": "박소현(리아)",
        "canonical": "리아",
        "memo": "박소현(리아) ↔ 리아 (영어)",
    },
    {
        "rms_name": "박소현(셀리나)",
        "sheet_name": "박소현",
        "canonical": "박소현(셀리나)",
        "memo": "박소현 ↔ 박소현(셀리나) (영어)",
    },
    {
        "rms_name": "민정",
        "sheet_name": "김민정",
        "canonical": "민정",
        "memo": "김민정 ↔ 민정 (사회탐구)",
    },
    {
        "rms_name": "이진리",
        "sheet_name": "이지원",
        "canonical": "이진리",
        "memo": "이지원 ↔ 이진리 (수학)",
    },
    {
        "rms_name": "현자의 돌",
        "sheet_name": "임수민",
        "canonical": "현자의 돌",
        "memo": "임수민 ↔ 현자의 돌 (통합사회)",
    },
    {
        "rms_name": "이현호(미친국어)",
        "sheet_name": "이현호",
        "canonical": "이현호(미친국어)",
        "memo": "이현호 ↔ 이현호(미친국어) (국어)",
    },
]


class InstructorMapper:
    """강사명 정규화 및 양방향 매핑 사전 관리 클래스"""

    def __init__(self, config_path: Union[str, Path] = "config/instructor_mapping.json"):
        self.config_path = Path(config_path)
        self.mappings: List[Dict[str, str]] = []
        self.lookup: Dict[str, str] = {}
        self.protected_names: Set[str] = set()
        self.load()

    def load(self) -> List[Dict[str, str]]:
        if not self.config_path.exists():
            self.mappings = [dict(item) for item in DEFAULT_MAPPING_LIST]
            self.save(self.mappings)
        else:
            try:
                raw_text = self.config_path.read_text(encoding="utf-8-sig")
                data = json.loads(raw_text)
                if isinstance(data, dict) and "mappings" in data and isinstance(data["mappings"], list):
                    self.mappings = data["mappings"]
                elif isinstance(data, list):
                    self.mappings = data
                elif isinstance(data, dict):
                    converted = []
                    for k, v in data.items():
                        converted.append(
                            {
                                "rms_name": str(v).strip(),
                                "sheet_name": str(k).strip(),
                                "canonical": str(v).strip(),
                                "memo": f"{k} ↔ {v}",
                            }
                        )
                    self.mappings = converted
                else:
                    self.mappings = [dict(item) for item in DEFAULT_MAPPING_LIST]
            except Exception:
                self.mappings = [dict(item) for item in DEFAULT_MAPPING_LIST]

        self._rebuild_lookup()
        return self.mappings

    def _rebuild_lookup(self) -> None:
        self.lookup = {}
        self.protected_names = set()
        for item in self.mappings:
            rms_n = str(item.get("rms_name", "")).strip()
            sht_n = str(item.get("sheet_name", "")).strip()
            canon = str(item.get("canonical", "")).strip() or rms_n or sht_n
            if not canon:
                continue
            if rms_n:
                self.lookup[rms_n] = canon
                self.lookup[re.sub(r"\s+", "", rms_n)] = canon
                self.protected_names.add(rms_n)
            if sht_n:
                self.lookup[sht_n] = canon
                self.lookup[re.sub(r"\s+", "", sht_n)] = canon
                self.protected_names.add(sht_n)
            self.lookup[canon] = canon
            self.lookup[re.sub(r"\s+", "", canon)] = canon

    def save(self, new_mappings: List[Dict[str, str]]) -> None:
        cleaned_list: List[Dict[str, str]] = []
        for item in new_mappings:
            rms_n = str(item.get("rms_name", "") or "").strip()
            sht_n = str(item.get("sheet_name", "") or "").strip()
            canon = str(item.get("canonical", "") or "").strip() or rms_n or sht_n
            memo = str(item.get("memo", "") or "").strip()
            if not rms_n and not sht_n:
                continue
            if not memo and rms_n and sht_n:
                memo = f"{sht_n} ↔ {rms_n}"
            cleaned_list.append(
                {
                    "rms_name": rms_n,
                    "sheet_name": sht_n,
                    "canonical": canon,
                    "memo": memo,
                }
            )

        self.mappings = cleaned_list
        self._rebuild_lookup()
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self.config_path.write_text(
            json.dumps({"mappings": self.mappings}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def normalize(self, raw_name: Any) -> str:
        if pd.isna(raw_name) or raw_name is None:
            return ""
        s = str(raw_name).strip()
        if not s:
            return ""

        s_clean = re.sub(r"\s*(선생님|강사님|강사|교수님|쌤)$", "", s).strip()
        if re.match(r"^[가-힣()]+T$", s_clean, flags=re.IGNORECASE):
            s_clean = s_clean[:-1].strip()

        if s_clean in self.lookup:
            return self.lookup[s_clean]
        s_nospace = re.sub(r"\s+", "", s_clean)
        if s_nospace in self.lookup:
            return self.lookup[s_nospace]

        return s_clean

    def match_instructors(self, name_a: Any, name_b: Any) -> bool:
        norm_a = self.normalize(name_a)
        norm_b = self.normalize(name_b)
        if not norm_a or not norm_b:
            return False
        if norm_a == norm_b:
            return True
        if re.sub(r"\s+", "", norm_a) == re.sub(r"\s+", "", norm_b):
            return True

        parts_a = [self.normalize(p) for p in re.split(r"[,/·&]", str(name_a)) if p.strip()]
        parts_b = [self.normalize(p) for p in re.split(r"[,/·&]", str(name_b)) if p.strip()]
        if len(parts_a) > 1 or len(parts_b) > 1:
            if set(parts_a) & set(parts_b):
                return True

        if "박소현" in str(name_a) or "박소현" in str(name_b) or "리아" in (norm_a, norm_b):
            return False

        def _extract_tokens(n_str: str) -> Set[str]:
            tokens = {n_str.strip()}
            m = re.match(r"^([^\(]+)\(([^\)]+)\)$", n_str.strip())
            if m:
                tokens.add(m.group(1).strip())
                tokens.add(m.group(2).strip())
            return {t for t in tokens if len(t) >= 2}

        tokens_a = _extract_tokens(norm_a)
        tokens_b = _extract_tokens(norm_b)
        return bool(tokens_a & tokens_b)


def compute_title_similarity(title_a: str, title_b: str) -> float:
    """두 강좌명(대괄호/꺾쇠 태그 제외) 간의 유사도(0.0 ~ 100.0)를 계산합니다."""
    t1, _ = strip_course_brackets(title_a)
    t2, _ = strip_course_brackets(title_b)
    if not t1 or not t2:
        return 0.0
    c1 = re.sub(r"[:\-_/,·~<>《》「」『』()\[\]]+", " ", t1)
    c2 = re.sub(r"[:\-_/,·~<>《》「」『』()\[\]]+", " ", t2)
    c1 = re.sub(r"\s+", " ", c1).strip()
    c2 = re.sub(r"\s+", " ", c2).strip()
    if not c1 or not c2:
        return 0.0
    score_token = float(fuzz.token_set_ratio(c1, c2))
    score_partial = float(fuzz.partial_ratio(c1, c2))
    return round(max(score_token, score_partial * 0.9), 1)


def verify_rms_vs_pdf_course(
    rms_row: pd.Series,
    pdf_row: pd.Series,
    target_ym: str = "2027-01",
    title_similarity_threshold: float = 55.0,
) -> Dict[str, Any]:
    """
    매칭된 RMS 강좌와 PDF 강좌 간의 7대 핵심 검수 규칙을 수행합니다:
    1. 수강료 & 회차당 단가 검증
    2. 회차 3중 검증 (RMS 회차 vs PDF 개요 회차 vs PDF 1월 진도표 날짜 수)
    3. 개강일 & 종강일 3중 검증 (RMS vs PDF 개요 vs PDF 진도표 첫날/끝날 + 개강일 요일 모순)
    4. 진도표 내 비정상 요일 섞임 검증
    5. 수업 요일/시간 및 바자관 시수 검증
    6. 과목 및 강좌명 비교
    """
    r_subj = str(rms_row.get("과목", "")).strip()
    r_area = str(rms_row.get("영역", "")).strip()
    r_grade = str(rms_row.get("학년", "")).strip()
    r_title = str(rms_row.get("강좌명", "")).strip()
    r_pure_title = str(rms_row.get("순수강좌명", "") or r_title).strip()
    r_tags = str(rms_row.get("말머리태그", "")).strip()

    r_open = str(rms_row.get("개강일", "")).strip()
    r_open_short = str(rms_row.get("개강일_짧은표시", "") or format_short_md_with_weekday(r_open)).strip()
    r_open_wk = str(rms_row.get("개강일_요일", "") or get_korean_weekday(r_open)).strip()

    r_close = str(rms_row.get("종강일", "")).strip()
    r_close_short = str(rms_row.get("종강일_짧은표시", "") or format_short_md_with_weekday(r_close)).strip()

    r_sess = int(rms_row.get("회차", 0) or 0)
    r_fee = int(rms_row.get("수강료", 0) or 0)
    r_unit_fee = int(rms_row.get("1회차 기준 수강료", 0) or 0)
    if r_unit_fee == 0 and r_fee > 0 and r_sess > 0:
        r_unit_fee = int(round(r_fee / r_sess))

    r_days_list: List[str] = list(rms_row.get("요일_리스트", []) or [])
    r_days_str = str(rms_row.get("수업요일", "")).strip()
    r_days_dot = str(rms_row.get("수업요일_점표시", "") or "·".join(r_days_list)).strip()
    r_start = str(rms_row.get("전체시작시간", "") or rms_row.get("시작시간", "")).strip()
    r_full_keys = list(rms_row.get("요일전체시간키_리스트", []) or [])
    r_compact_sched = str(rms_row.get("간결수업시간", "") or rms_row.get("표준수업시간", "")).strip()
    r_baza = normalize_baza_hours(rms_row.get("바자관시수", rms_row.get("바자관 인정기준", "-")))

    # PDF 추출 데이터
    p_subj = str(pdf_row.get("과목", "")).strip()
    p_grade = str(pdf_row.get("대상학년", "")).strip()
    p_title = str(pdf_row.get("강좌명", "")).strip()
    p_pure_title = str(pdf_row.get("순수강좌명", "") or p_title).strip()
    p_slogan = str(pdf_row.get("슬로건", "")).strip()

    p_open = str(pdf_row.get("개강일", "")).strip()
    p_open_short = format_short_md_with_weekday(p_open) if p_open else str(pdf_row.get("개강일_원본", "-")).strip()
    p_open_wk = str(pdf_row.get("개강일_실제요일", "") or get_korean_weekday(p_open)).strip()
    p_open_printed_wk = str(pdf_row.get("개강일_표기요일", "")).strip()

    p_close = str(pdf_row.get("종강일", "")).strip()
    p_close_short = format_short_md_with_weekday(p_close) if p_close else str(pdf_row.get("종강일_원본", "-")).strip()

    p_sess = int(pdf_row.get("회차", 0) or 0)
    p_fee = int(pdf_row.get("수강료", 0) or 0)
    p_unit_fee = int(pdf_row.get("회차당수강료", 0) or 0)

    p_days_list: List[str] = list(pdf_row.get("요일_리스트", []) or [])
    p_days_str = str(pdf_row.get("수업요일", "")).strip()
    p_days_dot = str(pdf_row.get("수업요일_점표시", "") or "·".join(p_days_list)).strip()
    p_start = str(pdf_row.get("전체시작시간", "") or pdf_row.get("시작시간", "")).strip()
    p_full_keys = list(pdf_row.get("요일전체시간키_리스트", []) or [])
    p_compact_sched = str(pdf_row.get("간결수업시간", "") or pdf_row.get("표준수업시간", "")).strip()
    p_baza = normalize_baza_hours(pdf_row.get("바자관시수", "-"))

    # PDF 1월 진도표 역산 데이터
    jan_sess = int(pdf_row.get("진도표_1월_실제회차", 0) or 0)
    jan_first = str(pdf_row.get("진도표_첫수업일", "")).strip()
    jan_first_short = str(pdf_row.get("진도표_첫수업일_표시", "-")).strip()
    jan_last = str(pdf_row.get("진도표_마지막수업일", "")).strip()
    jan_last_short = str(pdf_row.get("진도표_마지막수업일_표시", "-")).strip()
    spilled_jan_list: List[str] = list(pdf_row.get("진도표_2월표_1월날짜리스트", []) or [])
    jan_wk_err = bool(pdf_row.get("진도표_요일오류여부", False))
    jan_wk_err_detail = str(pdf_row.get("진도표_요일오류상세", "")).strip()

    error_tags: List[str] = []
    error_details: List[str] = []
    fix_requests: List[str] = []
    info_notes: List[str] = []
    error_columns: Set[str] = set()  # 엑셀 조건부 셀 강조용

    effective_days_list = p_days_list if p_days_list else r_days_list
    effective_days_dot = p_days_dot if p_days_dot else r_days_dot

    # ------------------------------------------------------------------
    # 규칙 1: 수강료 & 회차당 단가 검증
    # ------------------------------------------------------------------
    fee_mismatch = False
    if r_fee > 0 and p_fee > 0 and r_fee != p_fee:
        fee_mismatch = True
        error_tags.append("🔴 [수강료 불일치]")
        r_unit_str = f" (회차당 {r_unit_fee:,}원)" if r_unit_fee > 0 else ""
        p_unit_str = f" (회차당 {p_unit_fee:,}원)" if p_unit_fee > 0 else ""
        error_details.append(f"수강료 불일치: RMS {r_fee:,}원{r_unit_str} ≠ PDF {p_fee:,}원{p_unit_str}")
        fix_requests.append(
            f"수강료 {p_fee:,}원{p_unit_str} → {r_fee:,}원{r_unit_str}으로 수정 필요"
        )
        error_columns.update(["RMS 수강료", "PDF 수강료"])

    if p_fee > 0 and p_sess > 0 and p_unit_fee > 0 and (p_sess * p_unit_fee != p_fee):
        error_tags.append("🔴 [단가계산 오류]")
        calc_total = p_sess * p_unit_fee
        error_details.append(
            f"PDF 수강료 자가모순: {p_sess}회 × {p_unit_fee:,}원 = {calc_total:,}원 ≠ 표기 총액 {p_fee:,}원"
        )
        if not fee_mismatch:
            fix_requests.append(f"PDF 수강료 총액({p_fee:,}원)과 회차×단가({calc_total:,}원) 계산 불일치 수정 필요")
        error_columns.add("PDF 수강료")

    if not fee_mismatch and r_unit_fee > 0 and p_unit_fee > 0 and r_unit_fee != p_unit_fee:
        error_tags.append("🔴 [회차당단가 불일치]")
        error_details.append(f"회차당 단가 불일치: RMS {r_unit_fee:,}원 ≠ PDF {p_unit_fee:,}원")
        fix_requests.append(f"회차당 수강료 {p_unit_fee:,}원 → {r_unit_fee:,}원으로 수정 필요")
        error_columns.update(["RMS 수강료", "PDF 수강료"])

    # ------------------------------------------------------------------
    # 규칙 2: 회차 3중 검증 (RMS 회차 vs PDF 개요 회차 vs PDF 1월 진도표 날짜 수)
    # ------------------------------------------------------------------
    if r_sess > 0 and p_sess > 0 and r_sess != p_sess:
        error_tags.append("🔴 [회차 불일치]")
        error_details.append(
            f"회차 3중 불일치: RMS {r_sess}회 ≠ PDF 개요 {p_sess}회 (PDF 1월 진도표 {jan_sess}줄)"
        )
        fix_requests.append(
            f"PDF 강의개요 수강기간 회차를 {p_sess}회 → {r_sess}회로 수정 필요 (1월 진도표는 {jan_sess}줄)"
        )
        error_columns.update(["RMS 개강~종강(회차)", "PDF 개요 개강~종강(회차)", "PDF 1월진도표(첫날~끝날/회차)"])
    elif jan_sess > 0 and ((p_sess > 0 and jan_sess != p_sess) or (r_sess > 0 and jan_sess != r_sess)):
        error_tags.append("🟡 [진도표 회차 불일치]")
        if spilled_jan_list:
            spilled_str = ", ".join(spilled_jan_list)
            error_details.append(
                f"진도표 회차 불일치: 개요/RMS는 {p_sess}회이나 1월 진도표는 {jan_sess}줄 (1월 수업일 '{spilled_str}'이 2월 과정 표로 넘어감)"
            )
            fix_requests.append(
                f"PDF '2월 과정' 진도표 첫 줄에 들어간 '{spilled_str}' 수업을 '1월 과정' 표로 이동하여 1월 {r_sess}회 맞춤 필요"
            )
        else:
            error_details.append(
                f"진도표 회차 불일치: 개요/RMS {p_sess}회 ≠ PDF 1월 진도표 {jan_sess}줄"
            )
            fix_requests.append(
                f"PDF 1월 진도표 수업일 수({jan_sess}회)를 강의개요/RMS 회차({r_sess}회)에 맞게 수정 필요"
            )
        error_columns.update(["PDF 개요 개강~종강(회차)", "PDF 1월진도표(첫날~끝날/회차)"])

    # ------------------------------------------------------------------
    # 규칙 3: 개강일 & 종강일 3중 검증 (RMS vs PDF 개요 vs PDF 진도표 첫날/끝날)
    # ------------------------------------------------------------------
    # 3-1) 개강일 비교 및 개강일 요일 모순 검증
    open_rms_pdf_diff = bool(r_open and p_open and r_open != p_open)
    open_weekday_contradiction = bool(
        effective_days_list and p_open_wk and (p_open_wk not in effective_days_list)
    )
    open_printed_wk_contradiction = bool(
        p_open_printed_wk and p_open_wk and (p_open_printed_wk != p_open_wk)
    )
    open_plan_diff = bool(jan_first and p_open and (p_open != jan_first))

    if open_rms_pdf_diff:
        error_tags.append("🔴 [개강일 불일치]")
        error_details.append(
            f"개강일 불일치: RMS {r_open_short} ≠ PDF 개요 {p_open_short} (1월 진도표 첫날: {jan_first_short})"
        )
        fix_requests.append(
            f"개강일을 {p_open_short} → {r_open_short}(진도표 첫날 {jan_first_short})로 수정 필요"
        )
        error_columns.update(["RMS 개강~종강(회차)", "PDF 개요 개강~종강(회차)"])
    elif open_weekday_contradiction or open_plan_diff or open_printed_wk_contradiction:
        error_tags.append("🟡 [개강일 요일 모순]")
        if open_weekday_contradiction and open_plan_diff:
            error_details.append(
                f"개강일 요일 모순: 수업시간은 '{effective_days_dot}'이고 1월 진도표 첫날은 {jan_first_short}인데, "
                f"RMS와 PDF 개요 개강일이 '{p_open_short}'({p_open_wk}요일)로 오기재됨"
            )
            fix_requests.append(
                f"RMS 및 PDF 강의개요 개강일(수강기간 시작일)을 {p_open_short} → {jan_first_short}로 수정 필요 ({effective_days_dot} 수업)"
            )
        elif open_weekday_contradiction:
            error_details.append(
                f"개강일 요일 모순: 수업요일({effective_days_dot})과 개강일 요일({p_open_short})이 일치하지 않음"
            )
            fix_requests.append(
                f"개강일({p_open_short})이 수업요일({effective_days_dot})과 맞지 않으므로 수정 필요"
            )
        elif open_printed_wk_contradiction:
            error_details.append(
                f"개강일 요일 표기 오타: PDF에 '{p_open}({p_open_printed_wk})'로 표기되었으나 실제 달력은 '{p_open_wk}요일'임"
            )
            fix_requests.append(
                f"PDF 개강일 요일 표기를 ({p_open_printed_wk}) → ({p_open_wk})로 수정 필요"
            )
        else:
            error_details.append(
                f"개강일-진도표 첫날 불일치: 개요 개강일 {p_open_short} ≠ 1월 진도표 첫날 {jan_first_short}"
            )
            fix_requests.append(
                f"개강일({p_open_short})과 1월 진도표 첫 수업일({jan_first_short}) 일치 확인 필요"
            )
        error_columns.update(["RMS 개강~종강(회차)", "PDF 개요 개강~종강(회차)", "PDF 1월진도표(첫날~끝날/회차)"])

    # 3-2) 종강일 3중 비교 (RMS 종강일 vs PDF 개요 종강일 vs PDF 1월 진도표 마지막날)
    close_rms_pdf_diff = bool(r_close and p_close and r_close != p_close)
    close_plan_diff = bool(jan_last and p_close and (p_close != jan_last))

    if close_rms_pdf_diff:
        error_tags.append("🔴 [종강일 불일치]")
        error_details.append(
            f"종강일 3중 불일치: RMS {r_close_short} ≠ PDF 개요 {p_close_short} ≠ 1월 진도표 마지막날 {jan_last_short}"
        )
        fix_requests.append(
            f"PDF 강의개요 종강일을 {p_close_short} → {r_close_short}로 수정 필요"
        )
        error_columns.update(["RMS 개강~종강(회차)", "PDF 개요 개강~종강(회차)", "PDF 1월진도표(첫날~끝날/회차)"])
    elif close_plan_diff and not spilled_jan_list:
        error_tags.append("🟡 [진도표 종강일 불일치]")
        error_details.append(
            f"진도표 종강일 불일치: 개요/RMS 종강일 {p_close_short} ≠ 1월 진도표 마지막날 {jan_last_short}"
        )
        fix_requests.append(
            f"1월 진도표 마지막 수업일({jan_last_short})을 종강일({p_close_short})과 일치하도록 수정 필요"
        )
        error_columns.add("PDF 1월진도표(첫날~끝날/회차)")

    # ------------------------------------------------------------------
    # 규칙 4: 진도표 내 비정상 요일 섞임 검증
    # ------------------------------------------------------------------
    if jan_wk_err:
        error_tags.append("🔴 [진도표 날짜/요일 오타]")
        error_details.append(f"진도표 내 비정상 요일 섞임: {jan_wk_err_detail}")
        if r_close_short and jan_last_short and r_close != jan_last:
            fix_requests.append(
                f"PDF 1월 진도표 날짜/요일 오타 수정 필요 ({jan_last_short} → {r_close_short} 등 {effective_days_dot} 수업요일에 맞게 수정)"
            )
        else:
            fix_requests.append(f"PDF 1월 진도표 내 비정상 요일 오타 수정 필요 ({jan_wk_err_detail})")
        error_columns.add("PDF 1월진도표(첫날~끝날/회차)")

    # ------------------------------------------------------------------
    # 규칙 5: 수업 요일/시간 및 바자관 시수 검증
    # ------------------------------------------------------------------
    sched_matched = (r_full_keys == p_full_keys) if (r_full_keys and p_full_keys) else (
        (r_days_str == p_days_str) and (r_start == p_start)
    )
    if not sched_matched and (r_compact_sched and p_compact_sched):
        error_tags.append("🔴 [수업시간 불일치]")
        error_details.append(f"수업시간 불일치: RMS '{r_compact_sched}' ≠ PDF '{p_compact_sched}'")
        fix_requests.append(f"수업시간 표기 확인 및 수정 필요 (RMS: {r_compact_sched} vs PDF: {p_compact_sched})")
        error_columns.add("수업시간 비교")

    baza_matched = True
    if r_baza != "-" and p_baza != "-" and r_baza != p_baza:
        baza_matched = False
        error_tags.append("🔴 [바자관 시수 불일치]")
        error_details.append(f"바자관 시수 불일치: RMS {r_baza} ≠ PDF {p_baza}")
        fix_requests.append(f"PDF 상단 바자관 시수를 {p_baza} → {r_baza}로 수정 필요")
        error_columns.add("바자관시수 비교")

    # ------------------------------------------------------------------
    # 규칙 6: 과목 및 강좌명(슬로건 포함) 비교
    # ------------------------------------------------------------------
    if r_subj and p_subj and (r_subj not in p_subj and p_subj not in r_subj) and (
        r_area not in p_subj and p_subj not in r_area
    ):
        error_tags.append("🟡 [과목 표기 차이]")
        error_details.append(f"과목 표기 차이: RMS '{r_subj}' ≠ PDF '{p_subj}'")
        error_columns.add("과목")

    sim_title = compute_title_similarity(r_pure_title, p_pure_title)
    sim_slogan = compute_title_similarity(r_pure_title, p_slogan) if p_slogan else 0.0
    best_title_sim = max(sim_title, sim_slogan)

    if r_pure_title != p_pure_title:
        if sim_slogan >= 85.0 and sim_title < 60.0:
            info_notes.append(f"참고: RMS 강좌명('{r_pure_title}')이 PDF 슬로건에 표기됨 (PDF 강좌명: '{p_pure_title}')")
        elif best_title_sim < title_similarity_threshold:
            info_notes.append(f"참고: 강좌명 변경 확인 (유사도 {best_title_sim:.0f}%)")
        elif best_title_sim < 95.0:
            info_notes.append(f"참고: 강좌명 일부 표기 차이 (유사도 {best_title_sim:.0f}%)")

    # RMS 말머리태그 내 학년 오타 참고 알림 (예: 학년은 고1인데 말머리태그에 [고2]로 적힌 경우)
    if r_grade in ("고1", "고2", "고3") and r_tags:
        other_grades = [g for g in ("고1", "고2", "고3") if g != r_grade and g in r_tags]
        if other_grades:
            info_notes.append(f"참고: RMS 학년은 '{r_grade}'이나 RMS 강좌명 말머리태그에 '{other_grades[0]}'로 표기됨")

    # 비교 문자열 생성
    if sched_matched:
        sched_compare_str = f"✅ {r_compact_sched}"
    else:
        sched_compare_str = f"🔴 RMS: {r_compact_sched} ≠ PDF: {p_compact_sched}"

    if baza_matched:
        baza_compare_str = f"✅ {r_baza}" if r_baza != "-" else p_baza
    else:
        baza_compare_str = f"🔴 RMS:{r_baza} ≠ PDF:{p_baza}"

    has_error = len(error_tags) > 0
    has_red_error = any("🔴" in t for t in error_tags)

    if has_red_error:
        status_badge = "🔴 오류 (" + ", ".join(t.replace("🔴 ", "").replace("🟡 ", "") for t in error_tags) + ")"
    elif has_error:
        status_badge = "🟡 확인필요 (" + ", ".join(t.replace("🟡 ", "") for t in error_tags) + ")"
    else:
        status_badge = "✅ 정상 (완벽 일치)"

    all_detail_parts = error_details + info_notes
    detail_str = " / ".join(all_detail_parts) if all_detail_parts else "전 항목(수강료·회차·개강일·종강일·진도표·시간·바자관) 완벽 일치"
    fix_summary_str = " / ".join(fix_requests) if fix_requests else "-"

    return {
        "status_badge": status_badge,
        "has_error": has_error,
        "has_red_error": has_red_error,
        "error_tags": error_tags,
        "error_details": error_details,
        "detail_str": detail_str,
        "fix_summary_str": fix_summary_str,
        "info_notes": info_notes,
        "error_columns": list(error_columns),
        "title_similarity": best_title_sim,
        "sched_compare_str": sched_compare_str,
        "baza_compare_str": baza_compare_str,
        "fee_mismatch": fee_mismatch,
    }


def run_rms_pdf_crosscheck(
    rms_df: pd.DataFrame,
    syllabus_pdf_df: Optional[pd.DataFrame],
    mapper: InstructorMapper,
    target_ym: str = "2027-01",
    title_similarity_threshold: float = 55.0,
    max_char_threshold: int = 55,
    gemini_api_key: Optional[str] = None,
    stopped_rms_df: Optional[pd.DataFrame] = None,
) -> Dict[str, Any]:
    """
    [1. RMS 수업목록(.xls/.xlsx)] ↔ [2. 강의계획서 PDF] 양방향 정밀 검수를 수행합니다.
    - stopped_rms_df: '판매중지'로 분리된 RMS 강좌 목록 (PDF 수록 여부 대조 및 폐강 의심 강좌 감지)
    """
    if rms_df is None or rms_df.empty:
        return {
            "comparison_df": pd.DataFrame(),
            "error_df": pd.DataFrame(),
            "missing_pdf_df": pd.DataFrame(),
            "unmatched_pdf_df": pd.DataFrame(),
            "pdf_stopped_courses_df": pd.DataFrame(),
            "stopped_rms_df": stopped_rms_df if stopped_rms_df is not None else pd.DataFrame(),
            "pdf_schedule_df": pd.DataFrame(),
            "spelling_df": pd.DataFrame(),
            "long_sentence_df": pd.DataFrame(),
            "instructor_summary_df": pd.DataFrame(),
            "rms_summary": extract_rms_summary(None),
            "kpi": {
                "total_rms": 0,
                "total_pdf": 0,
                "pdf_matched_count": 0,
                "normal_count": 0,
                "error_count": 0,
                "missing_pdf_count": 0,
                "unmatched_pdf_count": 0,
                "pdf_stopped_count": 0,
                "fee_error_count": 0,
                "session_error_count": 0,
                "date_error_count": 0,
                "schedule_typo_count": 0,
                "spelling_error_count": 0,
                "long_sentence_count": 0,
            },
        }

    rms = rms_df.copy().reset_index(drop=True)
    rms["정규화강사명"] = rms["강사"].apply(mapper.normalize)

    pdf_df = (
        syllabus_pdf_df.copy().reset_index(drop=True)
        if syllabus_pdf_df is not None and not syllabus_pdf_df.empty
        else pd.DataFrame()
    )

    rms_to_pdf_idx: Dict[int, int] = {}
    pdf_to_rms_idx: Dict[int, int] = {}
    used_pdf_indices: Set[int] = set()

    if not pdf_df.empty:
        # [1차 정확 매칭]: 강사명 일치 + 수업요일 + 시작시간 일치 (동일 강사 다중 반은 강좌명/슬로건 유사도 + 학년으로 구분)
        for r_idx, r_row in rms.iterrows():
            r_inst = str(r_row["강사"]).strip()
            r_days = str(r_row.get("수업요일", "")).strip()
            r_start = str(r_row.get("전체시작시간", "") or r_row.get("시작시간", "")).strip()
            r_pure = str(r_row.get("순수강좌명", "")).strip()
            r_grade = str(r_row.get("학년", "")).strip()

            best_p_idx = None
            best_score = -1.0

            for p_idx, p_row in pdf_df.iterrows():
                if p_idx in used_pdf_indices:
                    continue
                p_inst_orig = str(p_row.get("강사명_원본", "")).strip()
                p_inst_clean = str(p_row.get("강사명", "")).strip()
                if not (
                    mapper.match_instructors(r_inst, p_inst_orig)
                    or mapper.match_instructors(r_inst, p_inst_clean)
                ):
                    continue

                p_days = str(p_row.get("수업요일", "")).strip()
                p_start = str(p_row.get("전체시작시간", "") or p_row.get("시작시간", "")).strip()

                if r_days and p_days and r_days == p_days and r_start and p_start and r_start == p_start:
                    sim_t = compute_title_similarity(r_pure, str(p_row.get("순수강좌명", "")))
                    sim_s = compute_title_similarity(r_pure, str(p_row.get("슬로건", "")))
                    sim = max(sim_t, sim_s)
                    p_grade = str(p_row.get("대상학년", "")).strip()
                    if r_grade and p_grade and (r_grade in p_grade or p_grade in r_grade):
                        sim += 15.0
                    if sim > best_score:
                        best_score = sim
                        best_p_idx = p_idx

            if best_p_idx is not None:
                rms_to_pdf_idx[r_idx] = best_p_idx
                pdf_to_rms_idx[best_p_idx] = r_idx
                used_pdf_indices.add(best_p_idx)

        # [2차 보조 매칭]: 요일/시간 일부 오기입 시 강사명 + 강좌명/슬로건 유사도 + 학년 기반 매칭
        for p_idx, p_row in pdf_df.iterrows():
            if p_idx in used_pdf_indices:
                continue
            p_inst_orig = str(p_row.get("강사명_원본", "")).strip()
            p_inst_clean = str(p_row.get("강사명", "")).strip()
            p_pure = str(p_row.get("순수강좌명", "")).strip()
            p_slogan = str(p_row.get("슬로건", "")).strip()
            p_grade = str(p_row.get("대상학년", "")).strip()

            best_r_idx = None
            best_score = -1.0
            for r_idx, r_row in rms.iterrows():
                if r_idx in rms_to_pdf_idx:
                    continue
                r_inst = str(r_row["강사"]).strip()
                if not (
                    mapper.match_instructors(r_inst, p_inst_orig)
                    or mapper.match_instructors(r_inst, p_inst_clean)
                ):
                    continue
                r_pure = str(r_row.get("순수강좌명", "")).strip()
                sim = max(
                    compute_title_similarity(r_pure, p_pure),
                    compute_title_similarity(r_pure, p_slogan),
                )
                r_grade = str(r_row.get("학년", "")).strip()
                if r_grade and p_grade and (r_grade in p_grade or p_grade in r_grade):
                    sim += 12.0
                if sim > best_score:
                    best_score = sim
                    best_r_idx = r_idx

            if best_r_idx is not None and best_score >= 45.0:
                rms_to_pdf_idx[best_r_idx] = p_idx
                pdf_to_rms_idx[p_idx] = best_r_idx
                used_pdf_indices.add(p_idx)

    # ---------------------------------------------------------
    # 전 강좌 대조 테이블 생성
    # ---------------------------------------------------------
    comparison_rows: List[Dict[str, Any]] = []

    kpi_normal = 0
    kpi_error = 0
    kpi_missing_pdf = 0
    kpi_fee_err = 0
    kpi_session_err = 0
    kpi_date_err = 0
    kpi_sched_typo = 0

    for r_idx, r_row in rms.iterrows():
        r_no = int(r_row.get("No", r_idx + 1) or (r_idx + 1))
        r_code = str(r_row.get("강좌코드", "")).strip()
        r_area = str(r_row.get("영역", "")).strip()
        r_subj = str(r_row.get("과목", "")).strip()
        r_grade = str(r_row.get("학년", "")).strip()
        r_inst = str(r_row.get("강사", "")).strip()
        r_norm_inst = str(r_row.get("정규화강사명", "")).strip()
        r_title = str(r_row.get("강좌명", "")).strip()
        r_pure_title = str(r_row.get("순수강좌명", "")).strip()
        r_compact_sched = str(r_row.get("간결수업시간", "") or r_row.get("표준수업시간", "")).strip()
        r_open = str(r_row.get("개강일", "")).strip()
        r_close = str(r_row.get("종강일", "")).strip()
        r_open_short = str(r_row.get("개강일_짧은표시", "") or format_short_md_with_weekday(r_open)).strip()
        r_close_short = str(r_row.get("종강일_짧은표시", "") or format_short_md_with_weekday(r_close)).strip()
        r_sessions = int(r_row.get("회차", 0) or 0)
        r_fee = int(r_row.get("수강료", 0) or 0)
        r_unit_fee = int(r_row.get("1회차 기준 수강료", 0) or 0)
        r_baza = normalize_baza_hours(r_row.get("바자관시수", r_row.get("바자관 인정기준", "-")))
        r_total_stu = int(r_row.get("가등록포함수강생수", 0) or 0)
        r_paid_stu = int(r_row.get("결제인원", 0) or 0)
        r_inst_stu = int(r_row.get("완납&분납수강생수", 0) or 0)
        r_unpaid_stu = int(r_row.get("가등록(미결제)", 0) or max(0, r_total_stu - r_paid_stu))
        r_pay_rate = float(r_row.get("결제율(%)", 0.0) or (round(r_paid_stu / r_total_stu * 100.0, 1) if r_total_stu > 0 else 0.0))

        rms_period_short_disp = f"{r_open_short}~{r_close_short} ({r_sessions}회)"
        rms_fee_disp = (
            f"{r_fee:,}원 (회차당 {r_unit_fee:,}원)"
            if (r_fee > 0 and r_unit_fee > 0)
            else (f"{r_fee:,}원" if r_fee > 0 else "-")
        )

        if r_idx in rms_to_pdf_idx:
            p_idx = rms_to_pdf_idx[r_idx]
            p_row = pdf_df.loc[p_idx]

            ver = verify_rms_vs_pdf_course(
                rms_row=r_row,
                pdf_row=p_row,
                target_ym=target_ym,
                title_similarity_threshold=title_similarity_threshold,
            )

            p_inst_orig = str(p_row.get("강사명_원본", "")).strip()
            inst_display = f"{r_inst} ({p_inst_orig})" if (p_inst_orig and p_inst_orig != r_inst) else r_inst
            p_title = str(p_row.get("강좌명", "")).strip()
            p_page_range = str(p_row.get("페이지_범위", "")).strip()
            p_ov_period_disp = str(p_row.get("PDF_개요_기간회차_표시", "-")).strip()
            p_jan_plan_disp = str(p_row.get("진도표_1월_기간회차_표시", "-")).strip()
            p_fee_disp = str(p_row.get("수강료_표시", "-")).strip()
            p_book_disp = str(p_row.get("교재비_표시", "-")).strip()
            p_jan_dates_str = str(p_row.get("진도표_1월_날짜문자열", "-")).strip()

            status_badge = ver["status_badge"]
            detail_str = ver["detail_str"]
            fix_summary_str = ver["fix_summary_str"]
            sched_compare_str = ver["sched_compare_str"]
            baza_compare_str = ver["baza_compare_str"]
            err_cols = ver["error_columns"]
            err_tags = ver["error_tags"]

            if ver["has_error"]:
                kpi_error += 1
                if any("수강료" in t or "단가" in t for t in err_tags):
                    kpi_fee_err += 1
                if any("회차" in t for t in err_tags):
                    kpi_session_err += 1
                if any("개강일" in t or "종강일" in t for t in err_tags):
                    kpi_date_err += 1
                if any("진도표 날짜/요일 오타" in t for t in err_tags):
                    kpi_sched_typo += 1
            else:
                kpi_normal += 1

            p_fee_int = int(p_row.get("수강료", 0) or 0)
            p_sess_int = int(p_row.get("회차", 0) or 0)
            p_jan_sess_int = int(p_row.get("진도표_1월_실제회차", 0) or 0)
        else:
            inst_display = r_inst
            p_title = "-"
            p_page_range = "-"
            p_ov_period_disp = "-"
            p_jan_plan_disp = "-"
            p_fee_disp = "-"
            p_book_disp = "-"
            p_jan_dates_str = "-"
            status_badge = "⚪ PDF 미수록"
            detail_str = "업로드된 강의계획서 PDF 내 미수록"
            fix_summary_str = "PDF 강의계획서 원고 수령 및 수록 확인 필요"
            sched_compare_str = f"RMS: {r_compact_sched} (PDF 미수록)"
            baza_compare_str = f"RMS: {r_baza} (PDF 미수록)"
            err_cols = []
            err_tags = []
            kpi_missing_pdf += 1
            p_fee_int = 0
            p_sess_int = 0
            p_jan_sess_int = 0

        comparison_rows.append(
            {
                "검수상태": status_badge,
                "No": r_no,
                "강좌코드": r_code,
                "영역": r_area,
                "과목": r_subj,
                "학년": r_grade,
                "강사명": inst_display,
                "RMS_강사명": r_inst,
                "정규화강사명": r_norm_inst,
                "RMS 강좌명": r_title,
                "PDF 강좌명": p_title,
                "PDF 페이지": p_page_range,
                "RMS 개강~종강(회차)": rms_period_short_disp,
                "PDF 개요 개강~종강(회차)": p_ov_period_disp,
                "PDF 1월진도표(첫날~끝날/회차)": p_jan_plan_disp,
                "RMS 수강료": rms_fee_disp,
                "PDF 수강료": p_fee_disp,
                "총수강생(가등록포함)": r_total_stu,
                "결제완료인원": r_paid_stu,
                "완납&분납인원": r_inst_stu,
                "가등록(미결제)": r_unpaid_stu,
                "결제율(%)": r_pay_rate,
                "수업시간 비교": sched_compare_str,
                "바자관시수 비교": baza_compare_str,
                "오류 상세 내역": detail_str,
                "자동 생성된 수정 요청 사항 요약": fix_summary_str,
                "PDF 교재비": p_book_disp,
                "PDF 1월 진도표 날짜리스트": p_jan_dates_str,
                "_오류컬럼리스트": err_cols,
                "_오류태그리스트": err_tags,
                "_RMS_수강료_정수": r_fee,
                "_PDF_수강료_정수": p_fee_int,
                "_RMS_회차_정수": r_sessions,
                "_PDF_개요회차_정수": p_sess_int,
                "_PDF_진도표회차_정수": p_jan_sess_int,
            }
        )

    comparison_df = pd.DataFrame(comparison_rows)

    # 오류 및 수정 필요 강좌 목록 (Sheet 2 전용)
    error_df = (
        comparison_df[comparison_df["검수상태"].str.startswith(("🔴", "🟡"))].copy().reset_index(drop=True)
        if not comparison_df.empty
        else pd.DataFrame()
    )

    # PDF 미수록 강좌 목록 (Sheet 3 전용)
    missing_pdf_df = (
        comparison_df[comparison_df["검수상태"] == "⚪ PDF 미수록"].copy().reset_index(drop=True)
        if not comparison_df.empty
        else pd.DataFrame()
    )

    # PDF 진도표 상세 검수 테이블 (Tab 3 & Sheet 4 전용)
    pdf_sched_rows: List[Dict[str, Any]] = []
    unmatched_pdf_rows: List[Dict[str, Any]] = []
    pdf_stopped_matches: List[Dict[str, Any]] = []

    if not pdf_df.empty:
        for p_idx, p_row in pdf_df.iterrows():
            p_dict = p_row.to_dict()
            p_seq = int(p_row.get("PDF_순번", p_idx + 1))
            p_page = str(p_row.get("페이지_범위", "")).strip()
            p_subj = str(p_row.get("과목", "")).strip()
            p_grade = str(p_row.get("대상학년", "")).strip()
            p_inst = str(p_row.get("강사명_원본", "") or p_row.get("강사명", "")).strip()
            p_title = str(p_row.get("강좌명", "")).strip()
            p_sched = str(p_row.get("간결수업시간", "") or p_row.get("수업시간_원본", "")).strip()
            p_baza = str(p_row.get("바자관시수", "-")).strip()
            p_ov_period = str(p_row.get("PDF_개요_기간회차_표시", "-")).strip()
            p_jan_period = str(p_row.get("진도표_1월_기간회차_표시", "-")).strip()
            p_jan_dates = str(p_row.get("진도표_1월_날짜문자열", "-")).strip()
            p_spilled = str(p_row.get("진도표_2월표_1월날짜문자열", "")).strip()
            p_fee_disp = str(p_row.get("수강료_표시", "-")).strip()
            p_wk_err = bool(p_row.get("진도표_요일오류여부", False))
            p_wk_err_msg = str(p_row.get("진도표_요일오류상세", "")).strip()

            if p_idx in pdf_to_rms_idx:
                r_idx = pdf_to_rms_idx[p_idx]
                r_row = rms.loc[r_idx]
                ver = verify_rms_vs_pdf_course(r_row, p_row, target_ym=target_ym)
                r_no = int(r_row.get("No", r_idx + 1))
                r_code = str(r_row.get("강좌코드", "")).strip()
                r_title = str(r_row.get("강좌명", "")).strip()
                r_open_short = str(r_row.get("개강일_짧은표시", "")).strip()
                r_close_short = str(r_row.get("종강일_짧은표시", "")).strip()
                r_sess = int(r_row.get("회차", 0) or 0)
                r_period = f"{r_open_short}~{r_close_short} ({r_sess}회)"
                r_fee = int(r_row.get("수강료", 0) or 0)
                r_fee_str = f"{r_fee:,}원" if r_fee > 0 else "-"
                status_badge = ver["status_badge"]
                detail_str = ver["detail_str"]
                fix_str = ver["fix_summary_str"]
            else:
                # 판매중지된 강좌와의 매칭 여부 정밀 확인
                stopped_matched_row = None
                if stopped_rms_df is not None and not stopped_rms_df.empty:
                    for s_idx, s_row in stopped_rms_df.iterrows():
                        s_inst = str(s_row.get("강사", "")).strip()
                        if mapper.match_instructors(s_inst, p_inst):
                            s_days = str(s_row.get("수업요일", "")).strip()
                            s_start = str(s_row.get("전체시작시간", "") or s_row.get("시작시간", "")).strip()
                            p_days = str(p_row.get("수업요일", "")).strip()
                            p_start = str(p_row.get("전체시작시간", "") or p_row.get("시작시간", "")).strip()
                            sim = compute_title_similarity(str(s_row.get("순수강좌명", "")), str(p_row.get("순수강좌명", "")))
                            if (s_days and p_days and s_days == p_days and s_start and p_start and s_start == p_start) or sim >= 45.0:
                                stopped_matched_row = s_row
                                break

                if stopped_matched_row is not None:
                    r_no = int(stopped_matched_row.get("No", 0) or 0)
                    r_code = str(stopped_matched_row.get("강좌코드", "")).strip()
                    r_title = str(stopped_matched_row.get("강좌명", "")).strip()
                    r_open_short = str(stopped_matched_row.get("개강일_짧은표시", "")).strip()
                    r_close_short = str(stopped_matched_row.get("종강일_짧은표시", "")).strip()
                    r_sess = int(stopped_matched_row.get("회차", 0) or 0)
                    r_period = f"{r_open_short}~{r_close_short} ({r_sess}회)"
                    r_fee = int(stopped_matched_row.get("수강료", 0) or 0)
                    r_fee_str = f"{r_fee:,}원" if r_fee > 0 else "-"
                    status_badge = "⚠️ RMS 판매중지 (폐강 의심)"
                    detail_str = f"RMS에서 '판매중지' 상태인 강좌입니다. (강좌코드: {r_code})"
                    fix_str = "사전 검수 시 상단 필터 해제 / 실제 폐강 강좌라면 PDF에서 삭제 필요"
                    pdf_stopped_matches.append(
                        {
                            "PDF_순번": p_seq,
                            "PDF_페이지": p_page,
                            "과목": p_subj,
                            "학년": p_grade,
                            "강사명": p_inst,
                            "PDF 강좌명": p_title,
                            "RMS 강좌명": r_title,
                            "강좌코드": r_code,
                            "RMS_No": r_no,
                            "RMS 판매상태": "판매중지",
                            "수업시간": p_sched,
                            "수강료": p_fee_disp,
                            "오류 상세 내역": detail_str,
                            "확인 및 조치사항": "접수 전 사전 검수 시 상단 필터 해제 / 폐강 시 PDF 삭제 필요",
                        }
                    )
                else:
                    r_no = 0
                    r_code = "미매칭"
                    r_title = "-"
                    r_period = "-"
                    r_fee_str = "-"
                    status_badge = "⚠️ RMS 미개설"
                    detail_str = "RMS 수업목록에서 일치하는 강좌를 찾지 못함"
                    fix_str = "RMS 개설 여부 확인 필요"
                    unmatched_pdf_rows.append(p_dict)

            # 진도표 요일 정합성 판정 요약
            if p_wk_err:
                wk_status = f"🔴 요일오타 ({p_wk_err_msg})"
            elif p_spilled:
                wk_status = f"🟡 2월표로 넘어감 ({p_spilled})"
            else:
                wk_status = "✅ 전 날짜 요일 일치"

            pdf_sched_rows.append(
                {
                    "PDF_순번": p_seq,
                    "PDF_페이지": p_page,
                    "검수상태": status_badge,
                    "RMS_No": r_no,
                    "강좌코드": r_code,
                    "과목": p_subj,
                    "학년": p_grade,
                    "강사명": p_inst,
                    "PDF 강좌명": p_title,
                    "RMS 강좌명": r_title,
                    "수업시간": p_sched,
                    "바자관시수": p_baza,
                    "RMS 개강~종강(회차)": r_period,
                    "PDF 개요 개강~종강(회차)": p_ov_period,
                    "PDF 1월진도표(첫날~끝날/회차)": p_jan_period,
                    "1월 진도표 추출 날짜 리스트": p_jan_dates,
                    "2월표 내 1월 날짜": p_spilled or "-",
                    "진도표 요일 정합성": wk_status,
                    "RMS 수강료": r_fee_str,
                    "PDF 수강료": p_fee_disp,
                    "오류 상세 내역": detail_str,
                    "자동 생성된 수정 요청 사항 요약": fix_str,
                    "주차별진도계획_1월리스트": p_row.get("주차별진도계획_1월리스트", []),
                    "주차별진도계획_2월리스트": p_row.get("주차별진도계획_2월리스트", []),
                    "주차별진도계획_리스트": p_row.get("주차별진도계획_리스트", []),
                    "원문전체텍스트": p_row.get("원문전체텍스트", ""),
                }
            )

    pdf_schedule_df = pd.DataFrame(pdf_sched_rows)
    unmatched_pdf_df = pd.DataFrame(unmatched_pdf_rows)

    # ---------------------------------------------------------
    # PDF 본문 텍스트 국어 맞춤법 & 긴 문장 간략화 제안 분석
    # ---------------------------------------------------------
    all_spelling_issues: List[Dict[str, Any]] = []
    all_long_proposals: List[Dict[str, Any]] = []

    if not pdf_df.empty:
        for p_idx, p_row in pdf_df.iterrows():
            c_title = str(p_row.get("강좌명", "")).strip()
            c_inst = str(p_row.get("강사명_원본", "") or p_row.get("강사명", "")).strip()
            raw_txts = p_row.get("강좌_텍스트리스트", [])
            if not raw_txts and p_row.get("원문전체텍스트"):
                raw_txts = [str(p_row.get("원문전체텍스트", ""))]

            # 업로드 시점에는 순수 0.1초 로컬 규칙(120종 맞춤법/오탈자 + 긴문장 축약)만 즉시 실행 (외부 API 호출 배제)
            analysis_res = analyze_course_text_content(
                course_name=c_title,
                instructor=c_inst,
                raw_texts=raw_txts,
                max_char_threshold=max_char_threshold,
                gemini_api_key=None,
            )
            all_spelling_issues.extend(analysis_res["맞춤법_상세목록"])
            all_long_proposals.extend(analysis_res["긴문장_상세목록"])

    spelling_df = pd.DataFrame(all_spelling_issues)
    long_sentence_df = pd.DataFrame(all_long_proposals)
    pdf_stopped_df = pd.DataFrame(pdf_stopped_matches)

    stopped_cnt = len(stopped_rms_df) if (stopped_rms_df is not None and not stopped_rms_df.empty) else 0

    return {
        "comparison_df": comparison_df,
        "error_df": error_df,
        "missing_pdf_df": missing_pdf_df,
        "unmatched_pdf_df": unmatched_pdf_df,
        "pdf_stopped_courses_df": pdf_stopped_df,
        "stopped_rms_df": stopped_rms_df if stopped_rms_df is not None else pd.DataFrame(),
        "pdf_schedule_df": pdf_schedule_df,
        "spelling_df": spelling_df,
        "long_sentence_df": long_sentence_df,
        "instructor_summary_df": generate_instructor_student_summary(rms),
        "rms_summary": extract_rms_summary(rms, stopped_count=stopped_cnt),
        "kpi": {
            "total_rms": len(rms),
            "total_pdf": len(pdf_df),
            "pdf_matched_count": len(rms_to_pdf_idx),
            "normal_count": kpi_normal,
            "error_count": kpi_error,
            "missing_pdf_count": kpi_missing_pdf,
            "unmatched_pdf_count": len(unmatched_pdf_df),
            "pdf_stopped_count": len(pdf_stopped_matches),
            "fee_error_count": kpi_fee_err,
            "session_error_count": kpi_session_err,
            "date_error_count": kpi_date_err,
            "schedule_typo_count": kpi_sched_typo,
            "spelling_error_count": len(all_spelling_issues),
            "long_sentence_count": len(all_long_proposals),
        },
    }


def generate_pdf_fix_request_messages(
    error_df: pd.DataFrame,
    target_ym: str = "2027-01",
) -> Dict[str, str]:
    """
    오류가 발견된 강좌들에 대해 디자인팀/강사님께 전달할 [PDF 수정 요청 텍스트]를 자동 생성합니다.
    반환값: {'__ALL__': 전체_통합_요청문, '강사명': 강사별_요청문, ...}
    """
    if error_df is None or error_df.empty:
        return {}

    ym_parts = str(target_ym).split("-")
    ym_label = f"{ym_parts[0]}년 {int(ym_parts[1])}월" if len(ym_parts) == 2 else target_ym

    messages: Dict[str, str] = {}
    summary_lines: List[str] = [
        f"[{ym_label} 러셀대치 단과 강의계획서 PDF 검수 — 수정 요청 사항 안내]",
        f"• 검수 일시: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"• 수정 필요 강좌: 총 {len(error_df)}건",
        "-" * 68,
    ]

    for idx, (_, row) in enumerate(error_df.iterrows(), start=1):
        no_val = row.get("No", idx)
        code = row.get("강좌코드", "")
        subj = row.get("과목", "")
        grade = row.get("학년", "")
        inst = row.get("강사명", "")
        p_page = row.get("PDF 페이지", "")
        r_title = row.get("RMS 강좌명", "")
        p_title = row.get("PDF 강좌명", "")
        r_period = row.get("RMS 개강~종강(회차)", "")
        p_ov_period = row.get("PDF 개요 개강~종강(회차)", "")
        p_jan_period = row.get("PDF 1월진도표(첫날~끝날/회차)", "")
        r_fee = row.get("RMS 수강료", "")
        p_fee = row.get("PDF 수강료", "")
        err_detail = row.get("오류 상세 내역", "")
        fix_req = row.get("자동 생성된 수정 요청 사항 요약", "")

        block = (
            f"{idx}. [{subj} / {grade}] {inst}T — {p_title} (PDF {p_page} / 강좌코드: {code}, RMS No.{no_val})\n"
            f"   • RMS 기준 정보 : 개강~종강 {r_period} | 수강료 {r_fee}\n"
            f"   • PDF 개요 표기 : 개강~종강 {p_ov_period} | 수강료 {p_fee}\n"
            f"   • PDF 1월진도표 : {p_jan_period}\n"
            f"   • 검출된 오류   : {err_detail}\n"
            f"   ▶ [수정 요청]  : {fix_req}"
        )
        summary_lines.append(block)

    summary_lines.append("-" * 68)
    summary_lines.append("위 항목 확인 후 배포용 강의계획서 PDF(및 필요 시 RMS) 수정 반영 부탁드립니다. 감사합니다.")
    messages["__ALL__"] = "\n\n".join(summary_lines)

    # 강사별 개별 메시지 생성
    for inst_name, grp in error_df.groupby("RMS_강사명", sort=False):
        inst_lines = [
            f"[메가스터디 러셀 대치학원 {ym_label} 단과 강의계획서 PDF 검수 확인 요청]",
            f"안녕하세요, {inst_name} 선생님! 러셀 대치학원 교무입시팀입니다.",
            f"{ym_label} 단과 강의계획서 PDF와 RMS 수업목록 대조 검수 결과, 아래 강좌에서 확인 및 수정이 필요한 항목이 발견되어 안내드립니다.\n",
        ]
        for c_i, (_, r) in enumerate(grp.iterrows(), start=1):
            inst_lines.append(
                f"■ {c_i}. [{r.get('학년', '')} {r.get('과목', '')}] {r.get('PDF 강좌명', '')} (PDF {r.get('PDF 페이지', '')})\n"
                f"  - RMS 등록 기준 : {r.get('RMS 개강~종강(회차)', '')} / 수강료 {r.get('RMS 수강료', '')}\n"
                f"  - PDF 개요 표기 : {r.get('PDF 개요 개강~종강(회차)', '')} / 수강료 {r.get('PDF 수강료', '')}\n"
                f"  - PDF 1월진도표 : {r.get('PDF 1월진도표(첫날~끝날/회차)', '')}\n"
                f"  - 확인 필요 사유 : {r.get('오류 상세 내역', '')}\n"
                f"  ▶ 수정 요청 사항 : {r.get('자동 생성된 수정 요청 사항 요약', '')}\n"
            )
        inst_lines.append("확인 후 회신 주시면 강의계획서 PDF 및 시스템에 즉시 수정 반영하겠습니다. 감사합니다!")
        messages[str(inst_name)] = "\n".join(inst_lines)

    return messages


def export_crosscheck_to_excel(crosscheck_result: Dict[str, Any], target_ym: str = "2027-01") -> bytes:
    """
    openpyxl을 사용하여 아래 4개 시트로 구성된 서식 포함 엑셀(.xlsx) 바이트를 생성합니다:
    - Sheet 1 (전체검수결과): 전 강좌 RMS vs PDF 비교표 (오류 행은 연한 빨간색 배경 #FADBD8, 불일치 세부 셀은 굵은 빨간 글씨 하이라이트, 열 너비 자동 맞춤)
    - Sheet 2 (오류및수정필요목록): 오류가 발견된 강좌만 추출 + 우측 [자동 생성된 수정 요청 사항 요약] 컬럼 포함
    - Sheet 3 (PDF누락강좌): RMS에는 있으나 PDF에서 빠진 강좌 리스트
    - Sheet 4 (PDF진도표상세검수): PDF 내 1월 진도표 날짜 리스트 및 요일 정합성 상세 내역
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws_s1 = wb.active
    ws_s1.title = "전체검수결과_유효강좌"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    base_font = Font(name="맑은 고딕", size=9, color="111111")

    # 사용자 지정 조건부 하이라이트 서식: 연한 빨간색 배경 #FADBD8, 불일치 세부 셀 굵은 빨간 글씨
    fill_error_row = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    font_error_row = Font(name="맑은 고딕", size=9, bold=False, color="5B1A18")
    font_error_cell = Font(name="맑은 고딕", size=9, bold=True, color="C00000")
    fill_error_cell = PatternFill(start_color="F5B7B1", end_color="F5B7B1", fill_type="solid")

    fill_normal_row = PatternFill(start_color="EAFAF1", end_color="EAFAF1", fill_type="solid")
    font_normal_row = Font(name="맑은 고딕", size=9, color="1E5631")

    fill_missing_row = PatternFill(start_color="F4F6F7", end_color="F4F6F7", fill_type="solid")
    font_missing_row = Font(name="맑은 고딕", size=9, color="5D6D7E")

    thin_side = Side(border_style="thin", color="D5D8DC")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def _write_styled_sheet(
        ws: Any,
        df: pd.DataFrame,
        columns_to_export: List[str],
        status_col_name: str = "검수상태",
        highlight_error_cells: bool = True,
    ) -> None:
        if df is None or df.empty:
            ws.append(["데이터가 없습니다."])
            return

        cols = [c for c in columns_to_export if c in df.columns]
        sub_df = df[cols].copy()

        for c in cols:
            sub_df[c] = sub_df[c].apply(
                lambda x: ", ".join(map(str, x)) if isinstance(x, (list, tuple, set)) else ("" if pd.isna(x) else str(x))
            )

        ws.append(cols)
        for col_idx in range(1, len(cols) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border
        ws.row_dimensions[1].height = 28

        for row_offset, (orig_idx, r_data) in enumerate(sub_df.iterrows(), start=2):
            row_vals = [r_data[c] for c in cols]
            ws.append(row_vals)
            ws.row_dimensions[row_offset].height = 22

            status_val = str(r_data.get(status_col_name, "")) if status_col_name in cols else ""
            err_cols_for_row: Set[str] = set()
            if highlight_error_cells and "_오류컬럼리스트" in df.columns:
                raw_ec = df.loc[orig_idx, "_오류컬럼리스트"]
                if isinstance(raw_ec, (list, set, tuple)):
                    err_cols_for_row = set(raw_ec)

            is_err_row = status_val.startswith(("🔴", "🟡")) or ("불일치" in status_val) or ("오타" in status_val) or ("모순" in status_val)
            is_ok_row = status_val.startswith("✅") or ("정상" in status_val)
            is_missing_row = status_val.startswith("⚪") or ("미수록" in status_val)

            row_fill = None
            row_font = base_font
            if is_err_row:
                row_fill = fill_error_row
                row_font = font_error_row
            elif is_ok_row:
                row_fill = fill_normal_row
                row_font = font_normal_row
            elif is_missing_row:
                row_fill = fill_missing_row
                row_font = font_missing_row

            for col_idx, col_name in enumerate(cols, start=1):
                c = ws.cell(row=row_offset, column=col_idx)
                c.border = cell_border
                if row_fill is not None:
                    c.fill = row_fill
                c.font = row_font

                # 불일치 세부 셀 굵은 빨간 글씨 하이라이트
                if is_err_row and (
                    col_name in err_cols_for_row
                    or col_name in ("검수상태", "오류 상세 내역", "자동 생성된 수정 요청 사항 요약")
                    or (col_name == "진도표 요일 정합성" and ("🔴" in str(c.value) or "🟡" in str(c.value)))
                ):
                    c.font = font_error_cell
                    if col_name in err_cols_for_row:
                        c.fill = fill_error_cell

                if any(k in col_name for k in ("강좌명", "내역", "요약", "리스트", "비교", "사유")):
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
                else:
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(sub_df) + 1}"

        # 열 너비 자동 맞춤
        for col_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = sum(1.8 if ord(ch) > 127 else 1.0 for ch in str(col_name)) + 2
            for val in sub_df[col_name].head(120):
                val_len = sum(1.65 if ord(ch) > 127 else 0.95 for ch in str(val))
                if val_len > max_len:
                    max_len = val_len
            ws.column_dimensions[col_letter].width = min(max(int(max_len) + 3, 10), 55)

    sheet1_cols = [
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
        "완납&분납인원",
        "결제율(%)",
        "수업시간 비교",
        "바자관시수 비교",
        "오류 상세 내역",
        "자동 생성된 수정 요청 사항 요약",
        "AI추천강좌명",
    ]
    _write_styled_sheet(
        ws_s1,
        crosscheck_result.get("comparison_df", pd.DataFrame()),
        columns_to_export=sheet1_cols,
        status_col_name="검수상태",
        highlight_error_cells=True,
    )

    # Sheet 2: 강사별_당월및누적수강생통계 (상단 요약 배너 포함)
    ws_s2 = wb.create_sheet(title="강사별_당월및누적수강생통계")
    rms_sum = crosscheck_result.get("rms_summary", {}) or {}
    banner_str = rms_sum.get(
        "headline_banner",
        f"총 강좌수 : {rms_sum.get('total_courses', 0)}건 | 총 수강생 : {rms_sum.get('total_students', 0)}명 (결제완료 : {rms_sum.get('paid_students', 0)}명)"
    )
    # 1행에 요약 헤더 삽입
    ws_s2.append([banner_str])
    ws_s2.cell(row=1, column=1).font = Font(name="맑은 고딕", size=11, bold=True, color="1E3A8A")
    ws_s2.row_dimensions[1].height = 24
    ws_s2.append([])  # 빈 행

    inst_summary_df = crosscheck_result.get("instructor_summary_df", pd.DataFrame())
    sheet2_cols = [
        "순위",
        "과목(영역)",
        "강사명",
        "개설 강좌수",
        "총 수강생수(가등록포함)",
        "결제완료 인원",
        "완납&분납 인원",
        "가등록(미결제) 인원",
        "결제율(%)",
        "강좌당 평균 수강생",
    ]
    if not inst_summary_df.empty:
        cols = [c for c in sheet2_cols if c in inst_summary_df.columns]
        ws_s2.append(cols)
        h_row = 3
        ws_s2.row_dimensions[h_row].height = 26
        for c_idx in range(1, len(cols) + 1):
            cell = ws_s2.cell(row=h_row, column=c_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = cell_border

        for r_offset, (_, r_data) in enumerate(inst_summary_df.iterrows(), start=h_row + 1):
            row_vals = [r_data[c] for c in cols]
            ws_s2.append(row_vals)
            ws_s2.row_dimensions[r_offset].height = 20
            for c_idx in range(1, len(cols) + 1):
                c_cell = ws_s2.cell(row=r_offset, column=c_idx)
                c_cell.border = cell_border
                c_cell.font = base_font
                c_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_s2.auto_filter.ref = f"A3:{get_column_letter(len(cols))}{len(inst_summary_df) + 3}"
        for col_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(col_idx)
            ws_s2.column_dimensions[col_letter].width = max(len(str(col_name)) * 2 + 4, 12)

    # Sheet 3: 데이터오류_및_판매중지확인
    ws_s3 = wb.create_sheet(title="데이터오류_및_판매중지확인")
    error_df = crosscheck_result.get("error_df", pd.DataFrame())
    pdf_stopped_df = crosscheck_result.get("pdf_stopped_courses_df", pd.DataFrame())

    combined_err_list: List[Dict[str, Any]] = []
    if not error_df.empty:
        for _, r in error_df.iterrows():
            d = dict(r)
            combined_err_list.append(d)

    if not pdf_stopped_df.empty:
        for _, r in pdf_stopped_df.iterrows():
            combined_err_list.append(
                {
                    "검수상태": "⚠️ RMS 판매중지 (폐강 의심)",
                    "No": r.get("RMS_No", 0),
                    "강좌코드": r.get("강좌코드", ""),
                    "과목": r.get("과목", ""),
                    "학년": r.get("학년", ""),
                    "강사명": r.get("강사명", ""),
                    "RMS 강좌명": r.get("RMS 강좌명", ""),
                    "PDF 강좌명": r.get("PDF 강좌명", ""),
                    "PDF 페이지": r.get("PDF_페이지", ""),
                    "수업시간 비교": f"PDF: {r.get('수업시간', '')}",
                    "바자관시수 비교": "-",
                    "오류 상세 내역": r.get("오류 상세 내역", "RMS에서 '판매중지' 상태로 등록된 강좌입니다."),
                    "자동 생성된 수정 요청 사항 요약": r.get("확인 및 조치사항", "사전 검수 시 필터 해제 / 폐강 시 PDF 삭제 필요"),
                }
            )

    combined_err_df = pd.DataFrame(combined_err_list)
    sheet3_cols = [
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
    _write_styled_sheet(
        ws_s3,
        combined_err_df,
        columns_to_export=sheet3_cols,
        status_col_name="검수상태",
        highlight_error_cells=True,
    )

    # Sheet 4: 맞춤법_및_긴문장축약
    ws_s4 = wb.create_sheet(title="맞춤법_및_긴문장축약")
    sheet4_cols = [
        "강좌명",
        "강사명",
        "섹션",
        "글자수",
        "원문문장",
        "제안1_서술형",
        "제안2_개조식",
        "생성엔진",
        "최종문장",
    ]
    _write_styled_sheet(
        ws_s4,
        crosscheck_result.get("long_sentence_df", pd.DataFrame()),
        columns_to_export=sheet4_cols,
        status_col_name="생성엔진",
        highlight_error_cells=False,
    )

    # Sheet 5: AI_초단축강의명제안
    ws_s5 = wb.create_sheet(title="AI_초단축강의명제안")
    ai_df = crosscheck_result.get("ai_titles_df", pd.DataFrame())
    sheet5_cols = [
        "순번",
        "과목",
        "학년",
        "강사명",
        "현재강좌명",
        "임팩트점수",
        "진단비평",
        "제안1_초단축직관형",
        "제안2_초단축시그니처",
        "제안3_초단축킬러형",
        "생성엔진",
        "최종선택강좌명",
    ]
    _write_styled_sheet(
        ws_s5,
        ai_df,
        columns_to_export=sheet5_cols,
        status_col_name="생성엔진",
        highlight_error_cells=False,
    )

    # Sheet 6: 판매중지_제외강좌목록
    ws_s6 = wb.create_sheet(title="판매중지_제외강좌목록")
    stopped_df = crosscheck_result.get("stopped_rms_df", pd.DataFrame())
    sheet6_cols = [
        "No",
        "강좌코드",
        "영역",
        "과목",
        "학년",
        "강사",
        "강좌명",
        "수업시간",
        "개강일",
        "종강일",
        "수강료",
        "판매",
    ]
    _write_styled_sheet(
        ws_s6,
        stopped_df,
        columns_to_export=sheet6_cols,
        status_col_name="판매",
        highlight_error_cells=False,
    )

    # Sheet 7: PDF진도표상세검수
    ws_s7 = wb.create_sheet(title="PDF진도표상세검수")
    sheet7_cols = [
        "PDF_순번",
        "PDF_페이지",
        "검수상태",
        "RMS_No",
        "강좌코드",
        "과목",
        "학년",
        "강사명",
        "PDF 강좌명",
        "수업시간",
        "바자관시수",
        "RMS 개강~종강(회차)",
        "PDF 개요 개강~종강(회차)",
        "PDF 1월진도표(첫날~끝날/회차)",
        "1월 진도표 추출 날짜 리스트",
        "2월표 내 1월 날짜",
        "진도표 요일 정합성",
        "RMS 수강료",
        "PDF 수강료",
        "오류 상세 내역",
        "자동 생성된 수정 요청 사항 요약",
    ]
    _write_styled_sheet(
        ws_s7,
        crosscheck_result.get("pdf_schedule_df", pd.DataFrame()),
        columns_to_export=sheet7_cols,
        status_col_name="검수상태",
        highlight_error_cells=True,
    )

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


EXPORT_DISPLAY_COLUMNS = [
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
    "완납&분납인원",
    "결제율(%)",
    "수업시간 비교",
    "바자관시수 비교",
    "오류 상세 내역",
    "자동 생성된 수정 요청 사항 요약",
]


def export_crosscheck_to_csv_bytes(comparison_df: pd.DataFrame) -> bytes:
    """구글 시트 및 엑셀에서 한글이 절대 깨지지 않는 UTF-8-SIG 인코딩 CSV 바이트를 반환합니다."""
    if comparison_df is None or comparison_df.empty:
        return "".encode("utf-8-sig")
    cols = [c for c in EXPORT_DISPLAY_COLUMNS if c in comparison_df.columns]
    sub = comparison_df[cols].copy()
    return sub.to_csv(index=False).encode("utf-8-sig")


def export_crosscheck_to_tsv_string(comparison_df: pd.DataFrame) -> str:
    """구글 시트(sheets.new)에 Ctrl+V로 즉시 붙여넣기할 수 있는 TSV 문자열을 반환합니다."""
    if comparison_df is None or comparison_df.empty:
        return ""
    cols = [c for c in EXPORT_DISPLAY_COLUMNS if c in comparison_df.columns]
    sub = comparison_df[cols].copy()
    # 셀 내부 탭/줄바꿈 제거하여 구글 시트 행열 밀림 방지
    for c in cols:
        sub[c] = (
            sub[c]
            .fillna("")
            .astype(str)
            .apply(lambda s: re.sub(r"[\t\r\n]+", " ", s).strip())
        )
    return sub.to_csv(index=False, sep="\t")


def export_results_to_google_sheets(
    sheet_url: str,
    crosscheck_result: Dict[str, Any],
    target_ym: str = "2027-01",
    tab_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    사용자가 입력한 Google Sheets URL의 시트에 `검수결과_YYMM` 탭(및 `오류및수정필요_YYMM` 탭)으로
    검수 결과지를 직접 업로드합니다.
    """
    import gspread
    import pydata_google_auth

    m = re.search(r"/d/([a-zA-Z0-9-_]+)", str(sheet_url or "").strip())
    if not m:
        raise ValueError("유효한 구글 시트 URL이 아닙니다. URL에 '/d/시트ID'가 포함되어야 합니다.")
    sheet_id = m.group(1)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = pydata_google_auth.get_user_credentials(scopes, auth_local_webserver=True)
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheet_id)

    ym_clean = re.sub(r"[^\d]", "", str(target_ym))
    yymm = ym_clean[-4:] if len(ym_clean) >= 4 else "2701"
    main_tab = (tab_name or f"검수결과_{yymm}").strip()
    err_tab = f"오류및수정필요_{yymm}"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    comparison_df: pd.DataFrame = crosscheck_result.get("comparison_df", pd.DataFrame())
    error_df: pd.DataFrame = crosscheck_result.get("error_df", pd.DataFrame())

    worksheets_map = {ws.title.strip(): ws for ws in sh.worksheets()}
    updated_tabs: List[str] = []
    logs: List[str] = []

    def _upload_df_to_tab(target_tab_title: str, df_to_up: pd.DataFrame) -> None:
        if df_to_up is None or df_to_up.empty:
            return
        cols = [c for c in EXPORT_DISPLAY_COLUMNS if c in df_to_up.columns]
        sub = df_to_up[cols].copy()
        sub["검수일시"] = now_str
        matrix = [list(sub.columns)] + sub.fillna("").astype(str).values.tolist()

        if target_tab_title in worksheets_map:
            ws = worksheets_map[target_tab_title]
            ws.clear()
            if ws.row_count < len(matrix) + 5 or ws.col_count < len(matrix[0]) + 2:
                ws.resize(rows=len(matrix) + 10, cols=len(matrix[0]) + 2)
        else:
            ws = sh.add_worksheet(
                title=target_tab_title,
                rows=len(matrix) + 20,
                cols=len(matrix[0]) + 2,
            )

        ws.update(values=matrix, range_name="A1")
        try:
            last_hdr_a1 = gspread.utils.rowcol_to_a1(1, len(matrix[0]))
            ws.format(
                f"A1:{last_hdr_a1}",
                {
                    "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
                    "textFormat": {
                        "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                        "bold": True,
                    },
                    "horizontalAlignment": "CENTER",
                },
            )
        except Exception:
            pass

        updated_tabs.append(target_tab_title)
        logs.append(f"✅ [{target_tab_title}] 탭에 {len(sub)}건 업로드 완료")

    _upload_df_to_tab(main_tab, comparison_df)
    if not error_df.empty:
        _upload_df_to_tab(err_tab, error_df)

    return {
        "spreadsheet_title": sh.title,
        "spreadsheet_url": sh.url,
        "updated_tabs": updated_tabs,
        "logs": logs,
    }

