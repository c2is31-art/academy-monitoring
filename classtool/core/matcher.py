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
from core.cumulative_db import (
    build_snapshot_pivot_table,
    compare_two_snapshots,
    format_short_snapshot_header,
    format_snapshot_label,
    get_available_snapshots,
    load_enrollment_db,
    normalize_snapshot_key,
)
from core.exporter import build_all_in_one_master_df

DEFAULT_MAPPING_LIST: List[Dict[str, str]] = [
    {
        "rms_name": "리아",
        "sheet_name": "박소현(리아)",
        "canonical": "리아",
        "memo": "박소현(리아) ↔ 리아 (영어)",
    },
    {
        "rms_name": "박소현(리아)",
        "sheet_name": "리아",
        "canonical": "리아",
        "memo": "박소현(리아) ↔ 리아 (영어)",
    },
    {
        "rms_name": "박소현(리아)",
        "sheet_name": "박소현(리아)",
        "canonical": "리아",
        "memo": "박소현(리아) ↔ 리아 (영어)",
    },
    {
        "rms_name": "박소현(셀리나)",
        "sheet_name": "박소현(셀리나)",
        "canonical": "박소현(셀리나)",
        "memo": "박소현(셀리나) ↔ 박소현(셀리나) (영어)",
    },
    {
        "rms_name": "박소현(셀리나)",
        "sheet_name": "셀리나",
        "canonical": "박소현(셀리나)",
        "memo": "셀리나 ↔ 박소현(셀리나) (영어)",
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
    {
        "rms_name": "차현호",
        "sheet_name": "차쌤",
        "canonical": "차현호",
        "memo": "차쌤 ↔ 차현호 (사회탐구)",
    },
    {
        "rms_name": "민정",
        "sheet_name": "민 정",
        "canonical": "민정",
        "memo": "민 정 ↔ 민정 (사회탐구)",
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

        # 리아 강사 별칭 정규화 ('박소현(리아)', '리아(박소현)', '리아T', '리아') -> '리아'
        s_ns = re.sub(r"\s+", "", s_clean)
        if any(k in s_ns for k in ["박소현(리아)", "리아(박소현)", "리아T"]) or s_ns == "리아":
            return "리아"
        # 셀리나 강사 별칭 정규화 ('박소현(셀리나)', '셀리나(박소현)', '셀리나', 'Selina') -> '박소현(셀리나)'
        if any(k in s_ns for k in ["박소현(셀리나)", "셀리나(박소현)", "셀리나", "Selina"]):
            return "박소현(셀리나)"

        if s_clean in self.lookup:
            return self.lookup[s_clean]
        if s_ns in self.lookup:
            return self.lookup[s_ns]

        # 2~4글자 순수 한글 이름 내부 공백 자동 제거 ('민 정' -> '민정', '차 현 호' -> '차현호')
        if re.match(r"^[가-힣]{2,4}$", s_ns):
            return s_ns

        return s_clean

    def match_instructors(self, name_a: Any, name_b: Any, text_a: str = "", text_b: str = "") -> bool:
        norm_a = self.normalize(name_a)
        norm_b = self.normalize(name_b)
        if not norm_a or not norm_b:
            return False

        # --- [박소현(리아) vs 박소현(셀리나) 동명이인 교차 키워드 판별] ---
        combo_a = f"{name_a} {norm_a} {text_a}".upper()
        combo_b = f"{name_b} {norm_b} {text_b}".upper()

        is_lia_a = ("리아" in norm_a) or ("박소현(리아)" in str(name_a)) or any(k in combo_a for k in ["MASTER", "한대부고 영어", "한대부고영어"])
        is_lia_b = ("리아" in norm_b) or ("박소현(리아)" in str(name_b)) or any(k in combo_b for k in ["MASTER", "한대부고 영어", "한대부고영어"])

        is_selina_a = ("셀리나" in norm_a) or ("박소현(셀리나)" in str(name_a)) or any(k in combo_a for k in ["셀리나", "SELINA", "명작과 필연", "명작과필연", "이제 시작이야", "이제시작이야", "명작", "필연"])
        is_selina_b = ("셀리나" in norm_b) or ("박소현(셀리나)" in str(name_b)) or any(k in combo_b for k in ["셀리나", "SELINA", "명작과 필연", "명작과필연", "이제 시작이야", "이제시작이야", "명작", "필연"])

        if is_lia_a or is_lia_b or is_selina_a or is_selina_b:
            # 리아 vs 셀리나 상호 충돌은 무조건 False
            if is_lia_a and is_selina_b:
                return False
            if is_selina_a and is_lia_b:
                return False
            if is_lia_a and is_lia_b:
                return True
            if is_selina_a and is_selina_b:
                return True
            # 한쪽만 키워드가 판별된 경우: 상대방 이름이 '박소현' 또는 해당 별칭이면 일치 허용
            if is_lia_a and not is_selina_b and ("박소현" in str(name_b) or norm_b == "리아"):
                return True
            if is_lia_b and not is_selina_a and ("박소현" in str(name_a) or norm_a == "리아"):
                return True
            if is_selina_a and not is_lia_b and ("박소현" in str(name_b) or norm_b == "박소현(셀리나)"):
                return True
            if is_selina_b and not is_lia_a and ("박소현" in str(name_a) or norm_a == "박소현(셀리나)"):
                return True
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

        def _extract_tokens(n_str: str) -> Set[str]:
            tokens = {n_str.strip()}
            m = re.match(r"^([^\(]+)\(([^\)]+)\)$", n_str.strip())
            if m:
                tokens.add(m.group(1).strip())
                tokens.add(m.group(2).strip())
            return {re.sub(r"\s+", "", t) for t in tokens if len(t) >= 2}

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


def extract_leading_brackets(title: str) -> List[str]:
    """강좌명 앞머리에 위치한 모든 [대괄호태그]를 순서대로 추출합니다."""
    if not title:
        return []
    m = re.match(r"^(\s*\[[^\]]+\])+", str(title))
    if m:
        return [b.strip() for b in re.findall(r"\[[^\]]+\]", m.group(0))]
    return []


def verify_course_title_match(rms_title: str, pdf_title: str) -> Dict[str, Any]:
    """
    RMS 강좌명과 PDF 강좌명이 앞단 대괄호([특강], [언어,문학] 등)를 포함하여
    완전히 동일한지 정밀 대조 검수를 수행합니다.
    """
    if not pdf_title or str(pdf_title).strip() in ("", "-"):
        return {
            "status": "⚪ PDF 미수록",
            "match_status": "⚪ PDF 미수록",
            "is_exact_match": False,
            "has_bracket_error": False,
            "error_detail": "",
            "fix_request": "",
            "rms_brackets": [],
            "pdf_brackets": [],
        }

    r_clean = re.sub(r"\s+", " ", str(rms_title)).strip()
    p_clean = re.sub(r"\s+", " ", str(pdf_title)).strip()

    r_brackets = extract_leading_brackets(r_clean)
    p_brackets = extract_leading_brackets(p_clean)

    # 1. 대괄호 포함 100% 완전 일치
    if r_clean == p_clean:
        return {
            "status": "✅ 완전 일치",
            "match_status": "✅ 완전 일치",
            "is_exact_match": True,
            "has_bracket_error": False,
            "error_detail": "",
            "fix_request": "",
            "rms_brackets": r_brackets,
            "pdf_brackets": p_brackets,
        }

    # 2. 불일치 분석
    bracket_diff = (r_brackets != p_brackets)

    r_body = r_clean[len("".join(r_brackets)):].strip()
    p_body = p_clean[len("".join(p_brackets)):].strip()
    body_diff = (r_body != p_body)

    r_br_str = "".join(r_brackets) if r_brackets else "(대괄호 없음)"
    p_br_str = "".join(p_brackets) if p_brackets else "(대괄호 없음)"

    if bracket_diff and body_diff:
        match_status = "🔴 대괄호·문구 불일치"
        err_detail = f"강좌명 대괄호 및 문구 불일치 (RMS: '{r_clean}' ≠ PDF: '{p_clean}')"
        fix_req = f"PDF 강좌명을 RMS와 대괄호 포함 100% 일치하도록 수정: '{r_clean}'"
    elif bracket_diff:
        match_status = "🔴 대괄호 불일치"
        err_detail = f"강좌명 앞단 대괄호 불일치 (RMS: {r_br_str} ≠ PDF: {p_br_str})"
        fix_req = f"PDF 강좌명 앞단 대괄호를 RMS('{r_br_str}')와 동일하게 수정: '{r_clean}'"
    else:
        match_status = "🟡 문구 표기 차이"
        err_detail = f"강좌명 본문 문구 차이 (RMS: '{r_body}' ≠ PDF: '{p_body}')"
        fix_req = f"PDF 강좌명을 RMS('{r_clean}')와 일치하도록 수정"

    return {
        "status": match_status,
        "match_status": match_status,
        "is_exact_match": False,
        "has_bracket_error": bracket_diff,
        "error_detail": err_detail,
        "fix_request": fix_req,
        "rms_brackets": r_brackets,
        "pdf_brackets": p_brackets,
    }


SUBJECT_ALIAS_MAP: Dict[str, str] = {
    "세시지": "세계시민과지리",
    "화반세": "화학반응의세계",
    "현사윤": "현대사회와윤리",
    "사문": "사회와문화",
    "윤사": "윤리와사상",
    "한지": "한국지리",
    "세지": "세계지리",
    "동사": "동아시아사",
    "정법": "정치와법",
    "생윤": "생활과윤리",
    "물1": "물리학",
    "화1": "화학",
    "생1": "생명과학",
    "지1": "지구과학",
    "물2": "물리학",
    "화2": "화학",
    "생2": "생명과학",
    "지2": "지구과학",
    "물": "물리학",
    "화": "화학",
    "생": "생명과학",
    "지": "지구과학",
}

MATCH_CORE_KEYWORDS: List[str] = [
    "개념",
    "특강",
    "BEGIN",
    "DEEP",
    "정규",
    "8주",
    "내신",
    "진로",
    "일반선택",
    "LEAD IN",
    "추월차선",
    "마스터",
    "DOPAMINE",
    "QUANTUM",
    "유전",
    "수능",
    "파이널",
    "심화",
    "기본",
    "실력",
]


def normalize_subject(subj: Any) -> str:
    """과목명 정규화: 약칭 및 개정 전후 과목명을 표준 과목명으로 매핑합니다."""
    if not subj:
        return ""
    s = str(subj).strip()
    s_clean = re.sub(r"[\s\[\]()]+", "", s)
    if s_clean in SUBJECT_ALIAS_MAP:
        return SUBJECT_ALIAS_MAP[s_clean]
    for alias in sorted(SUBJECT_ALIAS_MAP.keys(), key=len, reverse=True):
        if alias in s_clean:
            return SUBJECT_ALIAS_MAP[alias]
    return s_clean


def match_subjects(subj_a: Any, subj_b: Any) -> bool:
    """두 과목명이 동일하거나 상호 정규화/포함 관계인지 확인합니다."""
    if not subj_a or not subj_b:
        return False
    norm_a = normalize_subject(subj_a)
    norm_b = normalize_subject(subj_b)
    if not norm_a or not norm_b:
        return False
    if norm_a == norm_b or norm_a in norm_b or norm_b in norm_a:
        return True
    return False


def calculate_inquiry_match_score(
    r_row: pd.Series,
    p_row: pd.Series,
    mapper: InstructorMapper,
) -> float:
    """
    RMS 강좌와 PDF 강좌 간의 정합성 가중치 점수를 산출합니다.
    - 강사 일치 전제 (불일치 시 0.0)
    - 수업 요일 일치: +30점 (예: '일', '수', '금', '토')
    - 과목 정규화 일치: +25점 (세시지 ↔ 세계시민과지리 등)
    - 시작 시간 일치: +20점
    - 핵심 키워드 일치: 키워드당 +10점 ('개념', '특강', 'BEGIN', 'DEEP', '정규', '8주', '내신' 등)
    - 대상학년 일치: +10점
    - 강좌명/슬로건 유사도(0~100): * 0.3 (최대 30점)
    """
    r_inst = str(r_row.get("강사", "")).strip()
    p_inst_orig = str(p_row.get("강사명_원본", "")).strip()
    p_inst_clean = str(p_row.get("강사명", "")).strip()
    r_title = str(r_row.get("강좌명", "")).strip()
    p_title = f"{p_row.get('강좌명', '')} {p_row.get('슬로건', '')} {p_row.get('원문전체텍스트', '')[:300]}"

    if not (
        mapper.match_instructors(r_inst, p_inst_orig, r_title, p_title)
        or mapper.match_instructors(r_inst, p_inst_clean, r_title, p_title)
    ):
        return 0.0

    score = 100.0  # 기본 강사 일치 점수

    # 1. 수업 요일 일치 (+30점)
    r_days = str(r_row.get("수업요일", "")).strip()
    p_days = str(p_row.get("수업요일", "")).strip()
    if r_days and p_days:
        r_day_set = set(re.findall(r"[월화수목금토일]", r_days))
        p_day_set = set(re.findall(r"[월화수목금토일]", p_days))
        if r_day_set and p_day_set and (r_day_set == p_day_set or r_day_set.issubset(p_day_set) or p_day_set.issubset(r_day_set)):
            score += 30.0

    # 2. 과목 정규화 일치 (+25점)
    r_subj = str(r_row.get("과목", "")).strip()
    p_subj = str(p_row.get("과목", "")).strip()
    if r_subj and p_subj and match_subjects(r_subj, p_subj):
        score += 25.0
    elif r_subj and p_subj:
        p_title_full = f"{p_row.get('강좌명', '')} {p_row.get('강사명_원본', '')}"
        if normalize_subject(r_subj) in p_title_full:
            score += 20.0

    # 3. 시작 시간 일치 (+20점)
    r_start = str(r_row.get("전체시작시간", "") or r_row.get("시작시간", "")).strip()
    p_start = str(p_row.get("전체시작시간", "") or p_row.get("시작시간", "")).strip()
    if r_start and p_start and r_start == p_start:
        score += 20.0

    # 4. 핵심 키워드 일치 (키워드당 +10점)
    r_full_text = f"{r_row.get('강좌명', '')} {r_row.get('과목', '')}".upper()
    p_full_text = f"{p_row.get('강좌명', '')} {p_row.get('슬로건', '')} {p_row.get('과목', '')}".upper()
    for kw in MATCH_CORE_KEYWORDS:
        kw_u = kw.upper()
        if kw_u in r_full_text and kw_u in p_full_text:
            score += 10.0

    # 5. 대상학년 일치 (+10점)
    r_grade = str(r_row.get("학년", "")).strip()
    p_grade = str(p_row.get("대상학년", "")).strip()
    if r_grade and p_grade and (r_grade in p_grade or p_grade in r_grade):
        score += 10.0

    # 6. 강좌명/슬로건 유사도 (최대 30점)
    r_pure = str(r_row.get("순수강좌명", "")).strip()
    p_pure = str(p_row.get("순수강좌명", "")).strip()
    p_slogan = str(p_row.get("슬로건", "")).strip()
    sim = max(
        compute_title_similarity(r_pure, p_pure),
        compute_title_similarity(r_pure, p_slogan),
    )
    score += (sim * 0.3)

    return score


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

    # 6. 강좌명 및 앞단 대괄호 완전 일치 정밀 검수
    title_verif = verify_course_title_match(r_title, p_title)
    if not title_verif["is_exact_match"]:
        if title_verif["has_bracket_error"]:
            error_tags.append("🔴 [강좌명 대괄호 불일치]")
            error_details.append(title_verif["error_detail"])
            error_columns.add("강좌명")
            fix_requests.append(title_verif["fix_request"])
        else:
            error_tags.append("🟡 [강좌명 문구 차이]")
            error_details.append(title_verif["error_detail"])
            error_columns.add("강좌명")
            fix_requests.append(title_verif["fix_request"])

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
        "title_match_status": title_verif["match_status"],
        "title_has_bracket_error": title_verif["has_bracket_error"],
        "title_is_exact_match": title_verif["is_exact_match"],
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
        # [1차 정확 매칭]: 강사명 일치 + 수업요일 + 시작시간 일치 (calculate_inquiry_match_score 종합 가중치로 최적 선택)
        for r_idx, r_row in rms.iterrows():
            best_p_idx = None
            best_score = -1.0

            for p_idx, p_row in pdf_df.iterrows():
                if p_idx in used_pdf_indices:
                    continue
                sc = calculate_inquiry_match_score(r_row, p_row, mapper)
                if sc <= 0:
                    continue

                r_days = str(r_row.get("수업요일", "")).strip()
                p_days = str(p_row.get("수업요일", "")).strip()
                r_start = str(r_row.get("전체시작시간", "") or r_row.get("시작시간", "")).strip()
                p_start = str(p_row.get("전체시작시간", "") or p_row.get("시작시간", "")).strip()

                if r_days and p_days and r_days == p_days and r_start and p_start and r_start == p_start:
                    if sc > best_score:
                        best_score = sc
                        best_p_idx = p_idx

            if best_p_idx is not None:
                rms_to_pdf_idx[r_idx] = best_p_idx
                pdf_to_rms_idx[best_p_idx] = r_idx
                used_pdf_indices.add(best_p_idx)

        # [2차 가중치 종합 매칭]: 동일 강사 다중 강좌 최적 쌍 매칭 (과목 정규화 + 핵심 키워드 + 요일 가중치)
        candidate_pairs = []
        for r_idx, r_row in rms.iterrows():
            if r_idx in rms_to_pdf_idx:
                continue
            for p_idx, p_row in pdf_df.iterrows():
                if p_idx in used_pdf_indices:
                    continue
                sc = calculate_inquiry_match_score(r_row, p_row, mapper)
                if sc >= 120.0:  # 강사 일치 + 요일/과목/키워드 일치
                    candidate_pairs.append((sc, r_idx, p_idx))

        candidate_pairs.sort(key=lambda x: x[0], reverse=True)
        for sc, r_idx, p_idx in candidate_pairs:
            if r_idx in rms_to_pdf_idx or p_idx in used_pdf_indices:
                continue
            rms_to_pdf_idx[r_idx] = p_idx
            pdf_to_rms_idx[p_idx] = r_idx
            used_pdf_indices.add(p_idx)

        # [3차 보조 매칭]: 특정 강사의 잔여 미매칭 강좌가 RMS 1건, PDF 1건으로 유일한 경우 자동 매칭
        for p_idx, p_row in pdf_df.iterrows():
            if p_idx in used_pdf_indices:
                continue
            p_inst_orig = str(p_row.get("강사명_원본", "")).strip()
            p_inst_clean = str(p_row.get("강사명", "")).strip()

            candidate_r_indices = []
            for r_idx, r_row in rms.iterrows():
                if r_idx in rms_to_pdf_idx:
                    continue
                r_inst = str(r_row["강사"]).strip()
                r_title = str(r_row.get("강좌명", "")).strip()
                p_title = f"{p_row.get('강좌명', '')} {p_row.get('슬로건', '')}"
                if mapper.match_instructors(r_inst, p_inst_orig, r_title, p_title) or mapper.match_instructors(r_inst, p_inst_clean, r_title, p_title):
                    candidate_r_indices.append(r_idx)

            if len(candidate_r_indices) == 1:
                r_idx = candidate_r_indices[0]
                rms_to_pdf_idx[r_idx] = p_idx
                pdf_to_rms_idx[p_idx] = r_idx
                used_pdf_indices.add(p_idx)
                print(f"[Matcher 3차] 단일 강사 1:1 자동 매칭 성공: RMS '{rms.loc[r_idx, '강좌명']}' ↔ PDF '{p_row.get('강좌명')}' (강사: {p_inst_clean})")

        # [4차 유사도 Fallback 매칭]: 남은 강좌 중 점수 115.0 이상 최적 쌍 매칭
        fallback_pairs = []
        for r_idx, r_row in rms.iterrows():
            if r_idx in rms_to_pdf_idx:
                continue
            for p_idx, p_row in pdf_df.iterrows():
                if p_idx in used_pdf_indices:
                    continue
                sc = calculate_inquiry_match_score(r_row, p_row, mapper)
                if sc >= 115.0:
                    fallback_pairs.append((sc, r_idx, p_idx))

        fallback_pairs.sort(key=lambda x: x[0], reverse=True)
        for sc, r_idx, p_idx in fallback_pairs:
            if r_idx in rms_to_pdf_idx or p_idx in used_pdf_indices:
                continue
            rms_to_pdf_idx[r_idx] = p_idx
            pdf_to_rms_idx[p_idx] = r_idx
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
    kpi_title_mismatch = 0
    kpi_bracket_mismatch = 0
    kpi_title_exact = 0

    for r_idx, r_row in rms.iterrows():
        r_no = int(r_row.get("No", r_idx + 1) or (r_idx + 1))
        r_symbol = str(r_row.get("수강기호", "") or "").strip()
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

        r_sale_status = str(r_row.get("강의상태", "") or r_row.get("판매", "")).strip()
        if not r_sale_status or ("판매" in r_sale_status and "중지" not in r_sale_status):
            r_sale_status = "판매 중"
        elif "중지" in r_sale_status:
            r_sale_status = "[판매중지]"

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

            title_status = ver.get("title_match_status", "⚪ 미확인")
            if ver.get("title_is_exact_match"):
                kpi_title_exact += 1
            else:
                kpi_title_mismatch += 1
                if ver.get("title_has_bracket_error"):
                    kpi_bracket_mismatch += 1

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
            title_status = "⚪ PDF 미수록"
            kpi_missing_pdf += 1
            p_fee_int = 0
            p_sess_int = 0
            p_jan_sess_int = 0

        comparison_rows.append(
            {
                "검수상태": status_badge,
                "No": r_no,
                "강의상태": r_sale_status,
                "수강기호": r_symbol,
                "강좌코드": r_code,
                "영역": r_area,
                "과목": r_subj,
                "학년": r_grade,
                "강사명": inst_display,
                "RMS_강사명": r_inst,
                "정규화강사명": r_norm_inst,
                "강좌명 일치여부": title_status,
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

    # ---------------------------------------------------------
    # RMS ↔ PDF 미매칭 사유 디버깅 콘솔 로그 출력
    # ---------------------------------------------------------
    if not missing_pdf_df.empty:
        print(f"\n[RMS ↔ PDF 미매칭 분석] 총 {len(missing_pdf_df)}건의 RMS 강좌가 PDF 미수록으로 분류됨:")
        pdf_instructors_normalized = set()
        if not pdf_df.empty:
            for p_inst in pdf_df["강사명"].dropna():
                p_norm = mapper.normalize(p_inst)
                pdf_instructors_normalized.add(p_norm)
                pdf_instructors_normalized.add(re.sub(r"\s+", "", p_norm))

        for _, r_row in missing_pdf_df.iterrows():
            r_inst = str(r_row.get("RMS_강사명", "")).strip()
            r_norm = mapper.normalize(r_inst)
            r_norm_ns = re.sub(r"\s+", "", r_norm)
            r_title = str(r_row.get("RMS 강좌명", "")).strip()
            r_code = str(r_row.get("강좌코드", "")).strip()

            inst_in_pdf = bool(
                r_norm in pdf_instructors_normalized or r_norm_ns in pdf_instructors_normalized
            )
            if not inst_in_pdf:
                reason = "강사명 미발견 (PDF 내 해당 강사 강의계획서 미추출)"
            else:
                reason = "강사명은 PDF에 존재하나 요일/시간/강좌명 유사도 미달 또는 다른 분반에 우선 매칭됨"
            print(f"  - [{r_code}] 강사: {r_inst} | 강좌명: {r_title} -> 사유: {reason}")

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
                        s_title = str(s_row.get("강좌명", "")).strip()
                        p_title = str(p_row.get("강좌명", "")).strip()
                        if mapper.match_instructors(s_inst, p_inst, s_title, p_title):
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
                            "검수상태": status_badge,
                            "강의상태": "[판매중지]",
                            "수강기호": str(stopped_matched_row.get("수강기호", "") or "").strip(),
                            "과목": p_subj,
                            "학년": p_grade,
                            "강사명": p_inst,
                            "PDF 강좌명": p_title,
                            "RMS 강좌명": r_title,
                            "강좌코드": r_code,
                            "RMS_No": r_no,
                            "RMS 판매상태": "[판매중지]",
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
                    "강의상태": "[판매중지]" if "판매중지" in status_badge else ("미개설" if "미개설" in status_badge else "판매 중"),
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
    # PDF 본문 텍스트 국어 맞춤법, 문장부호('■') 통일 & 긴 문장 축약 분석
    # ---------------------------------------------------------
    all_spelling_issues: List[Dict[str, Any]] = []
    all_bullet_issues: List[Dict[str, Any]] = []
    all_disallowed_punct_issues: List[Dict[str, Any]] = []
    all_long_proposals: List[Dict[str, Any]] = []
    all_combined_proofread: List[Dict[str, Any]] = []
    course_unified_texts: Dict[Tuple[str, str], str] = {}

    if not pdf_df.empty:
        for p_idx, p_row in pdf_df.iterrows():
            c_title = str(p_row.get("강좌명", "")).strip()
            c_inst = str(p_row.get("강사명_원본", "") or p_row.get("강사명", "")).strip()
            raw_txts = p_row.get("강좌_텍스트리스트", [])
            if not raw_txts and p_row.get("원문전체텍스트"):
                raw_txts = [str(p_row.get("원문전체텍스트", ""))]

            analysis_res = analyze_course_text_content(
                course_name=c_title,
                instructor=c_inst,
                raw_texts=raw_txts,
                max_char_threshold=max_char_threshold,
                gemini_api_key=None,
            )
            all_spelling_issues.extend(analysis_res.get("맞춤법_상세목록", []))
            all_long_proposals.extend(analysis_res.get("긴문장_상세목록", []))
            all_combined_proofread.extend(analysis_res.get("통합교정_상세목록", []))
            course_unified_texts[(c_inst, c_title)] = analysis_res.get("통합완성본텍스트", "")

    spelling_df = pd.DataFrame(all_spelling_issues)
    bullet_issues_df = pd.DataFrame()
    disallowed_punct_df = pd.DataFrame()
    long_sentence_df = pd.DataFrame(all_long_proposals)
    combined_proofread_df = pd.DataFrame(all_combined_proofread)
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
        "bullet_issues_df": bullet_issues_df,
        "disallowed_punct_df": disallowed_punct_df,
        "long_sentence_df": long_sentence_df,
        "combined_proofread_df": combined_proofread_df,
        "course_unified_texts": course_unified_texts,
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
            "title_mismatch_count": kpi_title_mismatch,
            "title_bracket_mismatch_count": kpi_bracket_mismatch,
            "title_exact_count": kpi_title_exact,
            "spelling_error_count": len(all_spelling_issues),
            "bullet_error_count": 0,
            "disallowed_punct_count": 0,
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


def export_crosscheck_to_excel(
    crosscheck_result: Dict[str, Any],
    target_ym: str = "2027-01",
    snapshot_a_key: Optional[str] = None,
    snapshot_b_key: Optional[str] = None,
) -> bytes:
    """
    openpyxl을 사용하여 아래 시트들로 구성된 고품질 서식 포함 엑셀(.xlsx) 바이트를 생성합니다:
    - Sheet 0 (올인원_전체통합정보): 강좌별 기본/수강생/PDF검수/맞춤법/문장축약/AI강의명 가로 완전 병합 마스터 시트
    - Sheet 1 (매월20일_강사별비교): 선택한 두 시점(전월 20일 vs 당월 20일)의 강사별 1:1 비교표
      * 상단 1행: 비교 요약 배너 (A vs B, 강좌수/결제자수/총수강생 증감, 결제율 변화)
      * 조건부 서식: 결제자 증감 양수(+) 연한 초록/파랑(#D5F5E3), 음수(-) 연한 빨강(#FADBD8), 변동상태 배지
    - Sheet 2 (매월20일_연간누적추이): 매월 20일 기준 연간 시계열 누적 피벗 테이블 (월별 결제자수 추이)
    - Sheet 3 (당월20일_강사별상세): 당월 20일 기준 강사별 상세 순위표 및 강좌수/수강생/결제자/결제율
    - Sheet 4 (전체검수결과_유효강좌): 전 강좌 RMS vs PDF 비교표 (오류 행 연한 빨강, 불일치 셀 강조)
    - Sheet 5 (데이터오류_및_판매중지확인): 오류 및 판매중지 강좌 통합 시트
    - Sheet 6 (맞춤법_및_긴문장축약): 맞춤법 교정 및 긴 문장 축약 제안
    - Sheet 7 (AI_강의명제안): PDF 계획서 기반 스마트 패스 및 15자 이내 핵심 강의명 제안
    - Sheet 8 (판매중지_제외강좌목록): 판매중지 강좌 원본 리스트
    - Sheet 9 (PDF진도표상세검수): PDF 진도표 날짜 및 요일 정합성 검수 내역
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    base_font = Font(name="맑은 고딕", size=9, color="111111")
    bold_font = Font(name="맑은 고딕", size=9, bold=True, color="111111")

    # 배너 스타일
    banner_fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
    banner_font = Font(name="맑은 고딕", size=10, bold=True, color="1B4F72")

    # 사용자 지정 조건부 서식
    fill_error_row = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    font_error_row = Font(name="맑은 고딕", size=9, bold=False, color="5B1A18")
    font_error_cell = Font(name="맑은 고딕", size=9, bold=True, color="C00000")
    fill_error_cell = PatternFill(start_color="F5B7B1", end_color="F5B7B1", fill_type="solid")

    fill_normal_row = PatternFill(start_color="EAFAF1", end_color="EAFAF1", fill_type="solid")
    font_normal_row = Font(name="맑은 고딕", size=9, color="1E5631")

    fill_missing_row = PatternFill(start_color="F4F6F7", end_color="F4F6F7", fill_type="solid")
    font_missing_row = Font(name="맑은 고딕", size=9, color="5D6D7E")

    # 증감 전용 하이라이트 서식
    fill_inc_cell = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
    font_inc_cell = Font(name="맑은 고딕", size=9, bold=True, color="145A32")
    fill_dec_cell = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    font_dec_cell = Font(name="맑은 고딕", size=9, bold=True, color="78281F")

    # 변동상태 전용 배지 서식
    fill_badge_new = PatternFill(start_color="D1F2EB", end_color="D1F2EB", fill_type="solid")
    font_badge_new = Font(name="맑은 고딕", size=9, bold=True, color="0E6251")
    fill_badge_end = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
    font_badge_end = Font(name="맑은 고딕", size=9, color="5D6D7E")

    thin_side = Side(border_style="thin", color="D5D8DC")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # -------------------------------------------------------------------------
    # 헬퍼 함수: 표준 서식 시트 작성기
    # -------------------------------------------------------------------------
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
                lambda x: re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", ", ".join(map(str, x)) if isinstance(x, (list, tuple, set)) else ("" if pd.isna(x) else str(x)))
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

        for col_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = sum(1.8 if ord(ch) > 127 else 1.0 for ch in str(col_name)) + 2
            for val in sub_df[col_name].head(120):
                val_len = sum(1.65 if ord(ch) > 127 else 0.95 for ch in str(val))
                if val_len > max_len:
                    max_len = val_len
            ws.column_dimensions[col_letter].width = min(max(int(max_len) + 3, 10), 55)

    # =========================================================================
    # Sheet 0: 올인원_전체통합정보 (구글 시트 / 엑셀 통합 마스터 뷰)
    # =========================================================================
    ws_s0 = wb.active
    ws_s0.title = "올인원_전체통합정보"
    all_in_one_master_df = build_all_in_one_master_df(crosscheck_result, target_ym=target_ym)
    _write_styled_sheet(
        ws_s0,
        all_in_one_master_df,
        columns_to_export=list(all_in_one_master_df.columns),
        status_col_name="최종검수상태",
        highlight_error_cells=True,
    )

    # =========================================================================
    # Sheet 1: 매월20일_강사별비교
    # =========================================================================
    ws_s1 = wb.create_sheet(title="매월20일_강사별비교")

    avail_snaps = get_available_snapshots()
    norm_target = normalize_snapshot_key(target_ym)

    # 기준 스냅샷 A, B 자동 결정
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
    comp_df = comp_res.get("comparison_df", pd.DataFrame())
    short_a = comp_res.get("short_a", "전월")
    short_b = comp_res.get("short_b", "당월")

    # 1행 비교 요약 배너
    crs_diff = kpi.get("courses_diff", 0)
    crs_sign = "+" if crs_diff > 0 else ""
    paid_diff = kpi.get("paid_diff", 0)
    paid_sign = "+" if paid_diff > 0 else ""
    paid_pct = kpi.get("paid_pct", 0.0)
    paid_pct_sign = "+" if paid_pct > 0 else ""
    stu_diff = kpi.get("students_diff", 0)
    stu_sign = "+" if stu_diff > 0 else ""
    rate_diff = kpi.get("rate_diff", 0.0)
    rate_sign = "+" if rate_diff > 0 else ""

    banner_str = (
        f"[비교 기준: A({kpi.get('label_a', snap_a)}) vs B({kpi.get('label_b', snap_b)}) | "
        f"유효 강좌수: {kpi.get('courses_a', 0)}건 ➔ {kpi.get('courses_b', 0)}건 ({crs_sign}{crs_diff}건) | "
        f"결제완료 증감: {paid_sign}{paid_diff:,}명 ({paid_pct_sign}{paid_pct}%) | "
        f"총수강생 증감: {stu_sign}{stu_diff:,}명 | "
        f"평균 결제율 변화: {rate_sign}{rate_diff}%p]"
    )
    ws_s1.append([banner_str])
    ws_s1.row_dimensions[1].height = 26
    c_b1 = ws_s1.cell(row=1, column=1)
    c_b1.font = banner_font

    ws_s1.append([])  # 2행: 빈 행
    ws_s1.row_dimensions[2].height = 8

    sheet1_cols = [
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

    if not comp_df.empty:
        valid_cols = [c for c in sheet1_cols if c in comp_df.columns]
        ws_s1.append(valid_cols)
        h_row = 3
        ws_s1.row_dimensions[h_row].height = 26
        for c_idx in range(1, len(valid_cols) + 1):
            cell = ws_s1.cell(row=h_row, column=c_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border

        for r_offset, (_, r_data) in enumerate(comp_df.iterrows(), start=h_row + 1):
            row_vals = [r_data.get(c, "") for c in valid_cols]
            ws_s1.append(row_vals)
            ws_s1.row_dimensions[r_offset].height = 21

            p_diff_val = int(r_data.get("결제자_증감_정수", 0))
            status_val = str(r_data.get("변동 상태", ""))

            for c_idx, c_name in enumerate(valid_cols, start=1):
                c_cell = ws_s1.cell(row=r_offset, column=c_idx)
                c_cell.border = cell_border
                c_cell.font = base_font
                c_cell.alignment = Alignment(horizontal="center", vertical="center")

                # 결제자 증감 서식 (양수 연한 초록, 음수 연한 빨강)
                if c_name in ("결제자 증감(명)", "결제자 증감률(%)"):
                    if p_diff_val > 0:
                        c_cell.fill = fill_inc_cell
                        c_cell.font = font_inc_cell
                    elif p_diff_val < 0:
                        c_cell.fill = fill_dec_cell
                        c_cell.font = font_dec_cell
                elif c_name == "변동 상태":
                    if "신규개설" in status_val:
                        c_cell.fill = fill_badge_new
                        c_cell.font = font_badge_new
                    elif "미개설" in status_val or "종강" in status_val:
                        c_cell.fill = fill_badge_end
                        c_cell.font = font_badge_end
                    elif "증가" in status_val:
                        c_cell.fill = fill_inc_cell
                        c_cell.font = font_inc_cell
                    elif "감소" in status_val:
                        c_cell.fill = fill_dec_cell
                        c_cell.font = font_dec_cell

        ws_s1.freeze_panes = "A4"
        ws_s1.auto_filter.ref = f"A3:{get_column_letter(len(valid_cols))}{len(comp_df) + 3}"
        for col_idx, col_name in enumerate(valid_cols, start=1):
            col_letter = get_column_letter(col_idx)
            ws_s1.column_dimensions[col_letter].width = max(len(str(col_name)) * 2 + 3, 11)
    else:
        ws_s1.append(["비교할 스냅샷 데이터가 아직 등록되지 않았습니다."])

    # =========================================================================
    # Sheet 2: 매월20일_연간누적추이
    # =========================================================================
    ws_s2 = wb.create_sheet(title="매월20일_연간누적추이")
    pivot_df = build_snapshot_pivot_table("결제완료 인원")
    ws_s2.append(["[매월 20일 기준 강사별 결제완료 인원 연간 누적 추이 및 월평균 통계]"])
    ws_s2.cell(row=1, column=1).font = banner_font
    ws_s2.row_dimensions[1].height = 24
    ws_s2.append([])

    if not pivot_df.empty:
        p_cols = list(pivot_df.columns)
        ws_s2.append(p_cols)
        h_row = 3
        ws_s2.row_dimensions[h_row].height = 26
        for c_idx in range(1, len(p_cols) + 1):
            cell = ws_s2.cell(row=h_row, column=c_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border

        for r_offset, (_, r_data) in enumerate(pivot_df.iterrows(), start=h_row + 1):
            row_vals = [r_data.get(c, "") for c in p_cols]
            ws_s2.append(row_vals)
            ws_s2.row_dimensions[r_offset].height = 21

            for c_idx, c_name in enumerate(p_cols, start=1):
                c_cell = ws_s2.cell(row=r_offset, column=c_idx)
                c_cell.border = cell_border
                c_cell.font = base_font
                c_cell.alignment = Alignment(horizontal="center", vertical="center")

                if c_name == "최근월(20일) 증감":
                    val_str = str(c_cell.value)
                    if "+" in val_str:
                        c_cell.font = font_inc_cell
                    elif "-" in val_str and val_str != "-":
                        c_cell.font = font_dec_cell

        ws_s2.freeze_panes = "A4"
        ws_s2.auto_filter.ref = f"A3:{get_column_letter(len(p_cols))}{len(pivot_df) + 3}"
        for col_idx, col_name in enumerate(p_cols, start=1):
            col_letter = get_column_letter(col_idx)
            ws_s2.column_dimensions[col_letter].width = max(len(str(col_name)) * 2 + 3, 11)
    else:
        ws_s2.append(["누적 시계열 데이터가 아직 등록되지 않았습니다."])

    # =========================================================================
    # Sheet 3: 당월20일_강사별상세
    # =========================================================================
    ws_s3 = wb.create_sheet(title="당월20일_강사별상세")
    rms_sum = crosscheck_result.get("rms_summary", {}) or {}
    banner_detail = rms_sum.get(
        "headline_banner",
        f"총 강좌수 : {rms_sum.get('total_courses', 0)}건 | 총 수강생 : {rms_sum.get('total_students', 0)}명 (결제완료 : {rms_sum.get('paid_students', 0)}명)"
    )
    ws_s3.append([banner_detail])
    ws_s3.cell(row=1, column=1).font = banner_font
    ws_s3.row_dimensions[1].height = 24
    ws_s3.append([])

    inst_summary_df = crosscheck_result.get("instructor_summary_df", pd.DataFrame())
    sheet3_cols = [
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
        cols = [c for c in sheet3_cols if c in inst_summary_df.columns]
        ws_s3.append(cols)
        h_row = 3
        ws_s3.row_dimensions[h_row].height = 26
        for c_idx in range(1, len(cols) + 1):
            cell = ws_s3.cell(row=h_row, column=c_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = cell_border

        for r_offset, (_, r_data) in enumerate(inst_summary_df.iterrows(), start=h_row + 1):
            row_vals = [r_data[c] for c in cols]
            ws_s3.append(row_vals)
            ws_s3.row_dimensions[r_offset].height = 20
            for c_idx in range(1, len(cols) + 1):
                c_cell = ws_s3.cell(row=r_offset, column=c_idx)
                c_cell.border = cell_border
                c_cell.font = base_font
                c_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_s3.auto_filter.ref = f"A3:{get_column_letter(len(cols))}{len(inst_summary_df) + 3}"
        for col_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(col_idx)
            ws_s3.column_dimensions[col_letter].width = max(len(str(col_name)) * 2 + 4, 12)
    else:
        ws_s3.append(["강사별 상세 데이터가 없습니다."])

    # =========================================================================
    # Sheet 4: 전체검수결과_유효강좌
    # =========================================================================
    ws_s4 = wb.create_sheet(title="전체검수결과_유효강좌")
    sheet4_cols = [
        "검수상태",
        "No",
        "강의상태",
        "수강기호",
        "강좌코드",
        "과목",
        "학년",
        "강사명",
        "강좌명 일치여부",
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
        ws_s4,
        crosscheck_result.get("comparison_df", pd.DataFrame()),
        columns_to_export=sheet4_cols,
        status_col_name="검수상태",
        highlight_error_cells=True,
    )

    # =========================================================================
    # Sheet 5: 데이터오류_및_판매중지확인
    # =========================================================================
    ws_s5 = wb.create_sheet(title="데이터오류_및_판매중지확인")
    error_df = crosscheck_result.get("error_df", pd.DataFrame())
    pdf_stopped_df = crosscheck_result.get("pdf_stopped_courses_df", pd.DataFrame())

    combined_err_list: List[Dict[str, Any]] = []
    if not error_df.empty:
        for _, r in error_df.iterrows():
            d = dict(r)
            if "강의상태" not in d or not d["강의상태"]:
                d["강의상태"] = "판매 중"
            combined_err_list.append(d)

    if not pdf_stopped_df.empty:
        for _, r in pdf_stopped_df.iterrows():
            combined_err_list.append(
                {
                    "검수상태": "⚠️ RMS 판매중지 (폐강 의심)",
                    "No": r.get("RMS_No", 0),
                    "강의상태": "[판매중지]",
                    "수강기호": r.get("수강기호", ""),
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
    sheet5_cols = [
        "검수상태",
        "No",
        "강의상태",
        "수강기호",
        "강좌코드",
        "과목",
        "학년",
        "강사명",
        "강좌명 일치여부",
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
        ws_s5,
        combined_err_df,
        columns_to_export=sheet5_cols,
        status_col_name="검수상태",
        highlight_error_cells=True,
    )

    # =========================================================================
    # Sheet 6: 맞춤법_문장축약
    # =========================================================================
    ws_s6 = wb.create_sheet(title="맞춤법_문장축약")
    sheet6_cols = [
        "강사명",
        "강좌명",
        "교정 구분",
        "발견 항목",
        "PDF 원문 단락",
        "✅ 교정 완료 단락 (오타 수정본)",
        "사유",
    ]
    comb_proof_df = crosscheck_result.get("combined_proofread_df", pd.DataFrame())
    if comb_proof_df is None or comb_proof_df.empty:
        comb_proof_df = crosscheck_result.get("long_sentence_df", pd.DataFrame())
    _write_styled_sheet(
        ws_s6,
        comb_proof_df,
        columns_to_export=sheet6_cols,
        status_col_name="교정 구분",
        highlight_error_cells=False,
    )

    # =========================================================================
    # Sheet 7: AI_강의명제안
    # =========================================================================
    ws_s7 = wb.create_sheet(title="AI_강의명제안")
    ai_df = crosscheck_result.get("ai_titles_df")
    if ai_df is None or (isinstance(ai_df, pd.DataFrame) and ai_df.empty):
        comparison_df_for_ai = crosscheck_result.get("comparison_df", pd.DataFrame())
        if comparison_df_for_ai is not None and not comparison_df_for_ai.empty:
            from core.title_suggester import batch_suggest_course_titles
            courses_list = comparison_df_for_ai.to_dict(orient="records")
            ai_results = batch_suggest_course_titles(
                course_list=courses_list,
                api_key=None,
                direction="스마트 분석",
            )
            ai_df = pd.DataFrame(ai_results)
            crosscheck_result["ai_titles_df"] = ai_df
        else:
            ai_df = pd.DataFrame()

    if ai_df is not None and not ai_df.empty:
        ai_df = ai_df.copy()
        if "판정" not in ai_df.columns and "판정결과" in ai_df.columns:
            ai_df["판정"] = ai_df["판정결과"]
        if "현재 강좌명" not in ai_df.columns:
            ai_df["현재 강좌명"] = ai_df.get("현재강의명", ai_df.get("현재강좌명", ""))
        if "임팩트 점수" not in ai_df.columns and "임팩트점수" in ai_df.columns:
            ai_df["임팩트 점수"] = ai_df["임팩트점수"]
        if "판정 사유 및 피드백" not in ai_df.columns:
            ai_df["판정 사유 및 피드백"] = ai_df.get("판정사유", ai_df.get("진단비평", ""))
        if "⚡ 추천 강의명" not in ai_df.columns:
            ai_df["⚡ 추천 강의명"] = ai_df.get("⚡ 추천 강의명 1 (직관·시그니처형)", ai_df.get("제안1_초단축직관형", "-"))

    sheet7_cols = [
        "순번",
        "판정",
        "과목",
        "학년",
        "강사명",
        "현재 강좌명",
        "임팩트 점수",
        "판정 사유 및 피드백",
        "⚡ 추천 강의명",
    ]
    _write_styled_sheet(
        ws_s7,
        ai_df,
        columns_to_export=sheet7_cols,
        status_col_name="판정",
        highlight_error_cells=False,
    )

    # =========================================================================
    # Sheet 8: 판매중지_제외강좌목록
    # =========================================================================
    ws_s8 = wb.create_sheet(title="판매중지_제외강좌목록")
    stopped_df = crosscheck_result.get("stopped_rms_df", pd.DataFrame())
    sheet8_cols = [
        "No",
        "강의상태",
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
        ws_s8,
        stopped_df,
        columns_to_export=sheet8_cols,
        status_col_name="판매",
        highlight_error_cells=False,
    )

    # =========================================================================
    # Sheet 9: PDF진도표상세검수
    # =========================================================================
    ws_s9 = wb.create_sheet(title="PDF진도표상세검수")
    sheet9_cols = [
        "PDF_순번",
        "PDF_페이지",
        "검수상태",
        "RMS_No",
        "강의상태",
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
        ws_s9,
        crosscheck_result.get("pdf_schedule_df", pd.DataFrame()),
        columns_to_export=sheet9_cols,
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


def export_ai_titles_only_excel(ai_df: pd.DataFrame) -> bytes:
    """
    AI 추천 강의명 제안 결과만을 예쁘게 스타일링하여 단독 .xlsx 바이너리 바이트로 내보냅니다.
    """
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "AI_초단축강의명제안"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    base_font = Font(name="맑은 고딕", size=9, color="111111")
    bold_font = Font(name="맑은 고딕", size=9, bold=True, color="111111")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    fill_rev = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
    font_rev = Font(name="맑은 고딕", size=9, color="B78103", bold=True)
    fill_keep = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    font_keep = Font(name="맑은 고딕", size=9, color="2E7D32", bold=True)

    sheet_cols = [
        "순번",
        "판정",
        "과목",
        "학년",
        "강사명",
        "현재 강좌명",
        "임팩트 점수",
        "판정 사유 및 피드백",
        "⚡ 추천 강의명",
    ]

    if ai_df is None or ai_df.empty:
        ws.append(["데이터가 없습니다."])
    else:
        df_copy = ai_df.copy()
        if "판정" not in df_copy.columns and "판정결과" in df_copy.columns:
            df_copy["판정"] = df_copy["판정결과"]
        if "현재 강좌명" not in df_copy.columns:
            df_copy["현재 강좌명"] = df_copy.get("현재강의명", df_copy.get("현재강좌명", ""))
        if "임팩트 점수" not in df_copy.columns and "임팩트점수" in df_copy.columns:
            df_copy["임팩트 점수"] = df_copy["임팩트점수"]
        if "판정 사유 및 피드백" not in df_copy.columns:
            df_copy["판정 사유 및 피드백"] = df_copy.get("판정사유", df_copy.get("진단비평", ""))
        if "⚡ 추천 강의명" not in df_copy.columns:
            df_copy["⚡ 추천 강의명"] = df_copy.get("⚡ 추천 강의명 1 (직관·시그니처형)", df_copy.get("제안1_초단축직관형", "-"))

        cols = [c for c in sheet_cols if c in df_copy.columns]
        ws.append(cols)
        for col_idx in range(1, len(cols) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[1].height = 26

        for r_idx, (_, r_data) in enumerate(df_copy.iterrows(), start=2):
            status_val = str(r_data.get("판정", r_data.get("판정결과", "")))
            row_vals = [r_data.get(c, "") for c in cols]
            ws.append(row_vals)
            ws.row_dimensions[r_idx].height = 20

            for col_idx, col_name in enumerate(cols, start=1):
                cell = ws.cell(row=r_idx, column=col_idx)
                cell.font = base_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if col_name in ("순번", "판정", "판정결과", "과목", "학년", "강사명", "임팩트 점수", "임팩트점수", "생성엔진") else "left",
                    vertical="center",
                )
                if col_name in ("판정", "판정결과"):
                    if "수정" in status_val:
                        cell.fill = fill_rev
                        cell.font = font_rev
                    elif "유지" in status_val or "패스" in status_val:
                        cell.fill = fill_keep
                        cell.font = font_keep
                elif "추천 강의명" in col_name and cell.value and cell.value != "-":
                    cell.font = bold_font

        # 열 너비 자동 조정
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for c in col:
                val_str = str(c.value or "")
                korean_len = len(val_str.encode("utf-8"))
                max_len = max(max_len, int(korean_len * 0.7) + 2)
            ws.column_dimensions[col_letter].width = min(max(max_len, 10), 45)

    ws.freeze_panes = "A2"
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


