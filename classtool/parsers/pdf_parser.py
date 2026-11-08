"""
강의계획서 PDF 정밀 파서 모듈 (parsers/pdf_parser.py)
- 단일 통합 PDF(예: '27년 1월 단과 강의계획서.pdf') 또는 다중 개별 PDF 파일들을 스캔하여
  '강의개요'가 포함된 페이지부터 다음 '강의개요' 페이지 직전까지를 하나의 강좌 블록으로 묶어 파싱합니다.
- 상단 메타데이터(바자관 시수, 과목, 강사명, 대상학년, 강좌명, 슬로건),
  강의개요 테이블(개강일, 수강기간 및 회차, 수업시간, 휴강일, 수강료 및 회차당 단가, 교재비),
  주차별 진도계획(1월 과정 날짜 리스트, 실제 달력 요일 역산, 진도표 1월 실제회차, 첫수업일, 마지막수업일, 요일오류여부)을 정밀 추출합니다.
"""

from __future__ import annotations

import io
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd
from pypdf import PdfReader

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None


def _extract_page_texts_pymupdf(raw_bytes: bytes) -> List[str]:
    """
    PyMuPDF(pymupdf/fitz)의 words 좌표를 활용하여 Y좌표를 3pt 단위로 양자화(quantize)한 뒤
    X좌표 순으로 정렬하여 원본 레이아웃을 정확하게 보존하는 텍스트 블록을 초고속으로 구성합니다.
    (기존 pdfplumber 대비 15~30배 고속화)
    """
    if pymupdf is None:
        return []

    try:
        doc = pymupdf.open(stream=raw_bytes, filetype="pdf")
        page_texts: List[str] = []
        for page in doc:
            words = page.get_text("words")
            if not words:
                page_texts.append(page.get_text() or "")
                continue

            lines_dict: Dict[int, List[Tuple[float, str]]] = {}
            for w in words:
                y = int(round(w[1] / 3) * 3)
                lines_dict.setdefault(y, []).append((w[0], w[4]))

            sorted_y = sorted(lines_dict.keys())
            reconstructed_lines: List[str] = []
            for y in sorted_y:
                sorted_words = sorted(lines_dict[y], key=lambda x: x[0])
                line_str = " ".join(item[1] for item in sorted_words)
                reconstructed_lines.append(line_str)

            page_texts.append("\n".join(reconstructed_lines))
        return page_texts
    except Exception:
        return []


def _extract_page_texts_fast(raw_bytes: bytes) -> List[str]:
    """
    초고속 pypdf.PdfReader 및 pymupdf 엔진을 결합하여
    40페이지 분량의 PDF를 0.2초 이내에 텍스트로 추출합니다.
    (도형/테이블 좌표 연산으로 수십 초 소요되던 pdfplumber 전면 제거)
    """
    # 1) PyMuPDF 사용 가능 시 0.05초 초고속 우선 시도
    if pymupdf is not None:
        pymupdf_texts = _extract_page_texts_pymupdf(raw_bytes)
        if pymupdf_texts:
            return pymupdf_texts

    # 2) pypdf.PdfReader 고속 텍스트 추출 (순수 파이썬 환경 0.2~0.5초)
    reader = PdfReader(io.BytesIO(raw_bytes))
    return [(p.extract_text() or "") for p in reader.pages]

from parsers.rms_parser import (
    WEEKDAY_NAMES,
    format_short_md_with_weekday,
    get_korean_weekday,
    normalize_baza_hours,
    parse_schedule_text,
    strip_course_brackets,
)

KNOWN_SUBJECT_PREFIXES = (
    "통합사회",
    "통합과학",
    "사회탐구",
    "과학탐구",
    "인문논술",
    "수리논술",
    "과학논술",
    "논구술",
    "한국사",
    "국어",
    "영어",
    "수학",
    "사탐",
    "과탐",
    "탐구",
    "논술",
    "과학",
    "사회",
)


def convert_md_to_ymd(md_raw: str, target_ym: str = "2027-01") -> str:
    """
    '1/4(월)', '1/27(수)', '$1/5(9|)$', '1/12화()' 등의 월/일 문자열에서
    숫자 (M/D)를 추출하여 검수 대상 연월(예: '2027-01') 기준 'YYYY-MM-DD' 포맷으로 변환합니다.
    """
    if not md_raw:
        return ""
    s = str(md_raw).strip()

    m_full = re.search(r"(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})", s)
    if m_full:
        return f"{int(m_full.group(1)):04d}-{int(m_full.group(2)):02d}-{int(m_full.group(3)):02d}"

    m_yy = re.search(r"(\d{2})\s*[-.]\s*(\d{1,2})\s*[-.]\s*(\d{1,2})", s)
    if m_yy:
        yy = int(m_yy.group(1))
        yyyy = 2000 + yy if yy < 70 else 1900 + yy
        return f"{yyyy:04d}-{int(m_yy.group(2)):02d}-{int(m_yy.group(3)):02d}"

    m_md = re.search(r"(\d{1,2})\s*(?:/|월)\s*(\d{1,2})", s)
    if not m_md:
        return ""

    mm = int(m_md.group(1))
    dd = int(m_md.group(2))
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return ""

    base_year = 2027
    base_month = 1
    if re.match(r"^\d{4}-\d{2}$", str(target_ym).strip()):
        base_year, base_month = map(int, str(target_ym).strip().split("-"))

    if base_month == 1 and mm in (11, 12):
        yyyy = base_year - 1
    elif base_month == 12 and mm in (1, 2):
        yyyy = base_year + 1
    else:
        yyyy = base_year

    try:
        datetime(yyyy, mm, dd)
    except ValueError:
        return ""

    return f"{yyyy:04d}-{mm:02d}-{dd:02d}"


def extract_printed_weekday(raw_date_token: str) -> Tuple[str, bool]:
    """
    날짜 토큰(예: '1/4(월)', '1/12화()', '$1/16(E)$')에서
    표기된 요일 문자('월'~'일')와 폰트/OCR 깨짐 여부를 반환합니다.
    반환값: (표기요일_또는_빈문자열, 깨짐여부_bool)
    """
    if not raw_date_token:
        return "", False
    s = str(raw_date_token).strip()
    # 1) 정상 괄호 안 요일: 1/4(월)
    m_paren = re.search(r"\(\s*([월화수목금토일])\s*(?:요일)?\s*\)", s)
    if m_paren:
        return m_paren.group(1), False

    # 2) 괄호 밖 요일 깨짐 형태: 1/12화()
    m_outside = re.search(r"\d{1,2}\s*/\s*\d{1,2}\s*([월화수목금토일])\s*\(\s*\)", s)
    if m_outside:
        return m_outside.group(1), True

    # 3) 괄호는 있으나 한글 요일이 아닌 깨진 기호가 들어있는 경우: 1/5(9|), 1/16(E)
    m_any_paren = re.search(r"\(([^)]*)\)", s)
    if m_any_paren:
        inner = m_any_paren.group(1).strip()
        m_kor = re.search(r"[월화수목금토일]", inner)
        if m_kor:
            return m_kor.group(0), False
        return "", True

    return "", False


def clean_instructor_name(raw_inst: str) -> str:
    """
    '박소현(셀리나)', '심수현T' 등에서 괄호 별칭 및 불필요한 접미사 'T', '선생님'을 제거한 기본 강사명을 반환합니다.
    """
    if not raw_inst:
        return ""
    s = str(raw_inst).strip()
    s = re.sub(r"\([^)]*\)", "", s).strip()
    s = re.sub(r"(?:선생님|강사님|강사|교수님|T)$", "", s, flags=re.IGNORECASE).strip()
    return s


def _parse_header_lines(page_text: str) -> Dict[str, str]:
    """
    강좌 시작 페이지 상단('강의개요' 이전 영역)에서
    - 과정구분, 대상학년, 바자관 시수 (예: '6T', '3T')
    - 슬로건
    - PDF 강좌명
    - 과목, 강사명(및 강사명_원본)
    을 추출합니다.
    """
    info = {
        "상단태그줄": "",
        "과정구분": "",
        "대상학년": "",
        "바자관시수": "-",
        "슬로건": "",
        "강좌명": "",
        "과목": "",
        "강사명_원본": "",
        "강사명": "",
    }
    if not page_text:
        return info

    # 바자관 시수 전체 텍스트 우선 탐색
    m_baza_any = re.search(r"바자관\s*시수\s*(\d+\s*T)", page_text, flags=re.IGNORECASE)
    if m_baza_any:
        info["바자관시수"] = normalize_baza_hours(m_baza_any.group(1))

    parts = re.split(r"^\s*강의\s*개요\s*$", page_text, maxsplit=1, flags=re.MULTILINE)
    header_block = parts[0] if parts else page_text
    lines = [ln.strip() for ln in header_block.splitlines() if ln.strip()]
    if not lines:
        return info

    subj_inst_idx = -1
    subj_pattern = re.compile(
        r"^(" + "|".join(KNOWN_SUBJECT_PREFIXES) + r")\s*(.+)$"
    )
    for idx in range(len(lines) - 1, -1, -1):
        m = subj_pattern.match(lines[idx])
        if m:
            candidate_inst = m.group(2).strip()
            if 1 <= len(candidate_inst) <= 15 and " " not in candidate_inst:
                info["과목"] = m.group(1).strip()
                info["강사명_원본"] = candidate_inst
                info["강사명"] = clean_instructor_name(candidate_inst)
                subj_inst_idx = idx
                break

    pre_lines = lines[:subj_inst_idx] if subj_inst_idx != -1 else lines
    if not pre_lines:
        return info

    first_line = pre_lines[0]
    if any(k in first_line for k in ("바자관", "내신대비", "수능대비", "특강", "정규", "고1", "고2", "고3", "N수")):
        info["상단태그줄"] = first_line
        m_grade = re.search(r"(고1|고2|고3|N수|중3)", first_line)
        if m_grade:
            info["대상학년"] = m_grade.group(1)
        m_baza = re.search(r"바자관\s*시수\s*(\d+\s*T)", first_line, flags=re.IGNORECASE)
        if m_baza:
            info["바자관시수"] = normalize_baza_hours(m_baza.group(1))
        m_cat = re.match(r"^((?:내신대비|수능대비)?\s*(?:특강|정규|단과))", first_line)
        if m_cat:
            info["과정구분"] = re.sub(r"\s+", " ", m_cat.group(1)).strip()
        rem_lines = pre_lines[1:]
    else:
        rem_lines = pre_lines

    if len(rem_lines) == 1:
        info["강좌명"] = rem_lines[0]
    elif len(rem_lines) >= 2:
        info["슬로건"] = rem_lines[0]
        title_lines = rem_lines[1:]
        merged_title = title_lines[0]
        for tl in title_lines[1:]:
            if tl.startswith(("+", "-", ")", "]")) or merged_title.endswith(("+", "-", "(")):
                merged_title += tl
            else:
                merged_title += " " + tl
        info["강좌명"] = re.sub(r"\s+", " ", merged_title).strip()

    return info


def _parse_overview_section(page_text: str, target_ym: str = "2027-01") -> Dict[str, Any]:
    """
    '강의개요' 영역에서
    - PDF 개강일 (날짜 및 요일, OCR/폰트 깨짐 시 수강기간 시작일로 자동 보정)
    - PDF 수강기간 및 회차 (시작일, 종강일, PDF_회차)
    - PDF 수업시간
    - PDF 휴강일, 보강방법
    - PDF 수강료 & 회차당 수강료
    - 교재 및 교재비
    를 정밀 추출합니다.
    """
    res: Dict[str, Any] = {
        "개강일_원본": "",
        "개강일": "",
        "개강일_표기요일": "",
        "개강일_실제요일": "",
        "개강일_깨짐보정여부": False,
        "추천대상": "",
        "수강기간_원본": "",
        "수강시작일": "",
        "수강시작일_표기요일": "",
        "종강일_원본": "",
        "종강일": "",
        "종강일_표기요일": "",
        "종강일_실제요일": "",
        "회차": 0,
        "수업시간_원본": "",
        "휴강일": "",
        "보강방법": "",
        "수강료": 0,
        "회차당수강료": 0,
        "수강료_표시": "",
        "교재내역": "",
        "교재비_합계": 0,
        "교재비_표시": "",
    }
    if not page_text:
        return res

    m_ov = re.search(
        r"강의\s*개요(.*?)(?:강의\s*특징\s*/\s*관리|강의\s*특징|주차별\s*진도\s*계획|$)",
        page_text,
        flags=re.DOTALL,
    )
    ov_text = m_ov.group(1) if m_ov else page_text

    # 1. 수강 기간 (시작일 ~ 종강일 N회) 먼저 추출 (개강일 깨짐 시 보정 기준)
    m_period = re.search(
        r"수강\s*기간\s*(\d{1,2}\s*/\s*\d{1,2}\s*(?:[월화수목금토일]?\s*\([^)]*\))?)\s*[~\-∼～]\s*(\d{1,2}\s*/\s*\d{1,2}\s*(?:[월화수목금토일]?\s*\([^)]*\))?)\s*(\d+)\s*회",
        ov_text,
    )
    if m_period:
        p_start_raw = m_period.group(1).strip()
        p_end_raw = m_period.group(2).strip()
        p_cnt = int(m_period.group(3))

        p_start_ymd = convert_md_to_ymd(p_start_raw, target_ym=target_ym)
        p_end_ymd = convert_md_to_ymd(p_end_raw, target_ym=target_ym)
        p_start_wk, _ = extract_printed_weekday(p_start_raw)
        p_end_wk, _ = extract_printed_weekday(p_end_raw)

        res["수강기간_원본"] = f"{p_start_raw}~{p_end_raw} {p_cnt}회"
        res["수강시작일"] = p_start_ymd
        res["수강시작일_표기요일"] = p_start_wk or get_korean_weekday(p_start_ymd)
        res["종강일_원본"] = p_end_raw
        res["종강일"] = p_end_ymd
        res["종강일_실제요일"] = get_korean_weekday(p_end_ymd)
        res["종강일_표기요일"] = p_end_wk or res["종강일_실제요일"]
        res["회차"] = p_cnt

    # 2. 개강일 및 추천대상 추출
    m_open_line = re.search(
        r"개강일\s*([^\n\r]*?)(?:\s+추천\s*대상\s*([^\n\r]*))?(?=\n|\r|$)",
        ov_text,
    )
    if m_open_line:
        raw_open_token = m_open_line.group(1).strip()
        if m_open_line.group(2) is not None:
            res["추천대상"] = m_open_line.group(2).strip()

        open_ymd = convert_md_to_ymd(raw_open_token, target_ym=target_ym)
        open_wk, is_broken = extract_printed_weekday(raw_open_token)

        # OCR/폰트 깨짐이 있거나 날짜 추출 실패 시 수강 기간의 시작일로 자동 보정
        if (not open_ymd or is_broken) and res["수강시작일"]:
            res["개강일"] = res["수강시작일"]
            res["개강일_실제요일"] = get_korean_weekday(res["수강시작일"])
            res["개강일_표기요일"] = open_wk or res["수강시작일_표기요일"] or res["개강일_실제요일"]
            res["개강일_원본"] = format_short_md_with_weekday(res["개강일"])
            res["개강일_깨짐보정여부"] = True
        elif open_ymd:
            res["개강일"] = open_ymd
            res["개강일_실제요일"] = get_korean_weekday(open_ymd)
            res["개강일_표기요일"] = open_wk or res["개강일_실제요일"]
            res["개강일_원본"] = raw_open_token
            res["개강일_깨짐보정여부"] = is_broken

    if not res["개강일"] and res["수강시작일"]:
        res["개강일"] = res["수강시작일"]
        res["개강일_실제요일"] = get_korean_weekday(res["수강시작일"])
        res["개강일_표기요일"] = res["수강시작일_표기요일"] or res["개강일_실제요일"]
        res["개강일_원본"] = format_short_md_with_weekday(res["수강시작일"])

    # 3. 수업 시간
    m_sched = re.search(r"수업\s*시간\s*([^\n\r]+)", ov_text)
    if m_sched:
        res["수업시간_원본"] = m_sched.group(1).strip()

    # 4. 휴강일 및 보강 방법
    m_hol = re.search(
        r"휴강일\s*(.*?)\s*보강\s*방법\s*(.*?)(?=\n\s*수강료|\r|\Z)",
        ov_text,
        flags=re.DOTALL,
    )
    if m_hol:
        res["휴강일"] = re.sub(r"\s+", " ", m_hol.group(1)).strip()
        res["보강방법"] = re.sub(r"\s+", " ", m_hol.group(2)).strip()
    else:
        m_hol_only = re.search(r"휴강일\s*([^\n\r]+)", ov_text)
        if m_hol_only:
            res["휴강일"] = m_hol_only.group(1).strip()

    # 5. 수강료 & 회차당 수강료
    m_fee = re.search(
        r"수강료\s*([\d,]+)\s*원\s*(?:\(\s*회차당\s*([\d,]+)\s*원\s*\))?",
        ov_text,
    )
    if m_fee:
        fee_num = int(re.sub(r"[^\d]", "", m_fee.group(1)) or "0")
        unit_num = int(re.sub(r"[^\d]", "", m_fee.group(2) or "") or "0")
        res["수강료"] = fee_num
        res["회차당수강료"] = unit_num
        if unit_num > 0:
            res["수강료_표시"] = f"{fee_num:,}원 (회차당 {unit_num:,}원)"
        else:
            res["수강료_표시"] = f"{fee_num:,}원"

    # 6. 교재 및 교재비
    m_book = re.search(r"수강료[^\n\r]*?교재\s+(.*)", ov_text, flags=re.DOTALL)
    if not m_book:
        m_book = re.search(r"^\s*교재\s+(.*)", ov_text, flags=re.MULTILINE | re.DOTALL)
    if m_book:
        book_raw = m_book.group(1).strip()
        book_lines = [ln.strip() for ln in book_raw.splitlines() if ln.strip()]
        cleaned_book_text = re.sub(r"\s+", " ", " ".join(book_lines)).strip()
        res["교재내역"] = cleaned_book_text

        price_matches = re.findall(
            r"(?:▶|/)\s*([0-9,]{4,})\s*(?:원)?|([0-9,]{4,})\s*원",
            cleaned_book_text,
        )
        book_prices: List[int] = []
        for pm1, pm2 in price_matches:
            val_str = pm1 or pm2
            val_int = int(re.sub(r"[^\d]", "", val_str) or "0")
            if 1000 <= val_int <= 500000:
                book_prices.append(val_int)

        total_book_fee = sum(book_prices)
        res["교재비_합계"] = total_book_fee
        if total_book_fee > 0:
            if len(book_prices) > 1:
                each_str = "+".join(f"{p:,}" for p in book_prices)
                res["교재비_표시"] = f"{total_book_fee:,}원 ({each_str})"
            else:
                res["교재비_표시"] = f"{total_book_fee:,}원"
        else:
            if "무료" in cleaned_book_text and "미정" in cleaned_book_text:
                res["교재비_표시"] = "미정/무료"
            elif "무료" in cleaned_book_text:
                res["교재비_표시"] = "무료"
            elif "미정" in cleaned_book_text:
                res["교재비_표시"] = "미정"
            elif cleaned_book_text:
                res["교재비_표시"] = "별도 표기 없음"
            else:
                res["교재비_표시"] = "-"

    return res


def _extract_schedule_and_verify_january(
    page_texts: List[str],
    course_days_list: List[str],
    target_ym: str = "2027-01",
    plumber_pages: Optional[Any] = None,
) -> Dict[str, Any]:
    """
    강좌 블록 내 '주차별 진도계획'('1월 과정' 및 '2월 과정')을
    초고속 정규표현식 엔진으로 0.001초 만에 파싱합니다. (pdfplumber 테이블 연산 완전 제거)
    - '1월 과정' 아래부터 '2월 과정'(또는 '연간 커리큘럼') 직전까지 등장하는 모든 날짜를 추출합니다.
    - 폰트 인코딩 깨짐이 있더라도 (M/D) 숫자를 기준으로 실제 달력 요일을 datetime으로 계산합니다.
    - 진도표_1월_실제회차, 진도표_첫수업일, 진도표_마지막수업일, 진도표_요일오류여부, 2월 표로 넘어간 1월 날짜 여부를 산출합니다.
    """
    items: List[Dict[str, Any]] = []
    seen_keys = set()

    # 1) 텍스트 기반 '1월 과정' vs '2월 과정' 구간 내 날짜 소속 판별 맵 생성
    combined_text = "\n".join(page_texts)
    m_plan_start = re.search(r"주차별\s*진도\s*계획(.*)", combined_text, flags=re.DOTALL)
    plan_text = m_plan_start.group(1) if m_plan_start else combined_text
    # 연간 커리큘럼 및 하단 학원 정보 이전까지만 슬라이싱
    plan_text = re.split(r"연간\s*커리큘럼|○\s*러셀\s*대치", plan_text, maxsplit=1)[0]

    # '2월 과정' 키워드를 기준으로 1월 영역 텍스트와 2월 영역 텍스트 분리
    feb_split = re.split(r"2\s*월\s*과정", plan_text, maxsplit=1)
    jan_text_block = feb_split[0]
    feb_text_block = feb_split[1] if len(feb_split) > 1 else ""

    if "1월 과정" in jan_text_block or "1월과정" in jan_text_block:
        jan_text_block = re.split(r"1\s*월\s*과정", jan_text_block, maxsplit=1)[-1]

    # 정규표현식으로 jan_text_block과 feb_text_block에서 날짜(M/D) 집합 추출
    date_token_regex = re.compile(
        r"(?<!\d)(\d{1,2}\s*/\s*\d{1,2})\s*(?:[월화수목금토일]?\s*\([^)\n\r]{0,15}\)|\([월화수목금토일]\))"
    )
    jan_block_ymds = {
        convert_md_to_ymd(m.group(0), target_ym=target_ym)
        for m in date_token_regex.finditer(jan_text_block)
    } - {""}
    feb_block_ymds = {
        convert_md_to_ymd(m.group(0), target_ym=target_ym)
        for m in date_token_regex.finditer(feb_text_block)
    } - {""}

    # 2) 텍스트 정규표현식 기반 초고속 파싱
    for sec_label, sec_txt in [("1월 과정", jan_text_block), ("2월 과정", feb_text_block)]:
        if not sec_txt.strip():
            continue
        for m_line in re.finditer(
            r"(?<!\d)(\d{1,2}\s*/\s*\d{1,2}\s*(?:[월화수목금토일]?\s*\([^)\n\r]{0,15}\))?)\s+([^\n\r]+)",
            sec_txt,
        ):
            d_raw = m_line.group(1).strip()
            c_raw = m_line.group(2).strip()
            std_d = convert_md_to_ymd(d_raw, target_ym=target_ym)
            if not std_d:
                continue
            key = f"{std_d}|{c_raw}"
            if key in seen_keys:
                continue
            seen_keys.add(key)

            is_hol = ("휴강" in c_raw) and (
                len(c_raw) <= 15 or "설 연휴" in c_raw or "설연휴" in c_raw
            )
            printed_wk, wk_broken = extract_printed_weekday(d_raw)
            actual_wk = get_korean_weekday(std_d)
            effective_printed_wk = printed_wk if printed_wk else actual_wk
            calendar_wk_mismatch = bool(printed_wk and actual_wk and printed_wk != actual_wk)
            class_day_mismatch = bool(
                not is_hol
                and course_days_list
                and actual_wk
                and (actual_wk not in course_days_list)
            )

            items.append(
                {
                    "과정": sec_label,
                    "날짜_원본": d_raw,
                    "날짜_정제표시": format_short_md_with_weekday(std_d),
                    "날짜": std_d,
                    "표기요일": effective_printed_wk,
                    "실제요일": actual_wk,
                    "요일깨짐여부": wk_broken,
                    "달력요일불일치": calendar_wk_mismatch,
                    "수업요일불일치": class_day_mismatch,
                    "요일오류여부": calendar_wk_mismatch or class_day_mismatch,
                    "단원_및_수업내용": c_raw,
                    "휴강여부": is_hol,
                }
            )

    # 각 과정 내에서 날짜순 정렬
    jan_items = sorted(
        [it for it in items if it["과정"] == "1월 과정"],
        key=lambda x: x["날짜"],
    )
    feb_items = sorted(
        [it for it in items if it["과정"] != "1월 과정"],
        key=lambda x: x["날짜"],
    )
    all_items_sorted = jan_items + feb_items

    # 1월 과정 유효 수업일(휴강 제외)
    jan_active = [it for it in jan_items if not it["휴강여부"]]
    jan_holidays = [it for it in jan_items if it["휴강여부"]]

    # 2월 과정 표에 잘못 들어가 있는 1월(target_ym) 날짜 탐지 (예: 박성민T 연합반 1/30(토), 한대부고반 1/31(일))
    target_ym_prefix = str(target_ym).strip()
    spilled_jan_in_feb = [
        it
        for it in feb_items
        if not it["휴강여부"] and str(it["날짜"]).startswith(target_ym_prefix)
    ]

    # 요일 오류 상세 내역 수집
    weekday_error_details: List[str] = []
    for it in jan_active:
        if it["달력요일불일치"]:
            weekday_error_details.append(
                f"{it['날짜_원본']} → 실제 달력은 '{it['실제요일']}요일'임 (요일 오타)"
            )
        elif it["수업요일불일치"]:
            days_str = "·".join(course_days_list)
            weekday_error_details.append(
                f"{it['날짜_정제표시']} → 수업요일({days_str})이 아닌 '{it['실제요일']}요일' 날짜가 진도표에 기재됨"
            )

    jan_first_ymd = jan_active[0]["날짜"] if jan_active else ""
    jan_last_ymd = jan_active[-1]["날짜"] if jan_active else ""
    jan_dates_short_list = [it["날짜_정제표시"] for it in jan_active]
    spilled_dates_short_list = [it["날짜_정제표시"] for it in spilled_jan_in_feb]

    return {
        "all_items": all_items_sorted,
        "jan_items": jan_items,
        "feb_items": feb_items,
        "진도표_1월_실제회차": len(jan_active),
        "진도표_1월_휴강횟수": len(jan_holidays),
        "진도표_첫수업일": jan_first_ymd,
        "진도표_첫수업일_표시": format_short_md_with_weekday(jan_first_ymd) if jan_first_ymd else "-",
        "진도표_마지막수업일": jan_last_ymd,
        "진도표_마지막수업일_표시": format_short_md_with_weekday(jan_last_ymd) if jan_last_ymd else "-",
        "진도표_1월_날짜리스트": jan_dates_short_list,
        "진도표_1월_날짜문자열": ", ".join(jan_dates_short_list) if jan_dates_short_list else "-",
        "진도표_2월표_1월날짜리스트": spilled_dates_short_list,
        "진도표_2월표_1월날짜문자열": ", ".join(spilled_dates_short_list) if spilled_dates_short_list else "",
        "진도표_요일오류여부": bool(weekday_error_details),
        "진도표_요일오류상세": " / ".join(weekday_error_details),
    }


def _read_single_pdf_bytes(pdf_file: Union[str, Path, bytes, io.BytesIO, Any]) -> Tuple[bytes, str]:
    """단일 PDF 입력에서 (raw_bytes, filename)을 추출합니다."""
    fname = getattr(pdf_file, "name", "")
    if isinstance(pdf_file, (str, Path)):
        p = Path(pdf_file)
        return p.read_bytes(), p.name
    elif isinstance(pdf_file, bytes):
        return pdf_file, fname or "uploaded.pdf"
    elif hasattr(pdf_file, "getvalue"):
        return pdf_file.getvalue(), fname or "uploaded.pdf"
    elif hasattr(pdf_file, "read"):
        pdf_file.seek(0)
        return pdf_file.read(), fname or "uploaded.pdf"
    else:
        raise ValueError("지원하지 않는 PDF 파일 입력 형식입니다.")


def parse_syllabus_pdf(
    pdf_source: Union[str, Path, bytes, io.BytesIO, List[Any], Any],
    target_ym: str = "2027-01",
) -> pd.DataFrame:
    """
    단일 통합 PDF 또는 다중 개별 PDF 파일 리스트를 파싱하여
    강좌별 요약, 강의개요, 1월 진도표 역산 검증 데이터가 담긴 DataFrame을 반환합니다.
    """
    if isinstance(pdf_source, (list, tuple)):
        file_list = list(pdf_source)
    else:
        file_list = [pdf_source]

    records: List[Dict[str, Any]] = []
    global_course_idx = 0

    for f_item in file_list:
        if f_item is None:
            continue
        raw_bytes, pdf_filename = _read_single_pdf_bytes(f_item)

        # pypdf.PdfReader 기반 초고속 텍스트 추출 (40페이지 0.2초대 추출)
        fast_texts = _extract_page_texts_fast(raw_bytes)

        courses_raw_groups: List[Dict[str, Any]] = []
        for idx, txt in enumerate(fast_texts):
            p_num = idx + 1
            # '강의개요'가 포함된 페이지를 강좌 블록 시작점으로 식별
            is_course_start = ("강의개요" in txt or "강의 개요" in txt) and (
                "개강일" in txt or "수강 기간" in txt or "수강기간" in txt or "수강료" in txt
            )
            if is_course_start:
                courses_raw_groups.append(
                    {
                        "start_page": p_num,
                        "end_page": p_num,
                        "page_indices": [idx],
                        "texts": [txt],
                    }
                )
            else:
                if courses_raw_groups:
                    courses_raw_groups[-1]["end_page"] = p_num
                    courses_raw_groups[-1]["page_indices"].append(idx)
                    courses_raw_groups[-1]["texts"].append(txt)

        for grp in courses_raw_groups:
            global_course_idx += 1
            first_text = grp["texts"][0]
            all_texts = grp["texts"]
            full_text = "\n\n".join(
                f"[Page {grp['start_page'] + i}]\n{t}" for i, t in enumerate(all_texts)
            )

            hdr = _parse_header_lines(first_text)
            ov = _parse_overview_section(first_text, target_ym=target_ym)
            parsed_sched = parse_schedule_text(ov["수업시간_원본"])

            sched_ver = _extract_schedule_and_verify_january(
                page_texts=all_texts,
                course_days_list=parsed_sched["days_list"],
                target_ym=target_ym,
            )

            pure_title, bracket_tags = strip_course_brackets(hdr["강좌명"])
            sched_items = sched_ver["all_items"]
            active_sched = [it for it in sched_items if not it["휴강여부"]]

            sched_summary_parts = [
                f"{it['날짜_정제표시']}: {it['단원_및_수업내용']}" for it in sched_ver["jan_items"][:8]
            ]
            if len(sched_ver["jan_items"]) > 8:
                sched_summary_parts.append(f"... 외 {len(sched_ver['jan_items']) - 8}회")

            page_range_str = (
                f"p.{grp['start_page']}~{grp['end_page']}"
                if grp["start_page"] != grp["end_page"]
                else f"p.{grp['start_page']}"
            )

            # PDF 개요 개강~종강(회차) 및 1월 진도표(첫날~끝날/회차) 표준 표시 문자열
            ov_open_short = format_short_md_with_weekday(ov["개강일"]) if ov["개강일"] else (ov["개강일_원본"] or "-")
            ov_close_short = format_short_md_with_weekday(ov["종강일"]) if ov["종강일"] else (ov["종강일_원본"] or "-")
            pdf_overview_period_disp = f"{ov_open_short}~{ov_close_short} ({ov['회차']}회)"

            jan_cnt = sched_ver["진도표_1월_실제회차"]
            jan_first_disp = sched_ver["진도표_첫수업일_표시"]
            jan_last_disp = sched_ver["진도표_마지막수업일_표시"]
            if jan_cnt > 0:
                pdf_jan_plan_disp = f"{jan_first_disp}~{jan_last_disp} ({jan_cnt}회)"
                if sched_ver["진도표_2월표_1월날짜리스트"]:
                    spilled_str = ",".join(sched_ver["진도표_2월표_1월날짜리스트"])
                    pdf_jan_plan_disp += f" [※2월표에 {spilled_str} 포함]"
            else:
                pdf_jan_plan_disp = "진도표 미추출 (0회)"

            records.append(
                {
                    "PDF_순번": global_course_idx,
                    "PDF_파일명": pdf_filename,
                    "페이지_범위": page_range_str,
                    "시작페이지": grp["start_page"],
                    "종료페이지": grp["end_page"],
                    "과목": hdr["과목"],
                    "강사명": hdr["강사명"],
                    "강사명_원본": hdr["강사명_원본"],
                    "대상학년": hdr["대상학년"],
                    "과정구분": hdr["과정구분"],
                    "바자관시수": hdr["바자관시수"],
                    "슬로건": hdr["슬로건"],
                    "강좌명": hdr["강좌명"],
                    "순수강좌명": pure_title,
                    "말머리태그": bracket_tags,
                    "개강일": ov["개강일"],
                    "개강일_원본": ov["개강일_원본"],
                    "개강일_표기요일": ov["개강일_표기요일"],
                    "개강일_실제요일": ov["개강일_실제요일"],
                    "개강일_깨짐보정여부": ov["개강일_깨짐보정여부"],
                    "수강시작일": ov["수강시작일"],
                    "종강일": ov["종강일"],
                    "종강일_원본": ov["종강일_원본"],
                    "종강일_표기요일": ov["종강일_표기요일"],
                    "종강일_실제요일": ov["종강일_실제요일"],
                    "수강기간_원본": ov["수강기간_원본"],
                    "회차": ov["회차"],
                    "PDF_개요_기간회차_표시": pdf_overview_period_disp,
                    "수업시간_원본": ov["수업시간_원본"],
                    "수업요일": parsed_sched["days_str"],
                    "수업요일_점표시": parsed_sched["days_dot_str"],
                    "요일_리스트": parsed_sched["days_list"],
                    "시작시간": parsed_sched["primary_start"],
                    "전체시작시간": parsed_sched["start_times_str"],
                    "요일시간키_리스트": parsed_sched["day_start_keys"],
                    "요일전체시간키_리스트": parsed_sched["day_full_keys"],
                    "표준수업시간": parsed_sched["formatted_schedule"],
                    "간결수업시간": parsed_sched["compact_schedule"],
                    "수강료": ov["수강료"],
                    "회차당수강료": ov["회차당수강료"],
                    "수강료_표시": ov["수강료_표시"],
                    "교재내역": ov["교재내역"],
                    "교재비_합계": ov["교재비_합계"],
                    "교재비_표시": ov["교재비_표시"],
                    "추천대상": ov["추천대상"],
                    "휴강일": ov["휴강일"],
                    "보강방법": ov["보강방법"],
                    "진도표_1월_실제회차": jan_cnt,
                    "진도표_1월_휴강횟수": sched_ver["진도표_1월_휴강횟수"],
                    "진도표_첫수업일": sched_ver["진도표_첫수업일"],
                    "진도표_첫수업일_표시": jan_first_disp,
                    "진도표_마지막수업일": sched_ver["진도표_마지막수업일"],
                    "진도표_마지막수업일_표시": jan_last_disp,
                    "진도표_1월_기간회차_표시": pdf_jan_plan_disp,
                    "진도표_1월_날짜리스트": sched_ver["진도표_1월_날짜리스트"],
                    "진도표_1월_날짜문자열": sched_ver["진도표_1월_날짜문자열"],
                    "진도표_2월표_1월날짜리스트": sched_ver["진도표_2월표_1월날짜리스트"],
                    "진도표_2월표_1월날짜문자열": sched_ver["진도표_2월표_1월날짜문자열"],
                    "진도표_요일오류여부": sched_ver["진도표_요일오류여부"],
                    "진도표_요일오류상세": sched_ver["진도표_요일오류상세"],
                    "진도계획_당월회차": jan_cnt,
                    "진도계획_전체회차": len(active_sched),
                    "주차별진도계획_리스트": sched_items,
                    "주차별진도계획_1월리스트": sched_ver["jan_items"],
                    "주차별진도계획_2월리스트": sched_ver["feb_items"],
                    "주차별진도계획_요약": " | ".join(sched_summary_parts),
                    "원문전체텍스트": full_text,
                    "강좌_텍스트리스트": all_texts,
                }
            )

    return pd.DataFrame(records)
