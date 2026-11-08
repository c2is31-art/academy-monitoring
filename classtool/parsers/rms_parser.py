"""
[소스 1] RMS 수업목록 파일 (.xls / .xlsx) 파서 모듈
- 확장자가 .xls인 EUC-KR(CP949) HTML 테이블 문서(문서 내 세 번째 본 데이터 테이블 자동 추출)와
  일반 Excel(.xlsx) 파일을 모두 지원합니다.
"""

from __future__ import annotations

import html
import io
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd
from bs4 import BeautifulSoup

REQUIRED_RMS_COLUMNS = [
    "강좌코드",
    "영역",
    "학년",
    "과목",
    "강좌명",
    "강사",
    "회차",
    "수업시간",
    "개강일",
    "종강일",
]

WEEKDAY_ORDER = {"월": 0, "화": 1, "수": 2, "목": 3, "금": 4, "토": 5, "일": 6}
WEEKDAY_NAMES = ["월", "화", "수", "목", "금", "토", "일"]


def get_korean_weekday(ymd_str: Any) -> str:
    """'YYYY-MM-DD' 형식의 날짜 문자열에서 실제 달력 요일('월'~'일')을 반환합니다."""
    if not ymd_str or pd.isna(ymd_str):
        return ""
    s = str(ymd_str).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if not m:
        return ""
    try:
        dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return WEEKDAY_NAMES[dt.weekday()]
    except ValueError:
        return ""


def format_short_md_with_weekday(ymd_str: Any) -> str:
    """'2027-01-04' -> '1/4(월)' 형태로 변환합니다."""
    if not ymd_str or pd.isna(ymd_str):
        return "-"
    s = str(ymd_str).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if not m:
        return s
    try:
        yr, mo, dy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        dt = datetime(yr, mo, dy)
        wk = WEEKDAY_NAMES[dt.weekday()]
        return f"{mo}/{dy}({wk})"
    except ValueError:
        return s


def standardize_date(date_val: Any) -> str:
    """
    YY-MM-DD, YYYY.MM.DD, YYYY/MM/DD 등의 날짜 문자열을 YYYY-MM-DD 포맷으로 표준화합니다.
    예: '27-01-04' -> '2027-01-04', '26-12-23' -> '2026-12-23'
    """
    if pd.isna(date_val) or date_val is None:
        return ""
    if isinstance(date_val, datetime):
        return date_val.strftime("%Y-%m-%d")
    s = str(date_val).strip()
    if not s or s == "-":
        return ""

    # 시간 정보가 뒤에 붙어있는 경우 앞의 날짜만 추출
    s_date = s.split()[0].strip()
    # 구분자 통일 (., / -> -)
    s_date = re.sub(r"[./]", "-", s_date)

    # YY-MM-DD 패턴 (예: 27-01-04)
    m_short = re.match(r"^(\d{2})-(\d{1,2})-(\d{1,2})$", s_date)
    if m_short:
        yy, mm, dd = int(m_short.group(1)), int(m_short.group(2)), int(m_short.group(3))
        yyyy = 2000 + yy if yy < 70 else 1900 + yy
        return f"{yyyy:04d}-{mm:02d}-{dd:02d}"

    # YYYY-MM-DD 패턴 (예: 2027-1-4)
    m_full = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s_date)
    if m_full:
        yyyy, mm, dd = int(m_full.group(1)), int(m_full.group(2)), int(m_full.group(3))
        return f"{yyyy:04d}-{mm:02d}-{dd:02d}"

    # YYYYMMDD 8자리 숫자
    m_digits8 = re.match(r"^(\d{4})(\d{2})(\d{2})$", s_date)
    if m_digits8:
        return f"{m_digits8.group(1)}-{m_digits8.group(2)}-{m_digits8.group(3)}"

    # YYMMDD 6자리 숫자
    m_digits6 = re.match(r"^(\d{2})(\d{2})(\d{2})$", s_date)
    if m_digits6:
        yy = int(m_digits6.group(1))
        yyyy = 2000 + yy if yy < 70 else 1900 + yy
        return f"{yyyy:04d}-{m_digits6.group(2)}-{m_digits6.group(3)}"

    return s_date


def strip_course_brackets(course_name: Any) -> Tuple[str, str]:
    """
    강좌명에서 대괄호 태그(예: [고2 특강][실력반]) 또는 꺾쇠 태그(<윈터>, <Big Bang>)를 분리하여
    (순수강좌명, 추출된태그문자열) 튜플을 반환합니다.
    """
    if pd.isna(course_name) or course_name is None:
        return "", ""
    raw = html.unescape(str(course_name)).strip()
    tags = re.findall(r"\[.*?\]", raw)
    pure = re.sub(r"\[.*?\]", " ", raw)
    pure = re.sub(r"\s+", " ", pure).strip()
    return pure, "".join(tags)


def normalize_time_token(t_str: str) -> str:
    """'9:00' -> '09:00' 형태로 시간 포맷을 HH:MM 5자리로 정규화합니다."""
    t_str = t_str.strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", t_str)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return t_str


def normalize_baza_hours(val: Any) -> str:
    """바자관 시수/인정기준 문자열(예: '6T', '3T', '-')을 표준화합니다."""
    if pd.isna(val) or val is None:
        return "-"
    s = str(val).strip()
    if not s or s.lower() == "nan" or s == "-":
        return "-"
    m = re.search(r"(\d+)\s*T", s, flags=re.IGNORECASE)
    if m:
        return f"{int(m.group(1))}T"
    return s


def parse_schedule_text(schedule_val: Any, day_val: Any = None) -> Dict[str, Any]:
    """
    수업시간 및 요일 문자열을 파싱하여 구조화된 요일/시간 정보를 반환합니다.
    입력 예시:
    - RMS: '월 09:00~12:00 수 09:00~12:00' 또는 '화 18:00~21:00 토 13:00~16:00'
    - PDF: '월·수 09:00~12:00' 또는 '화 18:00~21:00,토 13:00~16:00'
    """
    raw_sched = "" if pd.isna(schedule_val) or schedule_val is None else str(schedule_val).strip()
    raw_day = "" if pd.isna(day_val) or day_val is None else str(day_val).strip()

    cleaned_sched = re.sub(r"[\r\n]+", " ", raw_sched)
    cleaned_sched = re.sub(r"\s+", " ", cleaned_sched).strip()

    slot_pattern = re.compile(
        r"([월화수목금토일](?:\s*[,/·ㆍ・]\s*[월화수목금토일])*)\s*(?:요일)?\s*(\d{1,2}:\d{2})\s*[~\-∼～]\s*(\d{1,2}:\d{2})"
    )
    matches = slot_pattern.findall(cleaned_sched)

    slots: List[Tuple[str, str, str]] = []  # (day, start, end)
    if matches:
        for days_part, start_t, end_t in matches:
            st = normalize_time_token(start_t)
            et = normalize_time_token(end_t)
            individual_days = re.findall(r"[월화수목금토일]", days_part)
            for d in individual_days:
                slots.append((d, st, et))
    else:
        combined_day_source = f"{raw_day} {cleaned_sched}"
        days_found = []
        for d in re.findall(r"[월화수목금토일]", combined_day_source):
            if d not in days_found:
                days_found.append(d)

        time_ranges = re.findall(r"(\d{1,2}:\d{2})\s*[~\-∼～]\s*(\d{1,2}:\d{2})", cleaned_sched)
        if time_ranges:
            if len(time_ranges) == len(days_found) and len(days_found) > 1:
                for d, (st, et) in zip(days_found, time_ranges):
                    slots.append((d, normalize_time_token(st), normalize_time_token(et)))
            else:
                st, et = normalize_time_token(time_ranges[0][0]), normalize_time_token(time_ranges[0][1])
                if days_found:
                    for d in days_found:
                        slots.append((d, st, et))
                else:
                    slots.append(("", st, et))
        else:
            single_times = re.findall(r"(\d{1,2}:\d{2})", cleaned_sched)
            st = normalize_time_token(single_times[0]) if single_times else ""
            et = normalize_time_token(single_times[1]) if len(single_times) > 1 else ""
            if days_found:
                for d in days_found:
                    slots.append((d, st, et))
            elif st:
                slots.append(("", st, et))

    slots_sorted = sorted(slots, key=lambda x: (WEEKDAY_ORDER.get(x[0], 99), x[1]))

    unique_days = []
    for d, _, _ in slots_sorted:
        if d and d not in unique_days:
            unique_days.append(d)

    start_times = []
    for _, st, _ in slots_sorted:
        if st and st not in start_times:
            start_times.append(st)

    day_start_pairs = [f"{d} {st}".strip() for d, st, _ in slots_sorted if d or st]
    day_full_pairs = [
        f"{d} {st}~{et}".strip() if (st and et) else f"{d} {st}".strip()
        for d, st, et in slots_sorted
        if d or st
    ]
    formatted_slots = []
    for d, st, et in slots_sorted:
        if d and st and et:
            formatted_slots.append(f"{d} {st}~{et}")
        elif d and st:
            formatted_slots.append(f"{d} {st}")
        elif st and et:
            formatted_slots.append(f"{st}~{et}")
        elif d:
            formatted_slots.append(d)

    # 간결한 수업시간 표시 (모든 요일의 시작~종료시간이 같으면 '월·수 09:00~12:00', 다르면 '화 18:00~21:00 / 토 13:00~16:00')
    unique_ranges = {(st, et) for _, st, et in slots_sorted if st or et}
    if len(unique_ranges) == 1 and unique_days:
        only_st, only_et = next(iter(unique_ranges))
        days_dot = "·".join(unique_days)
        compact_schedule = f"{days_dot} {only_st}~{only_et}" if only_et else f"{days_dot} {only_st}"
    else:
        compact_schedule = " / ".join(formatted_slots) if formatted_slots else cleaned_sched

    return {
        "days_list": unique_days,
        "days_str": ",".join(unique_days),
        "days_dot_str": "·".join(unique_days),
        "primary_start": start_times[0] if start_times else "",
        "start_times_str": ",".join(start_times),
        "day_start_keys": day_start_pairs,
        "day_full_keys": day_full_pairs,
        "formatted_schedule": " / ".join(formatted_slots) if formatted_slots else cleaned_sched,
        "compact_schedule": compact_schedule,
    }


def decode_rms_bytes(raw_bytes: bytes) -> str:
    """
    RMS .xls 파일 바이트를 디코딩하여 HTML 문자열로 반환합니다.
    EUC-KR(CP949) 우선 시도 후 UTF-8 계열로 폴백합니다.
    """
    for enc in ("cp949", "euc-kr", "utf-8-sig", "utf-8"):
        try:
            decoded = raw_bytes.decode(enc)
            if "강좌코드" in decoded or "강좌명" in decoded or "<table" in decoded.lower():
                return decoded
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("cp949", errors="replace")


def _normalize_candidate_df(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    추출된 DataFrame의 컬럼 또는 상단 행에 필수 컬럼이 모두 포함되어 있는지 확인하고,
    포함되어 있으면 컬럼명을 정규화하여 반환합니다.
    """
    if df is None or df.empty:
        return None

    cols_str = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    if all(req in cols_str for req in REQUIRED_RMS_COLUMNS):
        out = df.copy()
        out.columns = cols_str
        return out

    if isinstance(df.columns, pd.MultiIndex):
        last_level = [re.sub(r"\s+", " ", str(c[-1])).strip() for c in df.columns]
        if all(req in last_level for req in REQUIRED_RMS_COLUMNS):
            out = df.copy()
            out.columns = last_level
            return out

    first_row = [re.sub(r"\s+", " ", str(v)).strip() for v in df.iloc[0].values]
    if all(req in first_row for req in REQUIRED_RMS_COLUMNS):
        out = df.iloc[1:].copy().reset_index(drop=True)
        out.columns = first_row
        return out

    max_scan = min(10, len(df))
    for idx in range(1, max_scan):
        row_vals = [re.sub(r"\s+", " ", str(v)).strip() for v in df.iloc[idx].values]
        if all(req in row_vals for req in REQUIRED_RMS_COLUMNS):
            out = df.iloc[idx + 1 :].copy().reset_index(drop=True)
            out.columns = row_vals
            return out

    return None


def _parse_html_tables_Fallback_bs4(html_text: str) -> List[pd.DataFrame]:
    """pd.read_html 실패 시 BeautifulSoup으로 모든 테이블을 직접 DataFrame 리스트로 변환합니다."""
    soup = BeautifulSoup(html_text, "html.parser")
    dfs: List[pd.DataFrame] = []
    for tbl in soup.find_all("table"):
        if tbl.find("table") is not None:
            continue
        rows_data = []
        for tr in tbl.find_all("tr"):
            cells = [
                re.sub(r"\s+", " ", cell.get_text(separator=" ", strip=True)).strip()
                for cell in tr.find_all(["th", "td"])
            ]
            if cells:
                rows_data.append(cells)
        if rows_data:
            max_cols = max(len(r) for r in rows_data)
            padded = [r + [""] * (max_cols - len(r)) for r in rows_data]
            dfs.append(pd.DataFrame(padded))
    return dfs


def parse_rms_xls(file_source: Union[str, Path, bytes, io.BytesIO, Any]) -> pd.DataFrame:
    """
    RMS 수업목록 (.xls / .xlsx) 파일을 파싱하여 표준화된 DataFrame으로 반환합니다.
    - .xls (EUC-KR HTML 테이블 문서: 3번째 본 데이터 테이블 자동 추출)
    - .xlsx (일반 Excel 통합 문서) 모두 지원합니다.
    """
    filename = getattr(file_source, "name", str(file_source) if isinstance(file_source, (str, Path)) else "")
    try:
        if isinstance(file_source, (str, Path)):
            raw_bytes = Path(file_source).read_bytes()
        elif isinstance(file_source, bytes):
            raw_bytes = file_source
        elif hasattr(file_source, "getvalue"):
            raw_bytes = file_source.getvalue()
        elif hasattr(file_source, "read"):
            file_source.seek(0)
            raw_bytes = file_source.read()
        else:
            raise ValueError("지원하지 않는 파일 입력 형식입니다.")
    except PermissionError as e:
        raise RuntimeError(f"RMS 파일을 읽는 중 권한 오류가 발생했습니다. 파일이 다른 프로그램에서 열려있는지 확인해 주세요: {e}") from e
    except Exception as e:
        raise RuntimeError(f"RMS 파일 로드 실패: {e}") from e

    matched_candidates: List[pd.DataFrame] = []

    # 1) .xlsx (ZIP 헤더 PK\x03\x04) 또는 엑셀 바이너리 파일인 경우 우선 시도
    is_zip_xlsx = raw_bytes[:4] == b"PK\x03\x04" or str(filename).lower().endswith((".xlsx", ".xlsm"))
    if is_zip_xlsx:
        try:
            xls_file = pd.ExcelFile(io.BytesIO(raw_bytes))
            for s_name in xls_file.sheet_names:
                raw_sheet_df = pd.read_excel(xls_file, sheet_name=s_name)
                norm_df = _normalize_candidate_df(raw_sheet_df)
                if norm_df is not None:
                    matched_candidates.append(norm_df)
        except Exception:
            pass

    # 2) HTML 테이블 기반 .xls 파싱
    raw_banner_text = ""
    banner_counts = None
    if not matched_candidates:
        html_text = decode_rms_bytes(raw_bytes)
        preprocessed_html = re.sub(r"<br\s*/?>|<br[^>]*>", " ", html_text, flags=re.IGNORECASE)

        # 상단 요약 배너 텍스트 추출 (HTML 태그 제거 후 정규식 매칭)
        clean_text_for_summary = re.sub(r"<[^>]+>", " ", preprocessed_html)
        clean_text_for_summary = re.sub(r"\s+", " ", clean_text_for_summary)
        m_banner = re.search(
            r"총\s*강좌\s*수\s*:\s*(\d+)\s*건\s*,\s*총\s*수강생\s*:\s*(\d+)\s*명\s*\(\s*결제완료\s*:\s*(\d+)\s*명\s*\)",
            clean_text_for_summary,
        )
        if m_banner:
            banner_courses = int(m_banner.group(1))
            banner_total_stu = int(m_banner.group(2))
            banner_paid_stu = int(m_banner.group(3))
            banner_counts = {
                "total_courses": banner_courses,
                "total_students": banner_total_stu,
                "paid_students": banner_paid_stu,
            }
            raw_banner_text = f"총 강좌 수 : {banner_courses} 건 , 총 수강생 : {banner_total_stu} 명 ( 결제완료 : {banner_paid_stu} 명 )"

        tables: List[pd.DataFrame] = []
        for flavor in ("lxml", "bs4", "html5lib"):
            try:
                tables = pd.read_html(io.StringIO(preprocessed_html), flavor=flavor)
                if tables:
                    break
            except Exception:
                continue

        if not tables:
            tables = _parse_html_tables_Fallback_bs4(preprocessed_html)

        for tbl in tables:
            norm_df = _normalize_candidate_df(tbl)
            if norm_df is not None:
                matched_candidates.append(norm_df)

        if not matched_candidates:
            bs_tables = _parse_html_tables_Fallback_bs4(preprocessed_html)
            for tbl in bs_tables:
                norm_df = _normalize_candidate_df(tbl)
                if norm_df is not None:
                    matched_candidates.append(norm_df)

    if not matched_candidates:
        raise ValueError(
            f"업로드된 파일에서 필수 컬럼({', '.join(REQUIRED_RMS_COLUMNS)})이 포함된 RMS 수업목록 테이블을 찾지 못했습니다."
        )

    # 세 번째 테이블(가장 안쪽 본 데이터 테이블) 선택
    target_df = matched_candidates[-1].copy()

    # 헤더 행이 데이터 중간에 중복 삽입된 경우나 빈 행 제거
    target_df = target_df[target_df["강좌코드"].notna()].copy()
    target_df["강좌코드"] = target_df["강좌코드"].astype(str).str.strip()
    target_df = target_df[
        (target_df["강좌코드"] != "")
        & (target_df["강좌코드"] != "강좌코드")
        & (target_df["강좌코드"].str.lower() != "nan")
    ].reset_index(drop=True)

    # 주요 텍스트 컬럼 클리닝
    text_cols = [
        "강좌코드",
        "영역",
        "학년",
        "과목",
        "선택과목",
        "과정",
        "유형",
        "구분",
        "강좌명",
        "강사",
        "수업시간",
        "개강일",
        "종강일",
        "휴강일",
        "바자관 인정기준",
        "교재정보",
    ]
    for col in text_cols:
        if col in target_df.columns:
            target_df[col] = (
                target_df[col]
                .fillna("")
                .astype(str)
                .apply(lambda x: re.sub(r"\s+", " ", html.unescape(x)).strip())
                .replace({"nan": "", "None": ""})
            )

    def _to_int(val: Any) -> int:
        if pd.isna(val):
            return 0
        m = re.search(r"\d+", str(val))
        return int(m.group(0)) if m else 0

    def _to_money_int(val: Any) -> int:
        if pd.isna(val) or val is None:
            return 0
        s = re.sub(r"[^\d]", "", str(val))
        return int(s) if s else 0

    if "No" in target_df.columns:
        target_df["No"] = [
            _to_int(v) if _to_int(v) > 0 else (i + 1)
            for i, v in enumerate(target_df["No"])
        ]
    else:
        target_df["No"] = list(range(1, len(target_df) + 1))

    target_df["회차"] = target_df["회차"].apply(_to_int)
    if "수강료" in target_df.columns:
        target_df["수강료_원본"] = target_df["수강료"].astype(str)
        target_df["수강료"] = target_df["수강료"].apply(_to_money_int)
    else:
        target_df["수강료_원본"] = ""
        target_df["수강료"] = 0

    if "1회차 기준 수강료" in target_df.columns:
        target_df["1회차 기준 수강료"] = target_df["1회차 기준 수강료"].apply(_to_money_int)
    else:
        target_df["1회차 기준 수강료"] = [
            int(round(f / s)) if (f > 0 and s > 0) else 0
            for f, s in zip(target_df["수강료"], target_df["회차"])
        ]

    if "바자관 인정기준" in target_df.columns:
        target_df["바자관시수"] = target_df["바자관 인정기준"].apply(normalize_baza_hours)
    else:
        target_df["바자관시수"] = "-"

    # 개강일 / 종강일 YYYY-MM-DD 표준화 및 실제 달력 요일 계산
    target_df["개강일_원본"] = target_df["개강일"]
    target_df["종강일_원본"] = target_df["종강일"]
    target_df["개강일"] = target_df["개강일"].apply(standardize_date)
    target_df["종강일"] = target_df["종강일"].apply(standardize_date)
    target_df["개강일_요일"] = target_df["개강일"].apply(get_korean_weekday)
    target_df["종강일_요일"] = target_df["종강일"].apply(get_korean_weekday)
    target_df["개강일_짧은표시"] = target_df["개강일"].apply(format_short_md_with_weekday)
    target_df["종강일_짧은표시"] = target_df["종강일"].apply(format_short_md_with_weekday)

    # 강좌명 태그 분리 -> 순수강좌명, 말머리태그
    parsed_titles = target_df["강좌명"].apply(strip_course_brackets)
    target_df["순수강좌명"] = [pt[0] for pt in parsed_titles]
    target_df["말머리태그"] = [pt[1] for pt in parsed_titles]

    # 수업시간 구조화 파싱
    parsed_schedules = target_df["수업시간"].apply(parse_schedule_text)
    target_df["수업요일"] = [ps["days_str"] for ps in parsed_schedules]
    target_df["수업요일_점표시"] = [ps["days_dot_str"] for ps in parsed_schedules]
    target_df["요일_리스트"] = [ps["days_list"] for ps in parsed_schedules]
    target_df["시작시간"] = [ps["primary_start"] for ps in parsed_schedules]
    target_df["전체시작시간"] = [ps["start_times_str"] for ps in parsed_schedules]
    target_df["요일시간키_리스트"] = [ps["day_start_keys"] for ps in parsed_schedules]
    target_df["요일전체시간키_리스트"] = [ps["day_full_keys"] for ps in parsed_schedules]
    target_df["표준수업시간"] = [ps["formatted_schedule"] for ps in parsed_schedules]
    target_df["간결수업시간"] = [ps["compact_schedule"] for ps in parsed_schedules]

    # 수강생 관련 수치 컬럼 정수화 및 파생 컬럼 생성
    for stu_col in ("가등록포함수강생수", "완납&분납수강생수", "결제인원"):
        if stu_col in target_df.columns:
            target_df[stu_col] = (
                pd.to_numeric(target_df[stu_col], errors="coerce")
                .fillna(0)
                .astype(int)
            )
        else:
            target_df[stu_col] = 0

    # 가등록(미결제) 인원 및 결제율 계산
    target_df["가등록(미결제)"] = (
        target_df["가등록포함수강생수"] - target_df["결제인원"]
    ).apply(lambda v: max(0, int(v)))

    target_df["결제율(%)"] = [
        round((paid / total) * 100.0, 1) if total > 0 else 0.0
        for paid, total in zip(target_df["결제인원"], target_df["가등록포함수강생수"])
    ]

    # 전체 요약 통계 딕셔너리 생성
    sum_total_students = int(target_df["가등록포함수강생수"].sum())
    sum_paid_students = int(target_df["결제인원"].sum())
    sum_inst_students = int(target_df["완납&분납수강생수"].sum())
    sum_unpaid_students = max(0, sum_total_students - sum_paid_students)
    total_courses_count = len(target_df)

    # 원본 상단 배너 값이 있었으면 검증 및 기록
    if banner_counts:
        # 배너 상의 수치가 테이블 합계와 다를 경우 배너 값 보존
        disp_courses = banner_counts["total_courses"] or total_courses_count
        disp_students = banner_counts["total_students"] if banner_counts["total_students"] > 0 else sum_total_students
        disp_paid = banner_counts["paid_students"] if banner_counts["paid_students"] > 0 else sum_paid_students
    else:
        disp_courses = total_courses_count
        disp_students = sum_total_students
        disp_paid = sum_paid_students

    disp_unpaid = max(0, disp_students - disp_paid)
    calc_payment_rate = round((disp_paid / disp_students * 100.0), 1) if disp_students > 0 else 0.0

    summary_info = {
        "total_courses": disp_courses,
        "total_students": disp_students,
        "paid_students": disp_paid,
        "unpaid_students": disp_unpaid,
        "installment_students": sum_inst_students,
        "payment_rate": calc_payment_rate,
        "raw_summary_text": (
            raw_banner_text
            or f"총 강좌 수 : {disp_courses} 건 , 총 수강생 : {disp_students} 명 ( 결제완료 : {disp_paid} 명 )"
        ),
        "headline_banner": (
            f"📢 RMS 강의정보 요약 — 총 강좌수 : {disp_courses}건 | "
            f"총 수강생 : {disp_students:,}명 (결제완료 : {disp_paid:,}명 / 미결제 가등록 : {disp_unpaid:,}명)"
        ),
    }

    target_df.attrs["summary"] = summary_info

    return target_df


def detect_dominant_year_month(df: Optional[pd.DataFrame], default_ym: str = "2027-01") -> str:
    """
    RMS DataFrame의 '개강일'(또는 표준화된 날짜 컬럼)에서 YYYY-MM 형식을 추출한 뒤,
    가장 빈도가 높은(최빈값) 연월 문자열(예: "2027-01")을 반환합니다.
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return default_ym

    date_col = None
    for candidate_col in ("개강일", "개강일자", "첫수업일", "개강일_원본"):
        if candidate_col in df.columns:
            date_col = candidate_col
            break

    if date_col is None:
        return default_ym

    try:
        ym_list: List[str] = []
        for val in df[date_col].dropna():
            std_d = standardize_date(val)
            m = re.match(r"^(\d{4}-\d{2})", std_d)
            if m:
                ym_list.append(m.group(1))

        if not ym_list:
            return default_ym

        modes = pd.Series(ym_list).mode()
        return str(modes.iloc[0]) if not modes.empty else default_ym
    except Exception:
        return default_ym


def extract_rms_summary(df: Optional[pd.DataFrame]) -> Dict[str, Any]:
    """
    RMS DataFrame에서 요약 통계 딕셔너리를 추출합니다.
    df.attrs['summary']가 없으면 본 데이터 합계로 자동 계산합니다.
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return {
            "total_courses": 0,
            "total_students": 0,
            "paid_students": 0,
            "unpaid_students": 0,
            "installment_students": 0,
            "payment_rate": 0.0,
            "raw_summary_text": "데이터 없음",
            "headline_banner": "📢 RMS 강의정보 — 등록된 강좌 데이터가 없습니다.",
        }

    if "summary" in df.attrs and isinstance(df.attrs["summary"], dict):
        return df.attrs["summary"]

    total_courses = len(df)
    tot_stu = int(df["가등록포함수강생수"].sum()) if "가등록포함수강생수" in df.columns else 0
    paid_stu = int(df["결제인원"].sum()) if "결제인원" in df.columns else 0
    inst_stu = int(df["완납&분납수강생수"].sum()) if "완납&분납수강생수" in df.columns else 0
    unpaid_stu = max(0, tot_stu - paid_stu)
    rate = round((paid_stu / tot_stu * 100.0), 1) if tot_stu > 0 else 0.0

    return {
        "total_courses": total_courses,
        "total_students": tot_stu,
        "paid_students": paid_stu,
        "unpaid_students": unpaid_stu,
        "installment_students": inst_stu,
        "payment_rate": rate,
        "raw_summary_text": f"총 강좌 수 : {total_courses} 건 , 총 수강생 : {tot_stu} 명 ( 결제완료 : {paid_stu} 명 )",
        "headline_banner": (
            f"📢 RMS 강의정보 요약 — 총 강좌수 : {total_courses}건 | "
            f"총 수강생 : {tot_stu:,}명 (결제완료 : {paid_stu:,}명 / 미결제 가등록 : {unpaid_stu:,}명)"
        ),
    }


def generate_instructor_student_summary(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    강사별 수강생 현황 집계표를 생성합니다:
    - 컬럼: 순위, 과목(영역), 강사명, 개설 강좌수, 총 수강생수(가등록포함), 결제완료 인원,
            완납&분납 인원, 가등록(미결제) 인원, 결제율(%), 강좌당 평균 수강생
    - 정렬: 총 수강생수 내림차순 (동률 시 결제완료 인원 -> 개설 강좌수 내림차순)
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()

    inst_rows: List[Dict[str, Any]] = []
    # 강사 및 과목 기준으로 GroupBy
    group_cols = ["강사"]
    if "과목" in df.columns:
        group_cols.append("과목")

    for keys, grp in df.groupby(group_cols, sort=False):
        if isinstance(keys, tuple):
            inst_name = str(keys[0]).strip()
            subj_name = str(keys[1]).strip()
        else:
            inst_name = str(keys).strip()
            subj_name = str(grp["과목"].iloc[0]).strip() if "과목" in grp.columns else "-"

        area_name = str(grp["영역"].iloc[0]).strip() if "영역" in grp.columns else ""
        disp_subj = f"{subj_name} ({area_name})" if area_name and area_name != subj_name else subj_name

        c_count = len(grp)
        tot_stu = int(grp["가등록포함수강생수"].sum()) if "가등록포함수강생수" in grp.columns else 0
        paid_stu = int(grp["결제인원"].sum()) if "결제인원" in grp.columns else 0
        inst_stu = int(grp["완납&분납수강생수"].sum()) if "완납&분납수강생수" in grp.columns else 0
        unpaid_stu = max(0, tot_stu - paid_stu)
        pay_rate = round((paid_stu / tot_stu * 100.0), 1) if tot_stu > 0 else 0.0
        avg_stu = round((tot_stu / c_count), 1) if c_count > 0 else 0.0

        inst_rows.append(
            {
                "과목(영역)": disp_subj,
                "강사명": inst_name,
                "개설 강좌수": c_count,
                "총 수강생수(가등록포함)": tot_stu,
                "결제완료 인원": paid_stu,
                "완납&분납 인원": inst_stu,
                "가등록(미결제) 인원": unpaid_stu,
                "결제율(%)": pay_rate,
                "강좌당 평균 수강생": avg_stu,
            }
        )

    summary_df = pd.DataFrame(inst_rows)
    if summary_df.empty:
        return summary_df

    # 정렬: 총 수강생수 내림차순 -> 결제완료 인원 내림차순 -> 개설 강좌수 내림차순
    summary_df = summary_df.sort_values(
        by=["총 수강생수(가등록포함)", "결제완료 인원", "개설 강좌수"],
        ascending=[False, False, False],
    ).reset_index(drop=True)

    # 순위 컬럼 추가 (1위부터)
    summary_df.insert(0, "순위", range(1, len(summary_df) + 1))

    return summary_df
