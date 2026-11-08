"""
[소스 2] 강의계획서 수령 현황 (내 구글 계정 OAuth 직접 연동 & 시트 자동 쓰기 Write-back / CSV 백업) 모듈
- pydata-google-auth를 사용하여 서비스 계정 JSON 키 없이 내 구글 계정(편집자 권한)으로 원클릭 OAuth 로그인 및 로컬 캐시 재사용
- Google Sheets에서 '강좌목록' 및 '제출내역' 탭을 직접 읽어오고, 3단 교차 검수 결과를 시트에 즉시 기록(Write-back)하는 기능 제공
"""

from __future__ import annotations

import calendar
import io
import json
import re
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

import gspread
import pandas as pd
import pydata_google_auth
import requests

from parsers.rms_parser import (
    parse_schedule_text,
    standardize_date,
    strip_course_brackets,
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TEST_INSTRUCTORS = {"홍길동", "아자스", "야호야호", "심동준"}

SHEET_SHARE_GUIDE_MSG = (
    "시트가 '제한됨(비공개)' 상태입니다. "
    "구글 시트 우측 상단 [공유] -> 일반 액세스를 '링크가 있는 모든 사용자(뷰어)'로 변경해 주세요. "
    "(또는 아래에 시트 CSV 파일을 직접 올려주세요)"
)


class SheetAccessDeniedError(PermissionError):
    """구글 시트 접근 권한이 없거나 제한됨(비공개) 상태일 때 발생하는 예외"""


def get_user_gspread_client() -> gspread.Client:
    """
    내 구글 계정(OAuth)으로 gspread 클라이언트를 반환합니다.
    - 최초 1회만 브라우저 로그인 창이 뜨고, 이후에는 로컬 캐시(credentials)를 통해 자동 로그인됩니다.
    """
    credentials = pydata_google_auth.get_user_credentials(
        SCOPES,
        auth_local_webserver=True,
    )
    return gspread.authorize(credentials)


def fetch_google_sheet_tab(sheet_url: str, tab_name: str) -> pd.DataFrame:
    """
    requests 기반 다이렉트 로더 (SSL/인코딩/권한 체크 포함):
    - Google Visualization API CSV 엔드포인트를 통해 특정 시트 탭(예: '강좌목록', '제출내역')을 직접 로드합니다.
    - 시트가 '제한됨(비공개)' 상태여서 로그인 HTML이 반환되면 PermissionError를 발생시킵니다.
    """
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", str(sheet_url))
    if not match:
        raise ValueError("유효한 구글 시트 URL이 아닙니다. URL에 '/d/시트ID'가 포함되어야 합니다.")
    sheet_id = match.group(1)

    encoded_tab = urllib.parse.quote(tab_name)
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_tab}"

    resp = requests.get(csv_url, timeout=15)
    if resp.status_code in (401, 403):
        raise PermissionError(
            "시트가 '제한됨(비공개)' 상태입니다. "
            "구글 시트 우측 상단 [공유] -> 일반 액세스를 '링크가 있는 모든 사용자(뷰어)'로 변경해 주세요."
        )
    resp.raise_for_status()
    resp.encoding = "utf-8"

    # 시트가 '제한됨(비공개)' 상태이면 구글 로그인 HTML 페이지가 반환됨
    head_text = resp.text[:500]
    if (
        "<!DOCTYPE html>" in head_text[:300]
        or "<!doctype html>" in head_text[:300].lower()
        or "accounts.google.com" in head_text
    ):
        raise PermissionError(
            "시트가 '제한됨(비공개)' 상태입니다. "
            "구글 시트 우측 상단 [공유] -> 일반 액세스를 '링크가 있는 모든 사용자(뷰어)'로 변경해 주세요."
        )

    df = pd.read_csv(io.StringIO(resp.text))
    # 모두 비어 있는 Unnamed 컬럼 정리
    unnamed_empty = [
        c for c in df.columns if str(c).startswith("Unnamed:") and df[c].isna().all()
    ]
    if unnamed_empty:
        df = df.drop(columns=unnamed_empty)
    return df


def fetch_google_sheet_by_gid(sheet_url: str, gid: Optional[int] = None) -> pd.DataFrame:
    """
    시트 탭 이름이 '강좌목록'/'제출내역'과 다를 경우 URL 내 gid(또는 기본 시트 gid=0)로
    CSV Export 엔드포인트에서 직접 로드하는 2차 다이렉트 안전장치입니다.
    """
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", str(sheet_url))
    if not match:
        raise ValueError("유효한 구글 시트 URL이 아닙니다. URL에 '/d/시트ID'가 포함되어야 합니다.")
    sheet_id = match.group(1)

    if gid is None:
        gid = extract_gid_from_url(sheet_url)
    gid_param = f"&gid={gid}" if gid is not None else ""
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_param}"

    resp = requests.get(export_url, timeout=15)
    if resp.status_code in (401, 403):
        raise PermissionError(
            "시트가 '제한됨(비공개)' 상태입니다. "
            "구글 시트 우측 상단 [공유] -> 일반 액세스를 '링크가 있는 모든 사용자(뷰어)'로 변경해 주세요."
        )
    resp.raise_for_status()
    resp.encoding = "utf-8"

    head_text = resp.text[:500]
    if (
        "<!DOCTYPE html>" in head_text[:300]
        or "<!doctype html>" in head_text[:300].lower()
        or "accounts.google.com" in head_text
    ):
        raise PermissionError(
            "시트가 '제한됨(비공개)' 상태입니다. "
            "구글 시트 우측 상단 [공유] -> 일반 액세스를 '링크가 있는 모든 사용자(뷰어)'로 변경해 주세요."
        )

    if not resp.text.strip():
        return pd.DataFrame()

    df = pd.read_csv(io.StringIO(resp.text))
    unnamed_empty = [
        c for c in df.columns if str(c).startswith("Unnamed:") and df[c].isna().all()
    ]
    if unnamed_empty:
        df = df.drop(columns=unnamed_empty)
    return df


# 표준 컬럼 매핑 사전 (다양한 시트 컬럼명 변형 대응)
COLUMN_SYNONYMS: Dict[str, List[str]] = {
    "강사명": ["강사명", "강사", "선생님", "강사이름", "성명"],
    "과목": ["과목", "과목명", " 세부과목", "세부과목", "영역"],
    "대상학년": ["대상학년", "학년", "수강대상", "대상"],
    "강좌명": ["강좌명", "강의명", "수업명", "강좌제목"],
    "강좌구분": ["강좌구분", "구분", "과정구분", "반구분"],
    "수능유형": ["수능유형", "유형", "수업유형"],
    "수업요일": ["수업요일", "요일", "강의요일"],
    "수업시간": ["수업시간", "시간", "강의시간", "요일및시간"],
    "개강일": ["개강일", "첫수업일", "시작일", "개강일자"],
    "종강일": ["종강일", "마지막수업일", "종료일", "종강일자"],
    "총회차": ["총회차", "회차", "전체회차", "수업회차", "총수업회차"],
    "최근제출일시": ["최근제출일시", "제출일시", "타임스탬프", "Timestamp", "등록일시", "수정일시", "제출시간"],
    "회차별계획(JSON)": ["회차별계획(JSON)", "회차별계획", "주차별계획(JSON)", "주차별계획", "수업계획(JSON)", "plan_json"],
    "저장폴더URL": ["저장폴더URL", "폴더URL", "드라이브폴더URL", "폴더링크"],
    "문서URL": ["문서URL", "파일URL", "계획서URL", "구글독스URL", "문서링크", "링크"],
}


def extract_spreadsheet_id(sheet_url: str) -> str:
    """
    사용자가 입력한 Google Sheets URL에서 정규표현식 r'/d/([a-zA-Z0-9-_]+)'으로 스프레드시트 ID를 추출합니다.
    """
    if not sheet_url:
        return ""
    s = str(sheet_url).strip()
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", s)
    if m:
        return m.group(1)
    if re.match(r"^[a-zA-Z0-9-_]{20,}$", s):
        return s
    return ""


def extract_gid_from_url(sheet_url: str) -> Optional[int]:
    """URL에 포함된 gid 파라미터(#gid=1185470257 또는 ?gid=...)를 정수로 추출합니다."""
    if not sheet_url:
        return None
    m = re.search(r"[?#&]gid=(\d+)", str(sheet_url))
    return int(m.group(1)) if m else None


def extract_google_doc_id(url_or_id: Any) -> str:
    """Google Docs/Drive URL에서 문서 ID(doc_id)를 추출합니다."""
    if pd.isna(url_or_id) or not url_or_id:
        return ""
    s = str(url_or_id).strip()
    m = re.search(r"/d/([a-zA-Z0-9_-]{15,})", s)
    if m:
        return m.group(1)
    m2 = re.search(r"[?&]id=([a-zA-Z0-9_-]{15,})", s)
    if m2:
        return m2.group(1)
    m3 = re.search(r"/folders/([a-zA-Z0-9_-]{15,})", s)
    if m3:
        return m3.group(1)
    if re.match(r"^[a-zA-Z0-9_-]{20,}$", s):
        return s
    return ""


def _worksheet_to_dataframe(ws: gspread.Worksheet) -> pd.DataFrame:
    """
    gspread Worksheet의 모든 값을 읽어 안전하게 DataFrame으로 변환합니다.
    - 빈 헤더나 중복된 헤더 컬럼명이 있더라도 예외 없이 처리하며, 원본 시트 행 번호(_시트행번호)를 보존합니다.
    """
    all_values = ws.get_all_values()
    if not all_values or len(all_values) < 2:
        if all_values and len(all_values) == 1:
            headers = [str(h).strip() or f"비어있는컬럼_{i+1}" for i, h in enumerate(all_values[0])]
            return pd.DataFrame(columns=headers)
        return pd.DataFrame()

    raw_headers = all_values[0]
    clean_headers: List[str] = []
    seen_counts: Dict[str, int] = {}
    for idx, h in enumerate(raw_headers):
        h_str = str(h).strip()
        if not h_str:
            h_str = f"_empty_col_{idx + 1}"
        if h_str in seen_counts:
            seen_counts[h_str] += 1
            h_str = f"{h_str}_{seen_counts[h_str]}"
        else:
            seen_counts[h_str] = 1
        clean_headers.append(h_str)

    data_rows = all_values[1:]
    df = pd.DataFrame(data_rows, columns=clean_headers)
    # Google Sheets 상의 실제 행 번호 (헤더가 1행이므로 데이터는 2행부터 시작)
    df["_시트행번호"] = list(range(2, len(data_rows) + 2))
    df["_원본시트탭명"] = ws.title
    return df


def standardize_sheet_columns(df: pd.DataFrame) -> pd.DataFrame:
    """시트 또는 CSV의 컬럼명을 표준 컬럼명으로 매핑하고 누락된 표준 컬럼을 빈 값으로 생성합니다."""
    if df is None or df.empty:
        return pd.DataFrame(columns=list(COLUMN_SYNONYMS.keys()))

    out = df.copy()
    cleaned_cols = {col: re.sub(r"\s+", "", str(col)).strip() for col in out.columns}

    rename_map: Dict[str, str] = {}
    used_std_cols = set()

    for std_col, synonyms in COLUMN_SYNONYMS.items():
        norm_synonyms = [re.sub(r"\s+", "", syn).lower() for syn in synonyms]
        for orig_col, clean_col in cleaned_cols.items():
            if orig_col in rename_map:
                continue
            if clean_col.lower() in norm_synonyms and std_col not in used_std_cols:
                rename_map[orig_col] = std_col
                used_std_cols.add(std_col)
                break

    out = out.rename(columns=rename_map)

    for std_col in COLUMN_SYNONYMS.keys():
        if std_col not in out.columns:
            out[std_col] = ""

    return out


def parse_plan_json(plan_json_raw: Any, target_ym: str = "2027-01") -> Dict[str, Any]:
    """
    시트 '제출내역'의 '회차별계획(JSON)' 컬럼을 파싱하여
    - 전체 유효 수업일 목록 (holiday: false)
    - 검수 대상 연월(예: 2027-01)의 유효 수업 횟수 및 마지막 수업일
    - 조기 개강(전월 12월 개강 ~ 당월 말일) 포함 유효 수업 횟수 및 마지막 수업일
    을 계산하여 반환합니다.
    """
    result = {
        "is_valid_json": False,
        "all_active_dates": [],
        "all_holiday_dates": [],
        "total_active_count": 0,
        "json_first_date": "",
        "json_last_date": "",
        "month_active_dates": [],
        "month_active_count": 0,
        "month_first_date": "",
        "month_last_date": "",
        "early_incl_active_dates": [],
        "early_incl_active_count": 0,
        "early_incl_last_date": "",
        "topics_summary": "",
    }

    if pd.isna(plan_json_raw) or plan_json_raw is None:
        return result

    raw_str = str(plan_json_raw).strip()
    if not raw_str or raw_str in ("-", "[]", "{}", "nan", "None"):
        return result

    try:
        parsed = json.loads(raw_str)
    except Exception:
        try:
            m = re.search(r"(\[.*\]|\{.*\})", raw_str, flags=re.DOTALL)
            if m:
                candidate = m.group(1)
                parsed = json.loads(candidate)
            else:
                return result
        except Exception:
            return result

    items: List[Dict[str, Any]] = []
    if isinstance(parsed, list):
        items = [x for x in parsed if isinstance(x, dict)]
    elif isinstance(parsed, dict):
        for key in ("sessions", "plan", "items", "schedule", "data", "회차별계획"):
            if key in parsed and isinstance(parsed[key], list):
                items = [x for x in parsed[key] if isinstance(x, dict)]
                break
        if not items:
            items = [v for v in parsed.values() if isinstance(v, dict)]

    if not items:
        return result

    result["is_valid_json"] = True
    active_dates: List[str] = []
    holiday_dates: List[str] = []
    topics: List[str] = []

    for idx, item in enumerate(items, start=1):
        raw_d = (
            item.get("date")
            or item.get("날짜")
            or item.get("수업일")
            or item.get("day")
            or item.get("일자")
            or ""
        )
        std_d = standardize_date(raw_d)

        hol_val = item.get("holiday", item.get("is_holiday", item.get("isHoliday", item.get("휴강", False))))
        is_holiday = False
        if isinstance(hol_val, bool):
            is_holiday = hol_val
        elif isinstance(hol_val, (int, float)):
            is_holiday = bool(hol_val)
        elif isinstance(hol_val, str):
            is_holiday = hol_val.strip().lower() in ("true", "1", "y", "yes", "휴강")

        topic_val = str(
            item.get("topic")
            or item.get("content")
            or item.get("title")
            or item.get("주제")
            or item.get("내용")
            or item.get("단원")
            or ""
        ).strip()

        if "휴강" in topic_val and not is_holiday and topic_val in ("휴강", "[휴강]"):
            is_holiday = True

        if std_d:
            if is_holiday:
                holiday_dates.append(std_d)
            else:
                active_dates.append(std_d)
                if topic_val:
                    topics.append(f"{idx}회({std_d[5:]}): {topic_val}")
        else:
            if not is_holiday and topic_val:
                topics.append(f"{idx}회: {topic_val}")

    active_dates = sorted(set(active_dates))
    holiday_dates = sorted(set(holiday_dates))

    result["all_active_dates"] = active_dates
    result["all_holiday_dates"] = holiday_dates
    result["total_active_count"] = len(active_dates)
    result["json_first_date"] = active_dates[0] if active_dates else ""
    result["json_last_date"] = active_dates[-1] if active_dates else ""
    result["topics_summary"] = " | ".join(topics[:8]) + (" ..." if len(topics) > 8 else "")

    target_ym = target_ym.strip()
    month_dates = [d for d in active_dates if d.startswith(target_ym)]
    result["month_active_dates"] = month_dates
    result["month_active_count"] = len(month_dates)
    result["month_first_date"] = month_dates[0] if month_dates else ""
    result["month_last_date"] = month_dates[-1] if month_dates else ""

    if re.match(r"^\d{4}-\d{2}$", target_ym):
        yr, mo = map(int, target_ym.split("-"))
        last_day = calendar.monthrange(yr, mo)[1]
        target_end_date = f"{yr:04d}-{mo:02d}-{last_day:02d}"
        if mo == 1:
            early_start_floor = f"{yr - 1:04d}-12-01"
        else:
            early_start_floor = f"{yr:04d}-{mo - 1:02d}-15"

        early_incl_dates = [d for d in active_dates if early_start_floor <= d <= target_end_date]
        result["early_incl_active_dates"] = early_incl_dates
        result["early_incl_active_count"] = len(early_incl_dates)
        result["early_incl_last_date"] = early_incl_dates[-1] if early_incl_dates else ""

    return result


def read_local_csv_or_excel(file_source: Union[str, Path, bytes, io.BytesIO, Any]) -> pd.DataFrame:
    """로컬 CSV 또는 Excel(.xlsx) 파일을 인코딩 오류 없이 안전하게 DataFrame으로 로드합니다."""
    filename = getattr(file_source, "name", str(file_source) if isinstance(file_source, (str, Path)) else "")
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
        raise ValueError("지원하지 않는 시트 파일 형식입니다.")

    if str(filename).lower().endswith((".xlsx", ".xlsm")):
        xls = pd.ExcelFile(io.BytesIO(raw_bytes))
        sheet_names = xls.sheet_names
        df_list = None
        df_sub = None
        if "강좌목록" in sheet_names:
            df_list = pd.read_excel(xls, sheet_name="강좌목록")
        if "제출내역" in sheet_names:
            df_sub = pd.read_excel(xls, sheet_name="제출내역")
        if df_list is not None or df_sub is not None:
            return merge_course_and_submission_tabs(df_list, df_sub)
        return pd.read_excel(xls, sheet_name=0)

    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            text = raw_bytes.decode(enc)
            return pd.read_csv(io.StringIO(text))
        except Exception:
            continue

    text = raw_bytes.decode("cp949", errors="replace")
    return pd.read_csv(io.StringIO(text))


def merge_course_and_submission_tabs(
    course_df: Optional[pd.DataFrame],
    submission_df: Optional[pd.DataFrame],
) -> pd.DataFrame:
    """
    Google Sheets의 '강좌목록' 탭과 '제출내역' 탭을 병합/보완합니다.
    """
    c_std = standardize_sheet_columns(course_df) if course_df is not None and not course_df.empty else None
    s_std = standardize_sheet_columns(submission_df) if submission_df is not None and not submission_df.empty else None

    if c_std is None and s_std is None:
        return standardize_sheet_columns(pd.DataFrame())
    if c_std is None:
        s_std["데이터출처탭"] = "제출내역"
        return s_std
    if s_std is None:
        c_std["데이터출처탭"] = "강좌목록"
        return c_std

    if c_std.equals(s_std):
        s_std["데이터출처탭"] = "통합시트"
        return s_std

    c_std = c_std.copy()
    s_std = s_std.copy()
    c_std["데이터출처탭"] = "강좌목록"
    s_std["데이터출처탭"] = "제출내역"

    def _make_lookup_key(row: pd.Series) -> str:
        inst = str(row.get("강사명", "")).strip()
        day = str(row.get("수업요일", "")).strip()
        tm = str(row.get("수업시간", "")).strip()
        title = str(row.get("강좌명", "")).strip()
        return f"{inst}|{day}|{tm}|{title}"

    c_lookup: Dict[str, pd.Series] = {}
    for _, r in c_std.iterrows():
        c_lookup[_make_lookup_key(r)] = r
        short_k = f"{str(r.get('강사명', '')).strip()}|{str(r.get('강좌명', '')).strip()}"
        if short_k not in c_lookup:
            c_lookup[short_k] = r

    s_keys_seen = set()
    for idx, r in s_std.iterrows():
        k_full = _make_lookup_key(r)
        k_short = f"{str(r.get('강사명', '')).strip()}|{str(r.get('강좌명', '')).strip()}"
        s_keys_seen.add(k_full)
        s_keys_seen.add(k_short)
        match_r = c_lookup.get(k_full) if k_full in c_lookup else c_lookup.get(k_short)
        if match_r is not None:
            for col in ["저장폴더URL", "문서URL", "강좌구분", "수능유형", "대상학년", "과목"]:
                if not str(r.get(col, "")).strip() and str(match_r.get(col, "")).strip():
                    s_std.at[idx, col] = match_r.get(col, "")

    extra_c_rows = []
    for _, r in c_std.iterrows():
        k_full = _make_lookup_key(r)
        k_short = f"{str(r.get('강사명', '')).strip()}|{str(r.get('강좌명', '')).strip()}"
        if k_full not in s_keys_seen and k_short not in s_keys_seen:
            extra_c_rows.append(r)

    if extra_c_rows:
        combined = pd.concat([s_std, pd.DataFrame(extra_c_rows)], ignore_index=True)
    else:
        combined = s_std

    return combined


def load_google_sheets_data(sheet_url: str, use_oauth: bool = False) -> pd.DataFrame:
    """
    Google Sheets에서 '강좌목록' 및 '제출내역' 탭을 읽어 병합된 DataFrame으로 반환합니다.
    [3중 안전장치(Fallback) 구조]
    1순위: requests 기반 다이렉트 CSV 로더 (fetch_google_sheet_tab: '강좌목록', '제출내역' 탭 직접 추출)
    2순위: 탭 이름이 다를 경우 URL 내 gid 기반 CSV Export 다이렉트 로더 (fetch_google_sheet_by_gid)
    3순위: use_oauth=True 요청 시 내 구글 계정 OAuth(pydata-google-auth)를 통한 gspread 직접 로드
    """
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", str(sheet_url))
    if not match:
        raise ValueError("유효한 구글 시트 URL이 아닙니다. URL에 '/d/시트ID'가 포함되어야 합니다.")
    sheet_id = match.group(1)

    if not use_oauth:
        course_df: Optional[pd.DataFrame] = None
        submission_df: Optional[pd.DataFrame] = None
        tab_errors: List[str] = []

        # [1차 안전장치] '강좌목록' 및 '제출내역' 탭 requests 다이렉트 로드
        for tab_name in ("강좌목록", "제출내역"):
            try:
                df_tab = fetch_google_sheet_tab(sheet_url, tab_name)
                if df_tab is not None and not df_tab.empty:
                    if tab_name == "강좌목록":
                        course_df = df_tab
                    else:
                        submission_df = df_tab
            except PermissionError:
                # 비공개 시트인 경우 즉시 PermissionError 전파하여 화면에 정확한 안내 표시
                raise
            except Exception as e:
                tab_errors.append(f"[{tab_name}] 탭 로드 오류: {e}")

        if (course_df is not None and not course_df.empty) or (
            submission_df is not None and not submission_df.empty
        ):
            return merge_course_and_submission_tabs(course_df, submission_df)

        # [2차 안전장치] 탭 이름이 다를 경우 URL의 gid(또는 기본 시트)로 직접 Export CSV 로드
        try:
            gid_df = fetch_google_sheet_by_gid(sheet_url)
            if gid_df is not None and not gid_df.empty:
                return merge_course_and_submission_tabs(gid_df, None)
        except PermissionError:
            raise
        except Exception as e:
            tab_errors.append(f"[GID 기본 시트] 로드 오류: {e}")

        if tab_errors:
            raise RuntimeError(" / ".join(tab_errors))

    # [3차 안전장치] 내 구글 계정 OAuth (pydata-google-auth) 직접 로드
    client = get_user_gspread_client()
    try:
        sh = client.open_by_key(sheet_id)
    except Exception as e:
        raise SheetAccessDeniedError(f"구글 시트 열기 실패 ({e}). {SHEET_SHARE_GUIDE_MSG}") from e

    worksheets = {ws.title.strip(): ws for ws in sh.worksheets()}
    course_df = None
    submission_df = None

    if "강좌목록" in worksheets:
        course_df = _worksheet_to_dataframe(worksheets["강좌목록"])
    if "제출내역" in worksheets:
        submission_df = _worksheet_to_dataframe(worksheets["제출내역"])

    if (course_df is None or course_df.empty) and (submission_df is None or submission_df.empty):
        gid = extract_gid_from_url(sheet_url)
        target_ws = None
        if gid is not None:
            try:
                target_ws = sh.get_worksheet_by_id(gid)
            except Exception:
                target_ws = None
        if target_ws is None:
            target_ws = sh.get_worksheet(0)
        if target_ws is not None:
            course_df = _worksheet_to_dataframe(target_ws)

    return merge_course_and_submission_tabs(course_df, submission_df)


def write_crosscheck_to_google_sheets(
    sheet_url: str,
    crosscheck_result: Dict[str, Any],
    mapper: Any,
    target_ym: str = "2027-01",
    write_summary_tab: bool = True,
    summary_tab_name: str = "검수결과_3단대조",
    update_existing_tabs: bool = True,
) -> Dict[str, Any]:
    """
    내 구글 계정(편집자 권한)으로 3단 교차 검수 결과를 Google Sheets에 직접 기록(Write-back)합니다.

    1) update_existing_tabs=True 인 경우:
       - 시트 내 기존 '강좌목록' (및 '제출내역') 탭 우측 끝에
         [RMS_검수상태, RMS_검수상세사유, RMS_매칭강좌코드, RMS_개강~종강(회차), 검수일시] 컬럼을 자동으로 추가/업데이트합니다.
    2) write_summary_tab=True 인 경우:
       - 시트 내에 '검수결과_3단대조' 탭 및 '미제출_강좌목록' 탭을 자동 생성(또는 갱신)하여
         RMS 전체 강좌 기준 3단 대조표와 미제출 리스트를 한눈에 볼 수 있도록 기록합니다.
    """
    sheet_id = extract_spreadsheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("유효한 Google Sheets URL이 아닙니다.")

    client = get_user_gspread_client()
    sh = client.open_by_key(sheet_id)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    comparison_df: pd.DataFrame = crosscheck_result.get("comparison_df", pd.DataFrame())
    unsubmitted_df: pd.DataFrame = crosscheck_result.get("unsubmitted_df", pd.DataFrame())
    excluded_dup_df: pd.DataFrame = crosscheck_result.get("excluded_duplicates_df", pd.DataFrame())
    excluded_test_df: pd.DataFrame = crosscheck_result.get("excluded_test_df", pd.DataFrame())
    excluded_period_df: pd.DataFrame = crosscheck_result.get("excluded_period_df", pd.DataFrame())
    unopened_sheet_df: pd.DataFrame = crosscheck_result.get("unopened_sheet_df", pd.DataFrame())

    logs: List[str] = []
    updated_tabs: List[str] = []

    # -----------------------------------------------------------------
    # 1) 기존 '강좌목록' / '제출내역' 탭 우측에 검수 결과 컬럼 직접 기입 (Write-back)
    # -----------------------------------------------------------------
    if update_existing_tabs and not comparison_df.empty:
        # 매칭 조회용 룩업 구성: (정규화강사명, 시트강좌명) 또는 (시트_강사명, 시트 강좌명)
        comp_by_sheet_key: Dict[str, pd.Series] = {}
        comp_by_inst_sched: Dict[str, pd.Series] = {}

        for _, c_row in comparison_df.iterrows():
            s_inst = str(c_row.get("시트_강사명", "")).strip()
            norm_inst = str(c_row.get("정규화강사명", "")).strip()
            s_title = str(c_row.get("시트 강좌명", "")).strip()
            s_sched = str(c_row.get("시트 요일/시간", "")).strip()
            if s_inst and s_inst != "-":
                comp_by_sheet_key[f"{s_inst}|{s_title}"] = c_row
                comp_by_sheet_key[f"{norm_inst}|{s_title}"] = c_row
                if s_sched and s_sched != "-":
                    comp_by_inst_sched[f"{norm_inst}|{s_sched}"] = c_row

        # 제외된 행(중복/테스트/타시기/미개설) 사유 룩업
        excluded_lookup: Dict[str, Tuple[str, str]] = {}
        for ex_df, badge_label in [
            (excluded_dup_df, "⚪ 중복제출 제외 (구버전)"),
            (excluded_test_df, "⚪ 테스트 데이터 제외"),
            (excluded_period_df, "⚪ 타 시기 강좌 제외"),
            (unopened_sheet_df, "⚠️ RMS 미개설 (폐강/미등록 확인)"),
        ]:
            if ex_df is not None and not ex_df.empty:
                for _, erow in ex_df.iterrows():
                    e_inst = str(erow.get("강사명", "")).strip()
                    e_title = str(erow.get("강좌명", "")).strip()
                    e_ts = str(erow.get("최근제출일시", "")).strip()
                    e_reason = str(erow.get("제외사유", "")).strip() or badge_label
                    excluded_lookup[f"{e_inst}|{e_title}|{e_ts}"] = (badge_label, e_reason)
                    if f"{e_inst}|{e_title}" not in excluded_lookup:
                        excluded_lookup[f"{e_inst}|{e_title}"] = (badge_label, e_reason)

        worksheets_map = {ws.title.strip(): ws for ws in sh.worksheets()}
        target_existing_names = [t for t in ("강좌목록", "제출내역") if t in worksheets_map]

        for tab_name in target_existing_names:
            ws = worksheets_map[tab_name]
            all_vals = ws.get_all_values()
            if not all_vals or len(all_vals) < 1:
                continue

            headers = [str(h).strip() for h in all_vals[0]]
            data_rows = all_vals[1:]

            # 기존 시트에서 강사명, 강좌명, 수업요일, 수업시간, 제출일시 컬럼 인덱스 찾기
            def _find_col_idx(candidates: List[str]) -> int:
                norm_cands = [re.sub(r"\s+", "", c).lower() for c in candidates]
                for i, h in enumerate(headers):
                    if re.sub(r"\s+", "", h).lower() in norm_cands:
                        return i
                return -1

            idx_inst = _find_col_idx(COLUMN_SYNONYMS["강사명"])
            idx_title = _find_col_idx(COLUMN_SYNONYMS["강좌명"])
            idx_day = _find_col_idx(COLUMN_SYNONYMS["수업요일"])
            idx_time = _find_col_idx(COLUMN_SYNONYMS["수업시간"])
            idx_ts = _find_col_idx(COLUMN_SYNONYMS["최근제출일시"])

            if idx_inst == -1 and idx_title == -1:
                continue

            writeback_cols = [
                "RMS_검수상태",
                "RMS_검수상세사유",
                "RMS_매칭강좌코드",
                "RMS_개강~종강(회차)",
                "RMS_검수일시",
            ]

            # 이미 writeback_cols가 헤더에 있으면 해당 열 위치를 덮어쓰고, 없으면 우측 끝에 추가
            if "RMS_검수상태" in headers:
                start_col_idx = headers.index("RMS_검수상태")  # 0-based
            else:
                start_col_idx = len(headers)

            # 필요한 경우 시트 열 개수 확장
            needed_cols = start_col_idx + len(writeback_cols)
            if ws.col_count < needed_cols:
                ws.resize(rows=max(ws.row_count, len(all_vals) + 5), cols=needed_cols)

            out_matrix: List[List[str]] = [writeback_cols]
            matched_row_cnt = 0

            for r_vals in data_rows:
                def _get_cell(c_i: int) -> str:
                    return str(r_vals[c_i]).strip() if 0 <= c_i < len(r_vals) else ""

                r_inst = _get_cell(idx_inst)
                r_title = _get_cell(idx_title)
                r_day = _get_cell(idx_day)
                r_time = _get_cell(idx_time)
                r_ts = _get_cell(idx_ts)

                if not r_inst and not r_title:
                    out_matrix.append(["", "", "", "", ""])
                    continue

                norm_inst = mapper.normalize(r_inst) if mapper else r_inst
                ps = parse_schedule_text(r_time, r_day)
                sched_norm = ps["formatted_schedule"]

                # 1순위: 제외 목록(중복/테스트/타시기) 확인 (타임스탬프 포함 키 우선)
                if f"{r_inst}|{r_title}|{r_ts}" in excluded_lookup:
                    b_lbl, r_msg = excluded_lookup[f"{r_inst}|{r_title}|{r_ts}"]
                    out_matrix.append([b_lbl, r_msg, "-", "-", now_str])
                    continue

                # 2순위: 정상/불일치 매칭 결과 확인
                matched_c = (
                    comp_by_sheet_key.get(f"{r_inst}|{r_title}")
                    or comp_by_sheet_key.get(f"{norm_inst}|{r_title}")
                    or comp_by_inst_sched.get(f"{norm_inst}|{sched_norm}")
                )

                if matched_c is not None:
                    matched_row_cnt += 1
                    out_matrix.append(
                        [
                            str(matched_c.get("최종 검수 상태", "")),
                            str(matched_c.get("검수 상세 사유", "")),
                            str(matched_c.get("강좌코드", "")),
                            str(matched_c.get("RMS 개강~종강(회차)", "")),
                            now_str,
                        ]
                    )
                elif f"{r_inst}|{r_title}" in excluded_lookup:
                    b_lbl, r_msg = excluded_lookup[f"{r_inst}|{r_title}"]
                    out_matrix.append([b_lbl, r_msg, "-", "-", now_str])
                else:
                    out_matrix.append(
                        [
                            "⚠️ RMS 미개설 (확인 필요)",
                            "RMS 수업목록에서 일치하는 강좌를 찾지 못함",
                            "미매칭",
                            "-",
                            now_str,
                        ]
                    )

            start_col_letter = gspread.utils.rowcol_to_a1(1, start_col_idx + 1)
            end_col_letter = gspread.utils.rowcol_to_a1(
                len(out_matrix), start_col_idx + len(writeback_cols)
            )
            range_str = f"{start_col_letter}:{end_col_letter}"
            ws.update(values=out_matrix, range_name=range_str)

            # 헤더 셀 서식 적용
            header_range = f"{gspread.utils.rowcol_to_a1(1, start_col_idx + 1)}:{gspread.utils.rowcol_to_a1(1, start_col_idx + len(writeback_cols))}"
            try:
                ws.format(
                    header_range,
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

            updated_tabs.append(tab_name)
            logs.append(
                f"✅ 기존 시트 탭 [{tab_name}] 우측({start_col_letter}~)에 검수 결과 {len(data_rows)}행(RMS 매칭 {matched_row_cnt}건) 기입 완료"
            )

    # -----------------------------------------------------------------
    # 2) 전용 결과 탭 ('검수결과_3단대조' 및 '미제출_강좌목록') 생성/갱신
    # -----------------------------------------------------------------
    if write_summary_tab and not comparison_df.empty:
        worksheets_map = {ws.title.strip(): ws for ws in sh.worksheets()}
        export_cols = [
            "강좌코드",
            "과목",
            "학년",
            "강사명",
            "RMS 강좌명",
            "시트 강좌명",
            "강좌명_유사도(%)",
            "RMS 요일/시간",
            "시트 요일/시간",
            "RMS 개강~종강(회차)",
            "시트 개강~종강(총회차 / 당월회차)",
            "로컬 파일 존재여부(포맷)",
            "최종 검수 상태",
            "검수 상세 사유",
            "시트_문서URL",
            "최근제출일시",
        ]
        valid_cols = [c for c in export_cols if c in comparison_df.columns]
        sub_df = comparison_df[valid_cols].copy()
        sub_df["검수일시"] = now_str

        matrix = [list(sub_df.columns)] + sub_df.fillna("").astype(str).values.tolist()

        if summary_tab_name in worksheets_map:
            ws_sum = worksheets_map[summary_tab_name]
            ws_sum.clear()
            if ws_sum.row_count < len(matrix) + 5 or ws_sum.col_count < len(matrix[0]) + 2:
                ws_sum.resize(rows=len(matrix) + 10, cols=len(matrix[0]) + 2)
        else:
            ws_sum = sh.add_worksheet(
                title=summary_tab_name,
                rows=len(matrix) + 20,
                cols=len(matrix[0]) + 2,
            )

        ws_sum.update(values=matrix, range_name="A1")
        try:
            last_hdr_a1 = gspread.utils.rowcol_to_a1(1, len(matrix[0]))
            ws_sum.format(
                f"A1:{last_hdr_a1}",
                {
                    "backgroundColor": {"red": 0.06, "green": 0.17, "blue": 0.27},
                    "textFormat": {
                        "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                        "bold": True,
                    },
                    "horizontalAlignment": "CENTER",
                },
            )
        except Exception:
            pass

        updated_tabs.append(summary_tab_name)
        logs.append(f"✅ [{summary_tab_name}] 시트 탭에 전체 {len(sub_df)}개 강좌 3단 검수 대조표 기록 완료")

        # 미제출 강좌 탭도 함께 기록
        if not unsubmitted_df.empty:
            unsub_tab_name = "미제출_강좌목록"
            unsub_cols = [
                "강좌코드",
                "영역",
                "과목",
                "학년",
                "RMS_강사명",
                "RMS 강좌명",
                "RMS 요일/시간",
                "RMS 개강~종강(회차)",
                "최종 검수 상태",
            ]
            u_cols = [c for c in unsub_cols if c in unsubmitted_df.columns]
            u_df = unsubmitted_df[u_cols].rename(columns={"RMS_강사명": "강사명"}).copy()
            u_df["검수일시"] = now_str
            u_matrix = [list(u_df.columns)] + u_df.fillna("").astype(str).values.tolist()

            if unsub_tab_name in worksheets_map:
                ws_unsub = worksheets_map[unsub_tab_name]
                ws_unsub.clear()
            else:
                ws_unsub = sh.add_worksheet(
                    title=unsub_tab_name,
                    rows=len(u_matrix) + 20,
                    cols=len(u_matrix[0]) + 2,
                )
            ws_unsub.update(values=u_matrix, range_name="A1")
            updated_tabs.append(unsub_tab_name)
            logs.append(f"✅ [{unsub_tab_name}] 시트 탭에 미제출 강좌 {len(u_df)}건 기록 완료")

    return {
        "updated_tabs": updated_tabs,
        "logs": logs,
        "spreadsheet_title": sh.title,
        "spreadsheet_url": sh.url,
    }


def is_course_in_target_period(
    start_date: str,
    end_date: str,
    plan_info: Dict[str, Any],
    target_ym: str = "2027-01",
) -> Tuple[bool, str]:
    """
    검수 대상 연월(예: '2027-01')에 해당하는 강좌인지 판별합니다.
    - 12월 말에 조기 개강하여 1월까지 이어지는 윈터/정규 강좌는 포함합니다.
    - 10~11월 개강 및 종강 강좌 등 타 시기 강좌는 제외합니다.
    """
    target_ym = target_ym.strip()
    if not re.match(r"^\d{4}-\d{2}$", target_ym):
        return True, ""

    yr, mo = map(int, target_ym.split("-"))
    last_day = calendar.monthrange(yr, mo)[1]
    ym_start = f"{yr:04d}-{mo:02d}-01"
    ym_end = f"{yr:04d}-{mo:02d}-{last_day:02d}"

    if mo == 1:
        early_open_floor = f"{yr - 1:04d}-12-01"
    else:
        early_open_floor = f"{yr:04d}-{mo - 1:02d}-01"

    if plan_info.get("month_active_count", 0) > 0:
        return True, ""

    if not start_date and not end_date:
        return True, ""

    if start_date and start_date.startswith(target_ym):
        return True, ""

    if start_date and end_date:
        if early_open_floor <= start_date <= ym_end and end_date >= ym_start:
            return True, ""
        if end_date < ym_start:
            return False, f"당월 이전 종강 ({start_date} ~ {end_date})"
        if start_date > ym_end:
            return False, f"당월 이후 개강 ({start_date} ~ {end_date})"
        if start_date < early_open_floor:
            return False, f"이전 시기 개강 강좌 ({start_date} ~ {end_date})"

    elif start_date and not end_date:
        if start_date < early_open_floor:
            return False, f"이전 시기 개강 ({start_date})"
        if start_date > ym_end:
            return False, f"차기 월 개강 ({start_date})"

    elif end_date and not start_date:
        if end_date < ym_start:
            return False, f"당월 이전 종강 (~ {end_date})"

    return True, ""


def preprocess_sheet_dataframe(
    raw_df: pd.DataFrame,
    instructor_normalizer: Callable[[str], str],
    target_ym: str = "2027-01",
    filter_other_period: bool = True,
) -> Dict[str, pd.DataFrame]:
    """
    강의계획서 수령 현황 DataFrame을 표준화하고 3단계 전처리(테스트 제거, 타 시기 필터링, 중복 제출 제거)를 수행합니다.
    """
    std_df = standardize_sheet_columns(raw_df)
    if std_df.empty:
        empty = std_df.copy()
        return {
            "valid_df": empty,
            "excluded_duplicates_df": empty,
            "excluded_test_df": empty,
            "excluded_period_df": empty,
        }

    df = std_df.copy()

    for col in [
        "강사명",
        "과목",
        "대상학년",
        "강좌명",
        "강좌구분",
        "수능유형",
        "수업요일",
        "수업시간",
        "개강일",
        "종강일",
        "최근제출일시",
        "회차별계획(JSON)",
        "저장폴더URL",
        "문서URL",
    ]:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .apply(lambda v: "" if str(v).strip().lower() == "nan" else str(v).strip())
        )

    df = df[(df["강사명"] != "") | (df["강좌명"] != "")].reset_index(drop=True)

    def _to_int(val: Any) -> int:
        if pd.isna(val):
            return 0
        m = re.search(r"\d+", str(val))
        return int(m.group(0)) if m else 0

    df["총회차"] = df["총회차"].apply(_to_int)

    df["개강일_원본"] = df["개강일"]
    df["종강일_원본"] = df["종강일"]
    df["개강일"] = df["개강일"].apply(standardize_date)
    df["종강일"] = df["종강일"].apply(standardize_date)

    parsed_titles = df["강좌명"].apply(strip_course_brackets)
    df["순수강좌명"] = [pt[0] for pt in parsed_titles]
    df["말머리태그"] = [pt[1] for pt in parsed_titles]

    parsed_scheds = [
        parse_schedule_text(row["수업시간"], row["수업요일"]) for _, row in df.iterrows()
    ]
    df["표준요일"] = [ps["days_str"] for ps in parsed_scheds]
    df["요일_리스트"] = [ps["days_list"] for ps in parsed_scheds]
    df["시작시간"] = [ps["primary_start"] for ps in parsed_scheds]
    df["전체시작시간"] = [ps["start_times_str"] for ps in parsed_scheds]
    df["요일시간키_리스트"] = [ps["day_start_keys"] for ps in parsed_scheds]
    df["표준수업시간"] = [ps["formatted_schedule"] for ps in parsed_scheds]

    df["정규화강사명"] = df["강사명"].apply(instructor_normalizer)
    df["문서ID"] = df["문서URL"].apply(extract_google_doc_id)

    plan_parsed_list = [
        parse_plan_json(row["회차별계획(JSON)"], target_ym=target_ym) for _, row in df.iterrows()
    ]
    df["JSON_유효여부"] = [p["is_valid_json"] for p in plan_parsed_list]
    df["JSON_전체수업일"] = [p["all_active_dates"] for p in plan_parsed_list]
    df["JSON_휴강일"] = [p["all_holiday_dates"] for p in plan_parsed_list]
    df["JSON_총회차"] = [p["total_active_count"] for p in plan_parsed_list]
    df["JSON_첫수업일"] = [p["json_first_date"] for p in plan_parsed_list]
    df["JSON_마지막수업일"] = [p["json_last_date"] for p in plan_parsed_list]
    df["당월_회차"] = [p["month_active_count"] for p in plan_parsed_list]
    df["당월_개강일"] = [p["month_first_date"] for p in plan_parsed_list]
    df["당월_종강일"] = [p["month_last_date"] for p in plan_parsed_list]
    df["조기개강포함_당월회차"] = [p["early_incl_active_count"] for p in plan_parsed_list]
    df["조기개강포함_당월종강일"] = [p["early_incl_last_date"] for p in plan_parsed_list]
    df["회차별계획_요약"] = [p["topics_summary"] for p in plan_parsed_list]

    for idx, row in df.iterrows():
        if not row["개강일"] and row["JSON_첫수업일"]:
            df.at[idx, "개강일"] = row["JSON_첫수업일"]
        if not row["종강일"] and row["JSON_마지막수업일"]:
            df.at[idx, "종강일"] = row["JSON_마지막수업일"]
        if row["총회차"] == 0 and row["JSON_총회차"] > 0:
            df.at[idx, "총회차"] = row["JSON_총회차"]

    # 1. 테스트 데이터 필터링
    def _is_test_row(row: pd.Series) -> Tuple[bool, str]:
        inst_raw = str(row["강사명"]).strip()
        inst_norm = str(row["정규화강사명"]).strip()
        title_raw = str(row["강좌명"]).strip()
        if inst_raw in TEST_INSTRUCTORS or inst_norm in TEST_INSTRUCTORS:
            return True, f"테스트 강사명 ({inst_raw})"
        if "test" in title_raw.lower() or "테스트" in title_raw:
            return True, f"테스트 강좌명 ({title_raw})"
        return False, ""

    test_checks = [_is_test_row(r) for _, r in df.iterrows()]
    df["제외사유"] = [tc[1] for tc in test_checks]
    is_test_mask = pd.Series([tc[0] for tc in test_checks], index=df.index)

    excluded_test_df = df[is_test_mask].copy().reset_index(drop=True)
    rem_df = df[~is_test_mask].copy().reset_index(drop=True)

    # 2. 타 시기 강좌 필터링
    if filter_other_period and not rem_df.empty:
        period_checks = [
            is_course_in_target_period(
                row["개강일"],
                row["종강일"],
                {"month_active_count": row["당월_회차"]},
                target_ym=target_ym,
            )
            for _, row in rem_df.iterrows()
        ]
        keep_mask = pd.Series([pc[0] for pc in period_checks], index=rem_df.index)
        rem_df["제외사유"] = [pc[1] for pc in period_checks]
        excluded_period_df = rem_df[~keep_mask].copy().reset_index(drop=True)
        rem_df = rem_df[keep_mask].copy().reset_index(drop=True)
    else:
        excluded_period_df = df.iloc[0:0].copy()

    # 3. 중복 제출 제거
    if not rem_df.empty:
        def _parse_timestamp_key(ts_str: str) -> str:
            if not ts_str:
                return ""
            s = ts_str.strip()
            m_kr = re.search(
                r"(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})[.\-일\s]*(오전|오후|AM|PM)?\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?",
                s,
                flags=re.IGNORECASE,
            )
            if m_kr:
                yyyy, mm, dd = int(m_kr.group(1)), int(m_kr.group(2)), int(m_kr.group(3))
                ampm = (m_kr.group(4) or "").strip().upper()
                hh, mi = int(m_kr.group(5)), int(m_kr.group(6))
                ss = int(m_kr.group(7) or 0)
                if ampm in ("오후", "PM") and hh < 12:
                    hh += 12
                elif ampm in ("오전", "AM") and hh == 12:
                    hh = 0
                return f"{yyyy:04d}-{mm:02d}-{dd:02d} {hh:02d}:{mi:02d}:{ss:02d}"
            return s

        rem_df["_정렬용제출일시"] = rem_df["최근제출일시"].apply(_parse_timestamp_key)
        rem_df["_원본순번"] = range(len(rem_df))

        def _make_dedup_key(row: pd.Series) -> str:
            inst = str(row["정규화강사명"]).strip() or str(row["강사명"]).strip()
            days = str(row["표준요일"]).strip()
            st_times = str(row["전체시작시간"]).strip() or str(row["시작시간"]).strip()
            if days and st_times:
                return f"{inst}__{days}__{st_times}"
            return f"{inst}__{days}__{st_times}__{str(row['순수강좌명']).strip()}"

        rem_df["_중복그룹키"] = [_make_dedup_key(r) for _, r in rem_df.iterrows()]

        rem_df_sorted = rem_df.sort_values(
            by=["_중복그룹키", "_정렬용제출일시", "_원본순번"],
            ascending=[True, False, False],
        )

        is_dup = rem_df_sorted.duplicated(subset=["_중복그룹키"], keep="first")

        valid_df = rem_df_sorted[~is_dup].sort_values(by="_원본순번").copy()
        excluded_duplicates_df = rem_df_sorted[is_dup].sort_values(
            by=["_중복그룹키", "_정렬용제출일시"], ascending=[True, False]
        ).copy()

        latest_ts_by_key = dict(zip(valid_df["_중복그룹키"], valid_df["최근제출일시"]))
        excluded_duplicates_df["제외사유"] = [
            f"동일 강사·요일·시간 중복 제출 (최신본 유지: {latest_ts_by_key.get(k, '최신 행')})"
            for k in excluded_duplicates_df["_중복그룹키"]
        ]

        drop_temp_cols = ["_정렬용제출일시", "_원본순번", "_중복그룹키"]
        valid_df = valid_df.drop(columns=drop_temp_cols, errors="ignore").reset_index(drop=True)
        excluded_duplicates_df = excluded_duplicates_df.drop(
            columns=drop_temp_cols, errors="ignore"
        ).reset_index(drop=True)
    else:
        valid_df = rem_df.copy()
        excluded_duplicates_df = df.iloc[0:0].copy()

    return {
        "valid_df": valid_df,
        "excluded_duplicates_df": excluded_duplicates_df,
        "excluded_test_df": excluded_test_df,
        "excluded_period_df": excluded_period_df,
    }
