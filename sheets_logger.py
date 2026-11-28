import os
import re
import json
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. 설정 및 상수 정의
# ==========================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_HEADERS = [
    "수집일시(KST)",
    "학원명",
    "분류",
    "노출출처",
    "상태",
    "제목",
    "원문링크",
    "정규화키",
]

KST = timezone(timedelta(hours=9))

DEFAULT_LOCAL_KEY_FILES = [
    "service_account.json",
    "gcp_credentials.json",
    "google_credentials.json",
    "credentials.json",
]


# ==========================================
# 2. 정규화 및 인증 헬퍼
# ==========================================
def _fallback_normalize_title(text):
    if not text:
        return ""
    t = re.sub(r"https?://\S+", "", text)
    t = re.sub(r"\[.*?\]", "", t)
    t = re.sub(r"[^a-zA-Z0-9가-힣]", "", t)
    norm = t.lower().strip()
    if len(norm) < 2:
        t_fallback = re.sub(r"https?://\S+", "", text)
        t_fallback = re.sub(r"[^a-zA-Z0-9가-힣]", "", t_fallback).lower().strip()
        if t_fallback:
            return t_fallback
    return norm


def _load_credentials():
    """
    GitHub Actions 환경 변수(GCP_CREDENTIALS JSON 문자열) 또는
    로컬 서비스 계정 JSON 파일에서 Google Service Account Credentials 로드
    """
    raw_creds = os.environ.get("GCP_CREDENTIALS", "").strip()
    if raw_creds:
        # 환경 변수에 JSON 문자열이 들어온 경우
        if raw_creds.startswith("{"):
            creds_info = json.loads(raw_creds)
            if "private_key" in creds_info and isinstance(creds_info["private_key"], str):
                creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
            return Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        # 환경 변수에 파일 경로가 들어온 경우
        if os.path.isfile(raw_creds):
            return Credentials.from_service_account_file(raw_creds, scopes=SCOPES)

    # 로컬 환경 변수 지정 경로 탐색
    for env_key in ("GCP_CREDENTIALS_FILE", "GOOGLE_APPLICATION_CREDENTIALS"):
        env_path = os.environ.get(env_key, "").strip()
        if env_path and os.path.isfile(env_path):
            return Credentials.from_service_account_file(env_path, scopes=SCOPES)

    # 프로젝트 루트 / 현재 작업 디렉터리 내 기본 JSON 키 파일 탐색
    search_dirs = [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]
    for base_dir in search_dirs:
        for fname in DEFAULT_LOCAL_KEY_FILES:
            candidate = os.path.join(base_dir, fname)
            if os.path.isfile(candidate):
                return Credentials.from_service_account_file(candidate, scopes=SCOPES)

    return None


def _open_worksheet(gc, sheet_key, worksheet_name=None):
    """
    Spreadsheet Key 또는 URL을 통해 대상 워크시트 열기
    """
    if sheet_key.startswith("http://") or sheet_key.startswith("https://"):
        spreadsheet = gc.open_by_url(sheet_key)
    else:
        spreadsheet = gc.open_by_key(sheet_key)

    if worksheet_name:
        try:
            return spreadsheet.worksheet(worksheet_name)
        except gspread.WorksheetNotFound:
            return spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=len(SHEET_HEADERS))
    return spreadsheet.sheet1


def _ensure_headers_and_get_existing_keys(worksheet):
    """
    첫 번째 행에 헤더가 없으면 자동 생성하고,
    기존 시트에 등록된 (학원명, 정규화키) Set을 반환함
    """
    existing_values = worksheet.get_all_values()

    if not existing_values:
        worksheet.append_row(SHEET_HEADERS, value_input_option="USER_ENTERED")
        header_row = SHEET_HEADERS
        data_rows = []
    else:
        first_row = existing_values[0]
        has_header = ("학원명" in first_row and "정규화키" in first_row)
        if not has_header:
            if not any(cell.strip() for cell in first_row):
                worksheet.update("A1:H1", [SHEET_HEADERS])
                data_rows = existing_values[1:]
            else:
                worksheet.insert_row(SHEET_HEADERS, index=1, value_input_option="USER_ENTERED")
                data_rows = existing_values
            header_row = SHEET_HEADERS
        else:
            header_row = first_row
            data_rows = existing_values[1:]

    acad_idx = header_row.index("학원명") if "학원명" in header_row else 1
    norm_idx = header_row.index("정규화키") if "정규화키" in header_row else 7

    existing_keys = set()
    for row in data_rows:
        acad_val = row[acad_idx].strip() if len(row) > acad_idx else ""
        norm_val = row[norm_idx].strip() if len(row) > norm_idx else ""
        if acad_val and norm_val:
            existing_keys.add((acad_val, norm_val))

    return existing_keys


def _build_rows_to_append(results, existing_keys, timestamp_str, processed_results=None, prev_history=None):
    """
    수집 결과를 시트 행 형태로 변환하고 (학원명 + 정규화키) 중복을 제거함
    """
    rows = []

    if processed_results:
        source_map = processed_results
    else:
        source_map = {}
        prev_history = prev_history or {}
        for academy, items in results.items():
            is_error = items and any(
                ("수집 실패" in it if isinstance(it, str) else it.get("error", False))
                for it in items
            )
            if is_error:
                source_map[academy] = {"error": True, "items": []}
                continue

            prev_keys = set()
            for p_it in prev_history.get(academy, []):
                if isinstance(p_it, dict):
                    pk = p_it.get("norm_key") or _fallback_normalize_title(p_it.get("title", ""))
                    if pk:
                        prev_keys.add(pk)
                elif isinstance(p_it, str) and "수집 실패" not in p_it:
                    pk = _fallback_normalize_title(p_it)
                    if pk:
                        prev_keys.add(pk)

            acad_list = []
            for it in items:
                if isinstance(it, dict):
                    obj = dict(it)
                    if not obj.get("norm_key"):
                        obj["norm_key"] = _fallback_normalize_title(obj.get("title", ""))
                    obj["status"] = "new" if obj["norm_key"] not in prev_keys else "unchanged"
                    acad_list.append(obj)
                elif isinstance(it, str):
                    nk = _fallback_normalize_title(it)
                    acad_list.append({
                        "title": it[:100],
                        "norm_key": nk,
                        "source_badge": "[공지사항]",
                        "section": "event" if ("설명회" in it or "간담회" in it) else "notice",
                        "url": "",
                        "status": "new" if nk not in prev_keys else "unchanged",
                    })
            source_map[academy] = {"error": False, "items": acad_list}

    for academy, data in source_map.items():
        if data.get("error"):
            continue

        for item in data.get("items", []):
            title = (item.get("title") or "").strip()
            norm_key = (item.get("norm_key") or _fallback_normalize_title(title)).strip()
            if not academy or not norm_key:
                continue

            dedup_key = (academy.strip(), norm_key)
            if dedup_key in existing_keys:
                continue

            section_label = "설명회/간담회" if item.get("section") == "event" else "시간표/공지"
            source_badge = item.get("source_badge") or "[공지사항]"
            url = (item.get("url") or "").strip()

            status_raw = item.get("status")
            if status_raw == "new":
                status_label = "NEW"
            elif status_raw == "changed":
                status_label = "변경"
            else:
                status_label = "유지"

            rows.append([
                timestamp_str,
                academy.strip(),
                section_label,
                source_badge,
                status_label,
                title,
                url,
                norm_key,
            ])
            existing_keys.add(dedup_key)

    return rows


# ==========================================
# 3. Google Sheets 누적 저장 진입점
# ==========================================
def log_results_to_sheets(
    results,
    prev_history=None,
    collection_time_kst=None,
    processed_results=None,
    sheet_key=None,
    worksheet_name=None,
):
    """
    정제된 크롤링 결과를 Google Sheets에 누적 저장.
    인증 누락이나 API 일시적 오류가 발생하더라도 예외를 외부로 전파하지 않고 안전하게 격리함.
    """
    try:
        target_sheet_key = (sheet_key or os.environ.get("GOOGLE_SHEET_KEY", "")).strip()
        if not target_sheet_key:
            print("ℹ️ GOOGLE_SHEET_KEY 환경 변수가 설정되지 않아 Google Sheets 저장을 건너뜁니다.")
            return False

        credentials = _load_credentials()
        if credentials is None:
            print("ℹ️ Google 서비스 계정 인증 정보(GCP_CREDENTIALS 또는 로컬 JSON)가 없어 Google Sheets 저장을 건너뜁니다.")
            return False

        now_kst = collection_time_kst or datetime.now(KST)
        timestamp_str = now_kst.strftime("%Y-%m-%d %H:%M:%S")
        target_ws_name = worksheet_name or os.environ.get("GOOGLE_SHEET_WORKSHEET", "").strip() or None

        gc = gspread.authorize(credentials)
        worksheet = _open_worksheet(gc, target_sheet_key, worksheet_name=target_ws_name)

        existing_keys = _ensure_headers_and_get_existing_keys(worksheet)
        rows_to_append = _build_rows_to_append(
            results=results,
            existing_keys=existing_keys,
            timestamp_str=timestamp_str,
            processed_results=processed_results,
            prev_history=prev_history,
        )

        if rows_to_append:
            worksheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")
            print(f"📊 Google Sheets에 신규/미등록 항목 {len(rows_to_append)}건 누적 저장 완료!")
        else:
            print("📊 Google Sheets에 새로 추가할 미등록 항목이 없습니다. (기존 중복 건너뜀)")

        return True

    except Exception as e:
        print(f"⚠️ Google Sheets 기록 중 오류 발생 (무시하고 계속 진행): {e}")
        return False
