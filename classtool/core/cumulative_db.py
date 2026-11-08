"""
매월 20일 기준(As-of 20th of each month) 강사별 수강생·결제자수 스냅샷 누적 관리 및 비교 통계 모듈 (core/cumulative_db.py)
- 매월 RMS 수업목록 엑셀 업로드 시 파일명(YYMMDD) 또는 지정 기준일 기반 '매월 20일 기준'(YYYY-MM-20) 스냅샷 생성
- 저장 스키마: data/enrollment_history.json (키: "2027-01-20", "2027-02-20", 표시 라벨: "27년 1월 20일 기준")
- '전월 20일 vs 당월 20일'(예: 1월 20일 ↔ 2월 20일) 강사별 1:1 정밀 비교 대조 (Full Outer Join)
- 강사별 강좌수·결제자수·총수강생·결제율 증감(명/%/%p) 및 변동 상태(🔺증가, 🔻감소, ➖유지, 🆕신규개설, ⚪미개설/종강) 자동 산출
- 매월 20일 기준 연간 시계열 누적 피벗 테이블 (결제완료, 총수강생, 강좌수, 매출액)
- 엑셀 결과지(매월20일_강사별비교, 매월20일_연간누적추이, 당월20일_강사별상세) 자동 생성 및 백업/복원 지원
"""

from __future__ import annotations

import io
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd


def get_db_path() -> Path:
    """누적 DB 파일 경로를 반환하며 부모 디렉토리가 없으면 생성합니다."""
    base_dir = Path(__file__).resolve().parent.parent
    data_dir = base_dir / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir / "enrollment_history.json"


def normalize_snapshot_key(date_str: str) -> str:
    """
    임의의 날짜/연월 문자열을 'YYYY-MM-20' 형태의 표준 스냅샷 키로 정규화합니다.
    예: '2027-01-20' -> '2027-01-20'
        '2027-01' -> '2027-01-20'
        '27.01.20' -> '2027-01-20'
    """
    s = str(date_str).strip()
    # 1) YYYY-MM-DD
    m_ymd = re.match(r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$", s)
    if m_ymd:
        return f"{int(m_ymd.group(1)):04d}-{int(m_ymd.group(2)):02d}-20"

    # 2) YY-MM-DD or YY.MM.DD
    m_yymmdd_sep = re.match(r"^(\d{2})[-./](\d{1,2})[-./](\d{1,2})$", s)
    if m_yymmdd_sep:
        yy = int(m_yymmdd_sep.group(1))
        yyyy = 2000 + yy if yy < 70 else 1900 + yy
        return f"{yyyy:04d}-{int(m_yymmdd_sep.group(2)):02d}-20"

    # 3) YYYY-MM
    m_ym = re.match(r"^(\d{4})[-./](\d{1,2})$", s)
    if m_ym:
        return f"{int(m_ym.group(1)):04d}-{int(m_ym.group(2)):02d}-20"

    # 4) YYMMDD
    m_yymmdd = re.match(r"^(\d{2})(\d{2})(\d{2})$", s)
    if m_yymmdd:
        yy = int(m_yymmdd.group(1))
        yyyy = 2000 + yy if yy < 70 else 1900 + yy
        return f"{yyyy:04d}-{int(m_yymmdd.group(2)):02d}-20"

    return datetime.now().strftime("%Y-%m-20")


def format_snapshot_label(snapshot_key: str) -> str:
    """
    '2027-01-20' 형태의 키를 사람이 읽기 좋은 '27년 1월 20일 기준' 라벨로 변환합니다.
    """
    s = normalize_snapshot_key(snapshot_key)
    parts = s.split("-")
    if len(parts) >= 2:
        yy = parts[0][2:]
        mm = int(parts[1])
        return f"{yy}년 {mm}월 20일 기준"
    return f"{snapshot_key} 기준"


def format_short_snapshot_header(snapshot_key: str) -> str:
    """
    '2027-01-20' 형태의 키를 시계열 테이블 헤더용 '27.01.20' 형태로 변환합니다.
    """
    s = normalize_snapshot_key(snapshot_key)
    parts = s.split("-")
    if len(parts) >= 3:
        return f"{parts[0][2:]}.{parts[1]}.20"
    return snapshot_key


def detect_snapshot_date_from_filename(filename: str, fallback_ym: Optional[str] = None) -> str:
    """
    RMS 엑셀 파일명(예: 'Chr_SSDan_List_2611021410.xls', 'Chr_SSDan_List_2701201030.xls')에서
    YYMMDD 패턴을 추출하여 'YYYY-MM-20' 기준 스냅샷 날짜로 자동 매핑합니다.
    """
    fname = str(filename).strip()
    m = re.search(r"(\d{2})(\d{2})(\d{2})\d{4}", fname)
    if m:
        yy = int(m.group(1))
        mm = int(m.group(2))
        yyyy = 2000 + yy if yy < 70 else 1900 + yy
        return f"{yyyy:04d}-{mm:02d}-20"

    if fallback_ym:
        return normalize_snapshot_key(fallback_ym)

    return datetime.now().strftime("%Y-%m-20")


def load_enrollment_db() -> Dict[str, Any]:
    """누적 DB JSON을 안전하게 로드합니다."""
    db_path = get_db_path()
    if not db_path.exists():
        return {"version": "2.0", "updated_at": "", "months": {}}

    try:
        content = db_path.read_text(encoding="utf-8")
        if not content.strip():
            return {"version": "2.0", "updated_at": "", "months": {}}
        data = json.loads(content)
        if not isinstance(data, dict):
            return {"version": "2.0", "updated_at": "", "months": {}}
        if "months" not in data or not isinstance(data["months"], dict):
            data["months"] = {}

        # 기존 YYYY-MM 키가 있는 경우 YYYY-MM-20으로 키 자동 마이그레이션
        migrated_months = {}
        for k, v in data["months"].items():
            norm_k = normalize_snapshot_key(k)
            if isinstance(v, dict):
                v["snapshot_key"] = norm_k
                v["snapshot_label"] = format_snapshot_label(norm_k)
            migrated_months[norm_k] = v
        data["months"] = migrated_months
        return data
    except Exception:
        return {"version": "2.0", "updated_at": "", "months": {}}


def save_enrollment_db(data: Dict[str, Any]) -> None:
    """누적 DB JSON을 안전하게 저장합니다."""
    db_path = get_db_path()
    data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tmp_path = db_path.with_suffix(".tmp")
    tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(db_path)


def extract_snapshot_stats_from_rms(
    rms_df: pd.DataFrame,
    snapshot_key: str,
    filter_stopped: bool = True,
) -> Dict[str, Any]:
    """
    RMS DataFrame으로부터 '매월 20일 기준'(YYYY-MM-20)의 요약 통계 및 강사별 집계 딕셔너리를 추출합니다.
    """
    norm_key = normalize_snapshot_key(snapshot_key)
    label = format_snapshot_label(norm_key)

    if rms_df is None or rms_df.empty:
        return {
            "snapshot_key": norm_key,
            "snapshot_label": label,
            "year_month": norm_key[:7],
            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "filter_stopped": filter_stopped,
            "summary": {},
            "instructors": [],
        }

    df = rms_df.copy()

    # 판매중지 필터 적용
    total_raw_courses = len(df)
    stopped_count = 0
    if "판매" in df.columns:
        is_stopped = df["판매"].astype(str).str.strip() == "판매중지"
        stopped_count = int(is_stopped.sum())
        if filter_stopped:
            df = df[~is_stopped].copy()

    # 수치형 컬럼 정수화
    for col in ("가등록포함수강생수", "완납&분납수강생수", "결제인원", "수강료"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        else:
            df[col] = 0

    tot_courses = len(df)
    tot_students = int(df["가등록포함수강생수"].sum())
    paid_students = int(df["결제인원"].sum())
    inst_students = int(df["완납&분납수강생수"].sum())
    unpaid_students = max(0, tot_students - paid_students)
    total_revenue = int((df["수강료"] * df["결제인원"]).sum())
    pay_rate = round((paid_students / tot_students * 100.0), 1) if tot_students > 0 else 0.0

    summary = {
        "snapshot_key": norm_key,
        "snapshot_label": label,
        "year_month": norm_key[:7],
        "total_courses": tot_courses,
        "stopped_courses_count": stopped_count,
        "raw_total_courses": total_raw_courses,
        "total_students": tot_students,
        "paid_students": paid_students,
        "installment_students": inst_students,
        "unpaid_students": unpaid_students,
        "total_revenue": total_revenue,
        "payment_rate": pay_rate,
        "filter_stopped_applied": filter_stopped,
    }

    # 강사별 그룹화
    inst_list: List[Dict[str, Any]] = []
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

        c_cnt = len(grp)
        t_stu = int(grp["가등록포함수강생수"].sum())
        p_stu = int(grp["결제인원"].sum())
        i_stu = int(grp["완납&분납수강생수"].sum())
        u_stu = max(0, t_stu - p_stu)
        p_rate = round((p_stu / t_stu * 100.0), 1) if t_stu > 0 else 0.0
        rev = int((grp["수강료"] * grp["결제인원"]).sum())
        avg_s = round((t_stu / c_cnt), 1) if c_cnt > 0 else 0.0

        c_details = []
        for _, c_row in grp.iterrows():
            c_details.append(
                {
                    "강좌코드": str(c_row.get("강좌코드", "")).strip(),
                    "강좌명": str(c_row.get("강좌명", "")).strip(),
                    "수업시간": str(c_row.get("간결수업시간", "") or c_row.get("수업시간", "")).strip(),
                    "수강료": int(c_row.get("수강료", 0) or 0),
                    "회차": int(c_row.get("회차", 0) or 0),
                    "가등록수강생": int(c_row.get("가등록포함수강생수", 0) or 0),
                    "결제인원": int(c_row.get("결제인원", 0) or 0),
                    "판매상태": str(c_row.get("판매", "판매")).strip(),
                }
            )

        inst_list.append(
            {
                "강사명": inst_name,
                "과목": subj_name,
                "영역": area_name,
                "과목(영역)": disp_subj,
                "개설강좌수": c_cnt,
                "총수강생": t_stu,
                "결제완료": p_stu,
                "완납분납": i_stu,
                "가등록미결제": u_stu,
                "결제율(%)": p_rate,
                "매출액": rev,
                "강좌당평균수강생": avg_s,
                "강좌목록": c_details,
            }
        )

    # 총수강생 -> 결제완료 순 내림차순 정렬
    inst_list.sort(key=lambda x: (x["총수강생"], x["결제완료"], x["개설강좌수"]), reverse=True)

    return {
        "snapshot_key": norm_key,
        "snapshot_label": label,
        "year_month": norm_key[:7],
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "filter_stopped": filter_stopped,
        "summary": summary,
        "instructors": inst_list,
    }


def upsert_snapshot_data(
    rms_df: pd.DataFrame,
    snapshot_key: str,
    filter_stopped: bool = True,
) -> Tuple[bool, str]:
    """특정 20일 기준 스냅샷 데이터를 누적 DB에 Upsert합니다."""
    norm_key = normalize_snapshot_key(snapshot_key)
    data = extract_snapshot_stats_from_rms(rms_df, norm_key, filter_stopped=filter_stopped)
    db = load_enrollment_db()
    db["months"][norm_key] = data
    save_enrollment_db(db)
    label = format_snapshot_label(norm_key)
    return True, f"'{label}' 스냅샷 데이터가 누적 DB에 성공적으로 저장되었습니다."


# 기존 호환용 alias
def upsert_month_data(rms_df: pd.DataFrame, year_month: str, filter_stopped: bool = True) -> Tuple[bool, str]:
    return upsert_snapshot_data(rms_df, year_month, filter_stopped=filter_stopped)


def delete_snapshot_data(snapshot_key: str) -> bool:
    """누적 DB에서 특정 스냅샷 데이터를 삭제합니다."""
    norm_key = normalize_snapshot_key(snapshot_key)
    db = load_enrollment_db()
    deleted = False
    for k in (norm_key, norm_key[:7]):
        if k in db.get("months", {}):
            del db["months"][k]
            deleted = True
    if deleted:
        save_enrollment_db(db)
    return deleted


# 기존 호환용 alias
def delete_month_data(year_month: str) -> bool:
    return delete_snapshot_data(year_month)


def get_available_snapshots() -> List[str]:
    """저장된 스냅샷 키 목록을 날짜 오름차순(예: '2026-11-20', '2027-01-20')으로 반환합니다."""
    db = load_enrollment_db()
    keys = list(db.get("months", {}).keys())
    return sorted(keys)


# 기존 호환용 alias
def get_available_months() -> List[str]:
    return get_available_snapshots()


# -------------------------------------------------------------------------
# 2. '전월 20일 vs 당월 20일' 1:1 비교 핵심 함수
# -------------------------------------------------------------------------
def compare_two_snapshots(
    snapshot_a_key: str,
    snapshot_b_key: str,
) -> Dict[str, Any]:
    """
    저장된 두 시점(스냅샷 A: 전월 20일 vs 스냅샷 B: 당월 20일)의 데이터를 1:1로 정밀 대조하여
    전체 KPI 변화와 강사별 1:1 비교 데이터프레임을 생성합니다.
    """
    norm_a = normalize_snapshot_key(snapshot_a_key)
    norm_b = normalize_snapshot_key(snapshot_b_key)
    label_a = format_snapshot_label(norm_a)
    label_b = format_snapshot_label(norm_b)

    # 간결 표시 라벨 (예: '1월 20일', '2월 20일')
    mm_a = int(norm_a.split("-")[1]) if len(norm_a.split("-")) >= 2 else norm_a
    mm_b = int(norm_b.split("-")[1]) if len(norm_b.split("-")) >= 2 else norm_b
    short_a = f"{mm_a}월 20일"
    short_b = f"{mm_b}월 20일"

    db = load_enrollment_db()
    data_a = db["months"].get(norm_a, {})
    data_b = db["months"].get(norm_b, {})

    sum_a = data_a.get("summary", {})
    sum_b = data_b.get("summary", {})

    # 1. 전체 KPI 증감
    crs_a = sum_a.get("total_courses", 0)
    crs_b = sum_b.get("total_courses", 0)
    crs_diff = crs_b - crs_a

    stu_a = sum_a.get("total_students", 0)
    stu_b = sum_b.get("total_students", 0)
    stu_diff = stu_b - stu_a
    stu_pct = round((stu_diff / stu_a * 100.0), 1) if stu_a > 0 else (100.0 if stu_b > 0 else 0.0)

    paid_a = sum_a.get("paid_students", 0)
    paid_b = sum_b.get("paid_students", 0)
    paid_diff = paid_b - paid_a
    paid_pct = round((paid_diff / paid_a * 100.0), 1) if paid_a > 0 else (100.0 if paid_b > 0 else 0.0)

    rate_a = sum_a.get("payment_rate", 0.0)
    rate_b = sum_b.get("payment_rate", 0.0)
    rate_diff = round(rate_b - rate_a, 1)

    kpi_comparison = {
        "label_a": label_a,
        "label_b": label_b,
        "short_a": short_a,
        "short_b": short_b,
        "courses_a": crs_a,
        "courses_b": crs_b,
        "courses_diff": crs_diff,
        "students_a": stu_a,
        "students_b": stu_b,
        "students_diff": stu_diff,
        "students_pct": stu_pct,
        "paid_a": paid_a,
        "paid_b": paid_b,
        "paid_diff": paid_diff,
        "paid_pct": paid_pct,
        "rate_a": rate_a,
        "rate_b": rate_b,
        "rate_diff": rate_diff,
    }

    # 2. 강사별 Full Outer Join
    inst_a_map = {(i["과목"], i["강사명"]): i for i in data_a.get("instructors", [])}
    inst_b_map = {(i["과목"], i["강사명"]): i for i in data_b.get("instructors", [])}
    all_keys = sorted(list(set(inst_a_map.keys()) | set(inst_b_map.keys())), key=lambda x: (x[0], x[1]))

    comparison_rows: List[Dict[str, Any]] = []

    for subj, inst_name in all_keys:
        a_info = inst_a_map.get((subj, inst_name))
        b_info = inst_b_map.get((subj, inst_name))

        # 과목(영역) 표시
        disp_subj = (b_info.get("과목(영역)") if b_info else None) or (a_info.get("과목(영역)") if a_info else subj)

        c_a = int(a_info.get("개설강좌수", 0)) if a_info else 0
        c_b = int(b_info.get("개설강좌수", 0)) if b_info else 0
        c_diff = c_b - c_a
        c_diff_str = f"+{c_diff}" if c_diff > 0 else str(c_diff)

        p_a = int(a_info.get("결제완료", 0)) if a_info else 0
        p_b = int(b_info.get("결제완료", 0)) if b_info else 0
        p_diff = p_b - p_a
        if p_a > 0:
            p_pct = round((p_diff / p_a * 100.0), 1)
            p_pct_str = f"+{p_pct}%" if p_pct > 0 else f"{p_pct}%"
        elif p_b > 0:
            p_pct_str = "신규"
        else:
            p_pct_str = "-"

        t_a = int(a_info.get("총수강생", 0)) if a_info else 0
        t_b = int(b_info.get("총수강생", 0)) if b_info else 0
        t_diff = t_b - t_a
        t_diff_str = f"+{t_diff}명" if t_diff > 0 else (f"{t_diff}명" if t_diff < 0 else "0명")

        r_a = float(a_info.get("결제율(%)", 0.0)) if a_info else 0.0
        r_b = float(b_info.get("결제율(%)", 0.0)) if b_info else 0.0
        r_diff = round(r_b - r_a, 1)
        r_diff_str = f"+{r_diff}%p" if r_diff > 0 else f"{r_diff}%p"

        # 변동 상태 배지
        if a_info is None and b_info is not None:
            status_badge = "🆕 당월 신규개설"
            status_category = "신규"
        elif a_info is not None and b_info is None:
            status_badge = "⚪ 당월 미개설(종강)"
            status_category = "종강"
        elif p_diff > 0:
            status_badge = "🔺 증가"
            status_category = "증가"
        elif p_diff < 0:
            status_badge = "🔻 감소"
            status_category = "감소"
        else:
            status_badge = "➖ 유지"
            status_category = "유지"

        p_diff_disp = f"+{p_diff}명" if p_diff > 0 else (f"{p_diff}명" if p_diff < 0 else "0명")

        comparison_rows.append(
            {
                "과목(영역)": disp_subj,
                "과목": subj,
                "강사명": inst_name,
                f"[A] {short_a} 강좌수": c_a,
                f"[B] {short_b} 강좌수": c_b,
                "강좌수 증감": c_diff_str,
                f"[A] {short_a} 결제자수": p_a,
                f"[B] {short_b} 결제자수": p_b,
                "결제자 증감(명)": p_diff_disp,
                "결제자_증감_정수": p_diff,
                "결제자 증감률(%)": p_pct_str,
                f"[A] {short_a} 총수강생": t_a,
                f"[B] {short_b} 총수강생": t_b,
                "총수강생 증감(명)": t_diff_str,
                "총수강생_증감_정수": t_diff,
                f"[A] 결제율(%)": f"{r_a}%" if a_info else "-",
                f"[B] 결제율(%)": f"{r_b}%" if b_info else "-",
                "결제율 증감(%p)": r_diff_str if (a_info and b_info) else "-",
                "변동 상태": status_badge,
                "변동분류": status_category,
            }
        )

    # 기본 정렬: [B] 당월 20일 결제자수 내림차순 -> 증감 내림차순
    comparison_rows.sort(key=lambda x: (x[f"[B] {short_b} 결제자수"], x["결제자_증감_정수"]), reverse=True)
    df_compare = pd.DataFrame(comparison_rows)
    if not df_compare.empty:
        df_compare.insert(0, "순위", range(1, len(df_compare) + 1))

    return {
        "kpi": kpi_comparison,
        "comparison_df": df_compare,
        "label_a": label_a,
        "label_b": label_b,
        "short_a": short_a,
        "short_b": short_b,
    }


# -------------------------------------------------------------------------
# 3. 매월 20일 기준 연간 시계열 누적 피벗 테이블 생성
# -------------------------------------------------------------------------
def build_snapshot_pivot_table(metric_label: str = "결제완료 인원") -> pd.DataFrame:
    """
    모든 스냅샷(매월 20일 기준)의 강사별 통계를 피벗하여 연간 시계열 데이터프레임을 생성합니다.
    - 열 헤더: '26.11.20', '26.12.20', '27.01.20', '27.02.20' ...
    - 요약열: '기간 누적 합계', '매월 20일 평균 인원', '최근월(20일) 증감'
    """
    metric_key_map = {
        "결제완료 인원": ("결제완료", "명"),
        "총 수강생(가등록포함)": ("총수강생", "명"),
        "유효 강좌수": ("개설강좌수", "개"),
        "결제 매출액": ("매출액", "원"),
    }
    field_name, unit = metric_key_map.get(metric_label, ("결제완료", "명"))

    db = load_enrollment_db()
    snapshots = sorted(list(db.get("months", {}).keys()))
    if not snapshots:
        return pd.DataFrame()

    inst_keys: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for snap in snapshots:
        m_data = db["months"][snap]
        for inst in m_data.get("instructors", []):
            k = (inst["과목(영역)"], inst["강사명"])
            if k not in inst_keys:
                inst_keys[k] = {
                    "과목(영역)": inst["과목(영역)"],
                    "강사명": inst["강사명"],
                    "과목": inst["과목"],
                }

    rows: List[Dict[str, Any]] = []
    for (disp_subj, inst_name), info in inst_keys.items():
        row_dict: Dict[str, Any] = {
            "과목(영역)": disp_subj,
            "강사명": inst_name,
        }

        val_list: List[int] = []
        for snap in snapshots:
            col_hdr = format_short_snapshot_header(snap)
            m_data = db["months"][snap]
            found_val = 0
            for inst in m_data.get("instructors", []):
                if inst["강사명"] == inst_name and inst["과목(영역)"] == disp_subj:
                    found_val = int(inst.get(field_name, 0))
                    break
            row_dict[col_hdr] = found_val
            val_list.append(found_val)

        total_sum = sum(val_list)
        avg_val = round(total_sum / len(snapshots), 1)

        row_dict["기간 누적 합계"] = total_sum
        row_dict["매월 20일 평균"] = avg_val

        # 최근월 증감
        if len(snapshots) >= 2:
            last_v = val_list[-1]
            prev_v = val_list[-2]
            diff = last_v - prev_v
            if prev_v > 0:
                pct = round((diff / prev_v) * 100.0, 1)
                sign = "+" if diff > 0 else ""
                row_dict["최근월(20일) 증감"] = f"{sign}{diff:,}{unit} ({sign}{pct}%)"
            elif last_v > 0:
                row_dict["최근월(20일) 증감"] = f"신규 (+{last_v:,}{unit})"
            else:
                row_dict["최근월(20일) 증감"] = "-"
        else:
            row_dict["최근월(20일) 증감"] = "-"

        rows.append(row_dict)

    df_pivot = pd.DataFrame(rows)
    if df_pivot.empty:
        return df_pivot

    # 정렬: 최근월 수치 내림차순 -> 누적 합계 내림차순
    last_col_hdr = format_short_snapshot_header(snapshots[-1])
    sort_cols = [last_col_hdr, "기간 누적 합계"]
    df_pivot = df_pivot.sort_values(by=sort_cols, ascending=[False, False]).reset_index(drop=True)
    df_pivot.insert(0, "순위", range(1, len(df_pivot) + 1))

    return df_pivot


# 기존 호환용 alias
def build_pivot_table(metric_label: str = "결제완료 인원") -> pd.DataFrame:
    return build_snapshot_pivot_table(metric_label)


def get_monthly_summary_table() -> pd.DataFrame:
    """월별/스냅샷 전체 KPI 요약 테이블을 반환합니다."""
    db = load_enrollment_db()
    snapshots = sorted(list(db.get("months", {}).keys()))
    if not snapshots:
        return pd.DataFrame()

    rows = []
    for snap in snapshots:
        m_data = db["months"][snap]
        s = m_data.get("summary", {})
        label = m_data.get("snapshot_label") or format_snapshot_label(snap)
        rows.append(
            {
                "기준일": label,
                "스냅샷키": snap,
                "유효 강좌수": s.get("total_courses", 0),
                "판매중지 제외": s.get("stopped_courses_count", 0),
                "총 강좌수": s.get("raw_total_courses", s.get("total_courses", 0)),
                "총 수강생(가등록)": s.get("total_students", 0),
                "결제완료 인원": s.get("paid_students", 0),
                "미결제 가등록": s.get("unpaid_students", 0),
                "결제율(%)": s.get("payment_rate", 0.0),
                "총 매출액(원)": s.get("total_revenue", 0),
                "저장일시": m_data.get("saved_at", "-"),
            }
        )

    return pd.DataFrame(rows)


# -------------------------------------------------------------------------
# 4. 엑셀 백업 및 매월 20일 비교 통계 내보내기
# -------------------------------------------------------------------------
def export_enrollment_backup_excel() -> bytes:
    """
    누적 DB의 모든 데이터와 메타데이터를 서식 포함 Excel(.xlsx) 바이트로 내보냅니다.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws_p1 = wb.active
    ws_p1.title = "매월20일_피벗_결제완료"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    base_font = Font(name="맑은 고딕", size=9, color="111111")
    thin_side = Side(border_style="thin", color="D5D8DC")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def _append_df_to_ws(ws: Any, df: pd.DataFrame) -> None:
        if df is None or df.empty:
            ws.append(["데이터가 없습니다."])
            return
        cols = list(df.columns)
        ws.append(cols)
        ws.row_dimensions[1].height = 26
        for c_idx in range(1, len(cols) + 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_idx, (_, r_data) in enumerate(df.iterrows(), start=2):
            row_vals = [r_data[c] for c in cols]
            ws.append(row_vals)
            ws.row_dimensions[r_idx].height = 20
            for c_idx in range(1, len(cols) + 1):
                c_cell = ws.cell(row=r_idx, column=c_idx)
                c_cell.border = cell_border
                c_cell.font = base_font
                c_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(df) + 1}"
        for col_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = max(len(str(col_name)) * 2 + 5, 12)

    # 1. 결제완료 피벗
    df_paid = build_snapshot_pivot_table("결제완료 인원")
    _append_df_to_ws(ws_p1, df_paid)

    # 2. 총수강생 피벗
    ws_p2 = wb.create_sheet(title="매월20일_피벗_총수강생")
    df_tot = build_snapshot_pivot_table("총 수강생(가등록포함)")
    _append_df_to_ws(ws_p2, df_tot)

    # 3. 유효 강좌수 피벗
    ws_p3 = wb.create_sheet(title="매월20일_피벗_강좌수")
    df_crs = build_snapshot_pivot_table("유효 강좌수")
    _append_df_to_ws(ws_p3, df_crs)

    # 4. KPI 전체 요약
    ws_kpi = wb.create_sheet(title="스냅샷_KPI_전체요약")
    df_kpi = get_monthly_summary_table()
    _append_df_to_ws(ws_kpi, df_kpi)

    # 5. 전체 강사 상세 내역
    ws_detail = wb.create_sheet(title="스냅샷_강사상세_전체내역")
    db = load_enrollment_db()
    detail_rows = []
    for snap, m_data in sorted(db.get("months", {}).items()):
        label = m_data.get("snapshot_label") or format_snapshot_label(snap)
        for inst in m_data.get("instructors", []):
            detail_rows.append(
                {
                    "기준일": label,
                    "스냅샷키": snap,
                    "과목(영역)": inst.get("과목(영역)", ""),
                    "강사명": inst.get("강사명", ""),
                    "개설강좌수": inst.get("개설강좌수", 0),
                    "총수강생": inst.get("총수강생", 0),
                    "결제완료": inst.get("결제완료", 0),
                    "완납분납": inst.get("완납분납", 0),
                    "가등록미결제": inst.get("가등록미결제", 0),
                    "결제율(%)": inst.get("결제율(%)", 0.0),
                    "매출액": inst.get("매출액", 0),
                    "강좌당평균수강생": inst.get("강좌당평균수강생", 0.0),
                }
            )
    _append_df_to_ws(ws_detail, pd.DataFrame(detail_rows))

    # 6. RAW_JSON_BACKUP
    ws_json = wb.create_sheet(title="RAW_JSON_BACKUP")
    json_str = json.dumps(db, ensure_ascii=False)
    chunk_size = 30000
    for chunk_idx, i in enumerate(range(0, len(json_str), chunk_size), start=1):
        ws_json.cell(row=chunk_idx, column=1, value=json_str[i : i + chunk_size])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def import_enrollment_backup(file_source: Any) -> Tuple[bool, str, int]:
    """
    백업 파일(Excel .xlsx 또는 JSON .json)을 읽어 누적 DB를 복원합니다.
    """
    try:
        raw_bytes = None
        filename = getattr(file_source, "name", "")
        if isinstance(file_source, (str, Path)):
            raw_bytes = Path(file_source).read_bytes()
            filename = str(file_source)
        elif isinstance(file_source, bytes):
            raw_bytes = file_source
        elif hasattr(file_source, "getvalue"):
            raw_bytes = file_source.getvalue()
        elif hasattr(file_source, "read"):
            file_source.seek(0)
            raw_bytes = file_source.read()

        if not raw_bytes:
            return False, "파일 데이터가 비어있습니다.", 0

        # 1) JSON 파일인 경우
        if filename.lower().endswith(".json") or raw_bytes.strip().startswith(b"{"):
            data = json.loads(raw_bytes.decode("utf-8-sig", errors="replace"))
            if isinstance(data, dict) and "months" in data:
                save_enrollment_db(data)
                count = len(data["months"])
                return True, f"JSON 백업 파일로부터 {count}개의 스냅샷 데이터가 복원되었습니다.", count

        # 2) Excel 파일인 경우
        if raw_bytes[:4] == b"PK\x03\x04" or filename.lower().endswith((".xlsx", ".xlsm")):
            import openpyxl

            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
            if "RAW_JSON_BACKUP" in wb.sheetnames:
                ws = wb["RAW_JSON_BACKUP"]
                chunks = []
                for row in ws.iter_rows(values_only=True):
                    if row and row[0]:
                        chunks.append(str(row[0]))
                full_json = "".join(chunks)
                if full_json.strip():
                    data = json.loads(full_json)
                    if isinstance(data, dict) and "months" in data:
                        save_enrollment_db(data)
                        count = len(data["months"])
                        return True, f"엑셀 RAW 백업 시트로부터 {count}개의 스냅샷 데이터가 완벽 복원되었습니다.", count

            detail_sheet_name = next((s for s in wb.sheetnames if "강사상세" in s), None)
            if detail_sheet_name:
                df_detail = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=detail_sheet_name)
                date_col = next((c for c in ("스냅샷키", "기준일", "기준연월") if c in df_detail.columns), None)
                if not df_detail.empty and date_col:
                    db = load_enrollment_db()
                    restored_count = 0
                    for d_val, grp in df_detail.groupby(date_col):
                        norm_k = normalize_snapshot_key(str(d_val))
                        label = format_snapshot_label(norm_k)
                        inst_list = []
                        for _, r in grp.iterrows():
                            inst_list.append(
                                {
                                    "강사명": str(r.get("강사명", "")).strip(),
                                    "과목(영역)": str(r.get("과목(영역)", "")).strip(),
                                    "과목": str(r.get("과목", "")).strip(),
                                    "영역": str(r.get("영역", "")).strip(),
                                    "개설강좌수": int(r.get("개설강좌수", 0) or 0),
                                    "총수강생": int(r.get("총수강생", 0) or 0),
                                    "결제완료": int(r.get("결제완료", 0) or 0),
                                    "완납분납": int(r.get("완납분납", 0) or 0),
                                    "가등록미결제": int(r.get("가등록미결제", 0) or 0),
                                    "결제율(%)": float(r.get("결제율(%)", 0.0) or 0.0),
                                    "매출액": int(r.get("매출액", 0) or 0),
                                    "강좌당평균수강생": float(r.get("강좌당평균수강생", 0.0) or 0.0),
                                    "강좌목록": [],
                                }
                            )
                        t_stu = sum(i["총수강생"] for i in inst_list)
                        p_stu = sum(i["결제완료"] for i in inst_list)
                        db["months"][norm_k] = {
                            "snapshot_key": norm_k,
                            "snapshot_label": label,
                            "year_month": norm_k[:7],
                            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "filter_stopped": True,
                            "summary": {
                                "snapshot_key": norm_k,
                                "snapshot_label": label,
                                "year_month": norm_k[:7],
                                "total_courses": sum(i["개설강좌수"] for i in inst_list),
                                "total_students": t_stu,
                                "paid_students": p_stu,
                                "unpaid_students": max(0, t_stu - p_stu),
                                "payment_rate": round(p_stu / t_stu * 100.0, 1) if t_stu > 0 else 0.0,
                            },
                            "instructors": inst_list,
                        }
                        restored_count += 1
                    save_enrollment_db(db)
                    return True, f"엑셀 테이블로부터 {restored_count}개의 스냅샷 데이터가 복원되었습니다.", restored_count

        return False, "지원되지 않거나 손상된 백업 파일 형식입니다.", 0
    except Exception as e:
        return False, f"백업 파일 복원 실패: {e}", 0
