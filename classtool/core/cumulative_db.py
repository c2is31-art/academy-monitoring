"""
월별·강사별 수강생 및 결제자수 누적 통계 관리 DB 모듈 (core/cumulative_db.py)
- 매월 RMS 수업목록 엑셀 업로드 시 월별(YYYY-MM) 강사별 수강생/결제자수/매출액 누적 저장 (data/enrollment_history.json)
- '판매중지 강좌 제외' 필터 적용 유효 강좌 기준 자동 집계 (전체 기준 포함 지원)
- 다중 월별 파일 업로드 시 일괄 Upsert 지원
- 강사별·월별 누적 통계 피벗 테이블 생성 (결제완료 인원, 총 수강생, 유효 강좌수, 결제 매출액 전환)
- 전월 대비 증감(MoM, +00명 / +00.0%), 누적 합계, 월평균 자동 산출
- 클라우드 재시작 대비 누적 DB 백업 엑셀(.xlsx) 다운로드 및 복원(불러오기) 지원
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

DB_FILE_PATH = Path("data/enrollment_history.json")


def get_db_path() -> Path:
    """누적 DB 파일 경로를 반환하며 부모 디렉토리가 없으면 생성합니다."""
    # classtool 내부 data 디렉토리 우선
    base_dir = Path(__file__).resolve().parent.parent
    data_dir = base_dir / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir / "enrollment_history.json"


def load_enrollment_db() -> Dict[str, Any]:
    """누적 DB JSON을 안전하게 로드합니다."""
    db_path = get_db_path()
    if not db_path.exists():
        return {"version": "1.0", "updated_at": "", "months": {}}

    try:
        content = db_path.read_text(encoding="utf-8")
        if not content.strip():
            return {"version": "1.0", "updated_at": "", "months": {}}
        data = json.loads(content)
        if not isinstance(data, dict):
            return {"version": "1.0", "updated_at": "", "months": {}}
        if "months" not in data or not isinstance(data["months"], dict):
            data["months"] = {}
        return data
    except Exception:
        return {"version": "1.0", "updated_at": "", "months": {}}


def save_enrollment_db(data: Dict[str, Any]) -> None:
    """누적 DB JSON을 저장합니다."""
    db_path = get_db_path()
    data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tmp_path = db_path.with_suffix(".tmp")
    tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(db_path)


def extract_month_stats_from_rms(
    rms_df: pd.DataFrame,
    year_month: str,
    filter_stopped: bool = True,
) -> Dict[str, Any]:
    """
    RMS DataFrame으로부터 해당 연월(YYYY-MM)의 요약 통계 및 강사별 집계 딕셔너리를 추출합니다.
    """
    if rms_df is None or rms_df.empty:
        return {
            "year_month": year_month,
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
        "year_month": year_month,
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

        # 개별 강좌 간략 정보
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

    # 총 수강생 내림차순 정렬
    inst_list.sort(key=lambda x: (x["총수강생"], x["결제완료"], x["개설강좌수"]), reverse=True)

    return {
        "year_month": year_month,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "filter_stopped": filter_stopped,
        "summary": summary,
        "instructors": inst_list,
    }


def upsert_month_data(
    rms_df: pd.DataFrame,
    year_month: str,
    filter_stopped: bool = True,
) -> Tuple[bool, str]:
    """특정 연월의 데이터를 누적 DB에 Upsert(생성 또는 덮어쓰기)합니다."""
    ym = str(year_month).strip()
    if not re.match(r"^\d{4}-\d{2}$", ym):
        return False, f"기준 연월 포맷이 올바르지 않습니다: '{ym}' (YYYY-MM 형식이어야 함)"

    month_data = extract_month_stats_from_rms(rms_df, ym, filter_stopped=filter_stopped)
    db = load_enrollment_db()
    db["months"][ym] = month_data
    save_enrollment_db(db)
    return True, f"'{ym}' 기준 데이터가 누적 DB에 성공적으로 저장되었습니다."


def delete_month_data(year_month: str) -> bool:
    """누적 DB에서 특정 연월 데이터를 삭제합니다."""
    db = load_enrollment_db()
    if year_month in db.get("months", {}):
        del db["months"][year_month]
        save_enrollment_db(db)
        return True
    return False


def get_available_months() -> List[str]:
    """저장된 연월 목록을 시간 오름차순(예: '2026-11', '2026-12', '2027-01')으로 반환합니다."""
    db = load_enrollment_db()
    months = list(db.get("months", {}).keys())
    return sorted(months)


def build_pivot_table(metric_label: str = "결제완료 인원") -> pd.DataFrame:
    """
    모든 연월의 강사별 통계를 피벗하여 비교 데이터프레임을 생성합니다.
    - metric_label 옵션:
        '결제완료 인원' (기본값)
        '총 수강생(가등록포함)'
        '유효 강좌수'
        '결제 매출액'
    - 요약 컬럼: 누적 합계, 월평균, 전월 대비 증감(MoM, +00명 / +00.0%)
    """
    metric_key_map = {
        "결제완료 인원": ("결제완료", "명"),
        "총 수강생(가등록포함)": ("총수강생", "명"),
        "유효 강좌수": ("개설강좌수", "개"),
        "결제 매출액": ("매출액", "원"),
    }
    field_name, unit = metric_key_map.get(metric_label, ("결제완료", "명"))

    db = load_enrollment_db()
    months = sorted(list(db.get("months", {}).keys()))
    if not months:
        return pd.DataFrame()

    # 모든 강사 및 과목 목록 수집
    inst_keys: Dict[Tuple[str, str], Dict[str, Any]] = {}  # (과목(영역), 강사명) -> info

    for ym in months:
        m_data = db["months"][ym]
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
        for ym in months:
            m_data = db["months"][ym]
            found_val = 0
            for inst in m_data.get("instructors", []):
                if inst["강사명"] == inst_name and inst["과목(영역)"] == disp_subj:
                    found_val = int(inst.get(field_name, 0))
                    break
            row_dict[ym] = found_val
            val_list.append(found_val)

        # 누적 합계
        total_sum = sum(val_list)
        # 월평균 (데이터가 등록된 월 수 기준)
        active_months_count = len([v for v in val_list if v > 0]) or 1
        avg_val = round(total_sum / len(months), 1)

        row_dict["누적 합계"] = total_sum
        row_dict["월평균"] = avg_val

        # 전월 대비 증감 (MoM)
        if len(months) >= 2:
            last_val = val_list[-1]
            prev_val = val_list[-2]
            diff = last_val - prev_val
            if prev_val > 0:
                pct = round((diff / prev_val) * 100.0, 1)
                sign = "+" if diff > 0 else ""
                row_dict["전월 대비 증감(MoM)"] = f"{sign}{diff:,}{unit} ({sign}{pct}%)"
            elif last_val > 0:
                row_dict["전월 대비 증감(MoM)"] = f"신규 (+{last_val:,}{unit})"
            else:
                row_dict["전월 대비 증감(MoM)"] = "-"
        else:
            row_dict["전월 대비 증감(MoM)"] = "-"

        rows.append(row_dict)

    df_pivot = pd.DataFrame(rows)
    if df_pivot.empty:
        return df_pivot

    # 최신 월 수치 내림차순 -> 누적 합계 내림차순 정렬
    sort_cols = [months[-1], "누적 합계"]
    df_pivot = df_pivot.sort_values(by=sort_cols, ascending=[False, False]).reset_index(drop=True)
    df_pivot.insert(0, "순위", range(1, len(df_pivot) + 1))

    return df_pivot


def get_monthly_summary_table() -> pd.DataFrame:
    """월별 전체 KPI 요약 테이블을 반환합니다."""
    db = load_enrollment_db()
    months = sorted(list(db.get("months", {}).keys()))
    if not months:
        return pd.DataFrame()

    rows = []
    for ym in months:
        m_data = db["months"][ym]
        s = m_data.get("summary", {})
        rows.append(
            {
                "기준연월": ym,
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


def export_enrollment_backup_excel() -> bytes:
    """
    누적 DB의 모든 데이터와 메타데이터를 서식 포함 Excel(.xlsx) 바이트로 내보냅니다.
    - Sheet 1: 월별_피벗통계_결제완료
    - Sheet 2: 월별_피벗통계_총수강생
    - Sheet 3: 월별_피벗통계_유효강좌수
    - Sheet 4: 월별_KPI_전체요약
    - Sheet 5: 월별_강사상세_전체내역
    - Sheet 6: RAW_JSON_BACKUP (완벽한 복원을 위한 순수 JSON 저장)
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws_p1 = wb.active
    ws_p1.title = "월별_피벗_결제완료"

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
    df_paid = build_pivot_table("결제완료 인원")
    _append_df_to_ws(ws_p1, df_paid)

    # 2. 총수강생 피벗
    ws_p2 = wb.create_sheet(title="월별_피벗_총수강생")
    df_tot = build_pivot_table("총 수강생(가등록포함)")
    _append_df_to_ws(ws_p2, df_tot)

    # 3. 유효 강좌수 피벗
    ws_p3 = wb.create_sheet(title="월별_피벗_강좌수")
    df_crs = build_pivot_table("유효 강좌수")
    _append_df_to_ws(ws_p3, df_crs)

    # 4. 월별 KPI 전체 요약
    ws_kpi = wb.create_sheet(title="월별_KPI_전체요약")
    df_kpi = get_monthly_summary_table()
    _append_df_to_ws(ws_kpi, df_kpi)

    # 5. 전체 강사 상세 내역 (Flat 테이블)
    ws_detail = wb.create_sheet(title="월별_강사상세_전체내역")
    db = load_enrollment_db()
    detail_rows = []
    for ym, m_data in sorted(db.get("months", {}).items()):
        for inst in m_data.get("instructors", []):
            detail_rows.append(
                {
                    "기준연월": ym,
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
    # Excel 셀 글자 수 제한(32767자)을 고려하여 청크 단위로 나누어 기록
    chunk_size = 30000
    for chunk_idx, i in enumerate(range(0, len(json_str), chunk_size), start=1):
        ws_json.cell(row=chunk_idx, column=1, value=json_str[i : i + chunk_size])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def import_enrollment_backup(file_source: Any) -> Tuple[bool, str, int]:
    """
    백업 파일(Excel .xlsx 또는 JSON .json)을 읽어 누적 DB를 복원합니다.
    - 반환값: (성공여부, 안내메시지, 복원된월수)
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
                return True, f"JSON 백업 파일로부터 {count}개 월의 누적 데이터가 복원되었습니다.", count

        # 2) Excel 파일인 경우 (RAW_JSON_BACKUP 시트 또는 월별_강사상세_전체내역 시트 파싱)
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
                        return True, f"엑셀 RAW 백업 시트로부터 {count}개 월의 누적 데이터가 완벽 복원되었습니다.", count

            if "월별_강사상세_전체내역" in wb.sheetnames:
                df_detail = pd.read_excel(io.BytesIO(raw_bytes), sheet_name="월별_강사상세_전체내역")
                if not df_detail.empty and "기준연월" in df_detail.columns:
                    # Flat 테이블로부터 DB 재구성
                    db = load_enrollment_db()
                    restored_count = 0
                    for ym, grp in df_detail.groupby("기준연월"):
                        ym_str = str(ym).strip()
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
                        db["months"][ym_str] = {
                            "year_month": ym_str,
                            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "filter_stopped": True,
                            "summary": {
                                "year_month": ym_str,
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
                    return True, f"엑셀 테이블로부터 {restored_count}개 월의 데이터가 복원되었습니다.", restored_count

        return False, "지원되지 않거나 손상된 백업 파일 형식입니다.", 0
    except Exception as e:
        return False, f"백업 파일 복원 실패: {e}", 0
