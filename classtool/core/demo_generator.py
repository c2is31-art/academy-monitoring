"""
시뮬레이션/테스트용 데모 데이터 생성 모듈 (core/demo_generator.py)
- 실제 RMS 수업목록(Chr_SSDan_List_2611021410.xls 등)을 기반으로:
  1) 정상 1개월 단일 과정 일치 건
  2) 2개월 통합 과정 + 회차별계획(JSON) 당월 역산 일치 건
  3) 강사 필명/본명 매핑 건 (박소현(리아)↔리아, 박소현↔박소현(셀리나), 김민정↔민정, 이지원↔이진리, 임수민↔현자의 돌)
  4) 강좌명 변경 확인 건 (강사·요일·시간 동일, 강좌명 상이)
  5) 요일/시간/개강일 불일치 오류 건
  6) 중복 제출 자동 제외 건 (동일 강사·요일·시간 다중 제출)
  7) 테스트 데이터 (홍길동, 아자스, 야호야호, 심동준, test 강좌) 및 타 시기(10~11월) 강좌
  8) 로컬 드라이브 폴더 내 샘플 .gdoc, .docx, .pdf 파일 생성
을 지원합니다.
"""

from __future__ import annotations

import calendar
import json
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd

WEEKDAY_STR_TO_INT = {"월": 0, "화": 1, "수": 2, "목": 3, "금": 4, "토": 5, "일": 6}


def build_plan_json_for_dates(
    start_date_str: str,
    end_date_str: str,
    weekdays: List[str],
    holiday_dates: List[str] | None = None,
) -> str:
    """개강일~종강일 사이 지정 요일의 날짜 목록으로 회차별계획(JSON) 문자열을 생성합니다."""
    if not start_date_str or not end_date_str or not weekdays:
        return "[]"
    try:
        y1, m1, d1 = map(int, start_date_str.split("-"))
        y2, m2, d2 = map(int, end_date_str.split("-"))
        cur = date(y1, m1, d1)
        end = date(y2, m2, d2)
    except Exception:
        return "[]"

    target_wd = {WEEKDAY_STR_TO_INT[w] for w in weekdays if w in WEEKDAY_STR_TO_INT}
    holidays = set(holiday_dates or [])
    items = []
    seq = 1
    while cur <= end:
        if cur.weekday() in target_wd:
            ds = cur.strftime("%Y-%m-%d")
            is_hol = ds in holidays
            items.append(
                {
                    "session": seq,
                    "date": ds,
                    "holiday": is_hol,
                    "topic": "휴강" if is_hol else f"{seq}회차 핵심 개념 및 실전 문제풀이",
                }
            )
            seq += 1
        cur += timedelta(days=1)
    return json.dumps(items, ensure_ascii=False)


def generate_demo_sheet_from_rms(rms_df: pd.DataFrame) -> pd.DataFrame:
    """
    로드된 RMS DataFrame을 기반으로 다양한 검수 시나리오가 포함된 시트 제출내역 DataFrame을 생성합니다.
    """
    rows: List[Dict[str, Any]] = []
    if rms_df is None or rms_df.empty:
        return pd.DataFrame()

    reverse_alias = {
        "리아": "박소현(리아)",
        "박소현(셀리나)": "박소현",
        "민정": "김민정",
        "이진리": "이지원",
        "현자의 돌": "임수민",
        "이현호(미친국어)": "이현호",
    }

    for idx, r in rms_df.iterrows():
        inst_rms = str(r.get("강사", "")).strip()
        inst_sheet = reverse_alias.get(inst_rms, inst_rms)
        subj = str(r.get("과목", "")).strip()
        grade = str(r.get("학년", "")).strip()
        pure_title = str(r.get("순수강좌명", "") or r.get("강좌명", "")).strip()
        days_str = str(r.get("수업요일", "")).strip()
        days_list = list(r.get("요일_리스트", []) or [])
        sched_str = str(r.get("표준수업시간", "") or r.get("수업시간", "")).strip()
        open_d = str(r.get("개강일", "")).strip()
        close_d = str(r.get("종강일", "")).strip()
        sessions = int(r.get("회차", 0) or 0)
        code = str(r.get("강좌코드", f"{idx+1000}")).strip()
        doc_id = f"1DemoDocId_{code}_RussellDaechi2027"

        # 시나리오 분기
        # 0) 로컬 폴더에만 존재(시트 미제출) 시뮬레이션 대상 (idx 10, 22) 및 완전 미제출 대상 (idx % 9 == 8)
        if idx in (10, 22) or (idx % 9 == 8):
            continue
        # 1) 앞쪽 강좌 중 일부는 2개월 통합 과정(1~2월 8주~16회)으로 생성하되 1월 회차 역산이 일치하도록 구성
        elif idx % 7 == 1 and open_d.startswith("2027-01"):
            feb_close = "2027-02-25"
            plan_json = build_plan_json_for_dates(open_d, feb_close, days_list)
            parsed_items = json.loads(plan_json)
            total_2m_sessions = len(parsed_items)
            rows.append(
                {
                    "강사명": inst_sheet,
                    "과목": subj,
                    "대상학년": grade,
                    "강좌명": pure_title,
                    "강좌구분": "특강",
                    "수능유형": "수능/내신",
                    "수업요일": days_str,
                    "수업시간": sched_str,
                    "개강일": open_d,
                    "종강일": feb_close,
                    "총회차": total_2m_sessions,
                    "최근제출일시": "2026-11-01 15:30:00",
                    "회차별계획(JSON)": plan_json,
                    "저장폴더URL": f"https://drive.google.com/drive/folders/folder_{code}",
                    "문서URL": f"https://docs.google.com/document/d/{doc_id}/edit",
                }
            )
        # 2) 강좌명 변경 확인 케이스 (요일/시간/개강일/회차는 동일하나 강좌명이 완전히 바뀜)
        elif idx in (3, 14, 25):
            plan_json = build_plan_json_for_dates(open_d, close_d, days_list)
            rows.append(
                {
                    "강사명": inst_sheet,
                    "과목": subj,
                    "대상학년": grade,
                    "강좌명": f"[윈터 개편] 2028 수능 대비 {subj} 시그니처 마스터 클래스",
                    "강좌구분": "정규",
                    "수능유형": "수능",
                    "수업요일": days_str,
                    "수업시간": sched_str,
                    "개강일": open_d,
                    "종강일": close_d,
                    "총회차": sessions,
                    "최근제출일시": "2026-11-01 16:10:00",
                    "회차별계획(JSON)": plan_json,
                    "저장폴더URL": f"https://drive.google.com/drive/folders/folder_{code}",
                    "문서URL": f"https://docs.google.com/document/d/{doc_id}/edit",
                }
            )
        # 3) 불일치 케이스 (요일 오류, 시간 오류, 개강일 오류)
        elif idx in (7, 19, 31, 42):
            wrong_days = "목" if days_str != "목" else "월"
            wrong_sched = f"{wrong_days} 18:00~21:00" if idx in (7, 19) else sched_str
            wrong_open = "2027-01-11" if idx == 31 else open_d
            plan_json = build_plan_json_for_dates(wrong_open, close_d, days_list)
            rows.append(
                {
                    "강사명": inst_sheet,
                    "과목": subj,
                    "대상학년": grade,
                    "강좌명": pure_title,
                    "강좌구분": "특강",
                    "수능유형": "내신",
                    "수업요일": wrong_days if idx in (7, 19) else days_str,
                    "수업시간": wrong_sched,
                    "개강일": wrong_open,
                    "종강일": close_d,
                    "총회차": sessions + (2 if idx == 42 else 0),
                    "최근제출일시": "2026-11-01 17:05:00",
                    "회차별계획(JSON)": plan_json if idx != 42 else "[]",
                    "저장폴더URL": f"https://drive.google.com/drive/folders/folder_{code}",
                    "문서URL": f"https://docs.google.com/document/d/{doc_id}/edit",
                }
            )
        # 5) 정상 완벽 일치 케이스 (12월 말 조기 개강 건 포함)
        else:
            plan_json = build_plan_json_for_dates(open_d, close_d, days_list)
            rows.append(
                {
                    "강사명": inst_sheet,
                    "과목": subj,
                    "대상학년": grade,
                    "강좌명": pure_title,
                    "강좌구분": "정규",
                    "수능유형": "수능",
                    "수업요일": days_str,
                    "수업시간": sched_str,
                    "개강일": open_d,
                    "종강일": close_d,
                    "총회차": sessions,
                    "최근제출일시": "2026-11-01 14:00:00",
                    "회차별계획(JSON)": plan_json,
                    "저장폴더URL": f"https://drive.google.com/drive/folders/folder_{code}",
                    "문서URL": f"https://docs.google.com/document/d/{doc_id}/edit",
                }
            )

            # 중복 제출 테스트용 과거 제출 이력 추가 (상위 3개 강좌)
            if idx in (0, 5, 8):
                rows.append(
                    {
                        "강사명": inst_sheet,
                        "과목": subj,
                        "대상학년": grade,
                        "강좌명": f"(구버전 초안) {pure_title}",
                        "강좌구분": "정규",
                        "수능유형": "수능",
                        "수업요일": days_str,
                        "수업시간": sched_str,
                        "개강일": open_d,
                        "종강일": close_d,
                        "총회차": sessions,
                        "최근제출일시": "2026-10-28 09:15:00",
                        "회차별계획(JSON)": plan_json,
                        "저장폴더URL": f"https://drive.google.com/drive/folders/folder_{code}",
                        "문서URL": f"https://docs.google.com/document/d/{doc_id}_old/edit",
                    }
                )

    # 6) 테스트 데이터 추가 (자동 필터링 검증용: 홍길동, 아자스, 야호야호, 심동준, test 강좌명)
    for t_inst, t_title in [
        ("홍길동", "2027 국어 개념완성"),
        ("아자스", "수학 실전 테스트반"),
        ("야호야호", "영어 독해 올인원"),
        ("심동준", "물리학 개념 특강"),
        ("김수현", "[TEST] 강의계획서 제출 양식 test"),
    ]:
        rows.append(
            {
                "강사명": t_inst,
                "과목": "국어",
                "대상학년": "고3",
                "강좌명": t_title,
                "강좌구분": "특강",
                "수능유형": "수능",
                "수업요일": "월",
                "수업시간": "월 09:00~12:00",
                "개강일": "2027-01-04",
                "종강일": "2027-01-25",
                "총회차": 4,
                "최근제출일시": "2026-10-20 11:00:00",
                "회차별계획(JSON)": "[]",
                "저장폴더URL": "",
                "문서URL": "",
            }
        )

    # 7) 타 시기 강좌 추가 (10~11월 개강 및 종강 강좌 -> 자동 필터링 대상)
    rows.append(
        {
            "강사명": "심수현",
            "과목": "국어",
            "대상학년": "고2",
            "강좌명": "2학기 기말고사 내신 대비 파이널 특강",
            "강좌구분": "내신",
            "수능유형": "내신",
            "수업요일": "토",
            "수업시간": "토 14:00~17:00",
            "개강일": "2026-10-10",
            "종강일": "2026-11-28",
            "총회차": 8,
            "최근제출일시": "2026-09-25 13:00:00",
            "회차별계획(JSON)": "[]",
            "저장폴더URL": "",
            "문서URL": "",
        }
    )

    # 8) 시트에는 제출되었으나 RMS에는 미개설(폐강/미등록)된 강좌 추가
    rows.append(
        {
            "강사명": "박성민",
            "과목": "수학",
            "대상학년": "고2",
            "강좌명": "[미개설 확인용] 고2 기하 심화 마스터 특강 (일요반)",
            "강좌구분": "특강",
            "수능유형": "수능",
            "수업요일": "일",
            "수업시간": "일 09:00~12:00",
            "개강일": "2027-01-03",
            "종강일": "2027-01-31",
            "총회차": 5,
            "최근제출일시": "2026-11-01 18:20:00",
            "회차별계획(JSON)": build_plan_json_for_dates("2027-01-03", "2027-01-31", ["일"]),
            "저장폴더URL": "",
            "문서URL": "https://docs.google.com/document/d/1DemoUnopenedCourseDocId/edit",
        }
    )

    return pd.DataFrame(rows)


def create_demo_local_drive_folder(base_dir: Path, rms_df: pd.DataFrame) -> Path:
    """
    로컬 드라이브 스캔 및 문서 내용 검증(Tab 4) 테스트를 위해
    샘플 동기화 폴더(sample_drive_folder)에 .gdoc, .docx, .pdf 파일을 생성합니다.
    """
    base_dir.mkdir(parents=True, exist_ok=True)
    if rms_df is None or rms_df.empty:
        return base_dir

    # 1) .gdoc 파일 생성 (상위 5개 강좌)
    for idx in range(min(5, len(rms_df))):
        r = rms_df.iloc[idx]
        code = str(r.get("강좌코드", f"{idx}")).strip()
        inst = str(r.get("강사", "")).strip()
        subj = str(r.get("과목", "")).strip()
        grade = str(r.get("학년", "")).strip()
        pure_title = str(r.get("순수강좌명", "")).strip()
        safe_title = pure_title.replace(":", "_").replace("/", "_")
        sub_folder = base_dir / f"{subj}_{inst}"
        sub_folder.mkdir(parents=True, exist_ok=True)

        doc_id = f"1DemoDocId_{code}_RussellDaechi2027"
        gdoc_path = sub_folder / f"[{grade}_{subj}]_{inst}_{safe_title}.gdoc"
        gdoc_content = {
            "doc_id": doc_id,
            "url": f"https://docs.google.com/document/d/{doc_id}/edit",
            "email": "russel_daechi@megastudy.net",
        }
        gdoc_path.write_text(json.dumps(gdoc_content, ensure_ascii=False, indent=2), encoding="utf-8")

    # 2) 수동 수령 .docx 파일 생성 (idx 10, 22 -> 시트에는 미제출했으나 로컬 폴더에만 존재하는 강좌!)
    try:
        from docx import Document

        for idx in (10, 22):
            if idx >= len(rms_df):
                continue
            r = rms_df.iloc[idx]
            inst = str(r.get("강사", "")).strip()
            subj = str(r.get("과목", "")).strip()
            grade = str(r.get("학년", "")).strip()
            pure_title = str(r.get("순수강좌명", "")).strip()
            sched = str(r.get("표준수업시간", "") or r.get("수업시간", "")).strip()
            open_d = str(r.get("개강일", "")).strip()
            close_d = str(r.get("종강일", "")).strip()
            sessions = int(r.get("회차", 0) or 0)

            safe_title = pure_title.replace(":", "_").replace("/", "_")
            sub_folder = base_dir / "수동수령_개별제출폴더"
            sub_folder.mkdir(parents=True, exist_ok=True)
            docx_path = sub_folder / f"강의계획서_{subj}_{grade}_{inst}_{safe_title}.docx"

            doc = Document()
            doc.add_heading(f"2027년 1월 러셀 대치학원 강의계획서 ({subj})", level=1)
            doc.add_paragraph(f"담당강사: {inst} 선생님")
            doc.add_paragraph(f"강좌명: {pure_title}")
            doc.add_paragraph(f"수강대상: {grade} / 과목: {subj}")
            doc.add_paragraph(f"수업시간: {sched}")
            doc.add_paragraph(f"개강일: {open_d} / 종강일: {close_d} (총 {sessions}회)")

            table = doc.add_table(rows=1, cols=3)
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = "회차"
            hdr_cells[1].text = "수업일자"
            hdr_cells[2].text = "강의 주제 및 세부 진도 계획"
            for s_i in range(1, sessions + 1):
                row_cells = table.add_row().cells
                row_cells[0].text = f"{s_i}회차"
                row_cells[1].text = f"2027-01 ({s_i}주차)"
                row_cells[2].text = f"{pure_title} - {s_i}단원 핵심 개념 정리 및 실전 기출 적용"

            doc.save(str(docx_path))
    except Exception:
        pass

    return base_dir
