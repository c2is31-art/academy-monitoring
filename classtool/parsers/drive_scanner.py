"""
[소스 3] 로컬 PC Google Drive 동기화 폴더 스캔 및 문서 내용 추출 모듈
- 지원 포맷: .pdf (pdfplumber / PyMuPDF), .docx (python-docx), .gdoc (JSON doc_id + Google Docs/Drive API 또는 오프라인 메타데이터)
- 파일명 및 문서 본문에서 강사명/강좌명/요일/시간/개강일 정보를 추출하여 수동 수령 파일 감지 및 3단 대조 지원
- 검수 완료 파일 일괄 표준명([과목]_[학년]_[강사명]_[강좌명].확장자) 리네임/복사(Export) 기능 제공
"""

from __future__ import annotations

import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from parsers.rms_parser import (
    parse_schedule_text,
    standardize_date,
    strip_course_brackets,
)
from parsers.sheet_parser import extract_google_doc_id

SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".gdoc"}


def sanitize_windows_filename(name: str) -> str:
    r"""Windows 파일명에서 사용할 수 없는 특수문자(\ / : * ? " < > |)를 안전한 문자로 치환합니다."""
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", str(name))
    cleaned = re.sub(r"[\r\n\t]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned.strip(".")


def extract_text_from_pdf(file_path: Path, max_pages: int = 10) -> Tuple[str, str]:
    """
    .pdf 파일에서 텍스트 및 표 데이터를 추출합니다.
    1순위: pdfplumber (본문 + 표 데이터)
    2순위: PyMuPDF (fitz) 폴백
    반환값: (추출된_전체_텍스트, 추출_상태_메시지)
    """
    text_chunks: List[str] = []
    status_msg = "성공 (pdfplumber)"

    # 1순위: pdfplumber
    try:
        import pdfplumber

        with pdfplumber.open(str(file_path)) as pdf:
            pages_to_read = pdf.pages[:max_pages]
            for idx, page in enumerate(pages_to_read, start=1):
                page_txt = page.extract_text() or ""
                if page_txt.strip():
                    text_chunks.append(f"[Page {idx}]\n{page_txt.strip()}")

                # 표 데이터 추출
                try:
                    tables = page.extract_tables() or []
                    for t_idx, tbl in enumerate(tables, start=1):
                        rows_str = []
                        for row in tbl:
                            if not row:
                                continue
                            cells = [
                                re.sub(r"\s+", " ", str(c or "")).strip()
                                for c in row
                                if c is not None and str(c).strip()
                            ]
                            if cells:
                                rows_str.append(" | ".join(cells))
                        if rows_str:
                            text_chunks.append(f"[Page {idx} - 표 {t_idx}]\n" + "\n".join(rows_str))
                except Exception:
                    pass
    except Exception as e_plumber:
        status_msg = f"pdfplumber 예외 발생 ({e_plumber}) -> PyMuPDF 폴백"

    combined = "\n\n".join(text_chunks).strip()
    if combined:
        return combined, status_msg

    # 2순위: PyMuPDF (fitz)
    try:
        import fitz  # PyMuPDF

        doc = fitz.open(str(file_path))
        fitz_chunks: List[str] = []
        for idx in range(min(len(doc), max_pages)):
            page = doc.load_page(idx)
            p_text = page.get_text("text") or ""
            if p_text.strip():
                fitz_chunks.append(f"[Page {idx + 1}]\n{p_text.strip()}")
        doc.close()
        combined_fitz = "\n\n".join(fitz_chunks).strip()
        if combined_fitz:
            return combined_fitz, "성공 (PyMuPDF)"
        return "", "텍스트 없음 (스캔 이미지 PDF 가능성)"
    except Exception as e_fitz:
        return "", f"PDF 텍스트 추출 실패: {e_fitz}"


def extract_text_from_docx(file_path: Path) -> Tuple[str, str]:
    """
    .docx 파일에서 문단(Paragraphs) 및 표(Tables) 내부 텍스트를 추출합니다.
    반환값: (추출된_전체_텍스트, 추출_상태_메시지)
    """
    try:
        from docx import Document

        doc = Document(str(file_path))
        chunks: List[str] = []

        # 문단 추출
        para_lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
        if para_lines:
            chunks.append("[문단 본문]\n" + "\n".join(para_lines))

        # 표 내부 텍스트 추출 (중복 셀 방지 처리 포함)
        for t_idx, table in enumerate(doc.tables, start=1):
            table_rows: List[str] = []
            for row in table.rows:
                row_cells: List[str] = []
                prev_text = None
                for cell in row.cells:
                    c_text = re.sub(r"\s+", " ", cell.text or "").strip()
                    if c_text and c_text != prev_text:
                        row_cells.append(c_text)
                        prev_text = c_text
                if row_cells:
                    table_rows.append(" | ".join(row_cells))
            if table_rows:
                chunks.append(f"[표 {t_idx}]\n" + "\n".join(table_rows))

        full_text = "\n\n".join(chunks).strip()
        return full_text, "성공 (python-docx)" if full_text else "빈 워드 문서"
    except PermissionError as e:
        return "", f"권한 오류 (파일 열림 여부 확인): {e}"
    except Exception as e:
        return "", f"DOCX 텍스트 추출 실패: {e}"


def _read_google_doc_body_via_api(
    doc_id: str,
    service_account_path: Optional[str] = None,
    service_account_info: Optional[Dict[str, Any]] = None,
) -> Tuple[str, str]:
    """
    1순위: pydata-google-auth 로컬 캐시된 내 구글 계정 권한으로 Google Drive API 텍스트 Export 수행
    2순위: 무인증 Export 엔드포인트 시도 후, 비공개/오프라인인 경우 doc_id 메타데이터 매칭 모드로 동작
    """
    if not doc_id:
        return "", "doc_id 없음"

    # 1순위: pydata_google_auth 캐시된 자격증명으로 Drive API export 시도
    try:
        import pydata_google_auth
        from googleapiclient.discovery import build
        from parsers.sheet_parser import SCOPES

        creds = pydata_google_auth.get_user_credentials(
            SCOPES,
            auth_local_webserver=True,
        )
        if creds and creds.valid:
            drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)
            exported = (
                drive_service.files()
                .export(fileId=doc_id, mimeType="text/plain")
                .execute()
            )
            text = (
                exported.decode("utf-8-sig", errors="replace").strip()
                if isinstance(exported, bytes)
                else str(exported).strip()
            )
            if text:
                return text, "성공 (내 구글 계정 OAuth Drive API)"
    except Exception:
        pass

    return "", "메타데이터 매칭 모드 (doc_id 기준 시트 문서URL 대조)"


def extract_info_from_gdoc(
    file_path: Path,
    service_account_path: Optional[str] = None,
    service_account_info: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    """
    로컬 Google Drive 동기화 폴더 내 .gdoc 파일(JSON 포맷)을 읽어
    doc_id 및 url을 추출하고, 가능 시 API로 본문 텍스트를 가져옵니다.
    """
    doc_id = ""
    url = ""
    raw_content = ""

    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            raw_content = file_path.read_text(encoding=enc)
            break
        except Exception:
            continue

    if raw_content:
        try:
            data = json.loads(raw_content.strip())
            if isinstance(data, dict):
                doc_id = str(data.get("doc_id") or data.get("resource_id") or "").strip()
                url = str(data.get("url") or "").strip()
        except Exception:
            # JSON 파싱 실패 시 정규식으로 추출
            m_id = re.search(r'"(?:doc_id|resource_id)"\s*:\s*"([^"]+)"', raw_content)
            if m_id:
                doc_id = m_id.group(1).strip()
            m_url = re.search(r'"url"\s*:\s*"([^"]+)"', raw_content)
            if m_url:
                url = m_url.group(1).strip()

    if not doc_id and url:
        doc_id = extract_google_doc_id(url)
    if not url and doc_id:
        url = f"https://docs.google.com/document/d/{doc_id}/edit"

    body_text, status_msg = _read_google_doc_body_via_api(
        doc_id=doc_id,
        service_account_path=service_account_path,
        service_account_info=service_account_info,
    )

    if not body_text:
        body_text = (
            f"[Google 문서(.gdoc) 로컬 메타데이터]\n"
            f"- 파일명: {file_path.name}\n"
            f"- 문서 ID (doc_id): {doc_id or '미확인'}\n"
            f"- 문서 URL: {url or '미확인'}\n"
            f"- 상태: {status_msg}"
        )

    return {
        "doc_id": doc_id,
        "url": url,
        "text": body_text,
        "status": status_msg,
    }


def parse_metadata_from_filename_and_text(
    file_path: Path,
    extracted_text: str,
    known_instructors: Optional[List[str]] = None,
) -> Dict[str, str]:
    """
    파일명, 상위 폴더명, 추출된 문서 본문 텍스트에서
    강사명, 강좌명, 과목, 학년, 수업요일/시간, 개강일 단서를 추출합니다.
    """
    stem = file_path.stem
    parent_name = file_path.parent.name
    combined_title_source = f"{parent_name} {stem}"

    # 파일명에서 불필요한 공통 수식어 제거
    cleaned_stem = re.sub(r"(강의계획서|수업계획서|계획서|러셀대치|러셀|대치|2027년?|1월|윈터스쿨|윈터)", " ", stem)
    cleaned_stem = re.sub(r"[_\-]+", " ", cleaned_stem)
    cleaned_stem = re.sub(r"\s+", " ", cleaned_stem).strip()

    detected_instructor = ""
    if known_instructors:
        # 길이가 긴 강사명부터 우선 매칭 (예: '박소현(셀리나)' > '박소현')
        sorted_insts = sorted(
            [i for i in known_instructors if i and len(i.strip()) >= 2],
            key=lambda x: len(x),
            reverse=True,
        )
        for inst in sorted_insts:
            if inst in combined_title_source:
                detected_instructor = inst
                break
        if not detected_instructor and extracted_text:
            head_text = extracted_text[:1500]
            for inst in sorted_insts:
                if inst in head_text:
                    detected_instructor = inst
                    break

    # 본문 텍스트에서 주요 필드 정규식 탐색
    head_text = extracted_text[:2500] if extracted_text else ""

    if not detected_instructor and head_text:
        m_inst = re.search(r"(?:강\s*사\s*명?|선\s*생\s*님|담당\s*강사)\s*[:：|\t]\s*([^\n|]{2,15})", head_text)
        if m_inst:
            cand = re.sub(r"(선생님|강사|T)$", "", m_inst.group(1).strip()).strip()
            detected_instructor = cand

    detected_course_title = ""
    if head_text:
        m_title = re.search(r"(?:강\s*좌\s*명|강\s*의\s*명|수\s*업\s*명)\s*[:：|\t]\s*([^\n|]{3,60})", head_text)
        if m_title:
            detected_course_title = m_title.group(1).strip()

    if not detected_course_title:
        detected_course_title = cleaned_stem if cleaned_stem else stem

    # 본문에서 요일/시간 추출
    detected_schedule = ""
    if head_text:
        sched_matches = re.findall(
            r"[월화수목금토일](?:\s*[,/·]\s*[월화수목금토일])*\s*(?:요일)?\s*\d{1,2}:\d{2}\s*[~\-∼～]\s*\d{1,2}:\d{2}",
            head_text,
        )
        if sched_matches:
            ps = parse_schedule_text(" ".join(sched_matches[:3]))
            detected_schedule = ps["formatted_schedule"]

    # 본문에서 개강일 추출
    detected_start_date = ""
    if head_text:
        m_start = re.search(
            r"(?:개\s*강\s*일|첫\s*수\s*업\s*일|시\s*작\s*일)\s*[:：|\t]?\s*(\d{2,4}[.\-/년\s]+\d{1,2}[.\-/월\s]+\d{1,2})",
            head_text,
        )
        if m_start:
            raw_d = re.sub(r"[년월]", "-", m_start.group(1))
            raw_d = re.sub(r"일|\s+", "", raw_d)
            detected_start_date = standardize_date(raw_d)

    return {
        "detected_instructor": detected_instructor,
        "detected_course_title": detected_course_title,
        "detected_schedule": detected_schedule,
        "detected_start_date": detected_start_date,
        "cleaned_stem": cleaned_stem,
    }


def scan_local_drive_folder(
    folder_path: str,
    known_instructors: Optional[List[str]] = None,
    service_account_path: Optional[str] = None,
    service_account_info: Optional[Dict[str, Any]] = None,
) -> Tuple[pd.DataFrame, List[str]]:
    """
    로컬 PC의 Google Drive 동기화 폴더를 재귀 탐색하여
    .pdf, .docx, .gdoc 파일 목록 및 추출 텍스트/메타데이터를 DataFrame으로 반환합니다.

    반환값: (scanned_files_df, warnings_list)
    """
    warnings: List[str] = []
    columns = [
        "파일명",
        "확장자",
        "상위폴더",
        "전체경로",
        "파일크기(KB)",
        "수정일시",
        "문서ID",
        "문서URL",
        "추출_강사명",
        "추출_강좌명",
        "추출_수업시간",
        "추출_개강일",
        "텍스트_추출상태",
        "추출_텍스트",
        "파일명_정제",
    ]

    if not folder_path or not str(folder_path).strip():
        warnings.append("스캔할 로컬 폴더 경로가 입력되지 않았습니다.")
        return pd.DataFrame(columns=columns), warnings

    root = Path(str(folder_path).strip())
    if not root.exists():
        warnings.append(f"지정한 폴더 경로가 존재하지 않습니다: {root}")
        return pd.DataFrame(columns=columns), warnings
    if not root.is_dir():
        warnings.append(f"지정한 경로가 폴더(디렉토리)가 아닙니다: {root}")
        return pd.DataFrame(columns=columns), warnings

    records: List[Dict[str, Any]] = []

    def _on_walk_error(err: OSError) -> None:
        warnings.append(f"하위 폴더 접근 권한/탐색 오류 (건너뜀): {err}")

    for dirpath, dirnames, filenames in os.walk(str(root), onerror=_on_walk_error):
        # 숨김 폴더나 임시 폴더 제외
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d != "__pycache__"]
        for fname in filenames:
            # Word 임시 잠금 파일(~$...) 제외
            if fname.startswith("~$") or fname.startswith("."):
                continue
            fpath = Path(dirpath) / fname
            ext = fpath.suffix.lower()
            if ext not in SUPPORTED_EXTENSIONS:
                continue

            try:
                stat = fpath.stat()
                size_kb = round(stat.st_size / 1024.0, 1)
                mtime_str = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            except Exception as e:
                size_kb = 0.0
                mtime_str = ""
                warnings.append(f"파일 메타데이터 조회 오류 ({fname}): {e}")

            doc_id = ""
            doc_url = ""
            extracted_text = ""
            status_msg = ""

            if ext == ".pdf":
                extracted_text, status_msg = extract_text_from_pdf(fpath)
            elif ext == ".docx":
                extracted_text, status_msg = extract_text_from_docx(fpath)
            elif ext == ".gdoc":
                gdoc_info = extract_info_from_gdoc(
                    fpath,
                    service_account_path=service_account_path,
                    service_account_info=service_account_info,
                )
                doc_id = gdoc_info["doc_id"]
                doc_url = gdoc_info["url"]
                extracted_text = gdoc_info["text"]
                status_msg = gdoc_info["status"]

            meta = parse_metadata_from_filename_and_text(
                file_path=fpath,
                extracted_text=extracted_text,
                known_instructors=known_instructors,
            )

            records.append(
                {
                    "파일명": fname,
                    "확장자": ext,
                    "상위폴더": fpath.parent.name,
                    "전체경로": str(fpath),
                    "파일크기(KB)": size_kb,
                    "수정일시": mtime_str,
                    "문서ID": doc_id,
                    "문서URL": doc_url,
                    "추출_강사명": meta["detected_instructor"],
                    "추출_강좌명": meta["detected_course_title"],
                    "추출_수업시간": meta["detected_schedule"],
                    "추출_개강일": meta["detected_start_date"],
                    "텍스트_추출상태": status_msg,
                    "추출_텍스트": extracted_text,
                    "파일명_정제": meta["cleaned_stem"],
                }
            )

    df = pd.DataFrame(records, columns=columns)
    return df, warnings


def export_standardized_files(
    export_items: List[Dict[str, str]],
    target_dir: str,
) -> Tuple[int, int, List[str]]:
    """
    검수 완료된 로컬 파일들을 `[과목]_[학년]_[강사명]_[강좌명].확장자` 표준 규칙으로
    지정된 폴더(target_dir)에 일괄 복사/리네임(Export)합니다.

    export_items 각 항목 형식:
    {
        'source_path': 원본 파일 전체 경로,
        'subject': 과목,
        'grade': 학년,
        'instructor': 강사명,
        'course_name': 순수강좌명(또는 강좌명),
    }
    반환값: (성공_건수, 실패_건수, 결과_로그_리스트)
    """
    out_dir = Path(target_dir.strip())
    out_dir.mkdir(parents=True, exist_ok=True)

    success_cnt = 0
    fail_cnt = 0
    logs: List[str] = []

    for item in export_items:
        src = Path(item.get("source_path", ""))
        if not src.exists() or not src.is_file():
            fail_cnt += 1
            logs.append(f"❌ 원본 파일 없음: {src}")
            continue

        subj = sanitize_windows_filename(item.get("subject", "미분류") or "미분류")
        grade = sanitize_windows_filename(item.get("grade", "전학년") or "전학년")
        inst = sanitize_windows_filename(item.get("instructor", "미상") or "미상")
        pure_title, _ = strip_course_brackets(item.get("course_name", "강좌명미상"))
        title_clean = sanitize_windows_filename(pure_title or item.get("course_name", "강좌명미상"))
        # 파일명 길이 제한 보호
        if len(title_clean) > 60:
            title_clean = title_clean[:60].strip()

        ext = src.suffix.lower()
        new_filename = f"[{subj}]_[{grade}]_[{inst}]_[{title_clean}]{ext}"
        dest = out_dir / new_filename

        # 동일 파일명 존재 시 순번 부여
        counter = 2
        while dest.exists() and dest.resolve() != src.resolve():
            new_filename = f"[{subj}]_[{grade}]_[{inst}]_[{title_clean}]_{counter}{ext}"
            dest = out_dir / new_filename
            counter += 1

        try:
            shutil.copy2(str(src), str(dest))
            success_cnt += 1
            logs.append(f"✅ 복사 완료: {src.name}  ➔  {dest.name}")
        except Exception as e:
            fail_cnt += 1
            logs.append(f"❌ 복사 실패 ({src.name}): {e}")

    return success_cnt, fail_cnt, logs
