import os
import re
import json
import smtplib
from difflib import SequenceMatcher
from datetime import datetime, timezone, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# 1. 설정 및 전역 변수
# ==========================================
SMTP_SERVER = "smtp.dooray.com"
SMTP_PORT = 465
SENDER_EMAIL = "c2is@megastudyedu.com"
SENDER_PASSWORD = "us9wwst7vhqyxxe"

RECEIVER_EMAILS = [
    "c2is@megastudy.net",
    "profilm@megastudyedu.com",
    "tocka@megastudyedu.com",
    "kjm99@megastudyedu.com",
    "hbkim@megastudyedu.com",
    "yipsung@megastudyedu.com",
    "ms20906@megastudyedu.com"
]

HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "crawl_history.json")
KST = timezone(timedelta(hours=9))
SIMILARITY_THRESHOLD = 0.82

# ==========================================
# 2. 디자인 및 타이포그래피 스타일 정의
# ==========================================
FONT_STACK = "'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont, system-ui, Roboto, 'Helvetica Neue', 'Segoe UI', 'Apple SD Gothic Neo', 'Noto Sans KR', 'Malgun Gothic', sans-serif"

# ==========================================
# 3. 텍스트 정규화 및 유사도 헬퍼
# ==========================================
def normalize_title(text):
    """
    제목의 공백, 특수문자, 대괄호 태그 등을 제거하여 핵심 내용 기준 정규화 텍스트 생성
    """
    if not text:
        return ""
    t = re.sub(r'https?://\S+', '', text)
    t = re.sub(r'\[.*?\]', '', t)
    t = re.sub(r'[^a-zA-Z0-9가-힣]', '', t)
    norm = t.lower().strip()
    if len(norm) < 2:
        t_fallback = re.sub(r'https?://\S+', '', text)
        t_fallback = re.sub(r'[^a-zA-Z0-9가-힣]', '', t_fallback).lower().strip()
        if t_fallback:
            return t_fallback
    return norm


def clean_display_title(text):
    """
    불필요한 카테고리 접두사를 제거한 대표 제목
    """
    if not text:
        return ""
    t = text.strip()
    t = re.sub(r'https?://\S+', '', t)
    t = re.sub(r'-\s*\(\s*\)', '', t)
    t = re.sub(r'\[(📢|📅|📌)?[^\]]*(설명회|간담회|공지사항|공지|시간표|신규자료|안내)[^\]]*\]', '', t)
    t = re.sub(r'\s+', ' ', t).strip(' -:|·')
    if len(t) < 3:
        t = re.sub(r'https?://\S+', '', text)
        t = re.sub(r'\s+', ' ', t).strip(' -:|·()')
    return t[:100]


def calc_title_similarity(title_a, title_b, norm_a=None, norm_b=None):
    """
    difflib.SequenceMatcher를 활용해 두 제목 간 유사도(0.0 ~ 1.0) 반환
    """
    ratios = []
    if title_a and title_b:
        ratios.append(SequenceMatcher(None, title_a.lower(), title_b.lower()).ratio())
    if norm_a and norm_b:
        ratios.append(SequenceMatcher(None, norm_a, norm_b).ratio())
    return max(ratios) if ratios else 0.0

# ==========================================
# 4. 데이터 이력 관리 (신규 및 변경 항목 판별)
# ==========================================
def load_previous_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_current_history(results, prev_history=None):
    """
    수집 성공한 학원 데이터를 crawl_history.json에 저장.
    특정 학원이 일시적 네트워크 장애 등으로 실패한 경우 이전 이력을 보존하여
    다음 날 불필요하게 모든 항목이 NEW로 인식되는 현상을 방지함.
    """
    history_to_save = dict(prev_history) if prev_history else {}
    for academy, items in results.items():
        is_error = items and any(
            ("수집 실패" in item if isinstance(item, str) else item.get("error", False))
            for item in items
        )
        if not is_error:
            history_to_save[academy] = items
            
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history_to_save, f, ensure_ascii=False, indent=2)


def get_history_norm_keys(prev_history, academy):
    """
    이전 이력에서 해당 학원의 정규화 키 목록을 Set으로 추출 (신규 딕셔너리 및 레거시 문자열 완전 호환)
    """
    keys = set()
    prev_items = prev_history.get(academy, [])
    if not isinstance(prev_items, list):
        return keys
        
    for item in prev_items:
        if isinstance(item, dict):
            k = item.get("norm_key") or normalize_title(item.get("title", ""))
            if k:
                keys.add(k)
        elif isinstance(item, str):
            if "수집 실패" in item:
                continue
            k = normalize_title(item)
            if k:
                keys.add(k)
    return keys


def _extract_history_entries(prev_history, academy):
    """
    이전 이력에서 비교 가능한 정규화 엔트리 목록 추출
    """
    entries = []
    prev_items = prev_history.get(academy, [])
    if not isinstance(prev_items, list):
        return entries

    for item in prev_items:
        if isinstance(item, dict):
            if item.get("error", False):
                continue
            title = item.get("title", "")
            norm_key = item.get("norm_key") or normalize_title(title)
            url = (item.get("url") or "").strip()
            if norm_key or title:
                entries.append({
                    "title": title,
                    "norm_key": norm_key,
                    "url": url,
                    "is_legacy": False
                })
        elif isinstance(item, str):
            if "수집 실패" in item:
                continue
            title = clean_display_title(item)
            norm_key = normalize_title(item)
            if norm_key or title:
                entries.append({
                    "title": title,
                    "norm_key": norm_key,
                    "url": "",
                    "is_legacy": True
                })
    return entries


def classify_item_status(curr_item, prev_entries, similarity_threshold=SIMILARITY_THRESHOLD):
    """
    이전 이력과 비교하여 항목의 변경 상태를 분류:
    - 'unchanged': 제목(정규화 키) 및 링크 모두 동일
    - 'changed': 유사도는 82% 이상(또는 제목 동일)이나 제목 내용이나 링크(URL)가 변경됨
    - 'new': 완전 신규 항목
    반환: (status, change_info_dict_or_None)
    """
    curr_title = curr_item.get("title", "")
    curr_norm = curr_item.get("norm_key") or normalize_title(curr_title)
    curr_url = (curr_item.get("url") or "").strip()

    # 1. 정확 일치 항목 탐색 (제목 또는 norm_key 동일)
    exact_matches = [
        p for p in prev_entries
        if (curr_norm and p["norm_key"] == curr_norm) or (curr_title and p["title"] == curr_title)
    ]

    if exact_matches:
        # 동일 URL 또는 레거시 이력이 하나라도 있으면 변동 없음 처리
        for p in exact_matches:
            if p["is_legacy"] or not p["url"] or not curr_url or p["url"] == curr_url:
                return "unchanged", None
        # 제목은 완전히 같지만 링크(URL)가 바뀐 경우
        prev_ref = exact_matches[0]
        return "changed", {
            "reason": "링크 변경",
            "prev_title": prev_ref["title"],
            "prev_url": prev_ref["url"]
        }

    # 2. 유사도(Fuzzy) 기반 내용/제목 변경 탐색 (SequenceMatcher >= 82%)
    best_ratio = 0.0
    best_prev = None
    for p in prev_entries:
        ratio = calc_title_similarity(curr_title, p["title"], curr_norm, p["norm_key"])
        if ratio >= similarity_threshold and ratio > best_ratio:
            best_ratio = ratio
            best_prev = p

    if best_prev is not None:
        reason = "제목 및 링크 변경" if (best_prev["url"] and curr_url and best_prev["url"] != curr_url) else "제목/내용 변경"
        return "changed", {
            "reason": reason,
            "prev_title": best_prev["title"],
            "prev_url": best_prev["url"],
            "similarity": round(best_ratio * 100, 1)
        }

    # 3. 완전 신규
    return "new", None

# ==========================================
# 5. HTML 보고서 생성 및 메일 발송 로직
# ==========================================
def format_item_to_html(item, is_new=False, status=None):
    """
    개별 공지 항목의 HTML 태그 생성
    """
    title = item.get("title", "")
    url = item.get("url", "")
    source_badge = item.get("source_badge", "")

    if status is None:
        status = "new" if is_new else "unchanged"

    if status == "new":
        status_badge_html = "<span style='display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: #ffffff; font-size: 10px; font-weight: 800; padding: 2px 7px; border-radius: 10px; margin-right: 6px; letter-spacing: 0.5px; vertical-align: middle;'>NEW</span>"
    elif status == "changed":
        status_badge_html = "<span style='display: inline-block; background-color: #fef3c7; color: #b45309; font-size: 10px; font-weight: 800; padding: 2px 7px; border-radius: 10px; margin-right: 6px; border: 1px solid #fcd34d; letter-spacing: 0.3px; vertical-align: middle;'>🔄 변경</span>"
    else:
        status_badge_html = ""

    source_badge_html = f"<span style='display: inline-block; background-color: #f1f5f9; color: #475569; font-size: 10.5px; font-weight: 600; padding: 2px 6px; border-radius: 4px; margin-right: 6px; border: 1px solid #e2e8f0; vertical-align: middle;'>{source_badge}</span>" if source_badge else ""
    
    link_html = f'<br><a href="{url}" target="_blank" style="display: inline-flex; align-items: center; margin-top: 6px; font-size: 12.5px; color: #2563eb; text-decoration: none; font-weight: 600;">웹사이트에서 확인하기 &rarr;</a>' if url else ""

    if item.get("section") == "event":
        return f"""
        <li style="margin: 0; padding: 12px 0; border-bottom: 1px solid #f1f5f9; list-style: none;">
            {status_badge_html}
            {source_badge_html}
            <span style="color: #0f172a; font-size: 14.5px; font-weight: 600; line-height: 1.5; letter-spacing: -0.02em; vertical-align: middle;">{title}</span>
            {link_html}
        </li>
        """
    else:
        return f"""
        <li style="margin: 0; padding: 12px 0; border-bottom: 1px solid #f1f5f9; list-style: none;">
            {status_badge_html}
            {source_badge_html}
            <span style="color: #334155; font-size: 14px; font-weight: 500; line-height: 1.5; letter-spacing: -0.02em; vertical-align: middle;">{title}</span>
            {link_html}
        </li>
        """


def preprocess_results(results, prev_history=None):
    """
    크롤링 결과와 이전 이력을 비교해 신규/변경/유지 상태를 판별하고 전처리된 결과를 반환
    """
    if prev_history is None:
        prev_history = {}

    processed_results = {}
    all_new_items = []
    all_changed_items = []
    total_count = 0

    for academy, items in results.items():
        is_error = items and any(
            ("수집 실패" in item if isinstance(item, str) else item.get("error", False))
            for item in items
        )
        if is_error:
            processed_results[academy] = {
                "error": True,
                "items": []
            }
            continue

        prev_entries = _extract_history_entries(prev_history, academy)
        acad_items = []

        for it in items:
            if isinstance(it, dict):
                item_obj = dict(it)
                if not item_obj.get("norm_key"):
                    item_obj["norm_key"] = normalize_title(item_obj.get("title", ""))
            elif isinstance(it, str):
                # 구버전 호환 처리
                item_obj = {
                    "title": clean_display_title(it),
                    "norm_key": normalize_title(it),
                    "sources": ["공지사항"],
                    "source_badge": "[공지사항]",
                    "section": "event" if ("설명회" in it or "간담회" in it or "[📢" in it) else "notice",
                    "url": ""
                }
            else:
                continue

            status, change_info = classify_item_status(item_obj, prev_entries)
            item_obj["status"] = status
            item_obj["is_new"] = (status == "new")
            item_obj["is_changed"] = (status == "changed")
            item_obj["change_info"] = change_info
            item_obj["academy"] = academy

            acad_items.append(item_obj)
            total_count += 1
            if status == "new":
                all_new_items.append(item_obj)
            elif status == "changed":
                all_changed_items.append(item_obj)

        processed_results[academy] = {
            "error": False,
            "items": acad_items
        }

    return processed_results, all_new_items, all_changed_items, total_count


def build_email_html(results, prev_history=None, health_checks=None):
    if prev_history is None:
        prev_history = {}
    if health_checks is None:
        health_checks = []
        
    now_kst = datetime.now(KST)
    collection_time = now_kst.strftime('%Y년 %m월 %d일 %H:%M')

    # 1. 신규 및 변경 항목 판별 & 전처리
    processed_results, all_new_items, all_changed_items, total_count = preprocess_results(results, prev_history)

    new_count = len(all_new_items)
    changed_count = len(all_changed_items)
    highlight_items = all_new_items + all_changed_items

    # KPI 배지 HTML 생성
    kpi_badges = []
    if new_count > 0:
        kpi_badges.append(
            f'<span style="display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: #ffffff; font-size: 12px; font-weight: 800; padding: 3px 10px; border-radius: 12px; margin-left: 6px;">🔥 오늘 신규 {new_count}건</span>'
        )
    if changed_count > 0:
        kpi_badges.append(
            f'<span style="display: inline-block; background-color: #fef3c7; color: #b45309; font-size: 12px; font-weight: 800; padding: 3px 10px; border-radius: 12px; margin-left: 6px; border: 1px solid #fcd34d;">🔄 변경 {changed_count}건</span>'
        )
    if not kpi_badges:
        kpi_badges.append(
            '<span style="display: inline-block; background-color: #f1f5f9; color: #64748b; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 12px; margin-left: 6px;">신규/변경 0건</span>'
        )
    kpi_badges_html = "".join(kpi_badges)

    # 2. HTML 템플릿 빌드
    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>[일일 리포트] 경쟁학원 모니터링 업데이트</title>
</head>
<body style="background-color: #f1f5f9; margin: 0; padding: 40px 10px; font-family: {FONT_STACK}; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale;">
    
    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 720px; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.05), 0 8px 10px -6px rgba(0,0,0,0.05);">
        
        <!-- 세련된 다크 헤더 -->
        <tr>
            <td style="background-color: #0f172a; background-image: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 44px 36px; text-align: center;">
                <span style="display: inline-block; padding: 5px 12px; background: rgba(255,255,255,0.08); border-radius: 20px; color: #94a3b8; font-size: 11px; font-weight: 700; letter-spacing: 2px; margin-bottom: 14px; text-transform: uppercase;">DAILY MONITORING REPORT</span>
                <h1 style="margin: 0; color: #ffffff; font-size: 25px; font-weight: 800; letter-spacing: -0.03em;">경쟁학원 모니터링 일일 리포트</h1>
                <p style="margin: 10px 0 0 0; color: #94a3b8; font-size: 13.5px; font-weight: 400; letter-spacing: -0.01em;">{collection_time} 기준 업데이트</p>
            </td>
        </tr>

        <!-- 요약 KPI 바 -->
        <tr>
            <td style="background-color: #ffffff; padding: 18px 36px; border-bottom: 1px solid #e2e8f0;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td align="left" style="vertical-align: middle;">
                            <span style="font-size: 13px; color: #64748b; font-weight: 500;">모니터링 대상: <strong>{len(results)}개 학원</strong></span>
                        </td>
                        <td align="right" style="vertical-align: middle;">
                            <span style="font-size: 13px; color: #64748b; font-weight: 500; margin-right: 6px;">수집 공지 <strong>{total_count}건</strong></span>
                            {kpi_badges_html}
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- 메인 콘텐츠 영역 -->
        <tr>
            <td style="padding: 28px 36px 40px 36px;">
    """

    # [TOP SECTION] 신규 및 변경 항목 최상단 분리 배치
    if highlight_items:
        summary_label_parts = []
        if new_count > 0:
            summary_label_parts.append(f"신규 {new_count}건")
        if changed_count > 0:
            summary_label_parts.append(f"변경 {changed_count}건")
        summary_count_text = " · ".join(summary_label_parts)

        html += f"""
                <!-- [TOP SECTION] 오늘의 신규 및 변경 공지/설명회 -->
                <div style="margin-bottom: 36px; background: #fffcf8; border: 1.5px solid #fed7aa; border-radius: 14px; padding: 22px 24px; box-shadow: 0 4px 12px rgba(251, 146, 60, 0.08);">
                    <div style="margin-bottom: 12px; padding-bottom: 10px; border-bottom: 1px solid #ffedd5;">
                        <span style="font-size: 16px; font-weight: 800; color: #c2410c; letter-spacing: -0.02em;">🔥 오늘의 신규 / 변경 공지 · 설명회 ({summary_count_text})</span>
                    </div>
                    <p style="margin: 0 0 16px 0; font-size: 13px; color: #7c2d12; line-height: 1.4;">이전 수집 이력 대비 오늘 새롭게 등록되었거나 내용·링크가 변경된 핵심 공지 및 설명회 목록입니다.</p>
                    <ul style="margin: 0; padding: 0; list-style: none;">
        """
        for item in highlight_items:
            acad_name = item.get("academy", "")
            badge_cat = '<span style="display: inline-block; background-color: #fee2e2; color: #991b1b; font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 4px; margin-right: 6px; vertical-align: middle;">📢 설명회</span>' if item.get("section") == "event" else '<span style="display: inline-block; background-color: #e0f2fe; color: #0369a1; font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 4px; margin-right: 6px; vertical-align: middle;">📅 시간표/공지</span>'
            source_badge = f'<span style="display: inline-block; background-color: #f1f5f9; color: #475569; font-size: 10.5px; font-weight: 600; padding: 2px 6px; border-radius: 4px; margin-right: 6px; border: 1px solid #e2e8f0; vertical-align: middle;">{item.get("source_badge", "")}</span>'

            if item.get("status") == "changed":
                status_badge = '<span style="display: inline-block; background-color: #fef3c7; color: #b45309; font-size: 10px; font-weight: 800; padding: 2px 7px; border-radius: 10px; margin-right: 6px; border: 1px solid #fcd34d; vertical-align: middle;">🔄 변경</span>'
                cinfo = item.get("change_info") or {}
                reason = cinfo.get("reason", "내용 변경")
                prev_t = cinfo.get("prev_title", "")
                change_note_html = f'<div style="margin-top: 6px; font-size: 12px; color: #92400e; background-color: #fef9c3; padding: 5px 10px; border-radius: 6px;">↳ {reason}{f" (기존: {prev_t})" if prev_t and prev_t != item.get("title") else ""}</div>'
            else:
                status_badge = '<span style="display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: #ffffff; font-size: 10px; font-weight: 800; padding: 2px 6px; border-radius: 10px; margin-right: 6px; vertical-align: middle;">NEW</span>'
                change_note_html = ""

            link_html = f'<br><a href="{item.get("url", "#")}" target="_blank" style="display: inline-flex; align-items: center; margin-top: 8px; font-size: 12.5px; color: #2563eb; text-decoration: none; font-weight: 600;">웹사이트에서 확인하기 &rarr;</a>' if item.get("url") else ""
            
            html += f"""
                        <li style="margin: 0; padding: 14px 0; border-bottom: 1px solid #ffedd5; list-style: none;">
                            <span style="display: inline-block; background-color: #0f172a; color: #ffffff; font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 4px; margin-right: 6px; vertical-align: middle;">{acad_name}</span>
                            {badge_cat}
                            {source_badge}
                            {status_badge}
                            <span style="color: #0f172a; font-size: 14.5px; font-weight: 700; line-height: 1.5; vertical-align: middle;">{item.get("title", "")}</span>
                            {change_note_html}
                            {link_html}
                        </li>
            """
        html += """
                    </ul>
                </div>
        """
    else:
        html += """
                <!-- [TOP SECTION] 오늘의 신규/변경 항목 없음 알림 -->
                <div style="margin-bottom: 30px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px 20px; text-align: center;">
                    <span style="font-size: 13.5px; color: #64748b; font-weight: 500;">✨ 오늘 새로 등록되거나 변경된 공지가 없습니다. (기존 공지 현황 유지 중)</span>
                </div>
        """

    # [BOTTOM SECTION] 학원별 내부 카테고리 그룹화
    html += """
                <!-- [BOTTOM SECTION] 학원별 전체 모니터링 현황 -->
                <div style="margin-top: 10px;">
                    <div style="margin-bottom: 18px;">
                        <span style="font-size: 17px; font-weight: 800; color: #0f172a; letter-spacing: -0.02em;">🏛️ 학원별 전체 모니터링 현황</span>
                    </div>
    """

    for academy, data in processed_results.items():
        is_error = data["error"]
        items = data["items"]
        
        events = [it for it in items if it.get("section") == "event"]
        notices = [it for it in items if it.get("section") == "notice"]
        acad_count = len(items)

        html += f"""
                    <div style="margin-bottom: 24px; background-color: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.03);">
                        <!-- 학원 헤더 -->
                        <div style="background-color: #f8fafc; padding: 14px 20px; border-bottom: 1px solid #e2e8f0; display: flex; align-items: center; justify-content: space-between;">
                            <span style="font-size: 15.5px; color: #0f172a; font-weight: 700; letter-spacing: -0.02em;">
                                🏫 {academy}
                            </span>
                            <span style="font-size: 12px; color: #64748b; font-weight: 600; background-color: #e2e8f0; padding: 2px 8px; border-radius: 10px;">
                                {acad_count}건
                            </span>
                        </div>
                        <div style="padding: 16px 20px;">
        """

        if is_error:
            html += '<p style="padding: 12px 0; margin: 0; font-size: 13.5px; color: #ef4444; font-weight: 600;">⚠️ 웹사이트 구조 변경 또는 접속 지연으로 데이터를 불러오지 못했습니다.</p>'
        elif not items:
            html += '<p style="padding: 16px 0; margin: 0; font-size: 13.5px; color: #94a3b8; text-align: center; font-weight: 400;">업데이트된 자료가 없습니다.</p>'
        else:
            # 1. [📢 설명회/간담회] 하위 그룹
            if events:
                html += f"""
                            <div style="margin-bottom: 18px;">
                                <div style="margin-bottom: 8px; padding-bottom: 6px; border-bottom: 1.5px solid #fee2e2;">
                                    <span style="font-size: 13px; font-weight: 700; color: #991b1b; letter-spacing: -0.02em;">📢 설명회 / 간담회</span>
                                    <span style="margin-left: 6px; font-size: 11px; font-weight: 700; background-color: #fee2e2; color: #991b1b; padding: 1px 6px; border-radius: 8px;">{len(events)}</span>
                                </div>
                                <ul style="margin: 0; padding: 0; list-style: none;">
                """
                for it in events:
                    html += format_item_to_html(it, status=it.get("status", "unchanged"))
                html += """
                                </ul>
                            </div>
                """

            # 2. [📅 시간표/공지] 하위 그룹
            if notices:
                html += f"""
                            <div style="margin-bottom: 6px;">
                                <div style="margin-bottom: 8px; padding-bottom: 6px; border-bottom: 1.5px solid #e2e8f0;">
                                    <span style="font-size: 13px; font-weight: 700; color: #334155; letter-spacing: -0.02em;">📅 시간표 / 공지</span>
                                    <span style="margin-left: 6px; font-size: 11px; font-weight: 700; background-color: #f1f5f9; color: #475569; padding: 1px 6px; border-radius: 8px;">{len(notices)}</span>
                                </div>
                                <ul style="margin: 0; padding: 0; list-style: none;">
                """
                for it in notices:
                    html += format_item_to_html(it, status=it.get("status", "unchanged"))
                html += """
                                </ul>
                            </div>
                """

        html += """
                        </div>
                    </div>
        """

    html += """
                </div>
            </td>
        </tr>
    """

    # [HEALTH CHECK SECTION] 접속 에러 또는 매칭 요소 0개 URL 자가 진단 리포팅 (푸터 위쪽)
    issue_checks = [
        h for h in health_checks
        if (not h.get("success", True)) or h.get("matched_count", 0) == 0
    ]
    if issue_checks:
        html += f"""
        <!-- 수집 상태 자가 진단 안내 박스 -->
        <tr>
            <td style="padding: 0 36px 28px 36px;">
                <div style="background-color: #fffbeb; border: 1px solid #fcd34d; border-radius: 12px; padding: 18px 20px;">
                    <div style="margin-bottom: 8px; font-size: 14px; font-weight: 800; color: #92400e; letter-spacing: -0.02em;">
                        ⚙️ 크롤링 상태 점검 필요 ({len(issue_checks)}건)
                    </div>
                    <p style="margin: 0 0 12px 0; font-size: 12.5px; color: #78350f; line-height: 1.5;">
                        아래 URL은 페이지 접속에 실패했거나 매칭된 원본 DOM 요소가 0개여서 사이트 구조 변경 또는 접속 차단이 의심됩니다.
                    </p>
                    <ul style="margin: 0; padding: 0; list-style: none;">
        """
        for chk in issue_checks:
            acad = chk.get("academy", "")
            cat = chk.get("category", "")
            url = chk.get("url", "")
            if not chk.get("success", True):
                err_msg = chk.get("error_message") or "페이지 접속 오류"
                reason_text = f"접속 실패 ({err_msg})"
            else:
                reason_text = "매칭된 원본 요소 0개 (사이트 구조 변경 의심)"

            html += f"""
                        <li style="margin: 0; padding: 8px 0; border-top: 1px solid #fde68a; font-size: 12.5px; color: #78350f; line-height: 1.45;">
                            <strong style="color: #92400e;">[{acad} · {cat}]</strong> {reason_text}<br>
                            <a href="{url}" target="_blank" style="color: #2563eb; text-decoration: none; word-break: break-all;">{url}</a>
                        </li>
            """
        html += """
                    </ul>
                </div>
            </td>
        </tr>
        """

    html += f"""
        <!-- 푸터 -->
        <tr>
            <td style="background-color: #f8fafc; padding: 28px 36px; text-align: center; border-top: 1px solid #e2e8f0;">
                <p style="margin: 0; font-size: 13px; color: #475569; font-weight: 700; margin-bottom: 6px; letter-spacing: -0.01em;">Megastudy Education 경쟁학원 모니터링 시스템</p>
                <p style="margin: 0; font-size: 12px; color: #94a3b8; line-height: 1.5; letter-spacing: -0.01em;">
                    본 리포트는 한국 표준시(KST) 기준으로 자동 수집 및 정제되어 발송됩니다.<br>
                    데이터 수집 문의 및 모니터링 대상 추가/수정은 시스템 관리자에게 문의 바랍니다.
                </p>
            </td>
        </tr>
    </table>

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr><td height="40"></td></tr>
    </table>
</body>
</html>
"""
    return html


def send_email_report(results, health_checks=None):
    """
    정제된 크롤링 결과를 이메일 리포트로 전송하고 수집 이력 갱신 및 Google Sheets에 누적 저장함
    """
    now_kst = datetime.now(KST)
    today_str = now_kst.strftime("%Y-%m-%d")
    
    prev_history = load_previous_history()
    processed_results, _, _, _ = preprocess_results(results, prev_history)
    html_content = build_email_html(results, prev_history, health_checks=health_checks)
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[일일 리포트] {today_str} 경쟁학원 모니터링 업데이트"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print(f"✅ 총 {len(RECEIVER_EMAILS)}명에게 모니터링 리포트 메일 발송 완료!")
        
        save_current_history(results, prev_history=prev_history)

        # 메일 발송 및 이력 저장 직후 Google Sheets 누적 DB 저장 (오류 시 메일/이력에 영향 없도록 격리)
        try:
            from sheets_logger import log_results_to_sheets
            log_results_to_sheets(
                results=results,
                prev_history=prev_history,
                collection_time_kst=now_kst,
                processed_results=processed_results,
            )
        except Exception as sheet_err:
            print(f"⚠️ Google Sheets 기록 중 예외 발생 (이메일 및 crawl_history.json 저장은 정상 완료): {sheet_err}")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")
