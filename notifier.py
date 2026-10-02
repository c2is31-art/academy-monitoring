import os
import re
import json
import smtplib
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
    "yipsung@megastudyedu.com"
]

HISTORY_FILE = "crawl_history.json"
KST = timezone(timedelta(hours=9))

# ==========================================
# 2. 디자인 및 타이포그래피 스타일 정의
# ==========================================
FONT_STACK = "'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont, system-ui, Roboto, 'Helvetica Neue', 'Segoe UI', 'Apple SD Gothic Neo', 'Noto Sans KR', 'Malgun Gothic', sans-serif"

# ==========================================
# 3. 텍스트 정규화 헬퍼 (구버전 이력 호환용)
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

# ==========================================
# 4. 데이터 이력 관리 (신규 항목 판별)
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

# ==========================================
# 5. HTML 보고서 생성 및 메일 발송 로직
# ==========================================
def format_item_to_html(item, is_new=False):
    """
    개별 공지 항목의 HTML 태그 생성
    """
    title = item.get("title", "")
    url = item.get("url", "")
    source_badge = item.get("source_badge", "")
    
    new_badge_html = "<span style='display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: #ffffff; font-size: 10px; font-weight: 800; padding: 2px 7px; border-radius: 10px; margin-right: 6px; letter-spacing: 0.5px; vertical-align: middle;'>NEW</span>" if is_new else ""
    source_badge_html = f"<span style='display: inline-block; background-color: #f1f5f9; color: #475569; font-size: 10.5px; font-weight: 600; padding: 2px 6px; border-radius: 4px; margin-right: 6px; border: 1px solid #e2e8f0; vertical-align: middle;'>{source_badge}</span>" if source_badge else ""
    
    link_html = f'<br><a href="{url}" target="_blank" style="display: inline-flex; align-items: center; margin-top: 6px; font-size: 12.5px; color: #2563eb; text-decoration: none; font-weight: 600;">웹사이트에서 확인하기 &rarr;</a>' if url else ""

    if item.get("section") == "event":
        return f"""
        <li style="margin: 0; padding: 12px 0; border-bottom: 1px solid #f1f5f9; list-style: none;">
            {new_badge_html}
            {source_badge_html}
            <span style="color: #0f172a; font-size: 14.5px; font-weight: 600; line-height: 1.5; letter-spacing: -0.02em; vertical-align: middle;">{title}</span>
            {link_html}
        </li>
        """
    else:
        return f"""
        <li style="margin: 0; padding: 12px 0; border-bottom: 1px solid #f1f5f9; list-style: none;">
            {new_badge_html}
            {source_badge_html}
            <span style="color: #334155; font-size: 14px; font-weight: 500; line-height: 1.5; letter-spacing: -0.02em; vertical-align: middle;">{title}</span>
            {link_html}
        </li>
        """


def build_email_html(results, prev_history=None):
    if prev_history is None:
        prev_history = {}
        
    now_kst = datetime.now(KST)
    collection_time = now_kst.strftime('%Y년 %m월 %d일 %H:%M')

    # 1. 신규 항목 판별 및 전처리
    processed_results = {}
    all_new_items = []
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

        prev_keys = get_history_norm_keys(prev_history, academy)
        acad_items = []

        for it in items:
            if isinstance(it, dict):
                norm_key = it.get("norm_key") or normalize_title(it.get("title", ""))
                is_new = norm_key not in prev_keys
                item_obj = dict(it)
                item_obj["is_new"] = is_new
                item_obj["academy"] = academy
                acad_items.append(item_obj)
                total_count += 1
                if is_new:
                    all_new_items.append(item_obj)
            elif isinstance(it, str):
                # 구버전 호환 처리
                norm_key = normalize_title(it)
                is_new = norm_key not in prev_keys
                item_obj = {
                    "title": clean_display_title(it),
                    "norm_key": norm_key,
                    "sources": ["공지사항"],
                    "source_badge": "[공지사항]",
                    "section": "event" if ("설명회" in it or "간담회" in it or "[📢" in it) else "notice",
                    "url": "",
                    "is_new": is_new,
                    "academy": academy
                }
                acad_items.append(item_obj)
                total_count += 1
                if is_new:
                    all_new_items.append(item_obj)

        processed_results[academy] = {
            "error": False,
            "items": acad_items
        }

    new_count = len(all_new_items)

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
                            {f'<span style="display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: #ffffff; font-size: 12px; font-weight: 800; padding: 3px 10px; border-radius: 12px; margin-left: 6px;">🔥 오늘 신규 {new_count}건</span>' if new_count > 0 else '<span style="display: inline-block; background-color: #f1f5f9; color: #64748b; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 12px; margin-left: 6px;">신규 0건</span>'}
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- 메인 콘텐츠 영역 -->
        <tr>
            <td style="padding: 28px 36px 40px 36px;">
    """

    # [요구사항 2] 신규 등록 항목 최상단 분리 배치
    if new_count > 0:
        html += f"""
                <!-- [TOP SECTION] 오늘의 신규 공지/설명회 -->
                <div style="margin-bottom: 36px; background: #fffcf8; border: 1.5px solid #fed7aa; border-radius: 14px; padding: 22px 24px; box-shadow: 0 4px 12px rgba(251, 146, 60, 0.08);">
                    <div style="margin-bottom: 12px; padding-bottom: 10px; border-bottom: 1px solid #ffedd5;">
                        <span style="font-size: 16px; font-weight: 800; color: #c2410c; letter-spacing: -0.02em;">🔥 오늘의 신규 공지 / 설명회 ({new_count}건)</span>
                    </div>
                    <p style="margin: 0 0 16px 0; font-size: 13px; color: #7c2d12; line-height: 1.4;">이전 수집 이력 대비 오늘 새롭게 포착된 핵심 공지 및 설명회 목록입니다.</p>
                    <ul style="margin: 0; padding: 0; list-style: none;">
        """
        for item in all_new_items:
            acad_name = item.get("academy", "")
            badge_cat = '<span style="display: inline-block; background-color: #fee2e2; color: #991b1b; font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 4px; margin-right: 6px; vertical-align: middle;">📢 설명회</span>' if item.get("section") == "event" else '<span style="display: inline-block; background-color: #e0f2fe; color: #0369a1; font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 4px; margin-right: 6px; vertical-align: middle;">📅 시간표/공지</span>'
            source_badge = f'<span style="display: inline-block; background-color: #f1f5f9; color: #475569; font-size: 10.5px; font-weight: 600; padding: 2px 6px; border-radius: 4px; margin-right: 6px; border: 1px solid #e2e8f0; vertical-align: middle;">{item.get("source_badge", "")}</span>'
            link_html = f'<br><a href="{item.get("url", "#")}" target="_blank" style="display: inline-flex; align-items: center; margin-top: 8px; font-size: 12.5px; color: #2563eb; text-decoration: none; font-weight: 600;">웹사이트에서 확인하기 &rarr;</a>' if item.get("url") else ""
            
            html += f"""
                        <li style="margin: 0; padding: 14px 0; border-bottom: 1px solid #ffedd5; list-style: none;">
                            <span style="display: inline-block; background-color: #0f172a; color: #ffffff; font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 4px; margin-right: 6px; vertical-align: middle;">{acad_name}</span>
                            {badge_cat}
                            {source_badge}
                            <span style="display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: #ffffff; font-size: 10px; font-weight: 800; padding: 2px 6px; border-radius: 10px; margin-right: 6px; vertical-align: middle;">NEW</span>
                            <span style="color: #0f172a; font-size: 14.5px; font-weight: 700; line-height: 1.5; vertical-align: middle;">{item.get("title", "")}</span>
                            {link_html}
                        </li>
            """
        html += """
                    </ul>
                </div>
        """
    else:
        html += """
                <!-- [TOP SECTION] 오늘의 신규 항목 없음 알림 -->
                <div style="margin-bottom: 30px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px 20px; text-align: center;">
                    <span style="font-size: 13.5px; color: #64748b; font-weight: 500;">✨ 오늘 새로 등록된 신규 공지가 없습니다. (기존 공지 현황 유지 중)</span>
                </div>
        """

    # [요구사항 3] 학원별 내부 카테고리 그룹화
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
                    html += format_item_to_html(it, is_new=it.get("is_new", False))
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
                    html += format_item_to_html(it, is_new=it.get("is_new", False))
                html += """
                                </ul>
                            </div>
                """

        html += """
                        </div>
                    </div>
        """

    html += f"""
                </div>
            </td>
        </tr>

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


def send_email_report(results):
    """
    정제된 크롤링 결과를 이메일 리포트로 전송하고 수집 이력을 갱신함
    """
    now_kst = datetime.now(KST)
    today_str = now_kst.strftime("%Y-%m-%d")
    
    prev_history = load_previous_history()
    html_content = build_email_html(results, prev_history)
    
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
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")