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

]

HISTORY_FILE = "crawl_history.json"
KST = timezone(timedelta(hours=9))

# ==========================================
# 2. 디자인 스타일 정의 (중복 제거)
# ==========================================
STYLES = {
    "li": "margin-bottom: 12px; line-height: 1.6; color: #334155; font-size: 14px; list-style: none;",
    "link": "display: inline-block; margin-left: 8px; font-size: 12px; color: #2563eb; text-decoration: none; font-weight: 600; padding: 2px 8px; background-color: #eff6ff; border-radius: 4px;",
    "badge_new": "<span style='background-color: #ef4444; color: #ffffff; font-size: 11px; font-weight: bold; padding: 3px 6px; border-radius: 4px; margin-right: 6px;'>🆕 NEW</span>",
    "badge_event": "<span style='background-color: #fee2e2; color: #991b1b; font-size: 12px; font-weight: bold; padding: 3px 8px; border-radius: 4px; margin-right: 6px;'>설명회/이벤트</span>",
    "badge_notice": "<span style='background-color: #f1f5f9; color: #475569; font-size: 12px; font-weight: bold; padding: 3px 8px; border-radius: 4px; margin-right: 6px;'>공지/시간표</span>"
}

# ==========================================
# 3. 데이터 이력 관리 (신규 항목 판별)
# ==========================================
def load_previous_history():
    """어제(또는 직전) 수집된 데이터를 불러옵니다."""
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_current_history(results):
    """현재 수집된 데이터를 파일로 저장합니다."""
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

# ==========================================
# 4. HTML 포맷팅
# ==========================================
def format_item_to_html(item_text, is_new=False):
    """텍스트 항목을 세련된 HTML로 변환"""
    url_match = re.search(r'\((https?://[^\s]+)\)', item_text)
    link_html = ""
    clean_text = item_text
    
    if url_match:
        url = url_match.group(1)
        clean_text = item_text.replace(f"- ({url})", "").strip()
        link_html = f'<a href="{url}" target="_blank" style="{STYLES["link"]}">자세히 보기 &rarr;</a>'

    new_badge_html = STYLES["badge_new"] if is_new else ""

    if "[📢" in clean_text:
        content_text = re.sub(r'\[📢[^\]]+\]', '', clean_text).strip()
        return f'<li style="{STYLES["li"]}">{new_badge_html}{STYLES["badge_event"]} <strong style="color: #0f172a;">{content_text}</strong> {link_html}</li>'
    
    elif any(tag in clean_text for tag in ["[📌", "[📄", "[📅"]):
        content_text = re.sub(r'\[[^\]]+\]', '', clean_text).strip()
        return f'<li style="{STYLES["li"]}">{new_badge_html}{STYLES["badge_notice"]} {content_text} {link_html}</li>'
    
    else:
        return f'<li style="{STYLES["li"]}">{new_badge_html}{clean_text} {link_html}</li>'

# ==========================================
# 5. 메일 발송 로직
# ==========================================
def send_email_report(results):
    now_kst = datetime.now(KST)
    today_str = now_kst.strftime("%Y-%m-%d")
    collection_time = now_kst.strftime('%Y-%m-%d %H:%M')
    
    prev_history = load_previous_history()
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📢 [{today_str}] 주요 학원 신규 모니터링 리포트"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)

    total_count = sum(len(items) for items in results.values() if isinstance(items, list) and not any("수집 실패" in i for i in items))
    
    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; background-color: #f8fafc; margin: 0; padding: 40px 20px;">
        
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="700" style="background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 10px 25px rgba(0,0,0,0.05);">
            <!-- 헤더 -->
            <tr>
                <td style="background-color: #1e293b; padding: 30px 40px; text-align: left;">
                    <p style="margin: 0 0 8px 0; font-size: 12px; color: #94a3b8; font-weight: 600; letter-spacing: 1px;">DAILY COMPETITOR REPORT</p>
                    <h1 style="margin: 0; font-size: 24px; color: #ffffff; font-weight: 700; line-height: 1.4;">경쟁학원 모니터링 통합 리포트</h1>
                </td>
            </tr>

            <!-- 요약 정보 -->
            <tr>
                <td style="background-color: #f1f5f9; padding: 16px 40px; border-bottom: 1px solid #e2e8f0;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <td style="font-size: 13px; color: #475569;">
                                🕒 <strong>수집 일시 (KST):</strong> {collection_time}
                            </td>
                            <td align="right" style="font-size: 14px; color: #0f172a; font-weight: 700;">
                                총 감지 <span style="color: #2563eb;">{total_count}</span>건
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>

            <!-- 본문 내용 -->
            <tr>
                <td style="padding: 30px 40px;">
    """

    for academy, items in results.items():
        is_error = items and any("수집 실패" in item for item in items)
        prev_items = prev_history.get(academy, [])
        
        # 디자인 변수 설정
        header_bg = "#fef2f2" if is_error else "#ffffff"
        header_color = "#b91c1c" if is_error else "#0f172a"
        border_color = "#fecaca" if is_error else "#e2e8f0"

        html += f"""
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 24px; border: 1px solid {border_color}; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.02);">
            <tr>
                <td style="background-color: {header_bg}; padding: 14px 20px; font-weight: 700; font-size: 16px; color: {header_color}; border-bottom: 1px solid {border_color};">
                    🏫 {academy}
                </td>
            </tr>
            <tr>
                <td style="padding: 20px; background-color: #ffffff;">
        """
        
        if is_error:
            html += '<p style="margin: 0; font-size: 14px; color: #ef4444; font-weight: 600;">⚠️ 접속 지연 또는 웹사이트 구조 변경으로 수집에 실패했습니다.</p>'
        elif items:
            html += '<ul style="margin: 0; padding: 0;">'
            for item in items:
                # 이전 데이터에 없던 항목이면 신규(True)로 판별
                is_new = item not in prev_items
                html += format_item_to_html(item, is_new=is_new)
            html += '</ul>'
        else:
            html += '<p style="margin: 0; font-size: 14px; color: #94a3b8;">최근 신규 등록된 공지/시간표가 없습니다.</p>'
            
        html += """
                </td>
            </tr>
        </table>
        """

    html += """
                </td>
            </tr>
            <!-- 푸터 -->
            <tr>
                <td style="background-color: #f8fafc; padding: 24px 40px; text-align: center; border-top: 1px solid #e2e8f0;">
                    <p style="margin: 0; font-size: 12px; color: #64748b; line-height: 1.6;">
                        본 리포트는 자동화 시스템에 의해 KST 기준으로 매일 발송됩니다.<br>
                        문의사항이 있으시다면 시스템 관리자에게 연락 바랍니다.
                    </p>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """

    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print(f"✅ 총 {len(RECEIVER_EMAILS)}명에게 이메일 발송 완료!")
        
        # 메일 발송이 성공한 후 현재 수집 데이터를 이력으로 덮어씌움
        save_current_history(results)
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")