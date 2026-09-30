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
# 모던하고 세련된 시스템 폰트 스택 최적화
FONT_STACK = "'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont, system-ui, Roboto, 'Helvetica Neue', 'Segoe UI', 'Apple SD Gothic Neo', 'Noto Sans KR', 'Malgun Gothic', sans-serif"

STYLES = {
    "list_item": f"margin: 0; padding: 20px 0; border-bottom: 1px solid #f1f5f9; list-style: none; font-family: {FONT_STACK};",
    "link": f"display: inline-flex; align-items: center; margin-top: 10px; font-size: 13px; color: #2563eb; text-decoration: none; font-weight: 600; font-family: {FONT_STACK};",
    "badge_new": "<span style='display: inline-block; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: #ffffff; font-size: 10px; font-weight: 800; padding: 3px 8px; border-radius: 12px; margin-right: 8px; letter-spacing: 0.5px; vertical-align: middle;'>NEW</span>",
    "badge_event": "<span style='display: inline-block; background-color: #fee2e2; color: #991b1b; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 6px; margin-right: 8px; vertical-align: middle; letter-spacing: -0.025em;'>📢 설명회</span>",
    "badge_notice": "<span style='display: inline-block; background-color: #f1f5f9; color: #475569; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 6px; margin-right: 8px; vertical-align: middle; letter-spacing: -0.025em;'>📅 시간표/공지</span>"
}

# ==========================================
# 3. 데이터 이력 관리 (신규 항목 판별)
# ==========================================
def load_previous_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_current_history(results):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

# ==========================================
# 4. HTML 포맷팅
# ==========================================
def format_item_to_html(item_text, is_new=False):
    url_match = re.search(r'\((https?://[^\s]+)\)', item_text)
    link_html = ""
    clean_text = item_text
    
    if url_match:
        url = url_match.group(1)
        clean_text = item_text.replace(f"- ({url})", "").strip()
        link_html = f'<br><a href="{url}" target="_blank" style="{STYLES["link"]}">웹사이트에서 확인하기 &rarr;</a>'

    new_badge_html = STYLES["badge_new"] if is_new else ""

    clean_text = re.sub(r'\[📢[^\]]+\]', '', clean_text).strip()
    clean_text = re.sub(r'\[[^\]]+\]', '', clean_text).strip()

    if "[📢" in item_text:
        return f'<li style="{STYLES["list_item"]}">{new_badge_html}{STYLES["badge_event"]} <span style="color: #0f172a; font-size: 15px; font-weight: 600; line-height: 1.6; letter-spacing: -0.02em; vertical-align: middle;">{clean_text}</span> {link_html}</li>'
    else:
        return f'<li style="{STYLES["list_item"]}">{new_badge_html}{STYLES["badge_notice"]} <span style="color: #334155; font-size: 14.5px; font-weight: 500; line-height: 1.6; letter-spacing: -0.02em; vertical-align: middle;">{clean_text}</span> {link_html}</li>'

# ==========================================
# 5. 메일 발송 로직
# ==========================================
def send_email_report(results):
    now_kst = datetime.now(KST)
    today_str = now_kst.strftime("%Y-%m-%d")
    collection_time = now_kst.strftime('%Y년 %m월 %d일 %H:%M')
    
    prev_history = load_previous_history()
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[일일 리포트] {today_str} 경쟁학원 모니터링 업데이트"
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
    <body style="background-color: #f1f5f9; margin: 0; padding: 40px 10px; font-family: {FONT_STACK}; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale;">
        
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 680px; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.05), 0 8px 10px -6px rgba(0,0,0,0.05);">
            
            <!-- 세련된 다크 헤더 -->
            <tr>
                <td style="background-color: #0f172a; background-image: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 48px 40px; text-align: center;">
                    <span style="display: inline-block; padding: 6px 14px; background: rgba(255,255,255,0.08); border-radius: 20px; color: #94a3b8; font-size: 11px; font-weight: 700; letter-spacing: 2px; margin-bottom: 16px; text-transform: uppercase;">DAILY INSIGHT</span>
                    <h1 style="margin: 0; color: #ffffff; font-size: 26px; font-weight: 800; letter-spacing: -0.03em;">경쟁학원 모니터링 리포트</h1>
                    <p style="margin: 10px 0 0 0; color: #94a3b8; font-size: 14px; font-weight: 400; letter-spacing: -0.01em;">{collection_time} 기준 업데이트</p>
                </td>
            </tr>

            <!-- 요약 바 -->
            <tr>
                <td style="background-color: #ffffff; padding: 20px 40px; border-bottom: 1px solid #e2e8f0; text-align: right;">
                    <span style="font-size: 13.5px; color: #64748b; font-weight: 500; letter-spacing: -0.02em;">총 수집된 신규/변경 항목 </span>
                    <span style="display: inline-block; background-color: #eff6ff; color: #2563eb; font-size: 14px; font-weight: 700; padding: 4px 12px; border-radius: 20px; margin-left: 8px;">{total_count}건</span>
                </td>
            </tr>

            <!-- 본문 콘텐츠 -->
            <tr>
                <td style="padding: 20px 40px 40px 40px;">
    """

    for academy, items in results.items():
        is_error = items and any("수집 실패" in item for item in items)
        prev_items = prev_history.get(academy, [])
        
        html += f"""
                    <div style="margin-top: 28px;">
                        <h2 style="margin: 0 0 14px 0; font-size: 17px; color: #0f172a; font-weight: 700; letter-spacing: -0.025em; display: flex; align-items: center;">
                            <span style="font-size: 19px; margin-right: 8px;">🏛️</span> {academy}
                        </h2>
                        <div style="background-color: #f8fafc; border-radius: 12px; padding: 0 24px; border: 1px solid #e2e8f0;">
        """
        
        if is_error:
            html += '<p style="padding: 20px 0; margin: 0; font-size: 14px; color: #ef4444; font-weight: 600;">⚠️ 웹사이트 구조 변경 또는 접속 지연으로 데이터를 불러오지 못했습니다.</p>'
        elif items:
            html += '<ul style="margin: 0; padding: 0;">'
            for item in items:
                is_new = item not in prev_items
                html += format_item_to_html(item, is_new=is_new)
            html += '</ul>'
        else:
            html += '<p style="padding: 24px 0; margin: 0; font-size: 14px; color: #94a3b8; text-align: center; font-weight: 400;">업데이트된 자료가 없습니다.</p>'
            
        html += """
                        </div>
                    </div>
        """

    # 푸터
    html += f"""
                </td>
            </tr>
            <tr>
                <td style="background-color: #f8fafc; padding: 30px 40px; text-align: center; border-top: 1px solid #e2e8f0;">
                    <p style="margin: 0; font-size: 13px; color: #475569; font-weight: 700; margin-bottom: 6px; letter-spacing: -0.01em;">Megastudy Education</p>
                    <p style="margin: 0; font-size: 12px; color: #94a3b8; line-height: 1.5; letter-spacing: -0.01em;">
                        본 리포트는 KST 기준으로 매일 자동 수집되어 발송됩니다.<br>
                        데이터 수집 문의 및 모니터링 대상 추가는 시스템 관리자에게 연락 바랍니다.
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

    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print(f"✅ 총 {len(RECEIVER_EMAILS)}명에게 세련된 폰트의 리포트 메일 발송 완료!")
        
        save_current_history(results)
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")