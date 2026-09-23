import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# 1. Dooray SMTP 발송자 및 수신자 설정 (465 SSL 전용)
SMTP_SERVER = "smtp.dooray.com"
SMTP_PORT = 465  # SSL 전용 포트
SENDER_EMAIL = "c2is@megastudyedu.com"       # 보내는 사람 이메일
SENDER_PASSWORD = "us9wwst7vhqyxxe"     # Dooray 비밀번호 / 앱 비밀번호

# 📧 보고서를 받을 이메일 목록
RECEIVER_EMAILS = [
       "c2is@megastudy.net",
       "profilm@megastudyedu.com",
    "tocka@megastudyedu.com",
    "kjm99@megastudyedu.com",
"hbkim@megastudyedu.com"
]

def send_email_report(results):
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📢 [{today_str}] 경쟁학원 신규 공지 및 시간표 일일 모니터링 보고서"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)

    total_count = sum(len(items) for items in results.values())
    
    # HTML 이메일 템플릿
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Apple SD Gothic Neo', Arial, sans-serif; line-height: 1.6; color: #333; }}
            .container {{ max-width: 650px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 10px; }}
            .header {{ background-color: #1e3a8a; color: white; padding: 15px; border-radius: 8px 8px 0 0; text-align: center; }}
            .academy-card {{ margin-top: 20px; padding: 15px; border: 1px solid #cbd5e1; border-radius: 8px; background-color: #f8fafc; }}
            .academy-title {{ font-size: 16px; font-weight: bold; color: #0f172a; margin-bottom: 10px; border-bottom: 2px solid #2563eb; padding-bottom: 5px; }}
            .notice-list {{ list-style-type: none; padding-left: 0; margin: 0; }}
            .notice-item {{ margin-bottom: 8px; font-size: 14px; }}
            .notice-link {{ color: #2563eb; text-decoration: none; font-weight: 500; }}
            .empty-text {{ color: #94a3b8; font-size: 13px; font-style: italic; }}
            .footer {{ margin-top: 25px; text-align: center; font-size: 12px; color: #64748b; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>📊 경쟁학원 일일 모니터링 리포트</h2>
                <p style="margin:0; font-size: 14px;">수집일자: {today_str} | 총 신규 업데이트: <strong>{total_count}건</strong></p>
            </div>
    """

    for academy, items in results.items():
        html_content += f"""
            <div class="academy-card">
                <div class="academy-title">🏢 {academy} ({len(items)}건)</div>
        """
        if not items:
            html_content += '<p class="empty-text">※ 새로 등록된 공지사항/시간표가 없습니다.</p>'
        else:
            html_content += '<ul class="notice-list">'
            for item in items:
    
    if "[📢" in item:
        html += f"<li style='margin-bottom: 8px; color: #d9534f; font-weight: bold;'>{item}</li>\n"
    else:
        html += f"<li style='margin-bottom: 8px;'>{item}</li>\n"
                """
            html_content += '</ul>'
        html_content += '</div>'

    html_content += """
            <div class="footer">
                본 메일은 경쟁학원 크롤링 자동화 시스템에 의해 매일 발송됩니다.
            </div>
        </div>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_content, "html"))

    # 💡 465번 포트 전용 SMTP_SSL 적용 부분
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print(f"✅ 총 {len(RECEIVER_EMAILS)}명에게 이메일 보고서 발송 완료!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")