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
    msg["Subject"] = f"📢 [{today_str}] 학원 신규 공지·시간표·설명회 요약"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)

    total_count = sum(len(items) for items in results.values())

    # 1. html 변수 선언 및 초기화 (이 부분이 빠져있어서 에러가 발생했습니다)
    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <h2 style="color: #1a73e8;">🎓 일일 학원 공지·시간표·설명회 통합 보고서</h2>
        <p style="font-size: 0.9em; color: #666;">수집 일시: {today_str} | 총 {total_count}건 감지</p>
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
    """

    # 2. 학원별 수집 결과 조립
    for academy, items in results.items():
        html += f"<h3 style='margin-bottom: 5px; color: #2c3e50;'>🏫 {academy}</h3>"
        if items:
            html += "<ul style='margin-top: 5px; padding-left: 20px;'>"
            for item in items:
                if "[📢" in item:
                    html += f"<li style='margin-bottom: 8px; color: #d9534f; font-weight: bold;'>{item}</li>"
                else:
                    html += f"<li style='margin-bottom: 8px;'>{item}</li>"
            html += "</ul>"
        else:
            html += "<p style='color: #888; font-size: 0.9em; margin-top: 5px;'>- 최근 변동 사항 없음 -</p>"
        html += "<br>"

    html += "</body></html>"

    msg.attach(MIMEText(html, "html", "utf-8"))

    # 3. Dooray / SMTP 전송 로직
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print("✅ 성공적으로 메일이 발송되었습니다!")
    
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