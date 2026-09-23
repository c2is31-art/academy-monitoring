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
       
]

def send_email_report(results):
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📢 [{today_str}] 주요 학원 통합 모니터링 리포트"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)

    # 전체 감지 건수 계산
    total_count = sum(len(items) for items in results.values() if isinstance(items, list))

    # HTML 메일 본문 시작 ('html' 변수 하나로 통합)
    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 10px;">🎓 일일 학원 공지·시간표·설명회 통합 보고서</h2>
        <p style="font-size: 0.9em; color: #666;">수집 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 총 <b>{total_count}</b>건 감지</p>
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
    """

    # 학원별 데이터 조립
    for academy, items in results.items():
        html += f"<h3 style='margin-bottom: 8px; color: #2c3e50; background-color: #f8f9fa; padding: 8px 12px; border-left: 4px solid #1a73e8;'>🏫 {academy}</h3>"
        
        if items:
            html += "<ul style='margin-top: 5px; padding-left: 20px;'>"
            for item in items:
                # 설명회/간담회 항목은 빨간색 강조 처리
                if "[📢" in item:
                    html += f"<li style='margin-bottom: 8px; color: #d9534f; font-weight: bold;'>{item}</li>"
                else:
                    html += f"<li style='margin-bottom: 8px;'>{item}</li>"
            html += "</ul>"
        else:
            html += "<p style='color: #888; font-size: 0.9em; margin-left: 10px;'>- 최근 변동 사항 없음 -</p>"
        html += "<br>"

    # 하단 푸터(Footer) 연결
    html += """
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
        <div style="font-size: 0.8em; color: #888; text-align: center;">
            본 메일은 경쟁학원 크롤링 자동화 시스템(GitHub Actions)에 의해 매일 발송됩니다.
        </div>
    </body>
    </html>
    """

    # 메일 본문 첨부
    msg.attach(MIMEText(html, "html", "utf-8"))

    # Dooray / SSL 메일 전송 (1회 실행)
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print(f"✅ 총 {len(RECEIVER_EMAILS)}명에게 성공적으로 이메일 보고서가 발송되었습니다!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")