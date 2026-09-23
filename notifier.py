import os
import re  # <--- 이 줄이 들어가 있는지 확인 후 추가해 주세요!
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 1. Dooray SMTP 발송자 및 수신자 설정 (465 SSL 전용)
SMTP_SERVER = "smtp.dooray.com"
SMTP_PORT = 465  # SSL 전용 포트
SENDER_EMAIL = "c2is@megastudyedu.com"       # 보내는 사람 이메일
SENDER_PASSWORD = "us9wwst7vhqyxxe"     # Dooray 비밀번호 / 앱 비밀번호

# 📧 보고서를 받을 이메일 목록
RECEIVER_EMAILS = [
       "c2is@megastudy.net",
       
]

def format_item_to_html(item_text):
    """크롤링된 텍스트 항목을 가독성 높은 HTML 태그/버튼으로 변환"""
    # URL 링크 추출
    url_match = re.search(r'\((https?://[^\s]+)\)', item_text)
    link_html = ""
    clean_text = item_text
    
    if url_match:
        url = url_match.group(1)
        clean_text = item_text.replace(f"- ({url})", "").strip()
        link_html = f'<a href="{url}" target="_blank" style="display: inline-block; margin-left: 8px; font-size: 11px; color: #1a73e8; text-decoration: none; font-weight: bold; background-color: #e8f0fe; padding: 2px 8px; border-radius: 4px;">자세히 보기 🔗</a>'

    # 설명회/간담회 강조 (붉은 태그)
    if "[📢" in clean_text:
        tag_badge = '<span style="background-color: #fce8e6; color: #d93025; font-size: 11px; font-weight: bold; padding: 3px 7px; border-radius: 4px; margin-right: 6px;">설명회/이벤트</span>'
        content_text = re.sub(r'\[📢[^\]]+\]', '', clean_text).strip()
        return f'<li style="margin-bottom: 10px; line-height: 1.5; color: #202124; font-size: 13px;">{tag_badge} <strong>{content_text}</strong> {link_html}</li>'
    
    # 시간표/일반 공지 (파란 태그)
    elif "[📌" in clean_text or "[📄" in clean_text or "[📅" in clean_text:
        tag_badge = '<span style="background-color: #e8f0fe; color: #1a73e8; font-size: 11px; font-weight: bold; padding: 3px 7px; border-radius: 4px; margin-right: 6px;">공지/시간표</span>'
        content_text = re.sub(r'\[[^\]]+\]', '', clean_text).strip()
        return f'<li style="margin-bottom: 10px; line-height: 1.5; color: #3c4043; font-size: 13px;">{tag_badge} {content_text} {link_html}</li>'
    
    else:
        return f'<li style="margin-bottom: 8px; line-height: 1.5; color: #3c4043; font-size: 13px;">{clean_text} {link_html}</li>'

def send_email_report(results):
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📢 [{today_str}] 주요 학원 통합 모니터링 리포트"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)

    # 전체 감지 건수 및 수집 성공 학원 수
    total_count = sum(len(items) for items in results.values() if isinstance(items, list) and not any("수집 실패" in i for i in items))
    
    # 본문 헤더 시작
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Pretendard', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; background-color: #f4f6f8; margin: 0; padding: 20px 10px; -webkit-font-smoothing: antialiased;">
        <div style="max-width: 680px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border: 1px solid #e5e8ec;">
            
            <!-- 상단 헤더 뱅크 -->
            <div style="background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%); padding: 30px 25px; color: #ffffff;">
                <h1 style="margin: 0 0 10px 0; font-size: 22px; font-weight: 700; letter-spacing: -0.5px;">🎓 학원 모니터링 통합 데일리 리포트</h1>
                <p style="margin: 0; font-size: 13px; opacity: 0.9;">주요 경쟁 학원 공지사항 및 설명회/시간표 신규 업데이트 현황</p>
            </div>

            <!-- 요약 대시보드 바 -->
            <div style="background-color: #f8f9fa; padding: 15px 25px; border-bottom: 1px solid #eef1f4; display: flex; align-items: center; justify-content: space-between;">
                <span style="font-size: 12px; color: #5f6368; font-weight: 500;">
                    📅 수집 시각: <strong>{datetime.now().strftime('%Y-%m-%d %H:%M')}</strong>
                </span>
                <span style="font-size: 12px; background-color: #1a73e8; color: #ffffff; padding: 4px 10px; border-radius: 20px; font-weight: bold;">
                    총 {total_count}건 감지
                </span>
            </div>

            <!-- 메인 컨텐츠 영역 -->
            <div style="padding: 25px;">
    """

    # 학원별 카드 레이아웃 생성
    for academy, items in results.items():
        is_error = items and any("수집 실패" in item for item in items)
        
        # 학원 카드 시작
        html += f"""
        <div style="margin-bottom: 20px; border: 1px solid #e8eaed; border-radius: 8px; overflow: hidden; background-color: #ffffff;">
            <div style="background-color: {'#fce8e6' if is_error else '#f8f9fa'}; padding: 12px 16px; border-bottom: 1px solid {'#f5c6cb' if is_error else '#e8eaed'}; font-weight: bold; font-size: 15px; color: {'#d93025' if is_error else '#202124'};">
                🏫 {academy}
            </div>
            <div style="padding: 16px;">
        """
        
        if is_error:
            html += '<p style="margin: 0; font-size: 13px; color: #d93025; font-weight: 500;">⚠️ 접속 지연 또는 학원 웹사이트 구조 변경으로 수집에 실패했습니다.</p>'
        elif items:
            html += '<ul style="margin: 0; padding-left: 0; list-style-type: none;">'
            for item in items:
                html += format_item_to_html(item)
            html += '</ul>'
        else:
            html += '<p style="margin: 0; font-size: 13px; color: #80868b; font-style: italic;"> 최근 2일간 신규 올라온 공지/시간표가 없습니다.</p>'
            
        html += """
            </div>
        </div>
        """

    # 푸터 영역
    html += """
            </div>
            
            <div style="background-color: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #eef1f4;">
                <p style="margin: 0; font-size: 11px; color: #9aa0a6; line-height: 1.5;">
                    본 메일은 경쟁학원 모니터링 자동화 시스템(GitHub Actions)에 의해 매일 자동 생성됩니다.<br>
                    문의사항이나 대상 학원 추가 요청은 관리자에게 연락 바랍니다.
                </p>
            </div>
        </div>
    </body>
    </html>
    """

    msg.attach(MIMEText(html, "html", "utf-8"))

    # Dooray / SSL 메일 전송
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print(f"✅ 총 {len(RECEIVER_EMAILS)}명에게 고품질 모니터링 이메일 발송 완료!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")