import os
import re
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
    "profilm@megastudyedu.com",
    "tocka@megastudyedu.com",
    "kjm99@megastudyedu.com",
"hbkim@megastudyedu.com"
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
        link_html = f'<a href="{url}" target="_blank" style="display: inline-block; margin-left: 6px; font-size: 11px; color: #1a73e8; text-decoration: none; font-weight: bold; background-color: #e8f0fe; padding: 2px 7px; border-radius: 4px;">자세히 보기 🔗</a>'

    # 설명회/간담회 강조 (붉은 태그)
    if "[📢" in clean_text:
        tag_badge = '<span style="background-color: #fce8e6; color: #d93025; font-size: 11px; font-weight: bold; padding: 3px 7px; border-radius: 4px; margin-right: 6px;">설명회/이벤트</span>'
        content_text = re.sub(r'\[📢[^\]]+\]', '', clean_text).strip()
        return f'<li style="margin-bottom: 10px; line-height: 1.5; color: #202124; font-size: 13px; list-style: none;">{tag_badge} <strong>{content_text}</strong> {link_html}</li>'
    
    # 시간표/일반 공지 (파란 태그)
    elif "[📌" in clean_text or "[📄" in clean_text or "[📅" in clean_text:
        tag_badge = '<span style="background-color: #e8f0fe; color: #1a73e8; font-size: 11px; font-weight: bold; padding: 3px 7px; border-radius: 4px; margin-right: 6px;">공지/시간표</span>'
        content_text = re.sub(r'\[[^\]]+\]', '', clean_text).strip()
        return f'<li style="margin-bottom: 10px; line-height: 1.5; color: #3c4043; font-size: 13px; list-style: none;">{tag_badge} {content_text} {link_html}</li>'
    
    else:
        return f'<li style="margin-bottom: 8px; line-height: 1.5; color: #3c4043; font-size: 13px; list-style: none;">{clean_text} {link_html}</li>'

def send_email_report(results):
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📢 [{today_str}] 주요 학원 통합 모니터링 리포트"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)

    # 전체 감지 건수
    total_count = sum(len(items) for items in results.values() if isinstance(items, list) and not any("수집 실패" in i for i in items))
    
    # 이메일 클라이언트 완벽 호환 레이아웃 (Table 기반 구조)
    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="font-family: 'Malgun Gothic', '맑은 고딕', helvetica, apple-system, sans-serif; background-color: #f4f6f8; margin: 0; padding: 20px 0;">
        
        <!-- 전체 컨테이너 테이블 -->
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="650" style="background-color: #ffffff; border-radius: 10px; overflow: hidden; border: 1px solid #e0e0e0; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
            
            <!-- 1. 메인 상단 헤더 (배경색 및 제목 가독성 강화) -->
            <tr>
                <td style="background-color: #1a73e8; padding: 25px 30px; text-align: left;">
                    <p style="margin: 0 0 6px 0; font-size: 12px; color: #bbdefb; font-weight: bold; letter-spacing: 1px;">COMPETITOR MONITORING REPORT</p>
                    <h1 style="margin: 0; font-size: 22px; color: #ffffff; font-weight: bold; line-height: 1.3;">🎓 학원 모니터링 통합 데일리 리포트</h1>
                </td>
            </tr>

            <!-- 2. 요약 정보 바 -->
            <tr>
                <td style="background-color: #f8f9fa; padding: 12px 30px; border-bottom: 1px solid #e0e0e0;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <td style="font-size: 13px; color: #5f6368;">
                                📅 <strong>수집 일시:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M')}
                            </td>
                            <td align="right" style="font-size: 13px; color: #1a73e8; font-weight: bold;">
                                총 <span style="background-color: #1a73e8; color: #ffffff; padding: 2px 8px; border-radius: 12px; font-size: 12px;">{total_count}건</span> 감지
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>

            <!-- 3. 메인 콘텐츠 (학원별 카드 목록) -->
            <tr>
                <td style="padding: 25px 30px;">
    """

    # 학원별 데이터 카드 생성
    for academy, items in results.items():
        is_error = items and any("수집 실패" in item for item in items)
        
        header_bg = "#fce8e6" if is_error else "#f1f3f4"
        header_color = "#d93025" if is_error else "#202124"
        border_color = "#f5c6cb" if is_error else "#dadce0"

        html += f"""
        <!-- {academy} 카드 -->
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 20px; border: 1px solid {border_color}; border-radius: 8px; overflow: hidden;">
            <tr>
                <td style="background-color: {header_bg}; padding: 10px 16px; font-weight: bold; font-size: 14px; color: {header_color}; border-bottom: 1px solid {border_color};">
                    🏫 {academy}
                </td>
            </tr>
            <tr>
                <td style="padding: 16px; background-color: #ffffff;">
        """
        
        if is_error:
            html += '<p style="margin: 0; font-size: 13px; color: #d93025; font-weight: bold;">⚠️ 접속 지연 또는 웹사이트 구조 변경으로 수집에 실패했습니다.</p>'
        elif items:
            html += '<ul style="margin: 0; padding: 0;">'
            for item in items:
                html += format_item_to_html(item)
            html += '</ul>'
        else:
            html += '<p style="margin: 0; font-size: 13px; color: #80868b; font-style: italic;"> 최근 2일간 신규 등록된 공지/시간표가 없습니다.</p>'
            
        html += """
                </td>
            </tr>
        </table>
        """

    # 4. 하단 푸터 영역
    html += """
                </td>
            </tr>
            <tr>
                <td style="background-color: #f8f9fa; padding: 15px 30px; text-align: center; border-top: 1px solid #e0e0e0;">
                    <p style="margin: 0; font-size: 11px; color: #70757a; line-height: 1.6;">
                        본 리포트는 주요 경쟁학원 모니터링 자동화 시스템에 의해 매일 자동 발송됩니다.<br>
                        수신 해지 및 모니터링 대상 추가 요청은 시스템 관리자에게 문의 바랍니다.
                    </p>
                </td>
            </tr>
        </table>

    </body>
    </html>
    """

    msg.attach(MIMEText(html, "html", "utf-8"))

    # Dooray / SSL 메일 전송
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print(f"✅ 총 {len(RECEIVER_EMAILS)}명에게 가독성이 개선된 이메일 발송 완료!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")