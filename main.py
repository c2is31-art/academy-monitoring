import os
import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
from notifier import send_email_report

# 1. 최근 게시글/정보 판단 함수 (2일 이내 기준)
def is_recent(date_str):
    if not date_str:
        return False
    
    today = datetime.today().date()
    yesterday = today - timedelta(days=1)
    
    numbers = re.findall(r'\d+', date_str)
    if not numbers:
        return False
    
    try:
        clean_date = re.sub(r'[^\d]', '', date_str)
        if len(clean_date) >= 8:
            target_date = datetime.strptime(clean_date[:8], "%Y%m%d").date()
            return target_date in [today, yesterday]
        elif len(clean_date) == 6:
            target_date = datetime.strptime(clean_date, "%y%m%d").date()
            return target_date in [today, yesterday]
    except Exception:
        pass
        
    return False

# 2. 학원별 수집 메인 로직 (Playwright 활용)
def crawl_academies():
    results = {}
    
    academies_info = {
        "시대인재(SdiJ)": {
            "url": "https://www.sdij.com",
            "selector": ".notice_list li, .timetable_list tr, .briefing_list li"
        },
        "메가스터디 MEXX": {
            "url": "https://www.megastudy.net",
            "selector": ".board_list tr, .schedule_wrap, .presentation_list li"
        },
        "두각(Dugak)": {
            "url": "https://dugak.com",
            "selector": ".notice_item, .timetable_area, .briefing_item"
        },
        "대찬학원": {
            "url": "https://daechan.com",
            "selector": ".board_table tr, .briefing_table tr"
        },
        "S&T학원": {
            "url": "https://snt.com",
            "selector": ".notice_list tr"
        },
        "미래탐구(Mirae)": {
            "url": "https://mirae.com",
            "selector": ".board_list li, .event_list li"
        },
        "세정학원": {
            "url": "https://sejeong.com",
            "selector": ".notice_table tr, .briefing_area tr"
        }
    }

    # 설명회 관련 핵심 키워드 리스트
    briefing_keywords = ["설명회", "입시설명회", "간담회", "세미나", "학부모 교실", "예약"]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800},
            ignore_https_errors=True
        )
        page = context.new_page()

        for name, info in academies_info.items():
            print(f"[{name}] 크롤링 시작: {info['url']}")
            updates = []
            
            try:
                # 대기 조건을 domcontentloaded로 설정하고 타임아웃을 15초로 설정
                page.goto(info['url'], wait_until="domcontentloaded", timeout=15000)
                page.wait_for_timeout(2000) # 화면 요소 렌더링을 위해 2초 추가 대기
                
                # --- [기능 1] 팝업창 및 배너 이미지 ---
                popups = page.query_selector_all(".popup_area, .modal_content, img[src*='timetable'], img[src*='notice'], img[src*='event'], img[src*='briefing']")
                for popup in popups:
                    img_src = popup.get_attribute("src") or ""
                    alt_text = popup.get_attribute("alt") or "팝업/배너 안내"
                    
                    if any(kw in alt_text.lower() or kw in img_src.lower() for kw in briefing_keywords):
                        updates.append(f"[📢 설명회/간담회 팝업 감지] {alt_text} - (링크: {img_src})")
                    elif any(kw in img_src.lower() for kw in ["schedule", "timetable", "notice"]):
                        updates.append(f"[📅 팝업/시간표 이미지 감지] {alt_text} - (링크: {img_src})")

                # --- [기능 2] PDF 및 시간표/설명회 안내 첨부문서 ---
                download_links = page.query_selector_all("a[href*='.pdf'], a[href*='.hwp'], a[href*='download']")
                for link in download_links:
                    link_text = link.inner_text().strip()
                    link_url = link.get_attribute("href")
                    
                    if any(kw in link_text for kw in briefing_keywords):
                        updates.append(f"[📢 설명회 안내문서 감지] {link_text} - ({link_url})")
                    elif any(kw in link_text.lower() for kw in ["시간표", "수강", "개강", "안내"]):
                        updates.append(f"[📄 다운로드 문서 감지] {link_text} - ({link_url})")

                # --- [기능 3] 게시판/페이지 내 텍스트 ---
                elements = page.query_selector_all(info['selector'])
                for el in elements:
                    text = el.inner_text().strip().replace("\n", " ")
                    if not text:
                        continue
                    
                    link_el = el.query_selector("a")
                    link = link_el.get_attribute("href") if link_el else info['url']

                    if any(kw in text for kw in briefing_keywords):
                        if is_recent(text) or "예약" in text:
                            updates.append(f"[📢 신규 설명회/간담회 소식] {text[:80]}... - ({link})")
                    elif is_recent(text):
                        updates.append(f"[📌 신규 공지/시간표] {text[:80]}... - ({link})")

            except Exception as e:
                print(f"[{name}] 수집 중 타임아웃/에러 발생: {e}")
                updates.append(f"수집 실패 (접속 지연 또는 구조 변경)")

            # 중복 제거 후 결과 저장
            results[name] = list(dict.fromkeys(updates))
            
            # 서버 부하 방지 대기 (2초)
            page.wait_for_timeout(2000)

        browser.close()
        
    return results

# 3. HTML 메일 리포트 생성 및 발송
def main():
    print("=== 학원 모니터링 크롤러 실행 시작 ===")
    
    # crawl_academies 함수 내부에서 Playwright를 직접 생성/종료하므로 인자 없이 호출합니다.
    crawl_results = crawl_academies()
    
    # 이메일 리포트 발송
    send_email_report(crawl_results)
    print("=== 전체 프로세스 정상 완료 ===")

if __name__ == "__main__":
    main()