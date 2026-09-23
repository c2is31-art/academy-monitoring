import re
from datetime import datetime, timedelta
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright
from notifier import send_email_report

# ==========================================
# 1. 학원별 상세 URL 및 셀렉터 설정
# ==========================================
academies_info = {
    "시대인재(SdiJ)": {
        "notice": "https://www.sdij.com",
        "briefing": "https://www.sdij.com/aca/briefing/status/",
        "timetable1": "https://www.sdij.com/aca/schd/",
        "timetable2": "https://www.sdij.com/aca/schd/default.asp",
        "selector": "a, li, tr, img, .notice_list li, .timetable_list tr"
    },
    "메가스터디 MEXX": {
        "notice": "https://mexx.megastudy.net/",
        "briefing": "https://mexx.megastudy.net/campus_common/2026/fair/list.asp",
        "timetable1": "https://mexx.megastudy.net/mexx/schedule/",
        "timetable2": "https://mexx.megastudy.net/mexx/schedule/?grd=2&sgrd=1",
        "selector": "a, li, tr, img, .board_list tr, .schedule_wrap"
    },
    "두각(Dugak)": {
        "notice": "https://www.dugak.net/info/notice_dt.do",
        "briefing": "https://www.dugak.net/pre/reserv.do",
        "timetable1": "https://www.dugak.net/time/?tid=51&mid=353",
        "timetable2": "https://www.dugak.net/time/?tid=29",
        "selector": "a, li, tr, img, .notice_item, .timetable_area"
    },
    "대찬학원": {
        "notice": "https://daechanedu.com",
        "briefing": "https://daechanedu.com/menu/?menu_str=0408",
        "timetable1": "https://daechanedu.com/menu/?menu_str=0304",
        "timetable2": "https://daechanedu.com/menu/?menu_str=0305",
        "selector": "a, li, tr, img, .board_table tr"
    },
    "SNT학원": {
        "notice": "https://www.sntedu.co.kr",
        "briefing": "https://www.sntedu.co.kr/presentation/newclass/",
        "timetable1": "https://www.sntedu.co.kr/class01/02/?cate03=1&cate04=allimg",
        "timetable2": "https://www.sntedu.co.kr/class01/02/?cate03=4&cate04=allimg",
        "selector": "a, li, tr, img, .notice_list tr"
    },
    "미래탐구(Mirae)": {
        "notice": "https://dh.mirae-academy.co.kr/customer/notice",
        "briefing": "https://dhres.mirae-academy.co.kr/front/reservation_cardType",
        "timetable1": "https://dh.mirae-academy.co.kr/study/schedule_booking?schyear_code=180",
        "timetable2": "https://dh.mirae-academy.co.kr/study/schedule_booking?schyear_code=190",
        "selector": "a, li, tr, img, .board_list li"
    },
    "세정학원": {
        "notice": "https://sejungedu.com",
        "briefing": "https://sejungedu.com/explain/presentationplan?co=%EC%A4%91%EB%93%B1",
        "timetable1": "https://sejungedu.com/schedule/timetable?co=%EA%B3%A03",
        "timetable2": "https://sejungedu.com/schedule/timetable?t=&co=%EA%B3%A02",
        "selector": "a, li, tr, img, .notice_table tr"
    }
}

# ==========================================
# 2. 최근 2일 이내 날짜 및 키워드 검증 함수
# ==========================================
def is_recent(text):
    """텍스트 내 날짜를 추출하여 최근 2일 이내 게시글인지 판별"""
    today = datetime.now()
    two_days_ago = today - timedelta(days=2)
    
    # 단순 메뉴 버튼 텍스트 필터링 (메뉴명만 들어있는 경우 스킵)
    if text.strip() in ["시간표", "설명회", "간담회", "시간표 안내", "설명회 신청", "공지사항"]:
        return False

    date_patterns = [
        r'(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})',
        r'(\d{2})[-.\/](\d{1,2})[-.\/](\d{1,2})'
    ]
    
    # 1. 날짜가 포함된 경우 날짜 우선 검증
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            try:
                groups = match.groups()
                if len(groups[0]) == 4:
                    year, month, day = int(groups[0]), int(groups[1]), int(groups[2])
                else:
                    year, month, day = 2000 + int(groups[0]), int(groups[1]), int(groups[2])
                
                item_date = datetime(year, month, day)
                if two_days_ago <= item_date <= today + timedelta(days=1):
                    return True
                else:
                    return False # 2일 이전 과거 데이터는 통과하지 못함
            except ValueError:
                continue

    # 2. 날짜 표기가 없지만 본문 길이가 적절하고 핵심 키워드가 포함된 신규 공지/배너 처리
    keywords = ["시간표", "설명회", "간담회", "개강", "신규반", "모집", "안내"]
    if len(text) >= 8 and any(kw in text for kw in keywords):
        return True
        
    return False

# ==========================================
# 3. 크롤링 메인 로직
# ==========================================
def crawl_academies():
    results = {}
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080},
            locale="ko-KR",
            ignore_https_errors=True
        )
        page = context.new_page()

        for name, info in academies_info.items():
            print(f"\n==========================================")
            print(f"🏫 [{name}] 모니터링 시작")
            print(f"==========================================")
            
            academy_updates = []
            targets = [
                ("공지사항", info.get("notice")),
                ("설명회", info.get("briefing")),
                ("시간표1", info.get("timetable1")),
                ("시간표2", info.get("timetable2"))
            ]
            
            has_error = False

            for cat_name, target_url in targets:
                if not target_url or target_url.strip() == "":
                    continue
                
                print(f"👉 [{cat_name}] 수집 시도: {target_url}")
                
                try:
                    # 두각학원 등 접속 대기 필요한 사이트 분기
                    if "두각" in name:
                        page.goto(target_url, wait_until="commit", timeout=30000)
                        page.wait_for_timeout(4000)
                    else:
                        page.goto(target_url, wait_until="domcontentloaded", timeout=15000)
                        page.wait_for_timeout(1500)

                    # --- [기능 1] 이미지/배너/링크 키워드 수집 ---
                    images = page.query_selector_all("img, a")
                    for img in images:
                        alt_text = img.get_attribute("alt") or ""
                        title_text = img.get_attribute("title") or ""
                        link_url = img.get_attribute("href") or target_url
                        
                        # urljoin을 사용하여 상대주소를 정확한 절대주소로 변환
                        link_url = urljoin(target_url, link_url)

                        combined_text = f"{alt_text} {title_text}".strip()
                        if any(kw in combined_text for kw in ["설명회", "시간표", "간담회", "개강"]):
                            if is_recent(combined_text):
                                prefix = "[📢 설명회]" if any(k in combined_text for k in ["설명회", "간담회"]) else "[📅 시간표]"
                                entry = f"{prefix} [{cat_name}] {combined_text} - ({link_url})"
                                if entry not in academy_updates:
                                    academy_updates.append(entry)

                    # --- [기능 2] 텍스트 요솟값 수집 ---
                    elements = page.query_selector_all(info.get("selector", "a, li, tr"))
                    for el in elements:
                        text = el.inner_text().strip().replace("\n", " ")
                        if len(text) > 5 and is_recent(text):
                            prefix = "[📢 설명회]" if any(k in text for k in ["설명회", "간담회"]) else "[📌 공지/시간표]"
                            
                            link_el = el.query_selector("a")
                            link_url = link_el.get_attribute("href") if link_el else target_url
                            
                            # 정확한 URL 경로 결합 (urljoin)
                            link_url = urljoin(target_url, link_url) if link_url else target_url

                            entry = f"{prefix} [{cat_name}] {text[:80]} - ({link_url})"
                            if entry not in academy_updates:
                                academy_updates.append(entry)

                except Exception as e:
                    print(f"❌ [{name} - {cat_name}] 수집 중 에러: {e}")
                    has_error = True

            # 결과 처리
            if has_error and not academy_updates:
                results[name] = ["⚠️ 접속 지연 또는 학원 웹사이트 구조 변경으로 수집 실패"]
            else:
                results[name] = academy_updates

        browser.close()
        
    return results

# ==========================================
# 4. 실행 진입점
# ==========================================
if __name__ == "__main__":
    print("🚀 주요 학원 통합 모니터링 크롤러를 실행합니다...")
    crawled_data = crawl_academies()
    
    print("\n📩 수집 완료! 이메일 리포트를 전송합니다...")
    send_email_report(crawled_data)