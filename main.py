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
        "notice": "https://snt-edu.com/",
        "briefing": "https://snt-edu.com/reserve",
        "timetable1": "https://snt-edu.com/timetable?grade=%EA%B3%A02&term=%EA%B8%B0%EB%A7%90%EA%B3%A0%EC%82%AC",
        "timetable2": "https://snt-edu.com/timetable?grade=%EA%B3%A01&term=%EA%B8%B0%EB%A7%90%EA%B3%A0%EC%82%AC",
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
# 2. 미래 날짜 및 특정 키워드 엄격 검증 함수 (기존 로직 100% 보존)
# ==========================================
def is_recent(text):
    """
    1) 날짜가 포함된 경우: 해당 날짜가 최근(7일 이내)이거나 미래면 키워드 상관없이 무조건 통과!
    2) 날짜가 없는 경우: 필수 키워드("2028", "윈터", "예비")가 포함되어 있으면 통과!
    """
    today_date = datetime.now().date()
    clean_text = text.strip()

    # 1. 고정 메뉴/단순 버튼 필터링
    ignore_menu_texts = [
        "시간표", "설명회", "간담회", "시간표 안내", "설명회 신청", 
        "공지사항", "학원소개", "오시는길", "수강신청", "마이페이지", "로그인", "전체보기"
    ]
    if clean_text in ignore_menu_texts or len(clean_text) < 5:
        return False

    # 2. 필수 키워드 포함 여부 확인
    target_keywords = ["2028", "윈터", "예비"]
    has_keyword = any(kw in clean_text for kw in target_keywords)

    # 3. 날짜 패턴 검사 (YYYY-MM-DD, YY.MM.DD, MM/DD 등)
    date_patterns = [
        r'(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})',
        r'(\d{2})[-.\/](\d{1,2})[-.\/](\d{1,2})'
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, clean_text)
        if match:
            try:
                groups = match.groups()
                year = int(groups[0]) if len(groups[0]) == 4 else 2000 + int(groups[0])
                month, day = int(groups[1]), int(groups[2])
                
                item_date = datetime(year, month, day).date()
                
                if item_date >= today_date - timedelta(days=7):
                    return True  
                else:
                    return False 
            except ValueError:
                continue

    return has_keyword

# ==========================================
# 3. 정규화 및 중복 병합(Deduplication) 헬퍼 함수
# ==========================================
def normalize_title(text):
    """
    핵심 내용(제목의 공백, 특수문자, 대괄호 태그 등을 제거한 정규화 텍스트) 추출
    예: '[공지사항] [고1] 2028 윈터스쿨 설명회 안내!' -> '2028윈터스쿨설명회안내'
    """
    if not text:
        return ""
    # URL 부분 제거
    t = re.sub(r'https?://\S+', '', text)
    # 대괄호 태그 [내용] 제거
    t = re.sub(r'\[.*?\]', '', t)
    # 특수문자 및 공백 제거 (영문, 숫자, 한글만 유지)
    t = re.sub(r'[^a-zA-Z0-9가-힣]', '', t)
    norm = t.lower().strip()
    
    # 대괄호 제거 후 내용이 비거나 너무 짧아진 경우 fallback
    if len(norm) < 2:
        t_fallback = re.sub(r'https?://\S+', '', text)
        t_fallback = re.sub(r'[^a-zA-Z0-9가-힣]', '', t_fallback).lower().strip()
        if t_fallback:
            return t_fallback
    return norm


def clean_display_title(text):
    """
    불필요한 카테고리성 대괄호 태그를 제거하여 읽기 쉬운 대표 제목 정제
    ([예비고1], [고2] 등 유용한 교육 대상 학년 태그는 보존)
    """
    if not text:
        return ""
    t = text.strip()
    t = re.sub(r'https?://\S+', '', t)
    t = re.sub(r'-\s*\(\s*\)', '', t)
    # 카테고리성 대괄호 태그 제거 ([📢 설명회], [📅 시간표], [📌 신규자료], [공지사항], [설명회], [시간표1], [시간표2], [공지], [안내] 등)
    t = re.sub(r'\[(📢|📅|📌)?[^\]]*(설명회|간담회|공지사항|공지|시간표|신규자료|안내)[^\]]*\]', '', t)
    # 연속 공백 정리 및 앞뒤 특수문자 정리
    t = re.sub(r'\s+', ' ', t).strip(' -:|·')
    
    if len(t) < 3:
        t = re.sub(r'https?://\S+', '', text)
        t = re.sub(r'\s+', ' ', t).strip(' -:|·()')
        
    return t[:100]


def canonical_category(cat_name):
    """
    수집 타깃 카테고리명 정규화 ('시간표1', '시간표2' -> '시간표')
    """
    cat = re.sub(r'\d+$', '', cat_name).strip()
    return cat if cat else cat_name


def get_url_priority(url, category):
    """
    가장 우선순위가 높은 대표 URL(설명회 전용 신청 링크, 상세 페이지 등) 우선순위 점수 산출
    """
    if not url:
        return -100
        
    if url.startswith("javascript:") or url.endswith("#") or not url.startswith("http"):
        return -50
        
    score = 0
    url_lower = url.lower()
    
    # 1. 예약/설명회/상세 페이지 키워드 가산점
    high_priority_keywords = [
        "reserv", "briefing", "explain", "fair", "booking", 
        "schedule", "view", "read", "detail", "tid=", "idx=", "menu_str="
    ]
    for kw in high_priority_keywords:
        if kw in url_lower:
            score += 25
            break
            
    # 2. 카테고리 가산점 (설명회 링크 우선 연결)
    if "설명회" in category:
        score += 20
    elif "시간표" in category:
        score += 10
    elif "공지사항" in category:
        score += 5
        
    # 3. 상세 경로 가산점 (메인 도메인 루트보다 상세 하위 URL 우선)
    path_part = re.sub(r'^https?://[^/]+', '', url).strip('/')
    if path_part:
        score += 15
        if "/" in path_part or "?" in url:
            score += 10
            
    return score


def classify_section(sources, title, raw_text):
    """
    수집 항목을 [📢 설명회/간담회] 또는 [📅 시간표/공지] 섹션으로 분류
    """
    event_keywords = ["설명회", "간담회", "입시설명", "설명회신청", "예약", "사전예약", "fair", "briefing"]
    if "설명회" in sources or any(kw in title for kw in event_keywords) or any(kw in raw_text for kw in event_keywords):
        return "event"
    return "notice"

# ==========================================
# 4. 크롤링 및 중복 병합 메인 로직
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
            
            raw_candidates = []
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
                    if "두각" in name:
                        page.goto(target_url, wait_until="commit", timeout=30000)
                        page.wait_for_timeout(4000)
                    else:
                        page.goto(target_url, wait_until="domcontentloaded", timeout=15000)
                        page.wait_for_timeout(1500)

                    # --- [기능 1] 이미지/배너/링크 속 키워드 수집 ---
                    images = page.query_selector_all("img, a")
                    for img in images:
                        alt_text = img.get_attribute("alt") or ""
                        title_text = img.get_attribute("title") or ""
                        href = img.get_attribute("href")
                        link_url = urljoin(target_url, href) if href and not href.startswith("javascript:") and not href.endswith("#") else target_url
                        
                        combined_text = f"{alt_text} {title_text}".strip()
                        
                        if is_recent(combined_text):
                            raw_candidates.append({
                                "text": combined_text,
                                "url": link_url,
                                "cat_name": cat_name,
                                "target_url": target_url
                            })

                    # --- [기능 2] 텍스트 요솟값 수집 ---
                    elements = page.query_selector_all(info.get("selector", "a, li, tr"))
                    for el in elements:
                        text = el.inner_text().strip().replace("\n", " ")
                        if is_recent(text):
                            link_el = el.query_selector("a")
                            href = link_el.get_attribute("href") if link_el else None
                            link_url = urljoin(target_url, href) if href and not href.startswith("javascript:") and not href.endswith("#") else target_url

                            raw_candidates.append({
                                "text": text,
                                "url": link_url,
                                "cat_name": cat_name,
                                "target_url": target_url
                            })

                except Exception as e:
                    print(f"❌ [{name} - {cat_name}] 수집 중 에러: {e}")
                    has_error = True

            # --- [정규화 기반 중복 공지 1건 병합 (Deduplication)] ---
            merged_items = {}
            for cand in raw_candidates:
                raw_text = cand["text"]
                norm_key = normalize_title(raw_text)
                if not norm_key or len(norm_key) < 2:
                    continue
                    
                canon_cat = canonical_category(cand["cat_name"])
                cand_url = cand["url"]
                
                if norm_key not in merged_items:
                    display_title = clean_display_title(raw_text)
                    merged_items[norm_key] = {
                        "title": display_title if display_title else raw_text[:80],
                        "norm_key": norm_key,
                        "sources": [canon_cat],
                        "url": cand_url,
                        "primary_cat": canon_cat,
                        "raw_text": raw_text
                    }
                else:
                    existing = merged_items[norm_key]
                    # 여러 경로에 동시 노출된 경우 해당 출처 통합
                    if canon_cat not in existing["sources"]:
                        existing["sources"].append(canon_cat)
                        
                    # 대표 URL 우선순위 비교 및 갱신
                    curr_score = get_url_priority(existing["url"], existing["primary_cat"])
                    new_score = get_url_priority(cand_url, canon_cat)
                    if new_score > curr_score:
                        existing["url"] = cand_url
                        existing["primary_cat"] = canon_cat
                        
                    # 더 길거나 완성도 높은 대표 제목 채택
                    candidate_display_title = clean_display_title(raw_text)
                    if len(candidate_display_title) > len(existing["title"]) and len(candidate_display_title) <= 90:
                        existing["title"] = candidate_display_title

            # 후처리: 출처 통합 표기 생성 및 내부 섹션 분류
            academy_updates = []
            category_order = ["공지사항", "설명회", "시간표"]

            for norm_key, item in merged_items.items():
                # 출처 정렬 (공지사항 -> 설명회 -> 시간표 순)
                item["sources"].sort(key=lambda s: category_order.index(s) if s in category_order else 99)
                # 출처 통합 표기 (예: [공지사항+설명회])
                item["source_badge"] = f"[{'+'.join(item['sources'])}]"
                # 내부 카테고리 그룹화 ([📢 설명회/간담회] vs [📅 시간표/공지])
                item["section"] = classify_section(item["sources"], item["title"], item["raw_text"])
                
                academy_updates.append(item)

            if has_error and not academy_updates:
                results[name] = [{"error": True, "message": "⚠️ 접속 지연 또는 학원 웹사이트 구조 변경으로 수집 실패"}]
                print(f"⚠️ [{name}] 수집 실패 처리")
            else:
                results[name] = academy_updates
                print(f"✨ [{name}] 총 {len(raw_candidates)}개 후보 수집 -> 정규화 중복 병합 후 {len(academy_updates)}건 정리 완료")
                for u in academy_updates:
                    print(f"   - {u['source_badge']} {u['title']} ({u['url']})")

        browser.close()
        
    return results

# ==========================================
# 5. 실행 진입점
# ==========================================
if __name__ == "__main__":
    print("🚀 주요 학원 통합 모니터링 크롤러를 실행합니다...")
    crawled_data = crawl_academies()
    
    print("\n📩 수집 완료! 이메일 리포트를 전송합니다...")
    send_email_report(crawled_data)