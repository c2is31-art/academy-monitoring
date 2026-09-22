import re
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
from datetime import datetime
from notifier import send_email_report

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

from datetime import datetime, timedelta

def is_recent(date_text, today_dt):
    """
    수집한 날짜 텍스트가 '오늘' 또는 '어제' (최근 2일 이내)인지 판별하는 함수
    23일 오후에 올라온 글도 24일 아침 8시 크롤링 시 정상적으로 수집됩니다.
    """
    if not date_text:
        return False
    
    # 숫자만 추출 (예: 20260923)
    clean_date = re.sub(r'[^0-9]', '', str(date_text))
    
    # 오늘 및 어제 날짜 생성
    yesterday_dt = today_dt - timedelta(days=1)
    
    valid_dates = [
        today_dt.strftime("%Y%m%d"),     # 20260924
        today_dt.strftime("%y%m%d"),     # 260924
        today_dt.strftime("%m%d"),       # 0924
        yesterday_dt.strftime("%Y%m%d"), # 20260923
        yesterday_dt.strftime("%y%m%d"), # 260923
        yesterday_dt.strftime("%m%d")    # 0923
    ]
    
    for v_date in valid_dates:
        if clean_date.endswith(v_date):
            return True
            
    return False

def crawl_all_academies():
    results = {}
    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    print(f"🚀 [{today_str}] 8개 경쟁학원 '오늘자 신규 게시글' 정밀 크롤링을 시작합니다...\n")

    # =============================================================
    # 1. 시대인재
    # =============================================================
    print("👉 시대인재 오늘자 신규 수집 중...")
    sdij_items = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # 시대인재 N재수 및 시간표
        for target_url in ["https://www.sdij.com/sdn/", "https://www.sdij.com/aca/schd/"]:
            try:
                page.goto(target_url, timeout=20000)
                page.wait_for_timeout(1500)
                
                rows = page.query_selector_all(".notice_item, .board_list tr, .tbl_schd tr, .schd_item")
                for r in rows:
                    text = r.inner_text().strip()
                    # 날짜 텍스트 또는 오늘 날짜 패턴 검색
                    if is_recent(text, now) or "NEW" in text.upper():
                        title_el = r.query_selector("a") or r
                        title = title_el.inner_text().strip().replace('\n', ' ')
                        href = title_el.get_attribute("href") or ""
                        link = href if href.startswith("http") else f"https://www.sdij.com{href}"
                        
                        if title and len(title) > 2:
                            sdij_items.append({"title": f"[신규] {title[:50]}", "link": link})
            except Exception as e:
                print(f"시대인재 수집 오류({target_url}): {e}")

        browser.close()
    results["시대인재"] = sdij_items

    # =============================================================
    # 2. 메가스터디 멕스 (MEXX)
    # =============================================================
    print("👉 메가스터디 멕스 오늘자 신규 수집 중...")
    mexx_items = []
    grd_list = [1, 2, 3, 4, 6]
    for grd in grd_list:
        try:
            url = f"https://mexx.megastudy.net/mexx/schedule/?grd={grd}&sgrd=1"
            res = requests.get(url, headers=HEADERS, timeout=10)
            soup = BeautifulSoup(res.text, "html.parser")
            
            rows = soup.select(".tbl_schedule tr, .sch_list li, .board_list tbody tr")
            for r in rows:
                date_el = r.select_one(".date, td.date, .regdate")
                date_text = date_el.get_text(strip=True) if date_el else r.get_text()
                
                # 오늘 날짜에 해당하거나 NEW 아이콘이 있는 경우만 필터링
                if is_recent(date_text, now) or r.select_one(".icon_new, .new"):
                    title_el = r.select_one("a, .title, .subject")
                    if title_el:
                        title = title_el.get_text(strip=True)
                        link = title_el.get("href", "")
                        full_link = link if link.startswith("http") else f"https://mexx.megastudy.net{link}"
                        mexx_items.append({"title": f"[Grd{grd}] {title[:50]}", "link": full_link})
        except Exception as e:
            print(f"메가스터디 멕스(grd={grd}) 오류: {e}")

    results["메가스터디 멕스"] = mexx_items

    # =============================================================
    # 3. 두각 학원
    # =============================================================
    print("👉 두각 학원 오늘자 신규 수집 중...")
    dugak_items = []
    dugak_urls = [
        "https://www.dugak.net/time/?tid=51",
        "https://www.dugak.net/time/?tid=29",
        "https://www.dugak.net/time/?tid=26&mid=370",
        "https://www.dugak.net/time/?tid=43&mid=369"
    ]
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        for d_url in dugak_urls:
            try:
                page.goto(d_url, timeout=20000)
                page.wait_for_timeout(1500)
                
                cards = page.query_selector_all(".time_list li, .time_box, .lecture_item, .tbl_list tr")
                for c in cards:
                    text = c.inner_text().strip()
                    # 오늘 작성 날짜 체크
                    if is_recent(text, now) or "NEW" in text.upper():
                        clean_title = text.replace('\n', ' ')
                        dugak_items.append({"title": clean_title[:60], "link": d_url})
            except Exception as e:
                print(f"두각 수집 오류({d_url}): {e}")

        browser.close()

    results["두각 학원"] = dugak_items

    # =============================================================
    # 4. 대찬 학원
    # =============================================================
    print("👉 대찬 학원 오늘자 신규 수집 중...")
    daechan_items = []
    try:
        res = requests.get("https://daechanedu.com/", headers=HEADERS, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")
        for r in soup.select("tr, .latest_list li, .board_item"):
            text = r.get_text()
            if is_recent(text, now) or r.select_one(".new, .icon_new"):
                a = r.select_one("a")
                if a:
                    title = a.get_text(strip=True)
                    link = a['href'] if a['href'].startswith("http") else f"https://daechanedu.com/{a['href']}"
                    daechan_items.append({"title": title[:60], "link": link})
    except Exception as e:
        print(f"대찬학원 오류: {e}")
    results["대찬 학원"] = daechan_items

    # =============================================================
    # 5. S&T 학원
    # =============================================================
    print("👉 S&T 학원 오늘자 신규 수집 중...")
    snt_items = []
    try:
        res = requests.get("https://www.sntedu.co.kr/", headers=HEADERS, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")
        for r in soup.select("tr, .notice_area li, .board_list tr"):
            text = r.get_text()
            if is_recent(text, now) or r.select_one(".new"):
                a = r.select_one("a")
                if a:
                    title = a.get_text(strip=True)
                    link = a['href'] if a['href'].startswith("http") else f"https://www.sntedu.co.kr/{a['href']}"
                    snt_items.append({"title": title[:60], "link": link})
    except Exception as e:
        print(f"S&T학원 오류: {e}")
    results["S&T 학원"] = snt_items

    # =============================================================
    # 6. 미래탐구 (대치)
    # =============================================================
    print("👉 미래탐구 대치 오늘자 신규 수집 중...")
    mirae_items = []
    try:
        res = requests.get("https://dh.mirae-academy.co.kr/", headers=HEADERS, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")
        for r in soup.select("tr, .notice_list li"):
            text = r.get_text()
            if is_recent(text, now) or r.select_one(".new"):
                a = r.select_one("a")
                if a:
                    title = a.get_text(strip=True)
                    link = a['href'] if a['href'].startswith("http") else f"https://dh.mirae-academy.co.kr/{a['href']}"
                    mirae_items.append({"title": title[:60], "link": link})
    except Exception as e:
        print(f"미래탐구 오류: {e}")
    results["미래탐구 (대치)"] = mirae_items

    # =============================================================
    # 7. 세정 학원
    # =============================================================
    print("👉 세정 학원 오늘자 신규 수집 중...")
    sejung_items = []
    try:
        res = requests.get("https://sejungedu.com/", headers=HEADERS, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")
        for r in soup.select("tr, .latest_item, .board_list tr"):
            text = r.get_text()
            if is_recent(text, now) or r.select_one(".new"):
                a = r.select_one("a")
                if a:
                    title = a.get_text(strip=True)
                    link = a['href'] if a['href'].startswith("http") else f"https://sejungedu.com/{a['href']}"
                    sejung_items.append({"title": title[:60], "link": link})
    except Exception as e:
        print(f"세정학원 오류: {e}")
    results["세정 학원"] = sejung_items

    # =============================================================
    # 최종 수집 결과 이메일 발송
    # =============================================================
    print("\n📧 오늘자 신규 리포트 이메일 발송 중...")
    send_email_report(results)

if __name__ == "__main__":
    crawl_all_academies()