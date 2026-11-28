import re
from difflib import SequenceMatcher
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, unquote
from playwright.sync_api import sync_playwright
from notifier import send_email_report

# ==========================================
# 1. 학원별 상세 URL 및 정밀 DOM 셀렉터 설정
# ==========================================
TITLE_SIMILARITY_THRESHOLD = 0.82

EXCLUDE_CONTAINER_SELECTOR = (
    "header, nav, footer, .gnb, #gnb, .mexx-gnb, .side-panel, "
    "#globalMenu, .all_menu, [id^='mobile-pane'], .snt-mcta, "
    ".snt-footer-grid, .ft_companyInfo, .b_snsBtn, .page-sub-list, "
    "div.fixed[class*='translate-x-full']"
)

academies_info = {
    "시대인재(SdiJ)": {
        "notice": "https://www.sdij.com",
        "briefing": "https://www.sdij.com/aca/briefing/status/",
        "timetable1": "https://www.sdij.com/aca/schd/",
        "timetable2": "https://www.sdij.com/aca/schd/default.asp",
        "selectors": {
            "공지사항": "section.col-start-1 a, section.col-start-1 img, main section, main a, main img, div[class*='popup'] a, div[class*='popup'] img",
            "설명회": "section.mt-10 > div, .briefing-list li, .board_list tr",
            "시간표": "main img[src*='aca_sch'], main img, main div[class*='scrollbar-hide'] a, main div[class*='scrollbar-hide'] button, .timetable_list tr"
        },
        "selector": "main section, main a, main img, .notice_list li, .timetable_list tr",
        "check_iframes": True
    },
    "메가스터디 MEXX": {
        "notice": "https://mexx.megastudy.net/",
        "briefing": "https://mexx.megastudy.net/campus_common/2026/fair/list.asp",
        "timetable1": "https://mexx.megastudy.net/mexx/schedule/",
        "timetable2": "https://mexx.megastudy.net/mexx/schedule/?grd=2&sgrd=1",
        "selectors": {
            "공지사항": ".swiper-slide a, .mexx-hero__banner, .mexx-focus__card, .board_list tr",
            "설명회": "div.lecture-list a.lecture-card, .swiper-slide a",
            "시간표": ".mexx-schedule-waitbtn a, .mexx-tab__item, .pdf-canvas-gallery img, .swiper-slide a, .schedule_wrap"
        },
        "selector": ".swiper-slide a, .mexx-hero__banner, .mexx-focus__card, a.lecture-card"
    },
    "두각(Dugak)": {
        "notice": "https://www.dugak.net/info/notice_dt.do",
        "briefing": "https://www.dugak.net/pre/reserv.do",
        "timetable1": "https://www.dugak.net/time/?tid=51&mid=353",
        "timetable2": "https://www.dugak.net/time/?tid=29",
        "selectors": {
            "공지사항": "table tbody tr, .board-list tr, .notice_item",
            "설명회": ".class-list-section .class-item-row",
            "시간표": ".img-block img, .class-list-section .class-item-row, .timetable_area"
        },
        "selector": "table tbody tr, .class-list-section .class-item-row, .img-block img, .notice_item, .timetable_area"
    },
    "대찬학원": {
        "notice": "https://daechanedu.com",
        "briefing": "https://daechanedu.com/menu/?menu_str=0408",
        "timetable1": "https://daechanedu.com/menu/?menu_str=0304",
        "timetable2": "https://daechanedu.com/menu/?menu_str=0305",
        "selectors": {
            "공지사항": "#hd_pop > div, #hd_pop img, .new_main .intro_ment, .new_main .main_btn li, .main_banner img, table.board_table tbody tr",
            "설명회": "table.board_table2 tbody tr, table.board_table tbody tr",
            "시간표": ".sub_left_border .image img, .sub_left_border img[src*='/data/editor/'], .sub_left_border .sel_tab_menu a.on, table.board_table tbody tr"
        },
        "selector": "#hd_pop > div, .new_main .main_btn li, table.board_table2 tbody tr, table.board_table tbody tr, .sub_left_border .image img"
    },
    "SNT학원": {
        "notice": "https://snt-edu.com/",
        "briefing": "https://snt-edu.com/reserve",
        "timetable1": "https://snt-edu.com/timetable?grade=%EA%B3%A02&term=%EA%B8%B0%EB%A7%90%EA%B3%A0%EC%82%AC",
        "timetable2": "https://snt-edu.com/timetable?grade=%EA%B3%A01&term=%EA%B8%B0%EB%A7%90%EA%B3%A0%EC%82%AC",
        "selectors": {
            "공지사항": "main.snt-main ul.hs-brief-list li.hs-brief, main.snt-main section.dg-hero .hx-now, .notice_list tr",
            "설명회": "div.rv-picker-wrap a.rv-picker-item",
            "시간표": "main ul.tta-list li.tta-item, div.ct-terms a"
        },
        "selector": "ul.hs-brief-list li.hs-brief, div.rv-picker-wrap a.rv-picker-item, ul.tta-list li.tta-item"
    },
    "미래탐구(Mirae)": {
        "notice": "https://dh.mirae-academy.co.kr/customer/notice",
        "briefing": "https://dhres.mirae-academy.co.kr/front/reservation_cardType",
        "timetable1": "https://dh.mirae-academy.co.kr/study/schedule_booking?schyear_code=180",
        "timetable2": "https://dh.mirae-academy.co.kr/study/schedule_booking?schyear_code=190",
        "selectors": {
            "공지사항": "ul.board-list li.li_notice_list",
            "설명회": "ul.reserv_view li.reserv_box",
            "시간표": "div.schedule-reserve a.tab, div.schedule-reserve strong a, div.schedule-reserve img"
        },
        "selector": "ul.board-list li.li_notice_list, ul.reserv_view li.reserv_box, div.schedule-reserve a.tab"
    },
    "세정학원": {
        "notice": "https://sejungedu.com",
        "briefing": "https://sejungedu.com/explain/presentationplan?co=%EA%B3%A01",
        "timetable1": "https://sejungedu.com/schedule/timetable?co=%EA%B3%A03",
        "timetable2": "https://sejungedu.com/schedule/timetable?t=&co=%EA%B3%A02",
        "selectors": {
            "공지사항": "main section article, main ul.flex.flex-col > li, main a[href*='/seminars'], main a[href*='/notice']",
            "설명회": "main section ul > li",
            "시간표": "main section ul.divide-y > li"
        },
        "selector": "main section article, main section ul > li, main section ul.divide-y > li"
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
# 3. 정규화 및 유사도 기반 중복 병합(Deduplication) 헬퍼 함수
# ==========================================
def normalize_title(text):
    """
    핵심 내용(제목의 공백, 특수문자, 대괄호 태그 등을 제거한 정규화 텍스트) 추출
    예: '[공지사항] [고1] 2028 윈터스쿨 설명회 안내!' -> '2028윈터스쿨설명회안내'
    """
    if not text:
        return ""
    t = re.sub(r'https?://\S+', '', text)
    t = re.sub(r'\[.*?\]', '', t)
    t = re.sub(r'[^a-zA-Z0-9가-힣]', '', t)
    norm = t.lower().strip()
    
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
    t = re.sub(r'\[(📢|📅|📌)?[^\]]*(설명회|간담회|공지사항|공지|시간표|신규자료|안내)[^\]]*\]', '', t)
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


def calc_title_similarity(title_a, title_b, norm_a=None, norm_b=None):
    """
    difflib.SequenceMatcher를 사용해 두 공지 제목 간 유사도(0.0 ~ 1.0)를 계산
    """
    ratios = []
    if title_a and title_b:
        ratios.append(SequenceMatcher(None, title_a.lower(), title_b.lower()).ratio())
    if norm_a and norm_b:
        ratios.append(SequenceMatcher(None, norm_a, norm_b).ratio())
    return max(ratios) if ratios else 0.0


def merge_candidates(raw_candidates, similarity_threshold=TITLE_SIMILARITY_THRESHOLD):
    """
    정규화 키 일치 및 제목 유사도(82% 이상)를 기반으로 동일 학원 내 중복 후보를 1건으로 병합
    """
    merged_list = []

    for cand in raw_candidates:
        raw_text = cand["text"]
        norm_key = normalize_title(raw_text)
        if not norm_key or len(norm_key) < 2:
            continue

        canon_cat = canonical_category(cand["cat_name"])
        cand_url = cand["url"]
        candidate_display_title = clean_display_title(raw_text) or raw_text[:80]

        # 1. 정확 일치 우선 탐색, 없으면 유사도 >= threshold 탐색
        matched_item = None
        for existing in merged_list:
            if existing["norm_key"] == norm_key:
                matched_item = existing
                break

        if matched_item is None:
            best_ratio = 0.0
            for existing in merged_list:
                ratio = calc_title_similarity(
                    candidate_display_title, existing["title"],
                    norm_key, existing["norm_key"]
                )
                if ratio >= similarity_threshold and ratio > best_ratio:
                    best_ratio = ratio
                    matched_item = existing

        if matched_item is None:
            merged_list.append({
                "title": candidate_display_title,
                "norm_key": norm_key,
                "sources": [canon_cat],
                "url": cand_url,
                "primary_cat": canon_cat,
                "raw_text": raw_text
            })
        else:
            if canon_cat not in matched_item["sources"]:
                matched_item["sources"].append(canon_cat)

            curr_score = get_url_priority(matched_item["url"], matched_item["primary_cat"])
            new_score = get_url_priority(cand_url, canon_cat)
            if new_score > curr_score:
                matched_item["url"] = cand_url
                matched_item["primary_cat"] = canon_cat

            if len(candidate_display_title) > len(matched_item["title"]) and len(candidate_display_title) <= 90:
                matched_item["title"] = candidate_display_title

    return merged_list


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
    
    high_priority_keywords = [
        "reserv", "briefing", "explain", "fair", "booking", 
        "schedule", "view", "read", "detail", "tid=", "idx=", "menu_str="
    ]
    for kw in high_priority_keywords:
        if kw in url_lower:
            score += 25
            break
            
    if "설명회" in category:
        score += 20
    elif "시간표" in category:
        score += 10
    elif "공지사항" in category:
        score += 5
        
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
# 4. 정밀 DOM 추출 및 빈 alt 이미지 보조 검사 헬퍼
# ==========================================
def extract_meaningful_filename(src):
    """
    이미지 src URL에서 의미 있는 파일명 텍스트 추출 (UUID/해시/순수 숫자 등 노이즈 제외)
    """
    if not src or src.startswith("data:image"):
        return ""
    path = urlparse(src).path
    filename = unquote(path.rsplit("/", 1)[-1])
    stem = re.sub(r'\.[a-zA-Z0-9]{2,5}$', '', filename)
    # 앞부분에 붙은 UUID 또는 16진수 해시 접두사 제거
    stem = re.sub(r'^[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}[_+\-\s]*', '', stem)
    stem = re.sub(r'^[0-9a-fA-F]{16,}[_+\-\s]*', '', stem)
    cleaned = re.sub(r'[_+\-]+', ' ', stem).strip()

    # 무의미한 UUID/해시/숫자/단순 코드성 영문 제외
    if not cleaned:
        return ""
    if re.fullmatch(r'[0-9a-fA-F\s]{8,}', cleaned) or re.fullmatch(r'[\d\s]+', cleaned):
        return ""
    # 한글이 포함되었거나 2글자 이상 의미 있는 영문/숫자 키워드만 허용
    has_korean = bool(re.search(r'[가-힣]', cleaned))
    has_keyword = any(kw in cleaned.lower() for kw in ["2028", "winter", "notice", "fair", "brief", "schd"])
    if not (has_korean or has_keyword):
        return ""
    return cleaned


JS_EXTRACT_FRAME_DOM = """
({ selector, excludeSelector }) => {
    const isExcluded = (el) => {
        if (!el) return true;
        try {
            if (excludeSelector && el.closest(excludeSelector)) return true;
        } catch (e) {}
        return false;
    };

    const normalizeSpace = (s) => (s || '').replace(/\\s+/g, ' ').trim();

    // 1. 정밀 컨테이너 셀렉터 기반 요소 수집 (상/하위 중복 제거)
    let rawNodes = [];
    try {
        rawNodes = Array.from(document.querySelectorAll(selector));
    } catch (e) {
        rawNodes = [];
    }

    const validNodes = rawNodes.filter(el => !isExcluded(el));
    const validSet = new Set(validNodes);

    // 조상 노드가 이미 validSet에 포함되어 있다면 하위 중복 수집 방지
    const dedupedNodes = validNodes.filter(el => {
        let p = el.parentElement;
        while (p) {
            if (validSet.has(p)) return false;
            p = p.parentElement;
        }
        return true;
    });

    const textItems = [];
    for (const el of dedupedNodes) {
        const tag = el.tagName.toLowerCase();
        if (tag === 'img') continue; // 이미지는 별도 이미지 파이프라인에서 통합 검사

        const text = normalizeSpace(el.innerText);
        if (!text) continue;

        let href = '';
        if (tag === 'a') {
            href = el.getAttribute('href') || '';
        } else {
            const linkEl = el.querySelector('a[href]');
            href = linkEl ? (linkEl.getAttribute('href') || '') : '';
        }
        textItems.push({ text, href });
    }

    // 2. 주요 공지/배너/팝업 이미지 수집 (빈 alt 이미지 보조 검사 포함)
    const activeTabText = Array.from(
        document.querySelectorAll('.sel_tab_menu a.on, .subCategory li.on, .page-sub-list a.active, .ct-terms a.active')
    ).map(el => normalizeSpace(el.innerText)).filter(Boolean).join(' ');

    const ignoreImgPatterns = [
        'ico_', 'icon', 'btn_', 'logo', 'facebook', 'twitter', 'kakao',
        'blog', 'indicator', 'skplanet', 'getuid', '.svg'
    ];
    const ignoreClsPatterns = ['tta-photo', 'hs-teach-photo', 'teacher-image'];

    const imgNodes = Array.from(document.querySelectorAll('img')).filter(img => !isExcluded(img));
    const imgItems = [];

    for (const img of imgNodes) {
        const src = img.getAttribute('src') || '';
        if (!src || src.startsWith('data:image')) continue;

        const srcLower = src.toLowerCase();
        if (ignoreImgPatterns.some(pat => srcLower.includes(pat))) continue;

        const parentCls = (img.parentElement && typeof img.parentElement.className === 'string') ? img.parentElement.className : '';
        const imgCls = typeof img.className === 'string' ? img.className : '';
        if (ignoreClsPatterns.some(pat => parentCls.includes(pat) || imgCls.includes(pat))) continue;

        let alt = normalizeSpace(img.getAttribute('alt'));
        let title = normalizeSpace(img.getAttribute('title'));

        // 무의미한 플레이스홀더 alt 정리
        if (['alt', 'image', 'img', 'close', '닫기', '뒤로가기'].includes(alt.toLowerCase())) alt = '';
        if (['title', 'image', 'img'].includes(title.toLowerCase())) title = '';

        // 주변 컨테이너 텍스트 추출 (최대 3단계 상위까지 의미 있고 과도하게 길지 않은 텍스트)
        let contextText = '';
        let cur = img.parentElement;
        for (let depth = 0; depth < 3 && cur; depth++) {
            if (isExcluded(cur)) break;
            const ct = normalizeSpace(cur.innerText);
            if (ct && ct.length >= 4 && ct.length <= 180) {
                contextText = ct;
                break;
            }
            cur = cur.parentElement;
        }

        // 링크 추출
        let href = '';
        const anchor = img.closest('a[href]');
        if (anchor) {
            href = anchor.getAttribute('href') || '';
        } else if (cur) {
            const innerA = cur.querySelector('a[href]');
            if (innerA) href = innerA.getAttribute('href') || '';
        }

        imgItems.push({
            alt,
            title,
            src,
            contextText,
            activeTabText,
            href
        });
    }

    return {
        matchedCount: dedupedNodes.length + imgItems.length,
        textItems,
        imgItems
    };
}
"""


def collect_frame_candidates(frame, selector, target_url, cat_name):
    """
    단일 프레임(메인 페이지 또는 iframe)에서 정밀 텍스트 및 이미지 후보 추출
    반환: (candidates_list, matched_count, passed_count)
    """
    extracted = frame.evaluate(
        JS_EXTRACT_FRAME_DOM,
        {"selector": selector, "excludeSelector": EXCLUDE_CONTAINER_SELECTOR}
    )

    matched_count = extracted.get("matchedCount", 0)
    candidates = []
    passed_count = 0

    def resolve_link(href):
        if href and not href.startswith("javascript:") and not href.endswith("#"):
            return urljoin(target_url, href)
        return target_url

    # 1. 텍스트 요소 검사
    for item in extracted.get("textItems", []):
        text = item.get("text", "").strip()
        if not text:
            continue
        if is_recent(text):
            passed_count += 1
            candidates.append({
                "text": text,
                "url": resolve_link(item.get("href")),
                "cat_name": cat_name,
                "target_url": target_url
            })

    # 2. 이미지 요소 검사 (빈 alt 이미지 파일명 + 주변 컨테이너 텍스트 보조 검사 포함)
    for img in extracted.get("imgItems", []):
        alt_text = img.get("alt", "")
        title_text = img.get("title", "")
        combined_parts = [p for p in [alt_text, title_text] if p]

        if not combined_parts:
            # 빈 alt 이미지: 주변 컨테이너 텍스트 우선 검사 후, 부족할 때만 파일명 및 활성 탭 텍스트 보조 결합
            fname_text = extract_meaningful_filename(img.get("src", ""))
            ctx_text = img.get("contextText", "")
            tab_text = img.get("activeTabText", "")

            if ctx_text and is_recent(ctx_text):
                combined_text = ctx_text
            else:
                aux_parts = []
                if tab_text:
                    aux_parts.append(tab_text)
                if ctx_text:
                    aux_parts.append(ctx_text)
                if fname_text and fname_text not in ctx_text:
                    aux_parts.append(fname_text)
                combined_text = " ".join(aux_parts).strip()
        else:
            combined_text = " ".join(combined_parts).strip()

        if combined_text and is_recent(combined_text):
            passed_count += 1
            candidates.append({
                "text": combined_text,
                "url": resolve_link(img.get("href")),
                "cat_name": cat_name,
                "target_url": target_url
            })

    return candidates, matched_count, passed_count


# ==========================================
# 5. 크롤링 및 중복 병합 메인 로직
# ==========================================
def crawl_academies(return_health=False):
    results = {}
    health_checks = []
    
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
            cat_selectors = info.get("selectors", {})
            default_selector = info.get("selector", "main a, main li, table tr")

            for cat_name, target_url in targets:
                if not target_url or target_url.strip() == "":
                    continue
                
                print(f"👉 [{cat_name}] 수집 시도: {target_url}")
                canon_cat = canonical_category(cat_name)
                selector = cat_selectors.get(canon_cat, default_selector)

                health_entry = {
                    "academy": name,
                    "category": cat_name,
                    "url": target_url,
                    "success": False,
                    "matched_count": 0,
                    "passed_count": 0,
                    "error_message": None
                }

                try:
                    if "두각" in name:
                        resp = page.goto(target_url, wait_until="commit", timeout=30000)
                        page.wait_for_timeout(4000)
                    else:
                        resp = page.goto(target_url, wait_until="domcontentloaded", timeout=15000)
                        page.wait_for_timeout(1500)

                    if resp is not None and resp.status >= 400:
                        raise RuntimeError(f"HTTP 상태 오류 ({resp.status})")

                    # 시대인재 설명회 등 iframe 내부 콘텐츠가 있는 경우 프레임 로딩 추가 대기
                    if info.get("check_iframes") and "설명회" in cat_name:
                        try:
                            page.wait_for_selector("iframe", timeout=3000)
                            page.wait_for_timeout(1500)
                        except Exception:
                            pass

                    # 메인 프레임 수집
                    cands, m_cnt, p_cnt = collect_frame_candidates(page.main_frame, selector, target_url, cat_name)
                    raw_candidates.extend(cands)
                    total_matched = m_cnt
                    total_passed = p_cnt

                    # 추가 iframe 수집 (시대인재 등 iframe 기반 페이지 대응)
                    if info.get("check_iframes"):
                        for fr in page.frames:
                            if fr == page.main_frame:
                                continue
                            fr_url = fr.url or ""
                            if not fr_url.startswith("http") or "googletagmanager" in fr_url or "doubleclick" in fr_url:
                                continue
                            try:
                                fr_cands, fr_m, fr_p = collect_frame_candidates(fr, selector, target_url, cat_name)
                                raw_candidates.extend(fr_cands)
                                total_matched += fr_m
                                total_passed += fr_p
                            except Exception:
                                continue

                    health_entry["success"] = True
                    health_entry["matched_count"] = total_matched
                    health_entry["passed_count"] = total_passed
                    print(f"   📊 [Health] 매칭 요소: {total_matched}개 / 필터 통과: {total_passed}건")

                except Exception as e:
                    print(f"❌ [{name} - {cat_name}] 수집 중 에러: {e}")
                    has_error = True
                    health_entry["success"] = False
                    health_entry["error_message"] = str(e)[:120]

                health_checks.append(health_entry)

            # --- [정규화 및 유사도(82%+) 기반 중복 공지 1건 병합] ---
            merged_list = merge_candidates(raw_candidates)

            academy_updates = []
            category_order = ["공지사항", "설명회", "시간표"]

            for item in merged_list:
                item["sources"].sort(key=lambda s: category_order.index(s) if s in category_order else 99)
                item["source_badge"] = f"[{'+'.join(item['sources'])}]"
                item["section"] = classify_section(item["sources"], item["title"], item["raw_text"])
                academy_updates.append(item)

            if has_error and not academy_updates:
                results[name] = [{"error": True, "message": "⚠️ 접속 지연 또는 학원 웹사이트 구조 변경으로 수집 실패"}]
                print(f"⚠️ [{name}] 수집 실패 처리")
            else:
                results[name] = academy_updates
                print(f"✨ [{name}] 총 {len(raw_candidates)}개 후보 수집 -> 유사도/정규화 병합 후 {len(academy_updates)}건 정리 완료")
                for u in academy_updates:
                    print(f"   - {u['source_badge']} {u['title']} ({u['url']})")

        browser.close()
        
    if return_health:
        return results, health_checks
    return results


# ==========================================
# 6. 실행 진입점
# ==========================================
if __name__ == "__main__":
    print("🚀 주요 학원 통합 모니터링 크롤러를 실행합니다...")
    crawled_data, health_report = crawl_academies(return_health=True)
    
    print("\n📩 수집 완료! 이메일 리포트를 전송합니다...")
    send_email_report(crawled_data, health_checks=health_report)
