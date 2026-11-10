"""
국어 맞춤법, 오탈자 검수 및 긴 문장 간략화 제안 모듈 (core/text_improver.py)
- PDF 텍스트 줄바꿈 자동 복원 (불릿/단락 병합)
- 입시/학원 전문용어 화이트리스트 필터링 (오탐 방지)
- 국어 맞춤법, 오탈자, 띄어쓰기, 폰트 깨짐 기호 검출 및 교정 제안
- 긴 문장(기본 55자 이상, 40~100자 가변) 감지 및 2가지 축약안 생성:
    1) [제안 1: 간결한 서술형 문장]
    2) [제안 2: 브로슈어 개조식 명사형 종결 요약]
- 선택적 Gemini API 연동 (API 키 제공 시 AI 카피라이팅, 미제공 시 내장 규칙 기반 100% 작동)
"""

from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, List, Optional, Tuple

# -------------------------------------------------------------------------
# 1. 입시/학원 전문용어 화이트리스트 (오탐 방지)
# -------------------------------------------------------------------------
ACADEMIC_WHITELIST = {
    # 학원/시스템
    "바자관", "시수", "실모", "기출", "학평", "모평", "킬러", "준킬러", "노베", "유베",
    "선지", "주간지", "클리닉", "단과", "재종", "윈터스쿨", "썸머스쿨", "러셀", "메가스터디",
    "시대인재", "대치", "서초", "강남", "중계", "영통", "평촌", "분당", "부천", "광주",
    # 과목/교재
    "공수1", "공수2", "미적1", "미적2", "미적", "확통", "수1", "수2", "기하", "수리논술",
    "인문논술", "과학논술", "논구술", "통사", "통과", "물1", "화1", "생1", "지1", "물2",
    "화2", "생2", "지2", "수특", "수완", "자이스토리", "마더텅", "블랙라벨", "일품",
    "쎈", "시발점", "뉴런", "드릴", "수분감", "킬캠", "서바이벌", "강대모고", "더프",
    "이투스", "대성", "이감", "상상", "바탕", "한수", "백설", "명월", "수국", "동백",
    # 고등학교 이름
    "휘문고", "중동고", "단대부고", "한대부고", "중산고", "경기고", "숙명여고", "은광여고",
    "진선여고", "경기여고", "현대고", "세화고", "세화여고", "상산고", "하나고", "외대부고",
    "대원외고", "한영외고", "서울고", "서초고", "반포고", "개포고", "보성고", "동북고",
    "영동고", "압구정고", "청담고", "배재고", "보인고", "선덕고", "대진고", "낙생고",
    # 단위 및 표기
    "6T", "3T", "4T", "8T", "12T", "9T", "2T", "1T", "T", "회차", "주차",
}

# -------------------------------------------------------------------------
# 2. 내장 맞춤법 / 오탈자 교정 사전 및 패턴
# -------------------------------------------------------------------------
SPELLING_CORRECTION_RULES: List[Dict[str, Any]] = [
    {
        "pattern": r"\b맡게\b",
        "replacement": "맞게",
        "rule_type": "맞춤법 오류",
        "reason": "'수준이나 기준에 알맞게'의 의미일 때는 '맞게'가 올바른 표현입니다.",
    },
    {
        "pattern": r"안내문제\b",
        "replacement": "안내문에",
        "rule_type": "오탈자",
        "reason": "'안내문'에 조사 '에'가 결합된 형태('안내문에')의 오탈자입니다.",
    },
    {
        "pattern": r"없에(다|고|서|면|서도)",
        "replacement": r"없애\1",
        "rule_type": "맞춤법 오류",
        "reason": "'없애다'가 표준어이며 '없에다'는 잘못된 표기입니다.",
    },
    {
        "pattern": r"할수\s*있",
        "replacement": "할 수 있",
        "rule_type": "띄어쓰기",
        "reason": "의존명사 '수'는 앞말과 띄어 써야 합니다.",
    },
    {
        "pattern": r"될수\s*있",
        "replacement": "될 수 있",
        "rule_type": "띄어쓰기",
        "reason": "의존명사 '수'는 앞말과 띄어 써야 합니다.",
    },
    {
        "pattern": r"볼수\s*있",
        "replacement": "볼 수 있",
        "rule_type": "띄어쓰기",
        "reason": "의존명사 '수'는 앞말과 띄어 써야 합니다.",
    },
    {
        "pattern": r"알수\s*있",
        "replacement": "알 수 있",
        "rule_type": "띄어쓰기",
        "reason": "의존명사 '수'는 앞말과 띄어 써야 합니다.",
    },
    {
        "pattern": r"뿐만아니라",
        "replacement": "뿐만 아니라",
        "rule_type": "띄어쓰기",
        "reason": "'뿐만 아니라'는 띄어 씀이 원칙입니다.",
    },
    {
        "pattern": r"\b둘다\b",
        "replacement": "둘 다",
        "rule_type": "띄어쓰기",
        "reason": "부사 '다'는 앞말과 띄어 써야 합니다.",
    },
    {
        "pattern": r"\b강의\s*입니다\b",
        "replacement": "강의입니다",
        "rule_type": "맞춤법/조사",
        "reason": "서술격 조사 '이다/입니다'는 앞 체언에 붙여 써야 합니다.",
    },
    {
        "pattern": r"\b학생\s*입니다\b",
        "replacement": "학생입니다",
        "rule_type": "맞춤법/조사",
        "reason": "서술격 조사는 앞말에 붙여 씁니다.",
    },
    {
        "pattern": r"\b수업\s*입니다\b",
        "replacement": "수업입니다",
        "rule_type": "맞춤법/조사",
        "reason": "서술격 조사는 앞말에 붙여 씁니다.",
    },
    {
        "pattern": r"\b어떻해\b",
        "replacement": "어떻게",
        "rule_type": "맞춤법 오류",
        "reason": "'어떻게 해'의 줄임말이 아닐 때는 '어떻게'가 맞습니다.",
    },
    {
        "pattern": r"\b몇일\b",
        "replacement": "며칠",
        "rule_type": "맞춤법 오류",
        "reason": "'몇 일'은 비표준어이며 '며칠'이 표준어입니다.",
    },
    {
        "pattern": r"\b오랫만에\b",
        "replacement": "오랜만에",
        "rule_type": "맞춤법 오류",
        "reason": "'오래간만에'의 준말은 '오랜만에'입니다.",
    },
    {
        "pattern": r"\b바램\b",
        "replacement": "바람",
        "rule_type": "맞춤법 오류",
        "reason": "'바라다'에서 파생된 명사는 '바람'이 올바릅니다.",
    },
    {
        "pattern": r"\b제작년\b",
        "replacement": "재작년",
        "rule_type": "맞춤법 오류",
        "reason": "지난해의 바로 전 해는 '재작년'입니다.",
    },
    {
        "pattern": r"\b왠만하면\b",
        "replacement": "웬만하면",
        "rule_type": "맞춤법 오류",
        "reason": "'우연한 일/어찌 된'의 뜻이 아닐 때는 '웬만하면'이 맞습니다.",
    },
    {
        "pattern": r"\b어의없",
        "replacement": "어이없",
        "rule_type": "맞춤법 오류",
        "reason": "'어이없다'가 표준어입니다.",
    },
    {
        "pattern": r"\bPima\b",
        "replacement": "Prima",
        "rule_type": "영문 교재 오탈자",
        "reason": "라틴어 'Prima'의 오탈자 표기로 추정됩니다.",
    },
    {
        "pattern": r"NE특급기술",
        "replacement": "NE특급기출",
        "rule_type": "교재명 오탈자",
        "reason": "능률(NE능률) 특급기출 교재명의 오탈자입니다.",
    },
    {
        "pattern": r"(\d+)분안에",
        "replacement": r"\1분 안에",
        "rule_type": "띄어쓰기",
        "reason": "'안에'는 앞 단어와 띄어 써야 합니다.",
    },
    {
        "pattern": r"\b([가-힣]{1,4})\s+\1\b",
        "replacement": r"\1",
        "rule_type": "단어 중복",
        "reason": "동일한 단어/조사가 불필요하게 연속 중복 기재되었습니다.",
    },
]

# -------------------------------------------------------------------------
# 3. PDF 깨진 줄바꿈 복원 및 문장 분리 로직
# -------------------------------------------------------------------------
# 허용된 문장부호: ■, -, [ ]
ALLOWED_PUNCTUATION = {"■", "-", "[", "]"}

# 비허용 특수 문장부호 (■, -, [ ] 제외) 탐지 정규식
DISALLOWED_PUNCTUATION_RE = re.compile(
    r"([▶▷▲▼◀◁►◄●○◎◐◑◆◇◈▣□▪▫★☆✦✧♥♡※☞👉➢➤➔➜→➡✓✔☑•*|｜＆]"
    r"|[〈〉《》〔〕【】『』「」]"
    r"|(?<![월화수목금토일])·(?![월화수목금토일])"
    r"|(?<!https)(?<!http)(?<!\d):(?!//)(?!\d))"
)

# 비표준 괄호 매핑 (대괄호 [ ]로 변환)
BRACKET_REPLACE_RULES = [
    (re.compile(r"[〈《〔【『「]\s*(.*?)\s*[〉》〕】』」]"), r"[\1]"),
    (re.compile(r"[〈《〔【『「]"), "["),
    (re.compile(r"[〉》〕】』」]"), "]"),
]

# 비표준 글머리 기호 (■, - 이외의 기호) 탐지 정규식
NON_STANDARD_BULLET_RE = re.compile(
    r"^(\s*)([●○◆◇▶▷★☆•·▪▫✓✔☞➢➤※~→➔➜➡:]+|\+(?=\s|[가-힣a-zA-Z]))\s*(.*)$"
)
# ■ 기호 뒤 공백 누락 탐지 정규식 (예: ■수업특징 -> ■ 수업특징)
SQUARE_NO_SPACE_RE = re.compile(r"^(\s*)■([^\s].*)$")

# 표준 ■ 글머리 기호 (이미 정상: ■ 뒤 공백 1칸 이상)
STANDARD_SQUARE_RE = re.compile(r"^(\s*)■\s+(.*)$")


def sanitize_punctuations_in_text(sent: str) -> str:
    """
    본문 텍스트 내에서 규정된 3가지 문장부호 (■, -, [ ]) 이외의
    비허용 부호들을 규정 부호 또는 자연스러운 연결 부호로 교정합니다.
    """
    if not sent:
        return ""
    res = sent.strip()

    # 1. 단락 맨 앞머리 비표준 기호(▶, →, :, ●, ○, ◆, ◇ 등)를 표준 ■ 기호로 우선 변환
    m_head = re.match(r"^([●○◆◇▶▷★☆•·▪▫✓✔☞➢➤※\+\~*→➔➜➡:]+)\s*(.*)$", res)
    if m_head:
        res = f"■ {m_head.group(2).strip()}"

    # 2. 비표준 괄호 -> [ ] 대괄호로 변환
    for pat, rep in BRACKET_REPLACE_RULES:
        res = pat.sub(rep, res)

    # 3. 전각 ＆ -> 및
    res = res.replace("＆", " 및 ").replace(" & ", " 및 ")

    # 4. 파이프 기호 | -> -
    res = res.replace("｜", " - ").replace(" | ", " - ")

    # 5. 본문 중간 화살표 ▶, ▷, → 등 연결 -> -
    res = re.sub(r"\s*[▶▷→➔➜➡]+\s*", " - ", res)

    # 6. 본문 중간 콜론(:) (시간 09:00, URL http://, 1:1 제외) -> -
    res = re.sub(r"\s*(?<!https)(?<!http)(?<!\d):(?!//)(?!\d)\s*", " - ", res)

    # 7. 불필요한 장식 기호(★, ☆, ♥, ♡, ※, ●, ○, ◆, ◇, •, ·, *) 제거 또는 정리
    res = re.sub(r"[★☆♥♡]", "", res)
    res = re.sub(r"\s*[•∙]\s*", " - ", res)
    res = re.sub(r"(?<![월화수목금토일])·(?![월화수목금토일])", ", ", res)

    # ■ 바로 뒤 공백 확보
    res = re.sub(r"^■([^\s])", r"■ \1", res)
    # - 바로 뒤 공백 확보
    res = re.sub(r"^-([^\s\d])", r"- \1", res)

    # 다중 공백 정리
    res = re.sub(r"[ \t]+", " ", res).strip()
    return res


def clean_pdf_broken_glyphs(text: str) -> str:
    """
    PDF 텍스트 내 글리프 깨짐, 잡음 기호 정제.
    원본 글머리 기호(■, ▶, ●, ★ 등)는 기호 통일 검수를 위해 보존합니다.
    """
    if not text:
        return ""
    s = text
    # 엑셀/XML 비허용 제어문자 제거
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)
    # 불필요한 깨진 기호만 정제 (글머리 기호 보존)
    s = re.sub(r"[§¡¿|]{2,}", " ", s)
    return s


def correct_all_spelling_in_sentence(text: str) -> str:
    """
    문장 내 등록된 맞춤법/오탈자 패턴을 일괄 교정하여 반환합니다.
    (입시/학원 전문용어 화이트리스트는 오탐 배제)
    """
    if not text:
        return ""
    res = text
    for rule in SPELLING_CORRECTION_RULES:
        pat = rule["pattern"]
        rep = rule["replacement"]
        matches = list(re.finditer(pat, res))
        if not matches:
            continue
        for m in reversed(matches):
            matched_str = m.group(0)
            if matched_str in ACADEMIC_WHITELIST:
                continue
            try:
                replacement_str = re.sub(pat, rep, matched_str)
                res = res[:m.start()] + replacement_str + res[m.end():]
            except Exception:
                pass
    res = re.sub(r"[§¡¿|]{2,}", " ", res)
    return res


def restore_pdf_sentences(raw_text: str) -> List[Dict[str, Any]]:
    """
    PDF에서 추출된 줄바꿈 분절 텍스트를 문맥, 불릿 기호, 종결 어미를 고려하여
    온전한 단락 및 문장 단위로 자동 결합합니다.
    """
    if not raw_text:
        return []

    lines = raw_text.splitlines()
    merged_blocks: List[str] = []
    curr_block: List[str] = []

    # 불릿 기호 정규식
    bullet_regex = re.compile(
        r"^(?:[•\-*·▶▷■□★☆○●◆◇→➔➜➡:]|(?:\d+|[가-하]|[A-Z])[\.\)])\s*"
    )
    # 문장 종결 표현 정규식 (. ! ? 다. 요. 함. 음. 등)
    ending_regex = re.compile(
        r"(?:[\.\!\?]|다\.|요\.|함\.|음\.|합니다|됩니다|있습니다|제공|배부|진행|완성|정리|체화|도약|정립|안내)$"
    )

    for line in lines:
        line_str = line.strip()
        if not line_str:
            if curr_block:
                merged_blocks.append(" ".join(curr_block))
                curr_block = []
            continue

        # 섹션 헤더성 키워드 감지 시 블록 분리
        if line_str in ("강의특징 / 관리", "강의특징", "관리프로그램", "수강대상", "수업특징", "보강방법", "주차별 진도계획"):
            if curr_block:
                merged_blocks.append(" ".join(curr_block))
                curr_block = []
            continue

        is_bullet = bool(bullet_regex.match(line_str))

        if is_bullet:
            if curr_block:
                merged_blocks.append(" ".join(curr_block))
                curr_block = []
            curr_block.append(line_str)
        else:
            if not curr_block:
                curr_block.append(line_str)
            else:
                last_line = curr_block[-1]
                # 앞 라인이 종결 부호나 완전한 어미로 끝난 경우 새 블록 고려
                if ending_regex.search(last_line) and len(last_line) >= 20:
                    merged_blocks.append(" ".join(curr_block))
                    curr_block = [line_str]
                else:
                    # 중간 줄바꿈으로 판단하여 공백 하나로 이어붙임
                    curr_block.append(line_str)

    if curr_block:
        merged_blocks.append(" ".join(curr_block))

    # 이제 각 블록을 온전한 문장 목록으로 정제
    results: List[Dict[str, Any]] = []
    for blk in merged_blocks:
        blk_clean = re.sub(r"\s+", " ", blk).strip()
        if not blk_clean or len(blk_clean) < 4:
            continue

        # 섹션 라벨 추정
        sec = "본문"
        if any(k in blk_clean for k in ("관리", "클리닉", "과제", "피드백", "테스트", "오답노트")):
            sec = "관리프로그램"
        elif any(k in blk_clean for k in ("특징", "개념", "수능", "내신", "기출", "유형", "독해")):
            sec = "강의특징"
        elif any(k in blk_clean for k in ("교재", "본교재", "주간지", "모의고사")):
            sec = "교재/자료"
        elif any(k in blk_clean for k in ("보강", "영상", "결석")):
            sec = "보강안내"

        results.append({
            "section": sec,
            "sentence": blk_clean,
            "char_count": len(blk_clean),
        })

    return results


# -------------------------------------------------------------------------
# 4. 국어 맞춤법 및 오탈자 검사 함수
# -------------------------------------------------------------------------
def check_spelling_and_typos(sentence: str) -> List[Dict[str, Any]]:
    """
    한 문장에서 맞춤법 오류, 오탈자, 띄어쓰기, 폰트 깨짐 등을 검출하고
    교정 제안을 생성합니다. (입시 화이트리스트 용어는 오탐 배제)
    """
    errors: List[Dict[str, Any]] = []
    if not sentence:
        return errors

    # 1) 내장 정규식 규칙 대조
    for rule in SPELLING_CORRECTION_RULES:
        pat = rule["pattern"]
        matches = list(re.finditer(pat, sentence))
        for m in matches:
            matched_str = m.group(0)

            # 화이트리스트에 포함된 경우 건너뜀
            if matched_str in ACADEMIC_WHITELIST:
                continue

            # 치환 제안 생성
            rep = re.sub(pat, rule["replacement"], matched_str)
            suggested_sent = sentence[:m.start()] + rep + sentence[m.end():]

            errors.append({
                "rule_type": rule["rule_type"],
                "target_word": matched_str,
                "replacement_word": rep,
                "reason": rule["reason"],
                "original_sentence": sentence,
                "suggested_sentence": suggested_sent,
            })

    # 2) 폰트 깨짐 특수문자 탐지
    glyph_matches = list(re.finditer(r"[§¡¿|]{1,}", sentence))
    for gm in glyph_matches:
        g_str = gm.group(0)
        # 단, 구분자로 쓰인 | 한 글자가 단어 사이에 있는 경우는 제외하고 특수문자만 단독/연속된 경우
        if len(g_str) >= 2 or g_str in ("§", "¡", "¿"):
            rep_clean = sentence[:gm.start()] + sentence[gm.end():]
            rep_clean = re.sub(r"\s+", " ", rep_clean).strip()
            errors.append({
                "rule_type": "폰트 글리프 깨짐",
                "target_word": g_str,
                "replacement_word": "(삭제)",
                "reason": "PDF 폰트 변환 시 비정상적으로 유입된 특수문자 글리프입니다.",
                "original_sentence": sentence,
                "suggested_sentence": rep_clean,
            })

    return errors


# -------------------------------------------------------------------------
# 5. 긴 문장 간략화 제안 엔진 (규칙 기반)
# -------------------------------------------------------------------------
def simplify_sentence_rule_based(sentence: str) -> Tuple[str, str]:
    """
    긴 문장에 대해 2가지 스타일의 축약안을 규칙 기반으로 생성합니다:
    1) 제안 1: 간결한 서술형 문장 (불필요한 피동/중첩/군더더기 어미 축소)
    2) 제안 2: 브로슈어 개조식 명사형 종결 요약 (핵심 불릿 포인트)
    """
    s = sentence.strip()
    # 선행 불릿 제거 후 처리
    bullet_prefix = ""
    m_bullet = re.match(r"^([•\-*·▶▷■□★☆○●◆◇]|\d+[\.\)])\s*", s)
    if m_bullet:
        bullet_prefix = m_bullet.group(0).strip() + " "
        s_body = s[len(m_bullet.group(0)):].strip()
    else:
        s_body = s

    # -----------------------------------------------------------------
    # [제안 1: 간결한 서술형 문장] 변환 규칙
    # -----------------------------------------------------------------
    prop1 = s_body

    # 군더더기 서술어 간결화 치환 테이블
    verbose_patterns = [
        (r"진행함으로써\s*", "진행하여 "),
        (r"진행하는\s*것이\s*중요하다고\s*볼\s*수\s*있습니다\.", "진행이 핵심입니다."),
        (r"진행할\s*수\s*있도록\s*(지도|교육|안내)합니다\.", "집중 지도합니다."),
        (r"체화할\s*수\s*있도록\s*합니다\.", "확실히 체화합니다."),
        (r"극대화할\s*수\s*있도록\s*합니다\.", "극대화합니다."),
        (r"도모할\s*수\s*있도록\s*합니다\.", "완벽히 완성합니다."),
        (r"이루어질\s*수\s*있도록\s*합니다\.", "체계적으로 진행합니다."),
        (r"기틀을\s*마련할\s*수\s*있도록\s*최선을\s*다해\s*지도합니다\.", "기틀을 확실히 완성합니다."),
        (r"~에\s*대하여\s*", ""),
        (r"~을\s*바탕으로\s*하여\s*", "기반으로 "),
        (r"효율성을\s*극대화합니다\.", "효율을 극대화합니다."),
        (r"자신감을\s*갖고\s*풀어낼\s*수\s*있는\s*명확한\s*문제\s*풀이\s*방법\s*습득", "자신감 있는 실전 문제 풀이법 체화"),
        (r"놓치는\s*부분\s*없이\s*이루어지는\s*국어\s*학습", "빈틈없는 1:1 맞춤형 국어 학습"),
        (r"최상위권으로\s*도약하게\s*만드는\s*가장\s*정확한\s*선택,\s*", "최상위권 도약을 위한 "),
        (r"반드시\s*되게\s*하는\s*본질적인\s*실력\s*향상의\s*국어\s*강의가\s*러셀\s*대치에\s*있습니다\.", "실력 향상을 입증하는 러셀 대치 대표 강의."),
    ]

    for pat, rep in verbose_patterns:
        prop1 = re.sub(pat, rep, prop1)

    prop1 = re.sub(r"\s+", " ", prop1).strip()
    if bullet_prefix:
        prop1 = bullet_prefix + prop1

    # -----------------------------------------------------------------
    # [제안 2: 브로슈어 개조식 명사형 종결 요약] 변환 규칙
    # -----------------------------------------------------------------
    prop2 = s_body

    # 개조식 명사 종결 패턴 치환
    noun_ending_patterns = [
        (r"진행합니다\.|진행함\.|진행\.", "진행"),
        (r"배부합니다\.|배부함\.|배부\.", "배부"),
        (r"제공합니다\.|제공함\.|제공\.", "제공"),
        (r"지도합니다\.|지도함\.", "집중 지도"),
        (r"학습합니다\.|학습함\.", "학습"),
        (r"체화합니다\.|체화함\.", "완벽 체화"),
        (r"극대화합니다\.|극대화함\.", "극대화"),
        (r"대비합니다\.|대비함\.", "철저 대비"),
        (r"완성합니다\.|완성함\.", "완성"),
        (r"정리합니다\.|정리함\.", "총정리"),
        (r"선점합니다\.|선점함\.", "선점"),
        (r"습득합니다\.|습득함\.", "체화"),
        (r"이루어집니다\.|이루어짐\.", "실시"),
        (r"관리합니다\.|관리함\.", "밀착 관리"),
        (r"확립합니다\.|확립함\.", "확립"),
    ]

    for pat, rep in noun_ending_patterns:
        prop2 = re.sub(pat, rep, prop2)

    # 연결 어미 및 불필요한 수식어 축약
    prop2 = re.sub(r"으로\s*인한\s*", "별 ", prop2)
    prop2 = re.sub(r"을\s*통한\s*|를\s*통한\s*", " 기반 ", prop2)
    prop2 = re.sub(r"에\s*맞춘\s*", " 맞춤 ", prop2)
    prop2 = re.sub(r"수강생\s*대상\s*무료\s*", "수강생 무료 ", prop2)
    prop2 = re.sub(r"놓치는\s*부분\s*없이\s*", "밀착 ", prop2)
    prop2 = re.sub(r"\s+", " ", prop2).strip()

    # 불릿 기호 '• '로 통일
    prop2 = f"• {prop2}"

    return prop1, prop2


# -------------------------------------------------------------------------
# 6. Gemini API를 활용한 고급 카피라이팅 간략화 (선택적)
# -------------------------------------------------------------------------
IMPROVER_MODEL_CANDIDATES = [
    "gemini-3.8-flash",
    "gemini-3.8-pro",
    "gemini-2.5-flash",
    "gemini-1.5-flash",
]

def simplify_sentence_with_gemini(
    sentence: str,
    api_key: str,
    model_name: Optional[str] = None,
) -> Optional[Tuple[str, str]]:
    """
    Gemini 3.8 최우선 API를 호출하여 대치동 브로슈어 스타일 카피라이팅 축약안을 생성합니다.
    - 제안 1: 군더더기 없는 간결한 서술형 문장 (~합니다, ~완성)
    - 제안 2: 학원 브로슈어 개조식 명사 종결 요약 (• 키워드 중심)
    실패 시 None을 반환하여 내장 규칙 기반으로 안전하게 fallback합니다.
    """
    if not api_key:
        return None

    prompt = f"""
당신은 대한민국 교육 특구 대치동 최상위권 단과 학원의 전문 카피라이터이자 국어 교열 전문가입니다.
아래 강의계획서 원문 문장은 길이가 너무 길어 학부모와 학생이 한눈에 파악하기 어렵습니다.
원문의 핵심 교육 가치와 메시지를 100% 보존하면서, 아래 2가지 형식으로 축약해 주세요.

[원문 문장]
"{sentence}"

[요구사항]
1. 제안 1 (간결한 서술형): 군더더기, 불필요한 피동형/중첩 어미를 없애고 35~45자 내외의 명료하고 자신감 넘치는 서술형 문장으로 축약.
2. 제안 2 (브로슈어 개조식): 학원 팜플렛/리플렛에 들어갈 불릿 포인트(• ) 형태로, 핵심 키워드 위주의 명사형 종결(예: ~제공, ~완벽 체화, ~1등급 완성) 20~30자 내외 요약.

반드시 아래 JSON 형식으로만 응답해 주세요. (마크다운 코드블록 없이 순수 JSON만 출력)
{{"proposal_1": "제안 1 서술형 문장", "proposal_2": "• 제안 2 개조식 요약"}}
"""

    models_to_try = [model_name] if model_name else IMPROVER_MODEL_CANDIDATES

    payload = {
        "contents": [
            {
                "parts": [{"text": prompt}]
            }
        ],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 300,
        }
    }

    req_data = json.dumps(payload).encode("utf-8")

    for cand_model in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{cand_model}:generateContent?key={api_key}"
        try:
            req = urllib.request.Request(
                url,
                data=req_data,
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=8) as response:
                res_body = response.read().decode("utf-8")
                res_json = json.loads(res_body)

            candidates = res_json.get("candidates", [])
            if candidates:
                cand = candidates[0]
                parts = cand.get("content", {}).get("parts", [])
                if parts:
                    text_out = parts[0].get("text", "").strip()
                    text_clean = re.sub(r"^```(?:json)?\s*", "", text_out)
                    text_clean = re.sub(r"\s*```$", "", text_clean).strip()
                    data = json.loads(text_clean)
                    p1 = data.get("proposal_1", "").strip()
                    p2 = data.get("proposal_2", "").strip()
                    if p1 and p2:
                        return p1, p2
        except Exception:
            continue

    return None


# -------------------------------------------------------------------------
# 7. 강좌 단위 종합 텍스트 분석 메인 함수
# -------------------------------------------------------------------------
def analyze_course_text_content(
    course_name: str,
    instructor: str,
    raw_texts: List[str],
    max_char_threshold: int = 55,
    gemini_api_key: Optional[str] = None,
) -> Dict[str, Any]:
    """
    개별 강좌의 전체 텍스트를 정밀 분석하여
    1) 규정 문장부호(■, -, [ ]) 외 비허용 부호 검수 리스트
    2) 맞춤법 및 오탈자 오류 리스트
    3) 긴 문장 감지 및 간략화(서술형/개조식) 제안 리스트
    4) 오타 및 '■, -, [ ]' 규정 부호 통일이 100% 완료된 강좌별 통합완성본텍스트를 반환합니다.
    """
    combined_raw = "\n".join(raw_texts)
    combined_clean = clean_pdf_broken_glyphs(combined_raw)
    sentences_data = restore_pdf_sentences(combined_clean)

    spelling_issues: List[Dict[str, Any]] = []
    bullet_issues: List[Dict[str, Any]] = []
    disallowed_punct_issues: List[Dict[str, Any]] = []
    long_sentence_proposals: List[Dict[str, Any]] = []
    combined_proofread: List[Dict[str, Any]] = []
    unified_paragraphs: List[str] = []

    seen_spelling_keys = set()
    seen_bullet_keys = set()
    seen_disallowed_keys = set()
    seen_long_keys = set()

    for item in sentences_data:
        sec = item["section"]
        sent = item["sentence"].strip()
        c_len = item["char_count"]

        if not sent:
            continue

        # 학원 푸터/법정 표기 제외
        is_footer = any(
            ft in sent for ft in ["학원등록번호", "러셀 대치학원", "원격교습", "교습과정", "수강료 및 교습비", "신고번호"]
        )

        # -------------------------------------------------------------
        # A. 비허용 문장부호(■, -, [ ] 외) 탐지 & 규정 부호 교정
        # -------------------------------------------------------------
        raw_disallowed = DISALLOWED_PUNCTUATION_RE.findall(sent) if not is_footer else []
        unique_found = [s for s in sorted(list(set(raw_disallowed))) if s not in ALLOWED_PUNCTUATION]
        m_non = NON_STANDARD_BULLET_RE.match(sent) if not is_footer else None
        m_no_sp = SQUARE_NO_SPACE_RE.match(sent) if not is_footer else None

        # 비허용 부호가 본문에 들어가 있거나 단락 시작이 비표준 기호인 경우
        if not is_footer and (unique_found or m_non):
            if m_non:
                head_sym = m_non.group(2).strip()
                if head_sym not in unique_found and head_sym not in ALLOWED_PUNCTUATION:
                    unique_found.insert(0, head_sym)

            # 규정 부호(■, -, [ ])로 교정하고 오타까지 수정한 완성형 단락 도출
            sanitized = sanitize_punctuations_in_text(sent)
            corrected_sent = correct_all_spelling_in_sentence(sanitized)

            symbols_str = ", ".join(unique_found) if unique_found else "비표준 기호"
            p_key = f"PUNCT_DISALLOWED|{symbols_str}|{sent}"
            if p_key not in seen_disallowed_keys:
                seen_disallowed_keys.add(p_key)
                p_info = {
                    "강좌명": course_name,
                    "강사명": instructor,
                    "섹션": sec,
                    "오류유형": "비허용 문장부호 (■, -, [ ] 외)",
                    "대상어구": symbols_str,
                    "교정제안": "■, -, [ ] 규정 부호로 통일",
                    "발견항목": f"{symbols_str} (규정 부호 외 사용)",
                    "사유": f"강의계획서 규정 문장부호(■, -, [ ]) 외의 비허용 부호 '{symbols_str}'가 사용되었습니다.",
                    "원문문장": sent,
                    "교정문장": corrected_sent,
                    "적용상태": "대기",
                }
                disallowed_punct_issues.append(p_info)
                combined_proofread.append({
                    "강사명": instructor,
                    "강좌명": course_name,
                    "교정 구분": "⬛ 비허용 문장부호 (■, -, [ ] 외)",
                    "발견 항목": f"{symbols_str}",
                    "PDF 원문 단락": sent,
                    "✅ 교정 완료 단락 (■ 통일 + 오타 수정본)": corrected_sent,
                    "사유": p_info["사유"],
                })
        elif not is_footer and m_no_sp:
            sanitized = sanitize_punctuations_in_text(sent)
            corrected_sent = correct_all_spelling_in_sentence(sanitized)
            b_key = f"BULLET_NOSPACE|■|{sent}"
            if b_key not in seen_bullet_keys:
                seen_bullet_keys.add(b_key)
                b_info = {
                    "강좌명": course_name,
                    "강사명": instructor,
                    "섹션": sec,
                    "오류유형": "기호통일 (■ 띄어쓰기)",
                    "대상어구": "■(붙여쓰기)",
                    "교정제안": "■ ",
                    "발견항목": "■(붙여쓰기) → ■ ",
                    "사유": "글머리 기호 '■' 뒤에 공백 1칸을 띄어 표준 서식을 준수합니다.",
                    "원문문장": sent,
                    "교정문장": corrected_sent,
                    "적용상태": "대기",
                }
                bullet_issues.append(b_info)
                combined_proofread.append({
                    "강사명": instructor,
                    "강좌명": course_name,
                    "교정 구분": "⬛ 기호통일 (■ 띄어쓰기)",
                    "발견 항목": "■(붙여쓰기) → ■ ",
                    "PDF 원문 단락": sent,
                    "✅ 교정 완료 단락 (■ 통일 + 오타 수정본)": corrected_sent,
                    "사유": b_info["사유"],
                })
        else:
            sanitized = sanitize_punctuations_in_text(sent)
            corrected_sent = correct_all_spelling_in_sentence(sanitized)

        unified_paragraphs.append(corrected_sent)

        # -------------------------------------------------------------
        # B. 맞춤법 및 오탈자 검사
        # -------------------------------------------------------------
        typo_hits = check_spelling_and_typos(sent) if not is_footer else []
        for th in typo_hits:
            s_key = f"{th['rule_type']}|{th['target_word']}|{sent}"
            if s_key not in seen_spelling_keys:
                seen_spelling_keys.add(s_key)
                spelling_issues.append({
                    "강좌명": course_name,
                    "강사명": instructor,
                    "섹션": sec,
                    "오류유형": th["rule_type"],
                    "대상어구": th["target_word"],
                    "교정제안": th["replacement_word"],
                    "발견항목": f"{th['target_word']} → {th['replacement_word']}",
                    "사유": th["reason"],
                    "원문문장": sent,
                    "교정문장": corrected_sent if corrected_sent != sent else th["suggested_sentence"],
                    "적용상태": "대기",
                })
                combined_proofread.append({
                    "강사명": instructor,
                    "강좌명": course_name,
                    "교정 구분": "🔴 맞춤법·오탈자",
                    "발견 항목": f"{th['target_word']} → {th['replacement_word']}",
                    "PDF 원문 단락": sent,
                    "✅ 교정 완료 단락 (■ 통일 + 오타 수정본)": corrected_sent if corrected_sent != sent else th["suggested_sentence"],
                    "사유": th["reason"],
                })

        # -------------------------------------------------------------
        # C. 긴 문장 감지 및 간략화 제안
        # -------------------------------------------------------------
        if not is_footer and c_len >= max_char_threshold:
            if sent not in seen_long_keys:
                seen_long_keys.add(sent)

                # Gemini API 호출 시도 (키가 있을 때)
                ai_proposals = None
                if gemini_api_key:
                    ai_proposals = simplify_sentence_with_gemini(
                        sent, api_key=gemini_api_key
                    )

                if ai_proposals:
                    p1, p2 = ai_proposals
                    engine_type = "Gemini AI"
                else:
                    p1, p2 = simplify_sentence_rule_based(sent)
                    engine_type = "내장 규칙 엔진"

                long_sentence_proposals.append({
                    "강좌명": course_name,
                    "강사명": instructor,
                    "섹션": sec,
                    "글자수": c_len,
                    "원문문장": sent,
                    "제안1_서술형": p1,
                    "제안2_개조식": p2,
                    "생성엔진": engine_type,
                    "선택제안": "제안1_서술형",
                    "최종문장": p1,
                })
                combined_proofread.append({
                    "강사명": instructor,
                    "강좌명": course_name,
                    "교정 구분": "✂️ 긴 문장 축약",
                    "발견 항목": f"글자수 {c_len}자 (기준 {max_char_threshold}자 초과)",
                    "PDF 원문 단락": sent,
                    "✅ 교정 완료 단락 (■ 통일 + 오타 수정본)": f"[제안 1 서술형] {p1}\n[제안 2 개조식] {p2}",
                    "사유": f"문장 길이가 {c_len}자로 길어 가독성 저하 우려 (간결 서술형 및 개조식 제안)",
                })

    unified_course_text = "\n\n".join(unified_paragraphs)

    return {
        "강좌명": course_name,
        "강사명": instructor,
        "맞춤법오류수": len(spelling_issues),
        "기호통일오류수": len(bullet_issues),
        "비허용부호오류수": len(disallowed_punct_issues),
        "긴문장수": len(long_sentence_proposals),
        "맞춤법_상세목록": spelling_issues,
        "기호통일_상세목록": bullet_issues,
        "비허용부호_상세목록": disallowed_punct_issues,
        "긴문장_상세목록": long_sentence_proposals,
        "통합교정_상세목록": combined_proofread,
        "통합완성본텍스트": unified_course_text,
    }


# -------------------------------------------------------------------------
# 8. 외부 정밀 맞춤법 검사기 (네이버 맞춤법 + 멀티스레딩 병렬 엔진)
# -------------------------------------------------------------------------
def check_naver_spelling_single(text: str) -> List[Dict[str, Any]]:
    """
    단일 문장에 대해 네이버 맞춤법 검사기 REST API 호출을 시도합니다.
    """
    clean = text.strip()
    if not clean or len(clean) < 4:
        return []

    # 네이버 맞춤법 모바일 엔드포인트
    url = (
        "https://m.search.naver.com/p/csearch/ocontent/spellchecker.nhn?"
        f"_callback=window.__jindo2_callback._$2582_0&q={urllib.parse.quote(clean[:500])}"
    )
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://m.search.naver.com/",
    }
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=2.5) as resp:
            data = resp.read().decode("utf-8")
            m = re.search(r"window\.__jindo2_callback\._\$2582_0\((.*)\);", data)
            if m:
                res_data = json.loads(m.group(1))
                res_obj = res_data.get("message", {}).get("result", {})
                err_cnt = int(res_obj.get("errata_count", 0))
                not_updated = res_obj.get("notag_html", "")
                if err_cnt > 0 and not_updated:
                    return [
                        {
                            "rule_type": "네이버 정밀 맞춤법",
                            "target_word": clean[:30] + ("..." if len(clean) > 30 else ""),
                            "replacement_word": not_updated[:40],
                            "reason": f"네이버 맞춤법 검사기에서 {err_cnt}건의 맞춤법/띄어쓰기 교정을 감지했습니다.",
                            "suggested_sentence": not_updated,
                        }
                    ]
    except Exception:
        pass
    return []


def batch_run_external_spelling_parallel(
    pdf_courses: List[Dict[str, Any]],
    gemini_api_key: Optional[str] = None,
    max_workers: int = 10,
    progress_callback: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    """
    concurrent.futures.ThreadPoolExecutor(max_workers=10)를 사용하여
    모든 강좌의 본문 문장들에 대해 외부 정밀 맞춤법 검사를 병렬로 고속 실행합니다.
    (20~40개 강좌 전체를 1~2초 만에 완수)
    """
    import concurrent.futures

    tasks = []
    for c in pdf_courses:
        c_title = str(c.get("강좌명", "")).strip()
        c_inst = str(c.get("강사명_원본", "") or c.get("강사명", "")).strip()
        raw_txts = c.get("강좌_텍스트리스트", [])
        if not raw_txts and c.get("원문전체텍스트"):
            raw_txts = [str(c.get("원문전체텍스트", ""))]

        combined_clean = clean_pdf_broken_glyphs("\n".join(raw_txts))
        sents = restore_pdf_sentences(combined_clean)
        for s_item in sents:
            tasks.append((c_title, c_inst, s_item["section"], s_item["sentence"]))

    results: List[Dict[str, Any]] = []
    seen_keys = set()
    total_tasks = len(tasks)
    completed_cnt = 0

    def _worker(task: Tuple[str, str, str, str]) -> List[Dict[str, Any]]:
        c_name, inst, sec, sentence = task
        hits = check_naver_spelling_single(sentence)
        task_res = []
        for h in hits:
            task_res.append(
                {
                    "강좌명": c_name,
                    "강사명": inst,
                    "섹션": sec,
                    "오류유형": h.get("rule_type", "외부 정밀 맞춤법"),
                    "대상어구": h.get("target_word", "-"),
                    "교정제안": h.get("replacement_word", "-"),
                    "사유": h.get("reason", "외부 정밀 맞춤법 교정 권장"),
                    "원문문장": sentence,
                    "교정문장": h.get("suggested_sentence", sentence),
                    "적용상태": "대기",
                }
            )
        return task_res

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(_worker, t): t for t in tasks}
        for fut in concurrent.futures.as_completed(future_map):
            completed_cnt += 1
            if progress_callback:
                try:
                    progress_callback(completed_cnt, total_tasks)
                except Exception:
                    pass
            try:
                res_list = fut.result()
                for r in res_list:
                    k = f"{r['오류유형']}|{r['대상어구']}|{r['원문문장']}"
                    if k not in seen_keys:
                        seen_keys.add(k)
                        results.append(r)
            except Exception:
                pass

    return results
