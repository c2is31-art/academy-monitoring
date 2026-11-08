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
def clean_pdf_broken_glyphs(text: str) -> str:
    """PDF 텍스트 내 글리프 깨짐, 잡음 기호 정제"""
    if not text:
        return ""
    s = text
    # 유니코드 불릿 기호 통일
    s = s.replace("▶", " • ").replace("▷", " • ").replace("■", " • ").replace("□", " • ")
    s = s.replace("★", " • ").replace("☆", " • ").replace("○", " • ").replace("●", " • ")
    s = s.replace("◆", " • ").replace("◇", " • ")
    # 불필요한 깨진 기호 제거
    s = re.sub(r"[§¡¿|]{2,}", " ", s)
    return s


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
        r"^(?:[•\-*·▶▷■□★☆○●◆◇]|(?:\d+|[가-하]|[A-Z])[\.\)])\s*"
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
def simplify_sentence_with_gemini(
    sentence: str,
    api_key: str,
    model_name: str = "gemini-2.5-flash",
) -> Optional[Tuple[str, str]]:
    """
    Gemini API를 호출하여 대치동 브로슈어 스타일 카피라이팅 축약안을 생성합니다.
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

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
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

    try:
        req_data = json.dumps(payload).encode("utf-8")
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
                # JSON 파싱 시 마크다운 백틱 제거
                text_clean = re.sub(r"^```(?:json)?\s*", "", text_out)
                text_clean = re.sub(r"\s*```$", "", text_clean).strip()
                data = json.loads(text_clean)
                p1 = data.get("proposal_1", "").strip()
                p2 = data.get("proposal_2", "").strip()
                if p1 and p2:
                    return p1, p2
    except Exception:
        # API 오류 또는 타임아웃 발생 시 fallback 허용
        pass

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
    1) 맞춤법 및 오탈자 오류 리스트
    2) 긴 문장 감지 및 간략화(서술형/개조식) 제안 리스트를 반환합니다.
    """
    combined_raw = "\n".join(raw_texts)
    combined_clean = clean_pdf_broken_glyphs(combined_raw)
    sentences_data = restore_pdf_sentences(combined_clean)

    spelling_issues: List[Dict[str, Any]] = []
    long_sentence_proposals: List[Dict[str, Any]] = []

    seen_spelling_keys = set()
    seen_long_keys = set()

    for item in sentences_data:
        sec = item["section"]
        sent = item["sentence"]
        c_len = item["char_count"]

        # 1) 맞춤법 및 오탈자 검사
        typo_hits = check_spelling_and_typos(sent)
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
                    "사유": th["reason"],
                    "원문문장": sent,
                    "교정문장": th["suggested_sentence"],
                    "적용상태": "대기",
                })

        # 2) 긴 문장 감지 및 간략화 제안
        if c_len >= max_char_threshold:
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

    return {
        "강좌명": course_name,
        "강사명": instructor,
        "맞춤법오류수": len(spelling_issues),
        "긴문장수": len(long_sentence_proposals),
        "맞춤법_상세목록": spelling_issues,
        "긴문장_상세목록": long_sentence_proposals,
    }
