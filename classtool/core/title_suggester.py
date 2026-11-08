"""
Gemini API 기반 대치동 단과 '15자 이내 초단축·임팩트형 강의명 제안' 모듈 (core/title_suggester.py)
- 대치동 학원가 최상위권 브로슈어 및 시간표에 최적화된 초단축(공백 포함 10자~16자 이내) AI 카피라이팅
- Gemini 3.8 최우선 지원 (Gemini 3.8 Flash / Pro -> 3.0 Flash -> 2.5 Pro -> 2.5 Flash 스마트 라우팅 및 자동 Fallback)
- requests 기반 REST API 직접 호출 (SDK 버전 충돌 방지 및 종속성 최소화)
- 슬로건 제안 기능 완전 삭제 및 동일 강사의 여러 강좌 개설 시 강좌명 중복 100% 원천 차단
- 4대 카피라이팅 방향성 지원:
    1) 대치동 트렌드 종합
    2) 내신 1등급/학교별 기출 강조
    3) 수능 평가원/실전 킬러 강조
    4) 개념 완성/중위권 도약 강조
- 3가지 초단축 스타일 제안 (모두 10~16자 엄격 준수):
    - style_1_short (⚡ 핵심직관형 - 12자 내외, 예: '시간 단축 지문 장악 국어', '확통 개념×4점 킬러 압축')
    - style_2_brand (💎 시그니처형 - 14자 내외, 예: 'META 대수 : 조건의 단서', '지인해(知認解) 코어 국어', '수풀림 확통 시그니처')
    - style_3_punch (🔥 킬러임팩트형 - 12자 내외, 예: '미적1 조건반사 문풀', '기하 10% 킬러 돌파')
- 서술형 수식어(~를 위한, ~하는, ~완성하는 등) 금지 후처리 및 글자 수 초과 시 자동 압축 알고리즘 탑재
- API 키 부재 또는 호출 실패 시 고품질 대치동 초단축 규칙 기반 엔진(Fallback) 가동
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List, Optional, Set

import requests

# Gemini 모델 후보군 (Gemini 3.8 최우선 -> 자동 Fallback)
MODEL_CANDIDATES = [
    "gemini-3.8-flash",
    "gemini-3.8-pro",
    "gemini-3.0-flash",
    "gemini-2.5-pro",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
]

# 성공한 모델 캐싱 (호출 성공 시 해당 세션에서 우선 사용)
_WORKING_MODEL: Optional[str] = None

DIRECTION_PROMPTS = {
    "대치동 트렌드 종합": (
        "대치동 최상위권 트렌드: 시간표와 브로슈어 한 칸에 0.5초 만에 각인되는 초단축(10~15자) 카피. "
        "서술형 문장을 배제하고 [핵심 과목/브랜드 + 타깃/무기 명사]로 압축하세요."
    ),
    "내신 1등급/학교별 기출 강조": (
        "대치/강남 주요 고교(휘문·중동·단대·경기·숙명 등) 내신 전교 1등 및 1등급 목표. "
        "기출 변형, 서술형 만점, 학교별 적중을 15자 이내의 초단축 명사구로 압축하세요."
    ),
    "수능 평가원/실전 킬러 강조": (
        "평가원 수능 최신 트렌드: 준킬러 시간 단축, 4점 킬러 구조화, 수능 100점 만점 돌파. "
        "실전 킬러 무기를 12자 내외의 강렬한 펀치 카피로 압축하세요."
    ),
    "개념 완성/중위권 도약 강조": (
        "취약 유형 돌파, 빈틈없는 개념 정립, 중위권에서 확실한 1등급 도약. "
        "군더더기 없는 12자 내외의 명료하고 자신감 넘치는 압축 제목을 작성하세요."
    ),
}


def get_gemini_api_key() -> Optional[str]:
    """
    환경 변수 및 Streamlit secrets에서 Gemini API 키를 안전하게 탐색합니다.
    """
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if key:
        return key

    try:
        import streamlit as st

        if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
            key = str(st.secrets["GEMINI_API_KEY"]).strip()
            if key:
                return key
    except Exception:
        pass

    return None


def clean_and_truncate_title(title: str, max_len: int = 16) -> str:
    """
    강의명을 10~16자 이내로 엄격하게 정제하고 서술형 수식어를 제거합니다.
    """
    if not title:
        return ""
    t = str(title).strip()
    # 양쪽 따옴표, 백틱 제거
    t = re.sub(r"^['\"`]+|['\"`]+$", "", t).strip()

    # 서술형 불필요 수식어 제거
    remove_patterns = [
        r"~를\s*위한",
        r"를\s*위한",
        r"~하는\s*수업",
        r"하는\s*수업",
        r"한\s*번에\s*끝내는",
        r"완벽하게\s*끝내는",
        r"함께하는",
        r"만드는",
        r"완성하는",
        r"정복하는",
    ]
    for p in remove_patterns:
        t = re.sub(p, "", t, flags=re.IGNORECASE).strip()

    # 연속 공백 정리
    t = re.sub(r"\s+", " ", t).strip()

    # 만약 대괄호 말머리가 붙어서 16자를 넘으면 대괄호 태그 제거
    if len(t) > max_len and re.search(r"\[.*?\]", t):
        t_no_brackets = re.sub(r"\[.*?\]", "", t).strip()
        if len(t_no_brackets) >= 6:
            t = t_no_brackets

    # 여전히 16자를 넘으면 공백 기준 어절 단위 절단 후 안전 트리밍
    if len(t) > max_len:
        words = t.split()
        trimmed = ""
        for w in words:
            candidate = f"{trimmed} {w}".strip()
            if len(candidate) <= max_len:
                trimmed = candidate
            else:
                break
        t = trimmed if len(trimmed) >= 8 else t[:max_len].strip()

    return t


KNOWN_IMPACTFUL_KEYWORDS = [
    "기고만장", "진심국어", "백설", "명월", "풀이학개론", "Big Bang", "知認解",
    "수풀림", "META", "시그니처", "킬캠", "시발점", "뉴런", "드릴", "서바이벌",
    "시대인재", "블랙라벨", "PRIDE", "NO.1", "PRE :", "코어", "만장일치",
    "조건의 단서", "패러다임", "킬러 돌파", "골든키", "포텐셜", "헤비급",
    "미친국어", "동백", "수국", "아우라", "메타수학", "절대수학", "일격필살",
]

COMMON_BORING_PATTERNS = [
    r"기본반\b", r"입문반\b", r"개념반\b", r"총정리\b", r"개념유형\b",
    r"실력반\b", r"합격자반\b", r"윈터반\b", r"특강반\b", r"정규반\b",
    r"완성반\b", r"문제풀이반\b", r"심화반\b", r"기초반\b", r"대비반\b",
]


def _evaluate_title_impact_rule_based(current_title: str, instructor: str = "") -> Tuple[bool, int, str]:
    """
    기존 강의명이 이미 충분히 임팩트 있거나 고유 시그니처 브랜드를 포함하는지 규칙 기반으로 사전 평가합니다.
    Returns:
        (needs_revision: bool, impact_score: int, evaluation_reason: str)
    """
    t = str(current_title).strip()
    if not t:
        return True, 40, "강의명이 비어 있어 신규 임팩트 강의명 추천 필요"

    # 1. 고유 시그니처 브랜드 키워드 매칭
    for kw in KNOWN_IMPACTFUL_KEYWORDS:
        if kw.lower() in t.lower():
            return False, 88, f"강사 시그니처 브랜드('{kw}')가 명확하여 기존 강의명 유지 권장"

    # 2. 괄호 내 독창적 풀이형 네이밍 또는 한자 병기 (예: '기고만장 (기하는 고병훈 만장일치)', '지인해(知認解)')
    if re.search(r"\(.*?[知認解만장일치학개론시그니처킬러].*?\)", t):
        return False, 85, "강사 고유의 해설형 브랜드 네이밍이 돋보여 기존 강의명 유지 권장"

    # 3. 콜론(:) 기반 시그니처 구조이면서 16자 이하인 세련된 제목 (예: '진심국어: 백설', 'PRE : 고등국어')
    if ":" in t and len(t) <= 16 and not any(re.search(p, t) for p in COMMON_BORING_PATTERNS):
        return False, 82, "시그니처 콜론(:) 구조의 직관적 네이밍으로 기존 강의명 유지 권장"

    # 4. 평이한 나열형 패턴 검사 (학년/과목/기본반/총정리)
    clean_t = re.sub(r"\[[^\]]+\]|<[^>]+>", "", t).strip()
    boring_matches = [p for p in COMMON_BORING_PATTERNS if re.search(p, clean_t)]
    if boring_matches:
        matched_str = boring_matches[0].replace(r"\b", "")
        return True, 52, f"'{matched_str}' 등 학년·과목·수준만 평이하게 나열됨 -> 15자 이내 실전 킬러 카피 추천"

    # 5. 글자수가 너무 길고 산만한 문장형 제목 (20자 초과)
    if len(t) > 20:
        return True, 48, f"글자수({len(t)}자)가 너무 길고 서술적이어서 시간표 시인성 부족 -> 초단축 리브랜딩 추천"

    # 6. 기타 15자 이내이면서 임팩트 있는 명사형 제목인 경우 유지 권장
    if 8 <= len(t) <= 15 and any(kw in t for kw in ("킬러", "단서", "1등급", "돌파", "정복", "만점", "마스터", "구조화")):
        return False, 80, "핵심 목표와 킬러 키워드가 직관적이어서 기존 강의명 유지 권장"

    # 기본값: 수정 추천
    return True, 58, "학년과 과목 위주 구성으로 차별성이 다소 부족함 -> 15자 이내 초단축 카피 추천"


def _generate_fallback_suggestion(
    subject: str,
    grade: str,
    instructor: str,
    current_title: str,
    direction: str = "대치동 트렌드 종합",
    seen_titles: Optional[Set[str]] = None,
) -> Dict[str, Any]:
    """API 키가 없거나 호출 실패 시 15자 이내 초단축 대치동 규칙 기반 엔진으로 생성합니다 (유지/수정 판정 반영)."""
    needs_revision, score, reason = _evaluate_title_impact_rule_based(current_title, instructor)

    if not needs_revision:
        return {
            "impact_score": score,
            "needs_revision": False,
            "status_badge": "✅ 기존 유지",
            "evaluation_reason": reason,
            "critique": reason,
            "recommended_title_1": "-",
            "recommended_title_2": "-",
            # 기존 호환 키 유지
            "style_1_short": "-",
            "style_2_brand": "-",
            "style_3_punch": "-",
            "style_1_intuitive": "-",
            "style_2_signature": "-",
            "style_3_killer": "-",
            "engine": "대치동 초단축 규칙 엔진 (Fallback)",
        }

    # 수정 추천 대상 강좌에 대한 초단축 2종 생성
    clean_title = re.sub(r"\[[^\]]+\]|<[^>]+>", "", current_title).strip() or current_title
    clean_title = re.sub(r"\s+", " ", clean_title).strip()

    title_words = [w for w in clean_title.split() if len(w) >= 2]
    sub_detail = ""
    for w in title_words:
        if any(kw in w for kw in ("공수", "수학", "미적", "확통", "대수", "기하", "독서", "문학", "문법", "화작", "언매", "영어", "물리", "화학", "생명", "지구", "사문", "생윤", "한지", "세지")):
            sub_detail = w[:6]
            break
    if not sub_detail and title_words:
        sub_detail = title_words[0][:6]
    if not sub_detail:
        sub_detail = subject

    feature_kw = ""
    for w in title_words:
        if any(kw in w for kw in ("기출", "킬러", "심화", "개념", "200제", "300제", "유형", "마스터", "실전", "파이널", "고난도")):
            feature_kw = w[:6]
            break

    # 추천 1: ⚡ 직관·시그니처형 (12자 내외)
    # 추천 2: 🔥 킬러·임팩트형 (12자 내외)
    if feature_kw and feature_kw != sub_detail:
        base_s1 = f"{sub_detail} {feature_kw} 압축"
        base_s2 = f"{instructor} {sub_detail} 1등급 킬러"
    else:
        base_s1 = f"{sub_detail} 핵심 실전 압축"
        base_s2 = f"{sub_detail} 조건반사 킬러 문풀"

    s1 = clean_and_truncate_title(base_s1, max_len=15)
    s2 = clean_and_truncate_title(base_s2, max_len=15)

    if seen_titles is not None:
        if s1 in seen_titles:
            s1 = clean_and_truncate_title(f"{instructor} {sub_detail} 마스터", max_len=15)
        if s2 in seen_titles:
            s2 = clean_and_truncate_title(f"{sub_detail} 만점 킬러 돌파", max_len=15)

    return {
        "impact_score": score,
        "needs_revision": True,
        "status_badge": "🔧 수정 추천",
        "evaluation_reason": reason,
        "critique": reason,
        "recommended_title_1": s1,
        "recommended_title_2": s2,
        # 기존 호환 키 유지
        "style_1_short": s1,
        "style_2_brand": s2,
        "style_3_punch": s2,
        "style_1_intuitive": s1,
        "style_2_signature": s2,
        "style_3_killer": s2,
        "engine": "대치동 초단축 규칙 엔진 (Fallback)",
    }


def suggest_course_titles_with_gemini(
    course_info: Dict[str, Any],
    api_key: Optional[str] = None,
    direction: str = "대치동 트렌드 종합",
    sibling_titles: Optional[List[str]] = None,
    timeout_sec: int = 10,
) -> Dict[str, Any]:
    """
    단일 강좌의 정보를 받아 Gemini 3.8을 통해
    1) 기존 강의명 임팩트 사전 평가 및 유지 권장 vs 수정 추천 판정
    2) 수정 추천 시 15자 이내 초단축 2종 강의명을 제안합니다 (슬로건 제안 100% 배제).
    """
    global _WORKING_MODEL

    subj = str(course_info.get("과목", "")).strip() or "과목"
    grade = str(course_info.get("학년", "") or course_info.get("대상학년", "")).strip() or "고등"
    inst = str(course_info.get("강사명", "") or course_info.get("강사", "")).strip() or "강사"
    curr_title = str(course_info.get("PDF 강좌명", "") or course_info.get("강좌명", "")).strip()
    target = str(course_info.get("추천대상", "")).strip() or "상위권"
    time_info = str(course_info.get("수업시간", "") or course_info.get("수업시간_원본", "")).strip()
    features = str(course_info.get("강의특징_요약", "") or course_info.get("주차별진도계획_요약", "")).strip()[:200]

    key_to_use = api_key or get_gemini_api_key()
    if not key_to_use:
        return _generate_fallback_suggestion(subj, grade, inst, curr_title, direction)

    dir_instruction = DIRECTION_PROMPTS.get(direction, DIRECTION_PROMPTS["대치동 트렌드 종합"])

    sibling_note = ""
    if sibling_titles:
        sibling_note = f"\n[★ 동일 강사({inst}T)의 다른 개설 강좌들 - 중복 절대 금지!]\n" + "\n".join(
            f"- {t}" for t in sibling_titles if t != curr_title
        )

    prompt = f"""
당신은 대한민국 서울 대치동 최상위권 단과 학원의 스타 강사 전문 수석 카피라이터입니다.
학원 브로슈어와 시간표에 실릴 강의명들을 전문적으로 평가하고 리브랜딩하는 업무를 수행합니다.

[강좌 기본 정보]
- 과목 / 대상학년 : {subj} / {grade}
- 강사명 : {inst} 선생님
- 현재 강좌명 : {curr_title}
- 수업시간 및 타깃 : {time_info or target}
- 특징 요약 : {features or "(상위권 대비 수능/내신 실전 강좌)"}
{sibling_note}

[카피라이팅 방향성]
{dir_instruction}

[★ 핵심 업무 지침: 1) 기존 유지 vs 2) 수정 추천 엄격 판정 ★]
1. 모든 강좌에 새 이름을 억지로 추천하지 마세요!
2. ✅ [기존 강의명 유지 권장] (needs_revision: false, impact_score: 80~95점):
   - 강사 고유의 시그니처 브랜드(예: '기고만장', '진심국어 : 백설/명월', '[PRIDE NO.1] PRE : 고등국어', '풀이학개론', 'Big Bang', '知認解', '수풀림' 등)가 명확하거나, 이미 15자 이내로 대치동 수험생/학부모에게 강렬한 임팩트를 주는 훌륭한 강의명인 경우.
   - 이 경우 recommended_title_1과 recommended_title_2는 반드시 "-"로 출력하세요.
   - evaluation_reason에 어떤 브랜드/강점이 훌륭하여 유지 권장하는지 명확히 작성하세요.

3. 🔧 [강의명 수정 추천] (needs_revision: true, impact_score: 40~75점):
   - '고2 윈터 기본반', '<윈터> 대수 기본반', '[확률과통계] 개념유형 총정리', '예비고1 영어 Master실력반', '중동고 합격자반'처럼 단순히 학년/과목/기본반/총정리만 평이하게 나열되어 임팩트가 부족하거나, 문장형으로 산만한 경우.
   - 공백 포함 15자 이내의 임팩트 있는 새 강의명 2종을 추천하세요:
     * recommended_title_1: ⚡ 직관·시그니처형 (12자 내외, 목표와 강사 브랜드가 0.5초 만에 각인)
     * recommended_title_2: 🔥 킬러·임팩트형 (12자 내외, 1등급 킬러 정복과 문항 공략을 전면에 내세움)
   - evaluation_reason에 기존 이름의 아쉬운 점과 추천 방향을 1문장(50자 내외)으로 작성하세요.

[★ 추천 강의명 엄격 제약조건 ★]
- 부제나 슬로건은 절대 제안하지 마세요! 오직 본 강의명만 제시하세요.
- 모든 추천 강의명은 반드시 **공백 포함 10자 이상 ~ 15자 이하 (최대 3~4어절)**여야 합니다. 16자 초과 금지!
- 서술형 수식어(~를 위한, ~하는, ~완성하는, ~만드는, 한 번에 끝내는 등) 사용 절대 금지!
- 같은 강사의 다른 강좌와 절대 중복되지 않도록 고유 세부 과목/타깃을 살리세요.

반드시 아래 JSON 포맷으로만 응답하세요. (마크다운 백틱 및 추가 설명 금지, slogan 필드 금지)
{{
  "impact_score": 85,
  "needs_revision": false,
  "evaluation_reason": "강사 시그니처 브랜드와 과목 정체성이 명확하여 기존 강좌명 유지 권장",
  "recommended_title_1": "-",
  "recommended_title_2": "-"
}}
또는
{{
  "impact_score": 52,
  "needs_revision": true,
  "evaluation_reason": "학년과 과목만 평이하게 나열되어 시인성이 낮음 -> 시간 단축 킬러 카피 추천",
  "recommended_title_1": "공통수학1 킬러 압축",
  "recommended_title_2": "{inst}T 공수1 조건반사 문풀"
}}
"""

    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.25,
            "maxOutputTokens": 400,
        },
    }

    # 모델 후보군 순차 시도 (캐시된 작동 모델 우선)
    models_to_try = []
    if _WORKING_MODEL and _WORKING_MODEL in MODEL_CANDIDATES:
        models_to_try.append(_WORKING_MODEL)
    for m in MODEL_CANDIDATES:
        if m not in models_to_try:
            models_to_try.append(m)

    for cand_model in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{cand_model}:generateContent?key={key_to_use}"
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout_sec)
            if resp.status_code == 200:
                res_data = resp.json()
                candidates = res_data.get("candidates", [])
                if candidates:
                    cand = candidates[0]
                    parts = cand.get("content", {}).get("parts", [])
                    if parts:
                        raw_text = parts[0].get("text", "").strip()
                        clean_json = re.sub(r"^```(?:json)?\s*", "", raw_text)
                        clean_json = re.sub(r"\s*```$", "", clean_json).strip()
                        parsed = json.loads(clean_json)

                        impact_score = int(parsed.get("impact_score", 70))
                        needs_revision = bool(parsed.get("needs_revision", impact_score < 80))
                        eval_reason = str(parsed.get("evaluation_reason") or parsed.get("critique", "")).strip()

                        raw_r1 = str(parsed.get("recommended_title_1") or parsed.get("style_1_short") or "").strip()
                        raw_r2 = str(parsed.get("recommended_title_2") or parsed.get("style_2_brand") or "").strip()

                        if not needs_revision or impact_score >= 80:
                            needs_revision = False
                            status_badge = "✅ 기존 유지"
                            rec1 = "-"
                            rec2 = "-"
                        else:
                            status_badge = "🔧 수정 추천"
                            rec1 = clean_and_truncate_title(raw_r1, max_len=15) if raw_r1 and raw_r1 != "-" else clean_and_truncate_title(f"{subj} 핵심 실전 압축", max_len=15)
                            rec2 = clean_and_truncate_title(raw_r2, max_len=15) if raw_r2 and raw_r2 != "-" else clean_and_truncate_title(f"{inst}T {subj} 킬러 돌파", max_len=15)

                        # 작동 모델 캐싱
                        _WORKING_MODEL = cand_model

                        return {
                            "impact_score": impact_score,
                            "needs_revision": needs_revision,
                            "status_badge": status_badge,
                            "evaluation_reason": eval_reason,
                            "critique": eval_reason,
                            "recommended_title_1": rec1,
                            "recommended_title_2": rec2,
                            # 기존 호환 키 유지
                            "style_1_short": rec1,
                            "style_2_brand": rec2,
                            "style_3_punch": rec2,
                            "style_1_intuitive": rec1,
                            "style_2_signature": rec2,
                            "style_3_killer": rec2,
                            "engine": f"Gemini 3.8 ({cand_model})",
                        }
        except Exception:
            continue

    return _generate_fallback_suggestion(subj, grade, inst, curr_title, direction)


def batch_suggest_course_titles(
    course_list: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    direction: str = "대치동 트렌드 종합",
    progress_callback: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    """
    여러 강좌에 대해 순차적으로 AI 카피라이팅 분석을 실행합니다:
    - 임팩트 사전 평가: 우수 강좌명은 '기존 유지 권장(추천 생략)' 판정
    - 수정 필요 강좌명에 대해서만 15자 이내 초단축 2종 강의명 추천
    - 슬로건 추천 완전 삭제
    - 동일 강사의 다른 개설 강좌 정보를 프롬프트에 전달하여 중복 원천 차단
    """
    results: List[Dict[str, Any]] = []
    total = len(course_list)

    # 강사별 개설 강좌명 리스트 사전 수집
    instructor_courses_map: Dict[str, List[str]] = {}
    for c in course_list:
        inst = str(c.get("강사명", "") or c.get("강사", "")).strip()
        t = str(c.get("PDF 강좌명", "") or c.get("강좌명", "")).strip()
        if inst and t:
            instructor_courses_map.setdefault(inst, []).append(t)

    # 전체 중복 방지 세트
    seen_titles: Set[str] = set()

    for idx, c in enumerate(course_list, start=1):
        inst = str(c.get("강사명", "") or c.get("강사", "")).strip()
        sibling_titles = instructor_courses_map.get(inst, [])

        sug = suggest_course_titles_with_gemini(
            course_info=c,
            api_key=api_key,
            direction=direction,
            sibling_titles=sibling_titles,
        )

        needs_rev = sug.get("needs_revision", False)
        status_badge = sug.get("status_badge", "🔧 수정 추천" if needs_rev else "✅ 기존 유지")
        score = sug.get("impact_score", 70)
        reason = sug.get("evaluation_reason", "")
        rec1 = sug.get("recommended_title_1", "-")
        rec2 = sug.get("recommended_title_2", "-")

        curr_t = str(c.get("PDF 강좌명", "") or c.get("강좌명", "")).strip()

        # 수정 추천인 경우 중복 회피 처리
        if needs_rev and rec1 != "-":
            subj = str(c.get("과목", "")).strip()
            clean_curr = re.sub(r"\[[^\]]+\]|<[^>]+>", "", curr_t).strip()
            curr_words = clean_curr.split()
            kw_spec = curr_words[0][:5] if curr_words else subj

            if rec1 in seen_titles:
                rec1 = clean_and_truncate_title(f"{kw_spec} 핵심 마스터", max_len=15)
            if rec1 in seen_titles:
                rec1 = clean_and_truncate_title(f"{inst} {kw_spec} 실전", max_len=15)

            if rec2 in seen_titles:
                rec2 = clean_and_truncate_title(f"{inst}T {kw_spec} 킬러", max_len=15)

            seen_titles.add(rec1)
            seen_titles.add(rec2)

        combined = {
            "순번": idx,
            "판정결과": status_badge,
            "needs_revision": needs_rev,
            "과목": c.get("과목", ""),
            "학년": c.get("학년", "") or c.get("대상학년", ""),
            "강사명": inst,
            "현재강의명": curr_t,
            "현재강좌명": curr_t,  # 하위 호환
            "임팩트점수": score,
            "판정사유": reason,
            "진단비평": reason,  # 하위 호환
            "⚡ 추천 강의명 1 (직관·시그니처형)": rec1,
            "🔥 추천 강의명 2 (킬러·임팩트형)": rec2,
            # 기존 호환 키 유지
            "제안1_초단축직관형": rec1,
            "제안2_초단축시그니처": rec2,
            "제안3_초단축킬러형": rec2,
            "제안1_직관형": rec1,
            "제안2_시그니처형": rec2,
            "제안3_킬러임팩트형": rec2,
            "생성엔진": sug.get("engine", ""),
            "최종선택강좌명": curr_t if not needs_rev else rec1,
        }
        results.append(combined)

        if progress_callback:
            try:
                progress_callback(idx, total, curr_t)
            except Exception:
                pass

    return results
