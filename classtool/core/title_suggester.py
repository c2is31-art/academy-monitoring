"""
Gemini 3.8 기반 대치동 단과 'PDF 계획서 기반 스마트 패스 & 수정 추천' AI 리브랜딩 모듈 (core/title_suggester.py)
- RMS 강좌명뿐만 아니라 강의계획서 PDF(추천대상, 강의특징, 관리프로그램) 내용을 자동 검토하여 스마트하게 판정
- [스마트 패스 (기존 유지)]: 시그니처 브랜드(진심국어, 기고만장, 수풀림, META 등)나 과목 정체성이 확실한 강좌는 유지
- [수정 추천]: 평이한 나열형(기본반, 입문반, 총정리 등) 강좌만 선별하여 대치동 브로슈어 시인성에 맞춘 15자 이내 핵심 명사형 리브랜딩 제안
- 원본 대괄호 태그('[...]')는 100% 무조건 보존 (글자 수 산정 시 제외)
- 본문 제목은 공백 포함 15자 이내 (2~3어절) 핵심 키워드 조합으로만 제안 (문장형 및 슬로건 완전 배제)
- Gemini 3.8 최우선 지원 (Gemini 3.8 Flash / Pro -> 3.0 Flash -> 2.5 Pro -> 2.5 Flash 자동 스마트 라우팅 및 Fallback)
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple

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

# 카피라이팅 기본 방향성 (하위 호환 유지)
DIRECTION_PROMPTS = {
    "스마트 분석": (
        "강의계획서(PDF)의 추천 대상, 강의 특징, 관리 프로그램의 핵심 무기(시간 단축, 킬러 정복, 조건 해석, 지문 장악 등)를 "
        "대치동 최상위권 브로슈어에 최적화된 15자 이내(2~3어절) 핵심 명사구로 완성하세요."
    ),
    "대치동 트렌드 종합": (
        "대치동 최상위권 브로슈어에 어울리는 15자 이내(2~3어절) 핵심 명사형 카피로 완성하세요."
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


def extract_bracket_prefix_and_body(title: str) -> Tuple[str, str]:
    """
    강좌명의 맨 앞에 위치한 모든 연속된 대괄호 태그(예: '[고1 특강][실력반]', '[고2 정규]', '[고1 중동고 정규]')와
    그 뒤에 이어지는 실제 본문 강의명을 정밀하게 분리합니다.
    대괄호 안의 학년/반 분류 태그를 훼손 없이 100% 보존하기 위한 핵심 함수입니다.
    """
    if not title:
        return "", ""
    t = str(title).strip()
    match = re.match(r"^(\s*\[[^\]]+\])+", t)
    if match:
        prefix = match.group(0).strip()
        body = t[match.end():].strip()
        return prefix, body
    return "", t


def extract_course_details_from_pdf_text(full_text: str) -> Dict[str, str]:
    """
    강의계획서(PDF) 전체 텍스트에서 추천대상, 강의특징, 관리프로그램을 정밀하게 추출합니다.
    """
    res = {"추천대상": "", "강의특징": "", "관리프로그램": ""}
    if not full_text:
        return res

    # 1. 추천 대상 추출
    m_target = re.search(r"추천\s*대상\s*[:：]?\s*([^\n\r]+)", full_text)
    if m_target:
        res["추천대상"] = m_target.group(1).strip()

    # 2. 강의 특징 추출
    m_feat = re.search(
        r"(?:강의\s*특징\s*/\s*관리|강의\s*특징|강의\s*목표|강좌\s*특징)(.*?)(?:관리\s*프로그램|수강생\s*관리|클리닉|주차별\s*진도|진도\s*계획|$)",
        full_text,
        re.DOTALL,
    )
    if m_feat:
        lines = [line.strip() for line in m_feat.group(1).splitlines() if line.strip() and len(line.strip()) > 3]
        res["강의특징"] = " ".join(lines[:4])[:200]

    # 3. 관리 프로그램 추출
    m_mgmt = re.search(
        r"(?:관리\s*프로그램|수강생\s*관리|클리닉\s*안내|학습\s*관리)(.*?)(?:주차별\s*진도|진도\s*계획|교재\s*안내|$)",
        full_text,
        re.DOTALL,
    )
    if m_mgmt:
        lines = [line.strip() for line in m_mgmt.group(1).splitlines() if line.strip() and len(line.strip()) > 3]
        res["관리프로그램"] = " ".join(lines[:3])[:150]

    return res


def clean_and_truncate_title(
    title: str,
    original_prefix: str = "",
    min_body_len: int = 4,
    max_body_len: int = 15,
) -> str:
    """
    강의명을 정제하고 대괄호 태그('[...]')는 100% 무조건 보존합니다.
    - 글자 수 계산 및 제한(15자 이내, 2~3어절)은 대괄호 태그를 완전히 제외한 '본문 강의명'에만 적용됩니다.
    - 서술형 수식어(~를 위한, ~하는 수업, 한 번에 끝내는 등)를 완전히 제거합니다.
    - 2~3어절 초단축 명사형 키워드 조합으로만 유지합니다.
    """
    if not title or title == "-":
        return title

    t = str(title).strip()
    t = re.sub(r"^['\"`]+|['\"`]+$", "", t).strip()

    extracted_prefix, body = extract_bracket_prefix_and_body(t)
    prefix_to_use = (original_prefix or extracted_prefix).strip()

    if prefix_to_use and body:
        sub_prefix, sub_body = extract_bracket_prefix_and_body(body)
        if sub_prefix:
            body = sub_body

    if not body and prefix_to_use and not original_prefix:
        body = prefix_to_use
        prefix_to_use = ""

    # 서술형 불필요 수식어 및 슬로건 잔재 제거
    remove_patterns = [
        r"^AI\s*(추천|제안)?\s*[:：\-]\s*",
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
        body = re.sub(p, "", body, flags=re.IGNORECASE).strip()

    body = re.sub(r"\s+", " ", body).strip()

    # 본문 글자 수 조절 (대괄호 제외 15자 이내, 2~3어절 유지)
    if len(body) > max_body_len:
        words = body.split()
        trimmed = ""
        for w in words[:3]:
            candidate = f"{trimmed} {w}".strip()
            if len(candidate) <= max_body_len:
                trimmed = candidate
            else:
                break
        body = trimmed if len(trimmed) >= min_body_len else body[:max_body_len].strip()

    if prefix_to_use:
        return f"{prefix_to_use} {body}".strip() if body else prefix_to_use
    return body


KNOWN_IMPACTFUL_KEYWORDS = [
    "진심국어", "기고만장", "수풀림", "META", "풀이학개론", "Big Bang", "知認解",
    "시그니처", "킬캠", "시발점", "뉴런", "드릴", "서바이벌", "시대인재",
    "블랙라벨", "PRIDE", "NO.1", "PRE :", "코어", "만장일치", "조건의 단서",
    "패러다임", "킬러 돌파", "골든키", "포텐셜", "헤비급", "미친국어", "동백",
    "수국", "아우라", "메타수학", "절대수학", "일격필살", "지인해",
]

COMMON_BORING_PATTERNS = [
    r"기본반\b", r"입문반\b", r"개념반\b", r"총정리\b", r"개념유형\b",
    r"실력반\b", r"합격자반\b", r"윈터반\b", r"특강반\b", r"정규반\b",
    r"완성반\b", r"문제풀이반\b", r"심화반\b", r"기초반\b", r"대비반\b",
]


def _evaluate_title_impact_rule_based(
    current_title: str,
    instructor: str = "",
    target: str = "",
    features: str = "",
) -> Tuple[bool, int, str]:
    """
    기존 강의명이 이미 충분히 임팩트 있거나 고유 시그니처 브랜드를 포함하는지 스마트하게 판정합니다.
    Returns:
        (needs_revision: bool, impact_score: int, evaluation_reason: str)
    """
    t = str(current_title).strip()
    if not t:
        return True, 40, "강의명이 비어 있어 신규 임팩트 강의명 추천 필요"

    prefix, body = extract_bracket_prefix_and_body(t)
    eval_target = body if body else t

    # 1. 고유 시그니처 브랜드 키워드 매칭
    for kw in KNOWN_IMPACTFUL_KEYWORDS:
        if kw.lower() in t.lower():
            return False, 90, f"강사 고유 시그니처 브랜드('{kw}')가 명확하여 기존 강의명 유지(패스) 권장"

    # 2. 괄호 내 독창적 풀이형 네이밍 또는 한자 병기
    if re.search(r"\(.*?[知認解만장일치학개론시그니처킬러].*?\)", t):
        return False, 88, "강사 고유의 해설형 브랜드 네이밍이 돋보여 기존 강의명 유지(패스) 권장"

    # 3. 콜론(:) 기반 시그니처 구조이면서 본문 15자 이하인 세련된 제목
    if ":" in eval_target and len(eval_target) <= 15 and not any(re.search(p, eval_target) for p in COMMON_BORING_PATTERNS):
        return False, 85, "시그니처 콜론(:) 구조의 직관적 네이밍으로 기존 강의명 유지(패스) 권장"

    # 4. 평이한 나열형 패턴 검사 (학년/과목/기본반/총정리)
    clean_target = re.sub(r"\[[^\]]+\]|<[^>]+>", "", eval_target).strip()
    boring_matches = [p for p in COMMON_BORING_PATTERNS if re.search(p, clean_target)]
    if boring_matches:
        matched_str = boring_matches[0].replace(r"\b", "")
        return True, 45, f"'{matched_str}' 등 학년·수준만 평이하게 나열됨 -> PDF 계획서 핵심 포인트를 살린 초단축 네이밍 추천"

    # 5. 글자수가 너무 길고 산만한 문장형 제목 (대괄호 제외 본문 18자 초과)
    if len(eval_target) > 18:
        return True, 50, f"대괄호 제외 본문 글자수({len(eval_target)}자)가 너무 길어 시인성 부족 -> 15자 이내 핵심 임팩트 추천"

    # 6. 대괄호 제외 본문이 8~15자 이내이면서 임팩트 있는 명사형/킬러 키워드 포함 시 유지
    if 8 <= len(eval_target) <= 15 and any(kw in eval_target for kw in ("킬러", "단서", "1등급", "돌파", "정복", "만점", "마스터", "구조화", "FinalCUT", "최상위", "시그니처")):
        return False, 85, "핵심 목표와 킬러 키워드가 직관적이어서 기존 강의명 유지(패스) 권장"

    # 기본값: 수정 추천
    return True, 55, "학년과 과목 위주 구성으로 차별성 다소 부족 -> PDF 특징 반영 15자 이내 임팩트 강의명 추천"


def _generate_fallback_suggestion(
    subject: str,
    grade: str,
    instructor: str,
    current_title: str,
    target: str = "",
    features: str = "",
    management: str = "",
    direction: str = "스마트 분석",
    seen_titles: Optional[Set[str]] = None,
) -> Dict[str, Any]:
    """API 키가 없거나 호출 실패 시 규칙 기반 엔진으로 스마트 패스 및 15자 이내 제안을 생성합니다."""
    needs_revision, score, reason = _evaluate_title_impact_rule_based(current_title, instructor, target, features)

    if not needs_revision:
        return {
            "impact_score": score,
            "needs_revision": False,
            "status_badge": "✅ 기존 유지(패스)",
            "evaluation_reason": reason,
            "critique": reason,
            "recommended_title": "-",
            # 하위 호환 키
            "recommended_title_1": "-",
            "recommended_title_2": "-",
            "style_1_short": "-",
            "style_2_brand": "-",
            "style_3_punch": "-",
            "style_1_intuitive": "-",
            "style_2_signature": "-",
            "style_3_killer": "-",
            "engine": "대치동 스마트 규칙 엔진 (Fallback)",
        }

    # 수정 추천 대상 강좌: 대괄호 보존 + 본문 15자 이내 (2~3어절) 생성
    prefix, body = extract_bracket_prefix_and_body(current_title)
    clean_body = re.sub(r"\[[^\]]+\]|<[^>]+>", "", body).strip() or body or current_title
    clean_body = re.sub(r"\s+", " ", clean_body).strip()

    title_words = [w for w in clean_body.split() if len(w) >= 2]
    sub_detail = ""
    for w in title_words:
        if any(kw in w for kw in ("공수", "수학", "미적", "확통", "대수", "기하", "독서", "문학", "문법", "화작", "언매", "영어", "물리", "화학", "생명", "지구", "사문", "생윤", "한지", "세지")):
            sub_detail = w[:6]
            break
    if not sub_detail and title_words:
        sub_detail = title_words[0][:6]
    if not sub_detail:
        sub_detail = subject

    # PDF 특징이나 텍스트에서 킬러 키워드 추출
    feature_kw = "핵심 실전"
    search_context = f"{features} {management} {target} {' '.join(title_words)}"
    if any(k in search_context for k in ("킬러", "4점", "최상위", "1등급")):
        feature_kw = "1등급 킬러"
    elif any(k in search_context for k in ("시간 단축", "지문 장악", "독해")):
        feature_kw = "지문 장악"
    elif any(k in search_context for k in ("조건 해석", "구조화", "개념")):
        feature_kw = "조건 해석 개념"

    base_s = f"{sub_detail} {feature_kw} 마스터"
    rec = clean_and_truncate_title(base_s, original_prefix=prefix, max_body_len=15)

    if seen_titles is not None and rec in seen_titles:
        rec = clean_and_truncate_title(f"{instructor} {sub_detail} 실전", original_prefix=prefix, max_body_len=15)

    return {
        "impact_score": score,
        "needs_revision": True,
        "status_badge": "🔧 수정 추천",
        "evaluation_reason": reason,
        "critique": reason,
        "recommended_title": rec,
        # 하위 호환 키
        "recommended_title_1": rec,
        "recommended_title_2": rec,
        "style_1_short": rec,
        "style_2_brand": rec,
        "style_3_punch": rec,
        "style_1_intuitive": rec,
        "style_2_signature": rec,
        "style_3_killer": rec,
        "engine": "대치동 스마트 규칙 엔진 (Fallback)",
    }


def suggest_course_titles_with_gemini(
    course_info: Dict[str, Any],
    api_key: Optional[str] = None,
    direction: str = "스마트 분석",
    sibling_titles: Optional[List[str]] = None,
    timeout_sec: int = 10,
) -> Dict[str, Any]:
    """
    단일 강좌의 정보를 받아 Gemini 3.8을 통해
    강의계획서 PDF 컨텍스트(추천대상, 강의특징, 관리프로그램)를 기반으로
    1) '스마트 패스(기존 유지)' vs '수정 추천'을 자동 판정하고,
    2) 수정 추천 시 대괄호 태그('[...]')를 100% 보존하며 15자 이내 핵심 초단축 강의명을 제안합니다.
    """
    global _WORKING_MODEL

    subj = str(course_info.get("과목", "")).strip() or "과목"
    grade = str(course_info.get("학년", "") or course_info.get("대상학년", "")).strip() or "고등"
    inst = str(course_info.get("강사명", "") or course_info.get("강사", "")).strip() or "강사"
    curr_title = str(course_info.get("PDF 강좌명", "") or course_info.get("강좌명", "") or course_info.get("RMS 강좌명", "")).strip()
    target = str(course_info.get("추천대상", "")).strip()
    features = str(course_info.get("강의특징_요약", "") or course_info.get("강의특징", "") or course_info.get("주차별진도계획_요약", "")).strip()[:200]
    management = str(course_info.get("관리프로그램_요약", "") or course_info.get("관리프로그램", "")).strip()[:150]

    prefix, body = extract_bracket_prefix_and_body(curr_title)

    key_to_use = api_key or get_gemini_api_key()
    if not key_to_use:
        return _generate_fallback_suggestion(subj, grade, inst, curr_title, target, features, management, direction)

    sibling_note = ""
    if sibling_titles:
        sibling_note = f"\n[동일 강사({inst}T)의 다른 개설 강좌들 - 중복 절대 금지!]\n" + "\n".join(
            f"- {t}" for t in sibling_titles if t != curr_title
        )

    prefix_guide = ""
    if prefix:
        prefix_guide = (
            f"\n[★ 대괄호 태그 100% 무조건 보존]\n"
            f"- 원본 대괄호 태그: '{prefix}'\n"
            f"- 위 대괄호 태그는 절대 수정·삭제하지 말고 추천 강의명 맨 앞에 100% 그대로 붙이세요!\n"
            f"- 글자 수 제한(15자 이내)은 대괄호 태그를 완전히 제외한 '대괄호 뒤의 본문 제목'에만 적용됩니다.\n"
        )

    prompt = f"""당신은 대한민국 서울 대치동 최상위권 단과 학원의 수석 카피라이터입니다.
학원 브로슈어 및 시간표에 실릴 강의명들을 전문적으로 평가하고, 강의계획서 PDF의 핵심 내용을 바탕으로 '스마트 패스 & 수정 추천' 업무를 수행합니다.

[강좌 기본 정보]
- 과목 / 대상학년: {subj} / {grade}
- 강사명: {inst} 선생님
- 현재 강좌명: {curr_title}
- 원본 대괄호 태그: {prefix or "(없음)"}
- 현재 본문 강의명: {body or curr_title}

[강의계획서(PDF) 컨텍스트 정보]
- 추천 대상: {target or "(상위권 및 1등급 목표 수험생)"}
- 강의 특징 / 목표: {features or "(수능/내신 실전 핵심 개념 및 킬러 문항 공략)"}
- 관리 프로그램: {management or "(1:1 맞춤형 피드백 및 취약점 분석 클리닉)"}
{sibling_note}
{prefix_guide}

[★ 핵심 판정 지침: 1) 기존 유지(패스) vs 2) 수정 추천 스마트 판정 ★]
1. ✅ [패스 대상 - 기존 유지] (needs_revision: false, impact_score: 80~95점):
   - 기존 강의명에 이미 강사 고유의 시그니처 브랜드(예: '진심국어', '기고만장', '수풀림', 'META', '풀이학개론', 'Big Bang', 'PRIDE', '知認解' 등)가 잘 녹아있거나, 과목의 특장점이 직관적으로 잘 드러나 있어 충분히 임팩트가 있다고 판단되는 강좌.
   - 이 경우 새 이름을 제안하지 말고 반드시 recommended_title을 "-"로 출력하세요.
   - evaluation_reason에 강사 시그니처 브랜드나 우수한 점을 명확히 작성하세요.

2. 🔧 [수정 추천 대상 - 초단축 리브랜딩] (needs_revision: true, impact_score: 40~75점):
   - '고2 윈터 기본반', '<윈터> 대수 기본반', '[확률과통계] 개념유형 총정리', '예비고1 영어 Master실력반', '중동고 합격자반'처럼 단순히 학년·과목·기본반·총정리만 평이하게 나열되어 학부모/학생의 시선을 끌기 부족한 평이한 강좌명만 선별.
   - [수정 제안 엄격 규칙]:
     * 원본 강좌명 앞에 있는 대괄호 태그(예: `[고1 특강][실력반]`, `[고2 정규]`)는 글자 수에서 완전히 제외하고 그대로 맨 앞에 보존(유지)할 것.
     * 대괄호 뒤의 본문 제목은 대치동 브로슈어 시인성에 맞게 반드시 **공백 포함 15자 이내 (2~3어절)**의 짧고 강렬한 핵심 명사형/키워드 조합으로만 제안할 것. (절대 문장형이나 30자가 넘는 장문 금지!)
     * 강의계획서 PDF의 [추천 대상, 강의 특징, 관리 프로그램]에서 강좌의 핵심 무기(예: '시간 단축', '지문 장악', '조건 해석', '4점 킬러 돌파', '최상위 기출 완성' 등)를 적극 반영할 것.
     * 부제나 슬로건 금지! 서술형 수식어(~를 위한, ~하는 수업, 한 번에 끝내는 등) 금지!
     * 동일 강사의 다른 개설 강좌와 절대 중복되지 않도록 고유 포인트를 살릴 것.

반드시 아래 JSON 포맷으로만 응답하세요. (마크다운 백틱 및 추가 설명 금지, slogan 필드 금지)
{{
  "impact_score": 90,
  "needs_revision": false,
  "evaluation_reason": "강사 시그니처 네이밍과 과목 정체성이 우수하여 기존 강의명 유지 권장",
  "recommended_title": "-"
}}
또는
{{
  "impact_score": 45,
  "needs_revision": true,
  "evaluation_reason": "학년과 수준만 평이하게 나열됨 -> 시간 단축 및 지문 장악 포인트를 살린 초단축 네이밍 추천",
  "recommended_title": "{prefix + ' ' if prefix else ''}{subj} 핵심 실전 마스터"
}}"""

    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 300,
            "responseMimeType": "application/json",
        },
    }

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

                        if isinstance(parsed, list) and len(parsed) > 0:
                            parsed = parsed[0]

                        impact_score = int(parsed.get("impact_score", 70))
                        needs_revision = bool(parsed.get("needs_revision", impact_score < 80))
                        eval_reason = str(parsed.get("evaluation_reason") or parsed.get("critique", "")).strip()

                        raw_rec = str(
                            parsed.get("recommended_title")
                            or parsed.get("recommended_title_1")
                            or parsed.get("style_1_short")
                            or ""
                        ).strip()

                        if not needs_revision or impact_score >= 80:
                            needs_revision = False
                            status_badge = "✅ 기존 유지(패스)"
                            rec_title = "-"
                        else:
                            status_badge = "🔧 수정 추천"
                            def_body = f"{inst} {subj} 킬러 마스터"
                            rec_title = (
                                clean_and_truncate_title(raw_rec, original_prefix=prefix, max_body_len=15)
                                if raw_rec and raw_rec != "-"
                                else clean_and_truncate_title(def_body, original_prefix=prefix, max_body_len=15)
                            )

                        _WORKING_MODEL = cand_model

                        return {
                            "impact_score": impact_score,
                            "needs_revision": needs_revision,
                            "status_badge": status_badge,
                            "evaluation_reason": eval_reason,
                            "critique": eval_reason,
                            "recommended_title": rec_title,
                            # 하위 호환
                            "recommended_title_1": rec_title,
                            "recommended_title_2": rec_title,
                            "style_1_short": rec_title,
                            "style_2_brand": rec_title,
                            "style_3_punch": rec_title,
                            "style_1_intuitive": rec_title,
                            "style_2_signature": rec_title,
                            "style_3_killer": rec_title,
                            "engine": f"Gemini 3.8 ({cand_model})",
                        }
        except Exception:
            continue

    return _generate_fallback_suggestion(subj, grade, inst, curr_title, target, features, management, direction)


def batch_suggest_course_titles(
    course_list: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    direction: str = "스마트 분석",
    progress_callback: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    """
    여러 강좌에 대해 순차적으로 PDF 컨텍스트 기반 스마트 패스 & 수정 추천 분석을 실행합니다:
    - PDF(추천대상, 강의특징, 관리프로그램) 내용을 자동 검토하여 스마트 패스 vs 수정 추천 자동 판정
    - 원본 대괄호 태그('[...]') 100% 무조건 보존
    - 수정 추천 대상은 본문 15자 이내 (2~3어절) 핵심 명사형 리브랜딩 적용
    - 슬로건 추천 완전 삭제 및 동일 강사 중복 원천 차단
    """
    results: List[Dict[str, Any]] = []
    total = len(course_list)

    # 강사별 개설 강좌명 리스트 사전 수집
    instructor_courses_map: Dict[str, List[str]] = {}
    for c in course_list:
        inst = str(c.get("강사명", "") or c.get("강사", "")).strip()
        t = str(c.get("PDF 강좌명", "") or c.get("강좌명", "") or c.get("RMS 강좌명", "")).strip()
        if inst and t:
            instructor_courses_map.setdefault(inst, []).append(t)

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
        status_badge = sug.get("status_badge", "🔧 수정 추천" if needs_rev else "✅ 기존 유지(패스)")
        score = sug.get("impact_score", 70)
        reason = sug.get("evaluation_reason", "")
        rec_title = sug.get("recommended_title", "-")

        curr_t = str(c.get("PDF 강좌명", "") or c.get("강좌명", "") or c.get("RMS 강좌명", "")).strip()
        orig_prefix, orig_body = extract_bracket_prefix_and_body(curr_t)

        # 수정 추천인 경우 중복 회피 처리 (대괄호 태그 100% 보존, 15자 이내)
        if needs_rev and rec_title != "-":
            subj = str(c.get("과목", "")).strip()
            clean_curr = re.sub(r"\[[^\]]+\]|<[^>]+>", "", orig_body).strip()
            curr_words = clean_curr.split()
            kw_spec = curr_words[0][:5] if curr_words else subj

            if rec_title in seen_titles:
                rec_title = clean_and_truncate_title(f"{kw_spec} 핵심 실전", original_prefix=orig_prefix, max_body_len=15)
            if rec_title in seen_titles:
                rec_title = clean_and_truncate_title(f"{inst} {kw_spec} 마스터", original_prefix=orig_prefix, max_body_len=15)

            seen_titles.add(rec_title)

        combined = {
            "순번": idx,
            "판정": status_badge,
            "과목": c.get("과목", ""),
            "학년": c.get("학년", "") or c.get("대상학년", ""),
            "강사명": inst,
            "현재 강좌명": curr_t,
            "임팩트 점수": score,
            "판정 사유 및 피드백": reason,
            "⚡ 추천 강의명": rec_title,
            # 하위 호환 키 완벽 지원
            "판정결과": status_badge,
            "needs_revision": needs_rev,
            "현재강의명": curr_t,
            "현재강좌명": curr_t,
            "임팩트점수": score,
            "판정사유": reason,
            "진단비평": reason,
            "⚡ 추천 강의명 1 (직관·시그니처형)": rec_title,
            "🔥 추천 강의명 2 (킬러·임팩트형)": rec_title,
            "제안1_초단축직관형": rec_title,
            "제안2_초단축시그니처": rec_title,
            "제안3_초단축킬러형": rec_title,
            "생성엔진": sug.get("engine", ""),
            "최종선택강좌명": curr_t if not needs_rev else rec_title,
        }
        results.append(combined)

        if progress_callback:
            try:
                progress_callback(idx, total, curr_t)
            except Exception:
                pass

    return results
