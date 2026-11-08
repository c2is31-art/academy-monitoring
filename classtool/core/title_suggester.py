"""
Gemini API 기반 대치동 단과 '15자 이내 초단축·임팩트형 강의명 & 슬로건 제안' 모듈 (core/title_suggester.py)
- 대치동 학원가 최상위권 브로슈어 및 시간표에 최적화된 초단축(공백 포함 10자~16자 이내) AI 카피라이팅
- requests 기반 Gemini 2.5 Flash REST API 직접 호출 (SDK 버전 충돌 방지)
- 4대 카피라이팅 방향성 지원:
    1) 대치동 트렌드 종합
    2) 내신 1등급/학교별 기출 강조
    3) 수능 평가원/실전 킬러 강조
    4) 개념 완성/중위권 도약 강조
- 3가지 초단축 스타일 제안 (모두 10~16자 엄격 준수):
    - style_1_short (⚡ 핵심직관형 - 12자 내외, 예: '시간 단축 지문 장악 국어', '확통 개념×4점 킬러 압축')
    - style_2_brand (💎 시그니처형 - 14자 내외, 예: 'META 대수 : 조건의 단서', '지인해(知認解) 코어 국어', '수풀림 확통 시그니처')
    - style_3_punch (🔥 킬러임팩트형 - 12자 내외, 예: '미적1 조건반사 문풀', '기하 10% 킬러 돌파')
- 상단 슬로건(slogan_suggestion): 20자 이내의 강렬한 카피
- 서술형 수식어(~를 위한, ~하는, ~완성하는 등) 금지 후처리 및 글자 수 초과 시 자동 압축 알고리즘 탑재
- API 키 부재 또는 호출 실패 시 고품질 대치동 초단축 규칙 기반 엔진(Fallback) 가동
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List, Optional

import requests

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


def clean_and_truncate_slogan(slogan: str, max_len: int = 20) -> str:
    """슬로건을 20자 이내로 정제합니다."""
    if not slogan:
        return ""
    s = re.sub(r"^['\"`]+|['\"`]+$", "", str(slogan)).strip()
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) > max_len:
        s = s[:max_len].rstrip()
    return s


def _generate_fallback_suggestion(
    subject: str,
    grade: str,
    instructor: str,
    current_title: str,
    slogan: str,
    direction: str = "대치동 트렌드 종합",
) -> Dict[str, Any]:
    """API 키가 없거나 호출 실패 시 10~15자 초단축 대치동 규칙 기반 엔진으로 생성합니다."""
    clean_title = re.sub(r"\[[^\]]+\]", "", current_title).strip() or current_title
    clean_title = re.sub(r"\s+", " ", clean_title).strip()

    # 핵심 키워드 1~2단어 추출
    title_words = clean_title.split()
    core_kw = title_words[0] if title_words else subject
    if len(core_kw) > 6:
        core_kw = core_kw[:6]

    # 과목별 초단축 시그니처 템플릿
    sub_presets = {
        "국어": {
            "short": f"{core_kw} 지문 장악 국어",
            "brand": f"{instructor} 코어 국어",
            "punch": f"{core_kw} 1등급 킬러 돌파",
            "slogan": "지문 장악의 압도적 시간 단축",
        },
        "수학": {
            "short": f"{core_kw} 개념×4점 킬러",
            "brand": f"{instructor} {core_kw} 시그니처",
            "punch": f"{core_kw} 조건반사 문풀",
            "slogan": "4점 킬러를 꿰뚫는 단서와 해법",
        },
        "영어": {
            "short": f"{core_kw} 구문 독해 압축",
            "brand": f"{instructor} 로직 영어",
            "punch": f"{core_kw} 빈칸 순서 정복",
            "slogan": "절대평가 1등급 실전 루틴",
        },
        "사회탐구": {
            "short": f"{core_kw} 선지 분석 마스터",
            "brand": f"{instructor} 사탐 코어",
            "punch": f"{core_kw} 만점 선지 돌파",
            "slogan": "오답률 0%에 도전하는 개념 완성",
        },
        "과학탐구": {
            "short": f"{core_kw} 고난도 자료해석",
            "brand": f"{instructor} 과탐 시그니처",
            "punch": f"{core_kw} 킬러 조건 공략",
            "slogan": "전교 1등을 결정짓는 킬러 정복",
        },
    }

    preset = sub_presets.get(
        subject,
        {
            "short": f"{core_kw} 핵심 실전 압축",
            "brand": f"{instructor} {subject} 시그니처",
            "punch": f"{core_kw} 1등급 킬러 돌파",
            "slogan": "상위 1%를 만드는 압도적 실전력",
        },
    )

    s1 = clean_and_truncate_title(preset["short"], max_len=16)
    s2 = clean_and_truncate_title(preset["brand"], max_len=16)
    s3 = clean_and_truncate_title(preset["punch"], max_len=16)
    slogan_clean = clean_and_truncate_slogan(preset["slogan"], max_len=20)

    return {
        "impact_score": 70,
        "critique": "기존 강좌명은 서술적이어서 시간표 및 브로슈어 시인성이 낮습니다. 15자 이내 명사형 초단축 카피로 개편을 제안합니다.",
        "slogan_suggestion": slogan_clean,
        "style_1_short": s1,
        "style_2_brand": s2,
        "style_3_punch": s3,
        # 기존 호환 키 유지
        "style_1_intuitive": s1,
        "style_2_signature": s2,
        "style_3_killer": s3,
        "engine": "대치동 초단축 규칙 엔진 (Fallback)",
    }


def suggest_course_titles_with_gemini(
    course_info: Dict[str, Any],
    api_key: Optional[str] = None,
    direction: str = "대치동 트렌드 종합",
    model_name: str = "gemini-2.5-flash",
    timeout_sec: int = 10,
) -> Dict[str, Any]:
    """
    단일 강좌의 정보를 받아 Gemini 2.5 Flash를 통해 15자 이내 초단축 3종 스타일의 강좌명 및 슬로건을 제안합니다.
    """
    subj = str(course_info.get("과목", "")).strip() or "과목"
    grade = str(course_info.get("학년", "") or course_info.get("대상학년", "")).strip() or "고등"
    inst = str(course_info.get("강사명", "") or course_info.get("강사", "")).strip() or "강사"
    curr_title = str(course_info.get("PDF 강좌명", "") or course_info.get("강좌명", "")).strip()
    curr_slogan = str(course_info.get("슬로건", "") or course_info.get("상단부제", "")).strip()
    target = str(course_info.get("추천대상", "")).strip() or "상위권"
    features = str(course_info.get("강의특징_요약", "") or course_info.get("주차별진도계획_요약", "")).strip()[:200]

    key_to_use = api_key or get_gemini_api_key()
    if not key_to_use:
        return _generate_fallback_suggestion(subj, grade, inst, curr_title, curr_slogan, direction)

    dir_instruction = DIRECTION_PROMPTS.get(direction, DIRECTION_PROMPTS["대치동 트렌드 종합"])

    prompt = f"""
당신은 대한민국 서울 대치동 최상위권 단과 학원의 스타 강사 전문 수석 카피라이터입니다.
학원 브로슈어 시간표 한 칸에 바로 들어갈 수 있도록, 아래 강좌의 제목을 반드시 **공백 포함 10자~15자 이내의 초단축 임팩트형**으로 리브랜딩하세요.

[강좌 기본 정보]
- 과목 / 대상학년 : {subj} / {grade}
- 강사명 : {inst} 선생님
- 현재 강좌명 : {curr_title}
- 현재 부제/슬로건 : {curr_slogan or "(없음)"}
- 특징 요약 : {features or "(상위권 대비 수능/내신 실전 강좌)"}

[카피라이팅 방향성]
{dir_instruction}

[★ 초단축 필수 제약조건 - 반드시 엄수 ★]
1. 모든 추천 강의명은 **공백 포함 10자 이상 ~ 16자 이하 (최대 3~4어절)**여야 합니다. 16자를 초과하면 절대 안 됩니다!
2. 서술형 수식어(~를 위한, ~하는, ~완성하는, ~만드는, 한 번에 끝내는 등)는 절대 사용 금지!
3. 오직 [핵심 과목/브랜드 + 타깃/무기 명사] 조합으로만 압축하세요.
4. style_1_short (⚡ 핵심직관형 - 12자 내외): 수업 목표와 대상이 0.5초 만에 각인되는 초단축 제목 (예: '시간 단축 지문 장악 국어', '확통 개념×4점 킬러 압축')
5. style_2_brand (💎 시그니처형 - 14자 내외): 강사명이나 고유 브랜드를 결합한 대표 시그니처 제목 (예: 'META 대수 : 조건의 단서', '지인해 코어 국어', '수풀림 확통 시그니처')
6. style_3_punch (🔥 킬러임팩트형 - 12자 내외): 킬러 문항 정복과 1등급 돌파를 전면에 내세운 강력한 제목 (예: '미적1 조건반사 문풀', '기하 10% 킬러 돌파')
7. slogan_suggestion: 강좌명 위에 작게 들어갈 **20자 이내**의 세련된 한 줄 카피
8. impact_score: 현재 강좌명의 대치동 트렌드 기준 임팩트 점수 (40~85점)
9. critique: 현재 강좌명의 아쉬운 점 1문장 (50자 내외)

반드시 아래 JSON 포맷으로만 응답하세요. (마크다운 백틱 및 기타 설명 금지)
{{
  "impact_score": 65,
  "critique": "현재 강좌명이 서술적이어서 시간표 시인성이 떨어지므로 15자 이내 초단축 명사형으로 개편이 필요합니다.",
  "slogan_suggestion": "지문 장악의 압도적 시간 단축",
  "style_1_short": "시간 단축 지문 장악 국어",
  "style_2_brand": "{inst} 코어 국어 시그니처",
  "style_3_punch": "국어 1등급 킬러 돌파"
}}
"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={key_to_use}"
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.3,
            "maxOutputTokens": 450,
        },
    }

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

                    raw_s1 = str(parsed.get("style_1_short") or parsed.get("style_1_intuitive", "")).strip()
                    raw_s2 = str(parsed.get("style_2_brand") or parsed.get("style_2_signature", "")).strip()
                    raw_s3 = str(parsed.get("style_3_punch") or parsed.get("style_3_killer", "")).strip()
                    raw_slogan = str(parsed.get("slogan_suggestion", "")).strip()

                    s1 = clean_and_truncate_title(raw_s1, max_len=16)
                    s2 = clean_and_truncate_title(raw_s2, max_len=16)
                    s3 = clean_and_truncate_title(raw_s3, max_len=16)
                    slogan_clean = clean_and_truncate_slogan(raw_slogan, max_len=20)

                    return {
                        "impact_score": int(parsed.get("impact_score", 70)),
                        "critique": str(parsed.get("critique", "")).strip(),
                        "slogan_suggestion": slogan_clean,
                        "style_1_short": s1,
                        "style_2_brand": s2,
                        "style_3_punch": s3,
                        # 기존 호환 키 유지
                        "style_1_intuitive": s1,
                        "style_2_signature": s2,
                        "style_3_killer": s3,
                        "engine": f"Gemini 2.5 Flash ({direction})",
                    }
    except Exception:
        pass

    return _generate_fallback_suggestion(subj, grade, inst, curr_title, curr_slogan, direction)


def batch_suggest_course_titles(
    course_list: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    direction: str = "대치동 트렌드 종합",
    progress_callback: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    """
    여러 강좌에 대해 순차적으로 15자 이내 초단축 AI 카피라이팅 추천을 생성합니다.
    """
    results: List[Dict[str, Any]] = []
    total = len(course_list)

    for idx, c in enumerate(course_list, start=1):
        sug = suggest_course_titles_with_gemini(
            course_info=c,
            api_key=api_key,
            direction=direction,
        )
        combined = {
            "순번": idx,
            "과목": c.get("과목", ""),
            "학년": c.get("학년", "") or c.get("대상학년", ""),
            "강사명": c.get("강사명", "") or c.get("강사", ""),
            "현재강좌명": c.get("PDF 강좌명", "") or c.get("강좌명", ""),
            "현재슬로건": c.get("슬로건", "") or c.get("상단부제", "") or "-",
            "임팩트점수": sug["impact_score"],
            "진단비평": sug["critique"],
            "추천슬로건": sug["slogan_suggestion"],
            "제안1_초단축직관형": sug["style_1_short"],
            "제안2_초단축시그니처": sug["style_2_brand"],
            "제안3_초단축킬러형": sug["style_3_punch"],
            # 기존 호환 컬럼 유지
            "제안1_직관형": sug["style_1_short"],
            "제안2_시그니처형": sug["style_2_brand"],
            "제안3_킬러임팩트형": sug["style_3_punch"],
            "생성엔진": sug["engine"],
            "최종선택강좌명": sug["style_1_short"],
        }
        results.append(combined)

        if progress_callback:
            try:
                progress_callback(idx, total, combined["현재강좌명"])
            except Exception:
                pass

    return results
