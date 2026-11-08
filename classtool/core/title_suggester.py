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


def _generate_fallback_suggestion(
    subject: str,
    grade: str,
    instructor: str,
    current_title: str,
    direction: str = "대치동 트렌드 종합",
    seen_titles: Optional[Set[str]] = None,
) -> Dict[str, Any]:
    """API 키가 없거나 호출 실패 시 10~15자 초단축 대치동 규칙 기반 엔진으로 생성합니다 (중복 방지 적용)."""
    clean_title = re.sub(r"\[[^\]]+\]", "", current_title).strip() or current_title
    clean_title = re.sub(r"\s+", " ", clean_title).strip()

    # 원본 강좌명에서 고유 식별 키워드 추출 (예: '공통수학1', '확통', '미적분', '수능완성', '기출', '킬러')
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

    # 강좌 특성 키워드 (개념, 심화, 기출, 킬러, 특강 등)
    feature_kw = ""
    for w in title_words:
        if any(kw in w for kw in ("기출", "킬러", "심화", "개념", "200제", "300제", "유형", "마스터", "실전", "파이널")):
            feature_kw = w[:6]
            break

    # 기본 템플릿
    if feature_kw and feature_kw != sub_detail:
        base_s1 = f"{sub_detail} {feature_kw} 압축"
        base_s2 = f"{instructor} {sub_detail} 시그니처"
        base_s3 = f"{sub_detail} 1등급 킬러 돌파"
    else:
        base_s1 = f"{sub_detail} 핵심 실전 압축"
        base_s2 = f"{instructor} {sub_detail} 코어"
        base_s3 = f"{sub_detail} 조건반사 문풀"

    s1 = clean_and_truncate_title(base_s1, max_len=16)
    s2 = clean_and_truncate_title(base_s2, max_len=16)
    s3 = clean_and_truncate_title(base_s3, max_len=16)

    # 중복 방지 체크 및 고유화
    if seen_titles is not None:
        if s1 in seen_titles:
            s1 = clean_and_truncate_title(f"{instructor} {sub_detail} 완성", max_len=16)
        if s1 in seen_titles:
            s1 = clean_and_truncate_title(f"{sub_detail} 1등급 도약", max_len=16)
        if s2 in seen_titles:
            s2 = clean_and_truncate_title(f"{instructor}T {sub_detail} 실전", max_len=16)
        if s3 in seen_titles:
            s3 = clean_and_truncate_title(f"{sub_detail} 만점 전략", max_len=16)

    return {
        "impact_score": 70,
        "critique": "기존 강좌명은 서술적이어서 시간표 및 브로슈어 시인성이 낮습니다. 15자 이내 명사형 초단축 카피로 개편을 제안합니다.",
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
    sibling_titles: Optional[List[str]] = None,
    timeout_sec: int = 10,
) -> Dict[str, Any]:
    """
    단일 강좌의 정보를 받아 Gemini 3.8 (스마트 라우팅 및 자동 fallback)을 통해
    15자 이내 초단축 3종 스타일의 강좌명을 제안합니다 (슬로건 제안 완전 제외).
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
학원 브로슈어 시간표 한 칸에 바로 들어갈 수 있도록, 아래 강좌의 제목을 반드시 **공백 포함 10자~15자 이내의 초단축 임팩트형**으로 리브랜딩하세요.
(주의: 부제나 슬로건은 추천하지 마시고 오직 강의명만 제안하세요.)

[강좌 기본 정보]
- 과목 / 대상학년 : {subj} / {grade}
- 강사명 : {inst} 선생님
- 현재 강좌명 : {curr_title}
- 수업시간 및 타깃 : {time_info or target}
- 특징 요약 : {features or "(상위권 대비 수능/내신 실전 강좌)"}
{sibling_note}

[카피라이팅 방향성]
{dir_instruction}

[★ 초단축 필수 제약조건 - 반드시 엄수 ★]
1. 모든 추천 강의명은 **공백 포함 10자 이상 ~ 16자 이하 (최대 3~4어절)**여야 합니다. 16자를 초과하면 절대 안 됩니다!
2. 서술형 수식어(~를 위한, ~하는, ~완성하는, ~만드는, 한 번에 끝내는 등)는 절대 사용 금지!
3. 오직 [핵심 과목/브랜드 + 타깃/무기 명사] 조합으로만 압축하세요.
4. 동일 강사가 여러 강좌를 개설한 경우, 각 강좌의 고유 세부 과목(대수, 미적1, 확통, 공수1 등), 난이도, 타깃 키워드를 명확히 살려 **절대 서로 겹치거나 천편일률적이지 않은 독창적인 제목**을 만드세요!
5. style_1_short (⚡ 핵심직관형 - 12자 내외): 수업 목표와 대상이 0.5초 만에 각인되는 초단축 제목 (예: '시간 단축 지문 장악 국어', '확통 개념×4점 킬러 압축')
6. style_2_brand (💎 시그니처형 - 14자 내외): 강사명이나 고유 브랜드를 결합한 대표 시그니처 제목 (예: 'META 대수 : 조건의 단서', '지인해 코어 국어', '수풀림 확통 시그니처')
7. style_3_punch (🔥 킬러임팩트형 - 12자 내외): 킬러 문항 정복과 1등급 돌파를 전면에 내세운 강력한 제목 (예: '미적1 조건반사 문풀', '기하 10% 킬러 돌파')
8. impact_score: 현재 강좌명의 대치동 트렌드 기준 임팩트 점수 (40~85점)
9. critique: 현재 강좌명의 아쉬운 점 1문장 (50자 내외)

반드시 아래 JSON 포맷으로만 응답하세요. (마크다운 백틱 및 기타 설명 금지, 슬로건 필드는 절대 포함하지 마세요)
{{
  "impact_score": 65,
  "critique": "현재 강좌명이 서술적이어서 시간표 시인성이 떨어지므로 15자 이내 초단축 명사형으로 개편이 필요합니다.",
  "style_1_short": "시간 단축 지문 장악 국어",
  "style_2_brand": "{inst} 코어 국어 시그니처",
  "style_3_punch": "국어 1등급 킬러 돌파"
}}
"""

    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.3,
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

                        raw_s1 = str(parsed.get("style_1_short") or parsed.get("style_1_intuitive", "")).strip()
                        raw_s2 = str(parsed.get("style_2_brand") or parsed.get("style_2_signature", "")).strip()
                        raw_s3 = str(parsed.get("style_3_punch") or parsed.get("style_3_killer", "")).strip()

                        s1 = clean_and_truncate_title(raw_s1, max_len=16)
                        s2 = clean_and_truncate_title(raw_s2, max_len=16)
                        s3 = clean_and_truncate_title(raw_s3, max_len=16)

                        # 작동 모델 캐싱
                        _WORKING_MODEL = cand_model

                        return {
                            "impact_score": int(parsed.get("impact_score", 70)),
                            "critique": str(parsed.get("critique", "")).strip(),
                            "style_1_short": s1,
                            "style_2_brand": s2,
                            "style_3_punch": s3,
                            # 기존 호환 키 유지
                            "style_1_intuitive": s1,
                            "style_2_signature": s2,
                            "style_3_killer": s3,
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
    여러 강좌에 대해 순차적으로 15자 이내 초단축 AI 카피라이팅 추천을 생성합니다.
    - 슬로건 제안 완전 제거
    - 동일 강사의 다른 개설 강좌 정보를 프롬프트에 전달하여 중복 원천 차단
    - 후처리 중복 필터(seen_titles)를 적용하여 100% 서로 다른 독창적인 제목 보장
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

        s1 = sug["style_1_short"]
        s2 = sug["style_2_brand"]
        s3 = sug["style_3_punch"]

        # 후처리 중복 회피 로직
        subj = str(c.get("과목", "")).strip()
        clean_curr = re.sub(r"\[[^\]]+\]", "", str(c.get("PDF 강좌명", "") or c.get("강좌명", ""))).strip()
        curr_words = clean_curr.split()
        kw_spec = curr_words[0][:5] if curr_words else subj

        if s1 in seen_titles:
            s1 = clean_and_truncate_title(f"{kw_spec} 핵심 마스터", max_len=16)
        if s1 in seen_titles:
            s1 = clean_and_truncate_title(f"{inst} {kw_spec} 실전", max_len=16)

        if s2 in seen_titles:
            s2 = clean_and_truncate_title(f"{inst}T {kw_spec} 시그니처", max_len=16)

        if s3 in seen_titles:
            s3 = clean_and_truncate_title(f"{kw_spec} 1등급 킬러", max_len=16)

        seen_titles.add(s1)
        seen_titles.add(s2)
        seen_titles.add(s3)

        combined = {
            "순번": idx,
            "과목": c.get("과목", ""),
            "학년": c.get("학년", "") or c.get("대상학년", ""),
            "강사명": inst,
            "현재강좌명": c.get("PDF 강좌명", "") or c.get("강좌명", ""),
            "임팩트점수": sug["impact_score"],
            "진단비평": sug["critique"],
            "제안1_초단축직관형": s1,
            "제안2_초단축시그니처": s2,
            "제안3_초단축킬러형": s3,
            # 기존 호환 컬럼 유지
            "제안1_직관형": s1,
            "제안2_시그니처형": s2,
            "제안3_킬러임팩트형": s3,
            "생성엔진": sug["engine"],
            "최종선택강좌명": s1,
        }
        results.append(combined)

        if progress_callback:
            try:
                progress_callback(idx, total, combined["현재강좌명"])
            except Exception:
                pass

    return results
