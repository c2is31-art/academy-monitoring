"""
Gemini API 기반 대치동 단과 '강의명 & 슬로건 임팩트 수정 제안' 모듈 (core/title_suggester.py)
- 대치동 학원가 최상위권 브로슈어 트렌드를 반영한 AI 카피라이팅
- requests 기반 Gemini 2.5 Flash REST API 직접 호출 (SDK 버전 충돌 방지)
- 4대 카피라이팅 방향성 지원:
    1) 대치동 트렌드 종합
    2) 내신 1등급/학교별 기출 강조
    3) 수능 평가원/실전 킬러 강조
    4) 개념 완성/중위권 도약 강조
- 3가지 차별화된 스타일 제안:
    - style_1_intuitive (🎯 타깃·성적향상 직관형)
    - style_2_signature (💎 시그니처·브랜딩형)
    - style_3_killer (🔥 최상위권·실전압도 임팩트형)
- 슬로건(부제) 및 1줄 비평, 임팩트 점수(100점 만점) 산출
- API 키 부재 또는 호출 실패 시 내장 규칙 기반 알고리즘으로 안전하게 Fallback
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List, Optional

import requests

DIRECTION_PROMPTS = {
    "대치동 트렌드 종합": (
        "대치동 단과 최상위권 트렌드에 맞게 학부모와 학생의 시선을 3초 안에 사로잡고, "
        "의치약한수 및 SKY 진학을 위한 압도적인 실전 경쟁력을 드러내는 카피라이팅을 적용하세요."
    ),
    "내신 1등급/학교별 기출 강조": (
        "휘문·중동·단대·경기·숙명·은광 등 대치/강남 주요 고교 내신 전교 1등 및 1등급 쟁취를 위한 "
        "학교별 기출 변형, 서술형 문항 완전 정복, 내신 적중력을 극대화하여 강조하세요."
    ),
    "수능 평가원/실전 킬러 강조": (
        "평가원 수능 최신 출제 트렌드, 준킬러 시간 단축, 킬러 문항의 구조적 풀이법 체화 및 "
        "수능 만점을 향한 실전 감각을 극대화하여 표현하세요."
    ),
    "개념 완성/중위권 도약 강조": (
        "흔들리는 개념의 빈틈을 완벽하게 메우고, 취약 유형을 집중 공략하여 "
        "중위권(2~3등급)에서 확실한 안정적 1등급으로 도약할 수 있는 직관적이고 친절하면서도 자신감 넘치는 표현을 쓰세요."
    ),
}


def get_gemini_api_key() -> Optional[str]:
    """
    환경 변수 및 Streamlit secrets에서 Gemini API 키를 안전하게 탐색합니다.
    (StreamlitSecretNotFoundError 방어)
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


def _generate_fallback_suggestion(
    subject: str,
    grade: str,
    instructor: str,
    current_title: str,
    slogan: str,
    direction: str = "대치동 트렌드 종합",
) -> Dict[str, Any]:
    """API 키가 없거나 호출 실패 시 대치동 단과 규칙 기반으로 3종 스타일을 생성합니다."""
    # 순수 강좌명 추출
    clean_title = re.sub(r"\[[^\]]+\]", "", current_title).strip() or current_title

    # 과목별 핵심 키워드 매핑
    sub_map = {
        "국어": ("독서·문학·문법", "1등급 독해 원리", "수능·내신 완벽 적중"),
        "수학": ("공통수학·실전개념", "킬러·준킬러 구조화", "오답률 0% 도전"),
        "영어": ("빈칸·순서·삽입", "구문 독해 마스터", "절대평가 1등급 완성"),
        "사회탐구": ("통합사회·사탐", "필수 개념·선지 분석", "내신·수능 만점 완성"),
        "과학탐구": ("통합과학·과탐", "자료해석·고난도 킬러", "전교 1등 등극"),
        "논술": ("논구술·인문수리", "채점 기준 부합 답안", "합격의 결정적 답안"),
    }
    k1, k2, k3 = sub_map.get(subject, ("핵심 개념", "실전 문제풀이", "1등급 도약"))

    # 슬로건 제안
    if slogan and len(slogan) > 8:
        slogan_sug = f"상위 1%를 만드는 {instructor}T의 {k2}"
    else:
        slogan_sug = f"대치동이 인정한 단 하나의 선택, {instructor}T {subject} 완성"

    # 스타일 1: 타깃/성적향상 직관형
    style_1 = f"[{grade} {subject}] {instructor}T의 성적으로 증명하는 {clean_title}"

    # 스타일 2: 시그니처/브랜딩형
    style_2 = f"[{instructor} 시그니처] {clean_title} : 대치 최상위 실전 완성"

    # 스타일 3: 최상위권/실전압도 임팩트형
    style_3 = f"[1등급 킬러정복] {instructor}T {clean_title} (상위 1% 실전압도)"

    return {
        "impact_score": 70,
        "critique": "기존 강좌명은 표준적이나, 대치동 최상위권 학부모의 이목을 끄는 타깃 및 실전 성적 향상 임팩트가 다소 부족합니다.",
        "slogan_suggestion": slogan_sug,
        "style_1_intuitive": style_1,
        "style_2_signature": style_2,
        "style_3_killer": style_3,
        "engine": "대치동 규칙 기반 엔진 (Fallback)",
    }


def suggest_course_titles_with_gemini(
    course_info: Dict[str, Any],
    api_key: Optional[str] = None,
    direction: str = "대치동 트렌드 종합",
    model_name: str = "gemini-2.5-flash",
    timeout_sec: int = 10,
) -> Dict[str, Any]:
    """
    단일 강좌의 정보를 받아 Gemini 2.5 Flash를 통해 3종 스타일의 강좌명 및 슬로건을 제안합니다.
    """
    subj = str(course_info.get("과목", "")).strip() or "과목"
    grade = str(course_info.get("학년", "") or course_info.get("대상학년", "")).strip() or "고등"
    inst = str(course_info.get("강사명", "") or course_info.get("강사", "")).strip() or "강사"
    curr_title = str(course_info.get("PDF 강좌명", "") or course_info.get("강좌명", "")).strip()
    curr_slogan = str(course_info.get("슬로건", "") or course_info.get("상단부제", "")).strip()
    target = str(course_info.get("추천대상", "")).strip() or "상위권"
    features = str(course_info.get("강의특징_요약", "") or course_info.get("주차별진도계획_요약", "")).strip()[:250]

    key_to_use = api_key or get_gemini_api_key()
    if not key_to_use:
        return _generate_fallback_suggestion(subj, grade, inst, curr_title, curr_slogan, direction)

    dir_instruction = DIRECTION_PROMPTS.get(direction, DIRECTION_PROMPTS["대치동 트렌드 종합"])

    prompt = f"""
당신은 대한민국 사교육 1번지 서울 대치동 최상위권 단과 학원의 전설적인 스타 강사 전담 마케팅 디렉터이자 카피라이터입니다.
아래 강좌의 기본 정보를 분석하여, 학생과 학부모가 강의계획서 브로슈어를 펼쳤을 때 3초 안에 압도적인 신뢰감과 수강 욕구를 느끼도록
[1] 현재 강좌명 진단, [2] 상단 슬로건(부제) 제안, [3] 3가지 차별화된 스타일의 강좌명을 생성하세요.

[강좌 기본 정보]
- 과목 / 학년 : {subj} / {grade}
- 강사명 : {inst} 선생님
- 현재 강좌명 : {curr_title}
- 현재 부제/슬로건 : {curr_slogan or "(없음)"}
- 추천 대상 : {target}
- 강의 특징 요약 : {features or "(수능 및 내신 상위권 대비 강의)"}

[카피라이팅 방향성 지침]
{dir_instruction}

[필수 요구사항]
1. impact_score: 현재 강좌명의 대치동 트렌드 기준 임팩트 점수 (40~85점 사이 정수)
2. critique: 현재 강좌명에서 아쉬운 점 또는 보완할 핵심 포인트 1문장 (50자 내외)
3. slogan_suggestion: 강좌명 위에 작게 들어갈 25자 내외의 강렬하고 세련된 한 줄 슬로건
4. style_1_intuitive (🎯 타깃·성적향상 직관형): 어떤 학생이 들으면 즉각 성적이 오르는지 직관적으로 꽂히는 강좌명 (30~45자)
5. style_2_signature (💎 시그니처·브랜딩형): 강사의 이름, 브랜드명, 또는 고유 철학이 담긴 대치동 시그니처 프리미엄 강좌명 (30~45자)
6. style_3_killer (🔥 최상위권·실전압도 임팩트형): 1등급, 킬러 문항 정복, 만점 도약을 전면에 내세운 강렬하고 자신감 넘치는 강좌명 (30~45자)

반드시 아래 JSON 포맷으로만 응답하세요. (마크다운 코드블록이나 불필요한 설명 금지)
{{
  "impact_score": 65,
  "critique": "현재 강좌명은 개념과 유형 중심이라 실전 수능 및 내신 1등급을 노리는 대치동 상위권의 심리적 결핍을 자극하기에 다소 밋밋합니다.",
  "slogan_suggestion": "단 1%의 빈틈도 허용하지 않는 압도적 실전력",
  "style_1_intuitive": "[{grade} {subj}] 8주 만에 끝내는 실전 개념과 1등급 킬러 완벽 정립",
  "style_2_signature": "[{inst} 시그니처] {curr_title} : 최상위권 도약 마스터피스",
  "style_3_killer": "[1등급 킬러격파] {inst}T 최상위 압도적 실전 전략"
}}
"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={key_to_use}"
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.35,
            "maxOutputTokens": 600,
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
                    # 마크다운 백틱 및 json 레이블 정제
                    clean_json = re.sub(r"^```(?:json)?\s*", "", raw_text)
                    clean_json = re.sub(r"\s*```$", "", clean_json).strip()
                    parsed = json.loads(clean_json)

                    return {
                        "impact_score": int(parsed.get("impact_score", 70)),
                        "critique": str(parsed.get("critique", "")).strip(),
                        "slogan_suggestion": str(parsed.get("slogan_suggestion", "")).strip(),
                        "style_1_intuitive": str(parsed.get("style_1_intuitive", "")).strip(),
                        "style_2_signature": str(parsed.get("style_2_signature", "")).strip(),
                        "style_3_killer": str(parsed.get("style_3_killer", "")).strip(),
                        "engine": f"Gemini 2.5 Flash ({direction})",
                    }
    except Exception:
        pass

    # 오류 또는 실패 시 fallback 반환
    return _generate_fallback_suggestion(subj, grade, inst, curr_title, curr_slogan, direction)


def batch_suggest_course_titles(
    course_list: List[Dict[str, Any]],
    api_key: Optional[str] = None,
    direction: str = "대치동 트렌드 종합",
    progress_callback: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    """
    여러 강좌에 대해 순차적으로 AI 카피라이팅 추천을 생성합니다.
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
            "제안1_직관형": sug["style_1_intuitive"],
            "제안2_시그니처형": sug["style_2_signature"],
            "제안3_킬러임팩트형": sug["style_3_killer"],
            "생성엔진": sug["engine"],
            "최종선택강좌명": sug["style_1_intuitive"],
        }
        results.append(combined)

        if progress_callback:
            try:
                progress_callback(idx, total, combined["현재강좌명"])
            except Exception:
                pass

    return results
