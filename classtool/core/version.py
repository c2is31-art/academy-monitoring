# -*- coding: utf-8 -*-
"""
메가스터디 러셀 대치학원 강의계획서 검수 시스템 버전 및 릴리스 메타데이터
수정 및 업데이트 시 이 파일의 정보를 갱신하여 UI 화면, 사이드바, 결과지 엑셀/CSV에 즉각 반영합니다.
"""

APP_VERSION = "v2.5.0"
BUILD_DATE = "2026-11-10 19:30"
RELEASE_NOTES = [
    "수강료 및 수강회차 불일치 중대 오류 한눈에 확인 (전용 KPI, 1클릭 필터, 엑셀 강조 서식)",
    "검수결과지 다운로드 시 현재 검수·생성 일시(KST) 및 버전 정보 전 시트/CSV 명시적 출력",
    "RMS 엑셀 다중 컬럼명('수강 기호' 등) 정규화 및 결과지 '수강기호' 100% 정상 출력 보장",
]


def get_version_info() -> dict:
    return {
        "version": APP_VERSION,
        "build_date": BUILD_DATE,
        "release_notes": RELEASE_NOTES,
        "badge_text": f"{APP_VERSION} ({BUILD_DATE} 최신 배포본)",
    }
