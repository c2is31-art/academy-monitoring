"""
Streamlit Community Cloud 루트 진입점 (app.py)
Streamlit Cloud가 리포지토리 루트의 'app.py'를 실행할 때,
classtool 디렉토리의 종합 검수 시스템(classtool/app.py)으로 안전하게 연결 및 실행합니다.
"""
import runpy
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
CLASSTOOL_DIR = ROOT_DIR / "classtool"

if str(CLASSTOOL_DIR) not in sys.path:
    sys.path.insert(0, str(CLASSTOOL_DIR))

target_app = CLASSTOOL_DIR / "app.py"
runpy.run_path(str(target_app), run_name="__main__")

