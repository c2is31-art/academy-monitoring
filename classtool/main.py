"""
Streamlit Community Cloud 진입점 호환성 래퍼 (main.py)
Streamlit Cloud 대시보드에서 Main file path가 'main.py' 또는 'app.py' 어느 것으로 설정되어 있어도
동일하게 app.py를 정상 실행합니다.
"""
import runpy
import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

app_path = CURRENT_DIR / "app.py"
runpy.run_path(str(app_path), run_name="__main__")

