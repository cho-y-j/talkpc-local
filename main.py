"""
TalkPC Local - 카카오톡 자동 발송 프로그램
서버 연동 없는 독립 실행 버전
"""

import sys
import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.resolve()
sys.path.insert(0, str(PROJECT_ROOT))
os.chdir(PROJECT_ROOT)


def check_dependencies():
    missing = []
    for pkg in ["customtkinter", "pyautogui", "pytesseract", "PIL", "openpyxl"]:
        try:
            __import__(pkg)
        except ImportError:
            name_map = {"PIL": "Pillow"}
            missing.append(name_map.get(pkg, pkg))
    if missing:
        print(f"누락된 패키지: pip install {' '.join(missing)}")
        sys.exit(1)


def main():
    print("\n" + "=" * 40)
    print("  TalkPC Local v1.1.0")
    print("  카카오톡 자동 발송 프로그램")
    print("=" * 40 + "\n")

    check_dependencies()

    for d in ["config", "data/templates", "logs/screenshots"]:
        (PROJECT_ROOT / d).mkdir(parents=True, exist_ok=True)

    from core.orchestrator import Orchestrator
    from ui.app import App

    orchestrator = Orchestrator(base_dir=str(PROJECT_ROOT))
    app = App(orchestrator=orchestrator)
    app.mainloop()


if __name__ == "__main__":
    main()
