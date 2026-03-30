"""
TalkPC Local - 카카오톡 자동 발송 프로그램
서버 연동 없는 독립 실행 버전
"""

import sys
import os
from pathlib import Path

def get_project_root():
    """exe 실행 시 exe가 있는 폴더, 스크립트 실행 시 스크립트 폴더"""
    if getattr(sys, 'frozen', False):
        # PyInstaller exe: exe 파일이 있는 폴더
        return Path(sys.executable).parent.resolve()
    else:
        return Path(__file__).parent.resolve()

def get_bundle_dir():
    """PyInstaller 번들 리소스 경로 (config/tessdata 등 읽기전용 리소스)"""
    if getattr(sys, 'frozen', False):
        # 폴더모드: _internal 폴더, onefile: _MEIPASS 임시폴더
        return Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent / '_internal')).resolve()
    else:
        return Path(__file__).parent.resolve()

PROJECT_ROOT = get_project_root()   # 데이터 저장 경로 (exe 옆)
BUNDLE_DIR = get_bundle_dir()       # 번들 리소스 경로 (읽기전용)
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
    print("  TalkPC Local v1.1.1")
    print("  카카오톡 자동 발송 프로그램")
    print("=" * 40)
    print(f"  frozen: {getattr(sys, 'frozen', False)}")
    print(f"  ROOT: {PROJECT_ROOT}")
    print(f"  data: {PROJECT_ROOT / 'data'}")
    print("=" * 40 + "\n")

    check_dependencies()

    # 필요한 디렉토리 생성
    for d in ["config", "data/templates", "logs/screenshots"]:
        (PROJECT_ROOT / d).mkdir(parents=True, exist_ok=True)

    # exe 첫 실행: 번들에서 초기 파일 복사 (config, tessdata 등)
    if getattr(sys, 'frozen', False):
        import shutil
        # tessdata 복사
        bundle_tessdata = BUNDLE_DIR / "config" / "tessdata"
        local_tessdata = PROJECT_ROOT / "config" / "tessdata"
        if bundle_tessdata.exists() and not local_tessdata.exists():
            shutil.copytree(str(bundle_tessdata), str(local_tessdata))
        # default_config.json 복사
        bundle_cfg = BUNDLE_DIR / "config" / "default_config.json"
        local_cfg = PROJECT_ROOT / "config" / "default_config.json"
        if bundle_cfg.exists() and not local_cfg.exists():
            shutil.copy2(str(bundle_cfg), str(local_cfg))

    from core.orchestrator import Orchestrator
    from ui.app import App

    orchestrator = Orchestrator(base_dir=str(PROJECT_ROOT))
    app = App(orchestrator=orchestrator)
    app.mainloop()


if __name__ == "__main__":
    main()
