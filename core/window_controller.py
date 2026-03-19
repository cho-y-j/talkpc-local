"""
WindowController - 카카오톡 창 관리 모듈
Mac (AppleScript) / Windows (win32gui) 크로스 플랫폼 지원
"""

import platform
import subprocess
import json
import os
from pathlib import Path

try:
    import pyautogui
except ImportError:
    pyautogui = None


class WindowController:
    """카카오톡 창 위치/크기를 강제 고정하는 컨트롤러"""

    KAKAO_APP_NAME_MAC = "KakaoTalk"
    KAKAO_APP_NAME_WIN = "카카오톡"
    KAKAO_WINDOW_TITLE_WIN = "카카오톡"

    def __init__(self, config: dict = None):
        self.system = platform.system()  # "Darwin" (Mac) or "Windows"
        self.config = config or {}
        self.screen_width = 0
        self.screen_height = 0
        self.dpi_scale = 1.0
        self.kakao_rect = {}  # {x, y, width, height}
        self._detect_screen()

    def _detect_screen(self):
        """모니터 해상도 및 DPI 스케일링 감지"""
        if pyautogui:
            self.screen_width, self.screen_height = pyautogui.size()
        else:
            # fallback
            self.screen_width = 1920
            self.screen_height = 1080

        if self.system == "Windows":
            try:
                import ctypes
                self.dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100.0
            except Exception:
                self.dpi_scale = 1.0
        elif self.system == "Darwin":
            try:
                result = subprocess.run(
                    ["system_profiler", "SPDisplaysDataType", "-json"],
                    capture_output=True, text=True, timeout=5
                )
                data = json.loads(result.stdout)
                displays = data.get("SPDisplaysDataType", [])
                for gpu in displays:
                    for disp in gpu.get("spdisplays_ndrvs", []):
                        res = disp.get("_spdisplays_resolution", "")
                        if "Retina" in res:
                            self.dpi_scale = 2.0
                            break
            except Exception:
                self.dpi_scale = 1.0

    def get_screen_info(self) -> dict:
        """현재 화면 정보 반환"""
        return {
            "system": self.system,
            "screen_width": self.screen_width,
            "screen_height": self.screen_height,
            "dpi_scale": self.dpi_scale
        }

    def calculate_kakao_position(self) -> dict:
        """카카오톡 창 최적 위치/크기 계산 - 우측 상단 고정"""
        kakao_w = self.config.get("kakao_window", {}).get("width", 420)
        kakao_h = self.config.get("kakao_window", {}).get("height", 700)
        margin_right = self.config.get("kakao_window", {}).get("margin_right", 20)
        margin_top = self.config.get("kakao_window", {}).get("margin_top", 40)

        kakao_x = self.screen_width - kakao_w - margin_right
        kakao_y = margin_top

        self.kakao_rect = {
            "x": kakao_x,
            "y": kakao_y,
            "width": kakao_w,
            "height": kakao_h
        }
        return self.kakao_rect

    def find_kakao_window(self) -> bool:
        """카카오톡 창이 열려있는지 확인"""
        if self.system == "Darwin":
            return self._find_kakao_mac()
        elif self.system == "Windows":
            return self._find_kakao_win()
        return False

    def _find_kakao_mac(self) -> bool:
        """Mac에서 카카오톡 프로세스 확인"""
        try:
            script = '''
            tell application "System Events"
                set appList to name of every process
                if appList contains "KakaoTalk" then
                    return "found"
                else
                    return "not_found"
                end if
            end tell
            '''
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=5
            )
            return "found" in result.stdout.strip()
        except Exception:
            return False

    def _find_kakao_win(self) -> bool:
        """Windows에서 카카오톡 창 확인"""
        try:
            import win32gui

            def callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if self.KAKAO_WINDOW_TITLE_WIN in title:
                        results.append(hwnd)

            results = []
            win32gui.EnumWindows(callback, results)
            return len(results) > 0
        except ImportError:
            # win32gui 없으면 pyautogui로 대체
            try:
                import pygetwindow as gw
                windows = gw.getWindowsWithTitle(self.KAKAO_WINDOW_TITLE_WIN)
                return len(windows) > 0
            except Exception:
                return False

    def activate_kakao(self) -> bool:
        """카카오톡 창 활성화 (포커스)"""
        if self.system == "Darwin":
            return self._activate_kakao_mac()
        elif self.system == "Windows":
            return self._activate_kakao_win()
        return False

    def _activate_kakao_mac(self) -> bool:
        """Mac에서 카카오톡 활성화"""
        try:
            script = '''
            tell application "KakaoTalk"
                activate
            end tell
            '''
            subprocess.run(["osascript", "-e", script], timeout=5)
            return True
        except Exception:
            return False

    def _activate_kakao_win(self) -> bool:
        """Windows에서 카카오톡 활성화"""
        try:
            import win32gui
            import win32con

            def callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if self.KAKAO_WINDOW_TITLE_WIN in title:
                        results.append(hwnd)

            results = []
            win32gui.EnumWindows(callback, results)
            if results:
                hwnd = results[0]
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return True
            return False
        except ImportError:
            try:
                import pygetwindow as gw
                windows = gw.getWindowsWithTitle(self.KAKAO_WINDOW_TITLE_WIN)
                if windows:
                    windows[0].activate()
                    return True
            except Exception:
                pass
            return False

    def position_kakao_window(self) -> bool:
        """카카오톡 창을 계산된 위치로 이동 및 크기 조정"""
        if not self.kakao_rect:
            self.calculate_kakao_position()

        x = self.kakao_rect["x"]
        y = self.kakao_rect["y"]
        w = self.kakao_rect["width"]
        h = self.kakao_rect["height"]

        if self.system == "Darwin":
            return self._position_kakao_mac(x, y, w, h)
        elif self.system == "Windows":
            return self._position_kakao_win(x, y, w, h)
        return False

    def _position_kakao_mac(self, x, y, w, h) -> bool:
        """Mac에서 카카오톡 창 위치/크기 설정"""
        try:
            script = f'''
            tell application "System Events"
                tell process "KakaoTalk"
                    set frontmost to true
                    delay 0.3
                    try
                        set position of window 1 to {{{x}, {y}}}
                        set size of window 1 to {{{w}, {h}}}
                    end try
                end tell
            end tell
            '''
            subprocess.run(["osascript", "-e", script], timeout=10)
            return True
        except Exception as e:
            print(f"[WindowController] Mac position error: {e}")
            return False

    def _position_kakao_win(self, x, y, w, h) -> bool:
        """Windows에서 카카오톡 창 위치/크기 설정"""
        try:
            import win32gui
            import win32con

            def callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if self.KAKAO_WINDOW_TITLE_WIN in title:
                        results.append(hwnd)

            results = []
            win32gui.EnumWindows(callback, results)
            if results:
                hwnd = results[0]
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.MoveWindow(hwnd, x, y, w, h, True)
                return True
            return False
        except ImportError:
            try:
                import pygetwindow as gw
                windows = gw.getWindowsWithTitle(self.KAKAO_WINDOW_TITLE_WIN)
                if windows:
                    win = windows[0]
                    win.moveTo(x, y)
                    win.resizeTo(w, h)
                    return True
            except Exception:
                pass
            return False

    def calculate_ui_coordinates(self) -> dict:
        """
        카카오톡 창 내부 UI 요소 좌표 계산

        Mac 카카오톡 레이아웃 (상단부터):
        - 타이틀바: ~28px
        - 1행 (ky+28~ky+65): 프로필아이콘 | "채팅 ▼" | ... | 🔍 돋보기 | 💬 새채팅 | ➕
        - 2행 (ky+65~ky+100): [전체] [안읽음 999+] [+]
        - 3행~ (ky+100~): 채팅 목록

        검색 클릭 후 레이아웃:
        - 1행: ← 뒤로 | [검색 입력 필드]
        - 2행~: 검색 결과 목록
        """
        if not self.kakao_rect:
            self.calculate_kakao_position()

        kx = self.kakao_rect["x"]
        ky = self.kakao_rect["y"]
        kw = self.kakao_rect["width"]
        kh = self.kakao_rect["height"]

        coords = {
            "search_icon": {
                "x": kx + kw - 110,
                "y": ky + 50,
                "description": "돋보기 검색 아이콘 (우상단)"
            },
            "search_input": {
                "x": kx + kw // 2,
                "y": ky + 50,
                "description": "검색 입력 필드 (검색 모드 진입 후)"
            },
            "first_result": {
                "x": kx + kw // 2,
                "y": ky + 130,
                "description": "첫 번째 검색 결과"
            },
            "message_input": {
                "x": kx + kw // 2,
                "y": ky + kh - 50,
                "description": "메시지 입력창 (채팅방 하단)"
            },
            "back_button": {
                "x": kx + 30,
                "y": ky + 50,
                "description": "뒤로가기 버튼 (좌상단)"
            },
            "search_result_area": {
                "x1": kx + 10,
                "y1": ky + 80,
                "x2": kx + kw - 10,
                "y2": ky + 350,
                "description": "검색 결과 영역 (OCR 대상)"
            },
            "chat_header_area": {
                "x1": kx,
                "y1": ky + 28,
                "x2": kx + kw,
                "y2": ky + 100,
                "description": "채팅 헤더 영역 (캘리브레이션용)"
            }
        }
        return coords

    def calibrate(self, screen_capture) -> dict:
        """
        스크린샷 기반 캘리브레이션
        카카오톡 창을 캡처하고 UI 요소 위치를 검증

        Returns:
            {
                "success": bool,
                "screenshot_path": str,
                "coordinates": dict,
                "kakao_rect": dict,
            }
        """
        if not self.kakao_rect:
            self.calculate_kakao_position()

        result = {
            "success": False,
            "screenshot_path": None,
            "coordinates": {},
            "kakao_rect": self.kakao_rect,
        }

        try:
            # 1. 카카오톡 창 캡처
            screenshot = screen_capture.capture_kakao_window(self.kakao_rect)
            screenshot_path = screen_capture.save_screenshot(screenshot, "calibration")
            result["screenshot_path"] = screenshot_path

            # 2. 좌표 계산
            coords = self.calculate_ui_coordinates()
            result["coordinates"] = coords

            # 3. 캘리브레이션 성공
            result["success"] = True

        except Exception as e:
            result["error"] = str(e)

        return result

    def setup(self) -> dict:
        """
        전체 초기 설정 실행
        1. 화면 감지
        2. 카카오톡 찾기
        3. 활성화
        4. 위치/크기 고정
        5. UI 좌표 계산
        Returns: 설정 결과 딕셔너리
        """
        result = {
            "screen": self.get_screen_info(),
            "kakao_found": False,
            "kakao_positioned": False,
            "coordinates": {}
        }

        # 카카오톡 찾기
        result["kakao_found"] = self.find_kakao_window()
        if not result["kakao_found"]:
            return result

        # 활성화
        self.activate_kakao()

        # 위치 계산 및 이동
        self.calculate_kakao_position()
        result["kakao_positioned"] = self.position_kakao_window()

        # UI 좌표 계산
        result["coordinates"] = self.calculate_ui_coordinates()

        return result
