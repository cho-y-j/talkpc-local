"""
KakaoSender - 카카오톡 자동 발송 모듈
pyautogui를 이용한 카카오톡 UI 자동 조작

안전장치:
- 모든 클릭/입력 전후 마우스 위치 검증
- pyautogui.FAILSAFE (좌측 상단 모서리 이동 시 긴급 정지)
- 에러 발생 시 즉시 정지 + 콜백
"""

import time
import random
import platform

try:
    import pyautogui
    pyautogui.FAILSAFE = True   # 마우스를 좌측 상단으로 이동하면 긴급 중지
    pyautogui.PAUSE = 0.3       # 각 동작 사이 기본 딜레이
except ImportError:
    pyautogui = None

from core.screen_capture import ScreenCapture
from core.ocr_engine import OCREngine
from pathlib import Path

_LOG_PATH = Path("/tmp/kakao_sender_debug.log")
_OCR_PATH = Path("/tmp/kakao_ocr_capture.png")

def _debug_log(msg):
    """디버그 로그를 파일에 기록"""
    try:
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%H:%M:%S')}] {msg}\n")
    except Exception:
        pass


class SendResult:
    """발송 결과 데이터"""

    SUCCESS = "success"
    FAILED_NOT_FOUND = "not_found"
    FAILED_OCR = "ocr_error"
    FAILED_SEND = "send_error"
    FAILED_SAFETY = "safety_stop"
    SKIPPED = "skipped"

    def __init__(self, contact_name: str, status: str, message: str = "", detail: str = ""):
        self.contact_name = contact_name
        self.status = status
        self.message = message
        self.detail = detail
        self.timestamp = time.time()

    def to_dict(self) -> dict:
        return {
            "contact_name": self.contact_name,
            "status": self.status,
            "message": self.message[:50] + "..." if len(self.message) > 50 else self.message,
            "detail": self.detail,
            "timestamp": self.timestamp
        }


class SafetyError(Exception):
    """안전 장치 발동 시 발생하는 예외"""
    pass


class KakaoSender:
    """카카오톡 자동 발송기 (안전장치 + 사람처럼 행동)"""

    def __init__(self, coordinates: dict, config: dict = None):
        if pyautogui is None:
            raise ImportError("pyautogui가 설치되지 않았습니다.")

        self.coords = coordinates
        self.config = config or {}
        self.capture = ScreenCapture()
        self.ocr = OCREngine()
        self.is_mac = platform.system() == "Darwin"

        # 딜레이 설정
        self.delay_min = self.config.get("sending", {}).get("delay_min", 30)
        self.delay_max = self.config.get("sending", {}).get("delay_max", 120)
        self.retry_count = self.config.get("sending", {}).get("retry_count", 2)

        # 계정 보호 설정
        anti_detect = self.config.get("anti_detect", {})
        self.action_delay_min = anti_detect.get("action_delay_min", 0.5)
        self.action_delay_max = anti_detect.get("action_delay_max", 1.5)
        self.rest_every = anti_detect.get("rest_every", 10)       # N명마다 쉬기
        self.rest_min = anti_detect.get("rest_min", 60)           # 쉬는 시간 최소(초)
        self.rest_max = anti_detect.get("rest_max", 180)          # 쉬는 시간 최대(초)
        self.daily_limit = anti_detect.get("daily_limit", 50)     # 일일 최대 발송 수

        # 상태 플래그
        self._stop_flag = False
        self._safety_error = None
        self._send_count = 0  # 연속 발송 횟수 (2번째부터 돋보기 2회 클릭)
        self._last_mouse_pos = None  # 마우스 이동 감지용

        # 에러 콜백
        self._on_safety_stop = None

    def on_safety_stop(self, callback):
        """안전 정지 콜백 등록"""
        self._on_safety_stop = callback

    def stop(self):
        self._stop_flag = True

    def resume(self):
        self._stop_flag = False
        self._safety_error = None
        self._send_count = 0
        self._last_mouse_pos = None

    # ── 사람처럼 행동 ──

    def _human_delay(self, min_s: float = None, max_s: float = None):
        """사람처럼 랜덤 딜레이 (모든 동작 사이에 사용)"""
        lo = min_s if min_s is not None else self.action_delay_min
        hi = max_s if max_s is not None else self.action_delay_max
        delay = random.uniform(lo, hi)
        time.sleep(delay)

    def _human_move(self, x: int, y: int):
        """사람처럼 마우스를 곡선으로 이동 (직선 X)"""
        duration = random.uniform(0.2, 0.5)
        # pyautogui의 easeOutQuad → 가속 후 감속 (사람 패턴)
        pyautogui.moveTo(x, y, duration=duration, tween=pyautogui.easeOutQuad)

    def should_rest(self) -> bool:
        """N명마다 긴 휴식이 필요한지"""
        return self.rest_every > 0 and self._send_count > 0 and self._send_count % self.rest_every == 0

    def take_rest(self) -> float:
        """긴 휴식 (N명 발송 후)"""
        rest = random.uniform(self.rest_min, self.rest_max)
        _debug_log(f"계정 보호 휴식: {rest:.0f}초 ({self._send_count}명 발송 후)")
        time.sleep(rest)
        return rest

    def check_daily_limit(self) -> bool:
        """일일 발송 한도 체크"""
        return self.daily_limit <= 0 or self._send_count < self.daily_limit

    # ── 안전한 조작 메서드 ──

    def _safe_click(self, x: int, y: int, clicks: int = 1):
        """
        안전한 클릭 - 클릭 전후 마우스 위치 검증
        macOS: Quartz CGEvent로 정확한 클릭/더블클릭
        """
        try:
            # 클릭 전: 사용자가 마우스를 움직였는지 감지
            self._check_mouse_moved()

            if self.is_mac and clicks >= 2:
                # macOS 더블클릭: Quartz CGEvent 사용
                from Quartz import (
                    CGEventCreateMouseEvent,
                    CGEventPost,
                    CGEventSetIntegerValueField,
                    kCGHIDEventTap,
                    kCGEventLeftMouseDown,
                    kCGEventLeftMouseUp,
                )
                from Quartz import CGPointMake

                point = CGPointMake(float(x), float(y))

                # 사람처럼 곡선 이동
                self._human_move(x, y)
                time.sleep(0.1)

                # 첫 번째 클릭 (clickCount=1)
                # CGEventSetIntegerValueField field 1 = kCGMouseEventClickState
                down1 = CGEventCreateMouseEvent(None, kCGEventLeftMouseDown, point, 0)
                CGEventSetIntegerValueField(down1, 1, 1)
                CGEventPost(kCGHIDEventTap, down1)
                time.sleep(0.02)
                up1 = CGEventCreateMouseEvent(None, kCGEventLeftMouseUp, point, 0)
                CGEventSetIntegerValueField(up1, 1, 1)
                CGEventPost(kCGHIDEventTap, up1)
                time.sleep(0.05)

                # 두 번째 클릭 (clickCount=2 → 더블클릭)
                down2 = CGEventCreateMouseEvent(None, kCGEventLeftMouseDown, point, 0)
                CGEventSetIntegerValueField(down2, 1, 2)
                CGEventPost(kCGHIDEventTap, down2)
                time.sleep(0.02)
                up2 = CGEventCreateMouseEvent(None, kCGEventLeftMouseUp, point, 0)
                CGEventSetIntegerValueField(up2, 1, 2)
                CGEventPost(kCGHIDEventTap, up2)

                _debug_log(f"Quartz 더블클릭: ({x}, {y})")
            else:
                # 단일 클릭: 사람처럼 곡선 이동
                self._human_move(x, y)
                time.sleep(random.uniform(0.05, 0.15))
                for _ in range(clicks):
                    pyautogui.mouseDown(x, y)
                    time.sleep(random.uniform(0.03, 0.08))
                    pyautogui.mouseUp(x, y)
                    time.sleep(random.uniform(0.05, 0.15))

            self._human_delay(0.2, 0.5)

            # 클릭 후: 마우스 위치 기록 (다음 클릭 전 이동 감지용)
            self._record_mouse_pos()

        except pyautogui.FailSafeException:
            raise SafetyError("긴급 정지! 마우스가 화면 좌측 상단 모서리로 이동했습니다.")

    def _activate_kakao(self):
        """카카오톡을 최전면으로 활성화"""
        if self.is_mac:
            import subprocess
            subprocess.run([
                "osascript", "-e",
                'tell application "KakaoTalk" to activate'
            ], timeout=5)
            time.sleep(0.5)

    def _run_applescript(self, script: str):
        """AppleScript 실행 (macOS 전용)"""
        import subprocess
        result = subprocess.run(
            ["osascript", "-e", script],
            timeout=10, capture_output=True, text=True
        )
        _debug_log(f"AppleScript rc={result.returncode} err='{result.stderr.strip()}'")
        return result

    def _safe_type_text(self, text: str):
        """안전한 텍스트 입력 (한글 지원)"""
        try:
            _debug_log(f"type_text 시작: '{text[:30]}'")
            if self.is_mac:
                import subprocess
                import os

                # 1. 클립보드에 복사 (os 환경변수 유지)
                env = os.environ.copy()
                env["LANG"] = "en_US.UTF-8"
                proc = subprocess.run(
                    ["pbcopy"], input=text.encode("utf-8"),
                    env=env, timeout=5
                )
                _debug_log(f"pbcopy rc={proc.returncode}")
                time.sleep(0.1)

                # 2. 검색창에 포커스가 있는 상태에서 바로 Cmd+V
                #    (AppleScript activate 하지 않음 - 이미 클릭으로 포커스 잡혀있음)
                from Quartz import (
                    CGEventCreateKeyboardEvent,
                    CGEventSetFlags,
                    CGEventPost,
                    kCGHIDEventTap,
                    kCGEventFlagMaskCommand,
                )

                # Cmd+V: keycode 9 = 'v'
                # key down
                event = CGEventCreateKeyboardEvent(None, 9, True)
                CGEventSetFlags(event, kCGEventFlagMaskCommand)
                CGEventPost(kCGHIDEventTap, event)
                time.sleep(0.05)
                # key up
                event = CGEventCreateKeyboardEvent(None, 9, False)
                CGEventSetFlags(event, kCGEventFlagMaskCommand)
                CGEventPost(kCGHIDEventTap, event)

                _debug_log("type_text: Quartz Cmd+V 완료")
            else:
                import subprocess
                process = subprocess.Popen(
                    ["clip"], stdin=subprocess.PIPE
                )
                process.communicate(text.encode("utf-16le"))
                pyautogui.hotkey("ctrl", "v")
            time.sleep(0.5)
        except pyautogui.FailSafeException:
            raise SafetyError("긴급 정지! (텍스트 입력 중)")
        except Exception as e:
            _debug_log(f"type_text 에러: {e}")
            raise

    def _quartz_key(self, keycode: int, flags: int = 0):
        """macOS Quartz로 키 이벤트 직접 전송 (포커스된 앱에 전달)"""
        from Quartz import (
            CGEventCreateKeyboardEvent,
            CGEventSetFlags,
            CGEventPost,
            kCGHIDEventTap,
        )
        # key down
        event = CGEventCreateKeyboardEvent(None, keycode, True)
        if flags:
            CGEventSetFlags(event, flags)
        CGEventPost(kCGHIDEventTap, event)
        time.sleep(0.05)
        # key up
        event = CGEventCreateKeyboardEvent(None, keycode, False)
        if flags:
            CGEventSetFlags(event, flags)
        CGEventPost(kCGHIDEventTap, event)
        time.sleep(0.05)

    def _safe_clear_input(self):
        """안전한 입력 필드 전체 삭제"""
        try:
            if self.is_mac:
                from Quartz import kCGEventFlagMaskCommand
                # Cmd+A (keycode 0 = 'a')
                self._quartz_key(0, kCGEventFlagMaskCommand)
                time.sleep(0.1)
                # Delete (keycode 51)
                self._quartz_key(51)
                _debug_log("clear_input: Quartz Cmd+A → Delete 완료")
            else:
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.2)
                pyautogui.press("delete")
            time.sleep(0.3)
        except pyautogui.FailSafeException:
            raise SafetyError("긴급 정지! (입력 삭제 중)")

    def _safe_press(self, key: str):
        """안전한 키 입력"""
        try:
            if self.is_mac:
                key_map = {
                    "enter": 36,
                    "return": 36,
                    "escape": 53,
                    "delete": 51,
                    "tab": 48,
                }
                keycode = key_map.get(key)
                if keycode is not None:
                    self._quartz_key(keycode)
                    _debug_log(f"press: Quartz keycode={keycode} ({key})")
                else:
                    pyautogui.press(key)
            else:
                pyautogui.press(key)
            time.sleep(0.3)
        except pyautogui.FailSafeException:
            raise SafetyError(f"긴급 정지! ({key} 키 입력 중)")

    def _check_stop(self):
        """중지 플래그 체크"""
        if self._stop_flag:
            raise SafetyError("사용자가 발송을 중지했습니다.")

    def _record_mouse_pos(self):
        """현재 마우스 위치 기록 (클릭 직후 호출)"""
        if pyautogui:
            self._last_mouse_pos = pyautogui.position()

    def _check_mouse_moved(self):
        """사용자가 마우스를 움직였는지 감지 (클릭 전 호출)"""
        if pyautogui and self._last_mouse_pos:
            curr = pyautogui.position()
            lx, ly = self._last_mouse_pos
            dist = abs(curr[0] - lx) + abs(curr[1] - ly)
            if dist > 50:  # 클릭 후 마우스가 50px 이상 이동 = 사용자 개입
                self._stop_flag = True
                self._safety_error = (
                    f"마우스가 이동했습니다 ({lx},{ly})→({curr[0]},{curr[1]}). "
                    "사용자 개입으로 발송을 중단합니다."
                )
                if self._on_safety_stop:
                    self._on_safety_stop(self._safety_error)
                raise SafetyError(self._safety_error)

    # ── 카카오톡 조작 ──

    def click_search_icon(self):
        """돋보기(검색) 아이콘 클릭 (2번째 발송부터 2회 클릭: 기존 검색 닫기 + 새 검색 열기)"""
        self._check_stop()
        coord = self.coords.get("search_icon", {})
        _debug_log(f"click_search_icon: ({coord.get('x')}, {coord.get('y')}) send_count={self._send_count}")

        if self._send_count > 0:
            self._safe_click(coord["x"], coord["y"])
            _debug_log("click_search_icon: 1차 클릭 (기존 검색 닫기)")
            self._human_delay(0.6, 1.2)

        self._safe_click(coord["x"], coord["y"])
        _debug_log("click_search_icon: 검색 모드 열기 클릭 완료")
        self._human_delay(1.0, 1.8)

    def search_contact(self, name: str) -> bool:
        """이름으로 연락처 검색"""
        self._check_stop()
        _debug_log(f"search_contact: '{name}' 검색 시작")
        coord = self.coords.get("search_input", {})
        _debug_log(f"search_input 좌표: ({coord.get('x')}, {coord.get('y')})")
        self._safe_click(coord["x"], coord["y"])
        self._human_delay(0.3, 0.7)
        _debug_log("search_input 클릭 완료, clear_input 시작")
        self._safe_clear_input()
        _debug_log("clear_input 완료, type_text 시작")
        self._safe_type_text(name)
        _debug_log("type_text 완료, 검색 결과 대기 중...")
        self._human_delay(1.2, 2.0)
        return True

    def verify_search_result(self, target_name: str) -> dict:
        """
        OCR로 검색 결과 확인
        검색 결과 영역에서 이름 텍스트 부분만 캡처하여 OCR
        프로필 사진을 제외하고 이름이 표시되는 오른쪽 영역만 캡처
        """
        try:
            search_input = self.coords.get("search_input", {})
            first_result = self.coords.get("first_result", {})

            if not search_input or not first_result:
                return {"found": False, "error": "좌표 없음"}

            # 캡처 영역: first_result 좌표 기준
            # 프로필 사진(왼쪽 ~80px)을 제외하고 이름 텍스트 영역만 캡처
            fx = first_result["x"]
            fy = first_result["y"]
            x1 = fx - 80     # 프로필 사진 제외, 이름 시작 부근
            y1 = fy - 25     # 결과 위쪽
            x2 = fx + 150    # 이름 끝 (날짜 영역 제외)
            y2 = fy + 25     # 결과 아래쪽

            _debug_log(f"OCR 캡처 영역: ({x1},{y1}) ~ ({x2},{y2})")
            screenshot = self.capture.capture_region(x1, y1, x2, y2)

            # 디버그: 캡처 이미지 저장
            screenshot.save(str(_OCR_PATH))
            _debug_log(f"OCR 캡처 저장: {_OCR_PATH} (크기: {screenshot.size})")

            result = self.ocr.verify_name_in_results(screenshot, target_name)
            _debug_log(f"OCR 결과: found={result.get('found')} "
                       f"text='{result.get('matched_text')}' "
                       f"extracted='{result.get('extracted_text', '')[:50]}'")
            return result

        except Exception as e:
            _debug_log(f"OCR 에러: {e}")
            return {"found": False, "error": str(e)}

    def click_search_result(self):
        """첫 번째 검색 결과 더블클릭 (채팅방 진입)"""
        self._check_stop()
        coord = self.coords.get("first_result", {})
        self._safe_click(coord["x"], coord["y"], clicks=2)
        self._human_delay(1.2, 2.0)  # 채팅방 열리기 대기

        # macOS: 새로 열린 채팅창을 학습된 위치에 맞게 배치
        if self.is_mac:
            self._position_chat_window()

    def _position_chat_window(self):
        """새로 열린 채팅창을 원래 카카오톡 창과 같은 위치에 겹쳐서 배치 (모든 창 대상)"""
        try:
            import subprocess

            # 원래 카카오톡 창 위치/크기 (config에서 가져옴)
            kw_config = self.config.get("kakao_window", {})
            chat_w = kw_config.get("width", 420)
            chat_h = kw_config.get("height", 700)
            margin_r = kw_config.get("margin_right", 20)
            margin_t = kw_config.get("margin_top", 40)

            if pyautogui:
                sw, sh = pyautogui.size()
            else:
                sw, sh = 1920, 1080

            chat_x = sw - chat_w - margin_r
            chat_y = margin_t

            # 모든 카카오톡 윈도우를 같은 위치에 배치 (새 채팅창 포함)
            script = f'''
            tell application "System Events"
                tell process "KakaoTalk"
                    set frontmost to true
                    delay 0.3
                    set winCount to count of windows
                    repeat with i from 1 to winCount
                        try
                            set position of window i to {{{chat_x}, {chat_y}}}
                            set size of window i to {{{chat_w}, {chat_h}}}
                        end try
                    end repeat
                end tell
            end tell
            '''
            subprocess.run(["osascript", "-e", script], timeout=10)
            _debug_log(f"채팅창 배치: 모든 창 → ({chat_x}, {chat_y}) {chat_w}x{chat_h}")
            time.sleep(0.5)
        except Exception as e:
            _debug_log(f"채팅창 배치 실패: {e}")

    def type_message(self, message: str):
        """메시지 입력"""
        self._check_stop()
        self._activate_kakao()
        coord = self.coords.get("message_input", {})
        _debug_log(f"type_message: message_input ({coord.get('x')}, {coord.get('y')})")
        self._safe_click(coord["x"], coord["y"])
        self._human_delay(0.2, 0.5)
        self._safe_type_text(message)
        self._human_delay(0.3, 0.8)

    def send_message(self):
        """메시지 전송 (전송 버튼 클릭)"""
        self._check_stop()
        coord = self.coords.get("send_enter", {})
        _debug_log(f"send_message: 전송 버튼 클릭 ({coord.get('x')}, {coord.get('y')})")
        self._activate_kakao()
        self._human_delay(0.2, 0.5)
        self._safe_click(coord["x"], coord["y"])
        self._human_delay(0.5, 1.0)

    def paste_image(self, image_path: str):
        """이미지 클립보드 복사 → 붙여넣기 → 전송 버튼 클릭"""
        self._check_stop()
        _debug_log(f"paste_image: {image_path}")

        from core.image_clipboard import copy_image_to_clipboard
        copy_image_to_clipboard(image_path)
        _debug_log("paste_image: 클립보드 복사 완료")
        time.sleep(0.3)

        # 메시지 입력창 클릭 (포커스)
        self._activate_kakao()
        coord = self.coords.get("message_input", {})
        self._safe_click(coord["x"], coord["y"])
        time.sleep(0.3)

        # Cmd+V / Ctrl+V 붙여넣기
        if self.is_mac:
            from Quartz import kCGEventFlagMaskCommand
            self._quartz_key(9, kCGEventFlagMaskCommand)  # keycode 9 = 'v'
        else:
            pyautogui.hotkey("ctrl", "v")
        _debug_log("paste_image: 붙여넣기 완료")
        time.sleep(2.0)  # 카카오톡 이미지 전송 확인 팝업 대기

        # 이미지 전송 버튼 클릭 (학습된 image_send 좌표 우선, 없으면 send_enter)
        img_coord = self.coords.get("image_send", {})
        if img_coord and "x" in img_coord:
            self._safe_click(img_coord["x"], img_coord["y"])
            _debug_log(f"paste_image: 이미지 전송 버튼 클릭 ({img_coord['x']}, {img_coord['y']})")
        else:
            send_coord = self.coords.get("send_enter", {})
            if send_coord and "x" in send_coord:
                self._safe_click(send_coord["x"], send_coord["y"])
                _debug_log(f"paste_image: send_enter 클릭 (fallback)")
            else:
                self._safe_press("enter")
                _debug_log("paste_image: Enter (fallback)")
        time.sleep(1.0)

    def go_back(self):
        """채팅방에서 나가기 → 메인 창 위치 재고정"""
        self._check_stop()
        self._human_delay(0.3, 0.8)
        coord = self.coords.get("back_button", {})
        if coord and "x" in coord:
            self._safe_click(coord["x"], coord["y"])
        else:
            self._safe_press("escape")
        self._human_delay(0.4, 0.8)
        # 닫기 후 메인 창도 원래 위치로 재고정
        if self.is_mac:
            self._position_chat_window()

    def send_to_contact(self, name: str, message: str, image_path: str = None) -> SendResult:
        """
        한 명에게 메시지 발송 (전체 프로세스)

        워크플로우:
        1. 돋보기 아이콘 클릭 (검색 모드 진입)
        2. 검색창에 이름 입력
        3. OCR로 검색 결과 검증
        4. 검색 결과 클릭 (채팅방 진입)
        5. 메시지 입력 → 전송
        6. (이미지 있으면) 이미지 붙여넣기 → 전송
        7. 뒤로가기 (채팅방 나가기)
        """
        try:
            # 0. 카카오톡 활성화
            self._activate_kakao()
            _debug_log(f"[DEBUG] === {name}에게 발송 시작 ===")

            # 1. 돋보기 클릭 → 검색 모드
            self.click_search_icon()

            # 2. 이름 검색
            self.search_contact(name)

            # 3. OCR 검증 (필수 - 검증 완료 후 더블클릭)
            retry = 0
            verified = False
            while retry <= self.retry_count and not verified:
                self._check_stop()
                verification = self.verify_search_result(name)
                if verification.get("found"):
                    verified = True
                    _debug_log(f"OCR 검증 성공: '{name}' ✅")
                else:
                    retry += 1
                    _debug_log(f"OCR 검증 실패 (시도 {retry}/{self.retry_count}): {verification}")
                    if retry <= self.retry_count:
                        time.sleep(1.0)

            if not verified:
                self._safe_press("escape")
                return SendResult(
                    name, SendResult.FAILED_NOT_FOUND,
                    detail=f"검색 결과에서 '{name}'을 찾을 수 없음"
                )

            # 4. 검증 완료 → 검색 결과 더블클릭 → 채팅방 진입
            _debug_log(f"검증 완료! '{name}' 더블클릭 시작")
            self.click_search_result()

            # 5. 메시지 입력 → 전송
            if message:
                self.type_message(message)
                self.send_message()

            # 6. 이미지 첨부 (있으면)
            if image_path:
                _debug_log(f"이미지 첨부: {image_path}")
                self.paste_image(image_path)

            # 7. 뒤로가기
            self.go_back()

            self._send_count += 1
            return SendResult(name, SendResult.SUCCESS, message=message)

        except SafetyError as e:
            # 안전 장치 발동 (긴급 정지/마우스 이동) → 즉시 정지
            self._stop_flag = True
            self._safety_error = str(e)
            if self._on_safety_stop:
                self._on_safety_stop(str(e))
            return SendResult(
                name, SendResult.FAILED_SAFETY,
                detail=f"안전 정지: {e}"
            )
        except Exception as e:
            # 기타 에러 → 실패 처리하고 다음으로 계속 진행 (멈추지 않음)
            _debug_log(f"[ERROR] {name} 발송 중 에러 (계속 진행): {e}")
            try:
                self._safe_press("escape")
            except Exception:
                pass
            return SendResult(
                name, SendResult.FAILED_SEND,
                detail=f"발송 실패: {str(e)}"
            )

    def random_delay(self):
        """랜덤 딜레이"""
        delay = random.uniform(self.delay_min, self.delay_max)
        time.sleep(delay)
        return delay
