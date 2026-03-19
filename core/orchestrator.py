"""
Orchestrator (팀장) - 전체 워크플로우 오케스트레이션
각 모듈을 조율하여 자동 발송 프로세스를 실행
"""

import json
import os
import time
import threading
from pathlib import Path
from datetime import datetime
from typing import Callable, Optional

from dotenv import load_dotenv

from core.window_controller import WindowController
from core.screen_capture import ScreenCapture
from core.ocr_engine import OCREngine
from core.contact_manager import ContactManager, Contact
from core.message_engine import MessageEngine
from core.kakao_sender import KakaoSender, SendResult
from core.report_generator import ReportGenerator
from core.scheduler import Scheduler
from core.sejong_sender import SejongSender, SejongConfig, SejongSendResult


class OrchestratorState:
    """오케스트레이터 상태"""
    IDLE = "idle"
    INITIALIZING = "initializing"
    READY = "ready"
    SENDING = "sending"
    PAUSED = "paused"
    COMPLETED = "completed"
    ERROR = "error"


class Orchestrator:
    """
    팀장 모듈 - 전체 자동화 워크플로우 관리

    워크플로우:
    1. 초기화 (카카오톡 창 설정)
    2. 발송 대상 로드 & 메시지 생성
    3. 순차 발송 (검색 → OCR 확인 → 전송)
    4. 리포트 생성
    """

    def __init__(self, base_dir: str = None):
        self.base_dir = Path(base_dir) if base_dir else Path(".")

        # 설정 로드
        self.config = self._load_config()

        # 모듈 초기화
        self.window_ctrl = WindowController(self.config)
        self.screen_capture = ScreenCapture(str(self.base_dir / "logs" / "screenshots"))
        self.contact_mgr = ContactManager(str(self.base_dir / "data" / "contacts.json"))
        self.message_engine = MessageEngine(str(self.base_dir / "data" / "templates"))
        self.report = ReportGenerator(str(self.base_dir / "logs"))

        self.sender: Optional[KakaoSender] = None
        self.sejong_sender: Optional[SejongSender] = None
        self.coordinates: dict = {}
        self.scheduler = Scheduler(str(self.base_dir / "data" / "schedules.json"), self)

        # 발송 방법: "kakao_bot" | "alimtalk" | "sms"
        self.send_method = "kakao_bot"

        # 상태
        self.state = OrchestratorState.IDLE
        self.current_index = 0
        self.total_count = 0
        self.send_queue: list[dict] = []  # [{"contact": Contact, "message": str}]

        # 콜백 (UI 업데이트용)
        self._on_state_change: Optional[Callable] = None
        self._on_progress: Optional[Callable] = None
        self._on_result: Optional[Callable] = None
        self._on_log: Optional[Callable] = None

        # 스레드
        self._send_thread: Optional[threading.Thread] = None

    def _load_config(self) -> dict:
        """설정 파일 로드 + .env 환경변수 병합"""
        # .env 로드
        env_path = self.base_dir / ".env"
        if env_path.exists():
            load_dotenv(env_path)

        config_path = self.base_dir / "config" / "default_config.json"
        config = {}
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)

        # .env → sejong 설정 병합 (환경변수가 있으면 우선)
        env_sejong = {}
        if os.getenv("SEJONG_DB_HOST"):
            env_sejong["db"] = {
                "host": os.getenv("SEJONG_DB_HOST", "localhost"),
                "port": int(os.getenv("SEJONG_DB_PORT", "3306")),
                "name": os.getenv("SEJONG_DB_NAME", "sms"),
                "user": os.getenv("SEJONG_DB_USER", ""),
                "password": os.getenv("SEJONG_DB_PASSWORD", ""),
            }
            env_sejong["kakao"] = {
                "sender_key": os.getenv("SEJONG_SENDER_KEY", ""),
                "callback": os.getenv("SEJONG_CALLBACK", ""),
                "template_code": os.getenv("SEJONG_TEMPLATE_CODE", ""),
            }
            # .env 값으로 덮어쓰기 (config에 sejong이 없거나 비어있으면)
            existing = config.get("sejong", {})
            if not existing.get("db", {}).get("user"):
                config["sejong"] = env_sejong

        return config

    # -- 콜백 등록 --

    def on_state_change(self, callback: Callable):
        self._on_state_change = callback

    def on_progress(self, callback: Callable):
        self._on_progress = callback

    def on_result(self, callback: Callable):
        self._on_result = callback

    def on_log(self, callback: Callable):
        self._on_log = callback

    def _emit_state(self, state: str):
        self.state = state
        if self._on_state_change:
            self._on_state_change(state)

    def _emit_progress(self, current: int, total: int, name: str = ""):
        if self._on_progress:
            self._on_progress(current, total, name)

    def _emit_result(self, result: dict):
        if self._on_result:
            self._on_result(result)

    def _emit_log(self, message: str, level: str = "info"):
        if self._on_log:
            self._on_log(message, level)

    # -- 초기화 --

    def initialize(self) -> dict:
        """
        시스템 초기화
        1. 카카오톡 창 찾기
        2. 위치/크기 고정
        3. 스크린샷 캡처 & 캘리브레이션
        4. 사용자 확인 대기 (UI에서 처리)
        """
        self._emit_state(OrchestratorState.INITIALIZING)
        self._emit_log("시스템 초기화 중...")

        # 화면 정보
        screen_info = self.window_ctrl.get_screen_info()
        self._emit_log(f"화면: {screen_info['screen_width']}x{screen_info['screen_height']} "
                       f"(DPI: {screen_info['dpi_scale']}x)")

        # 카카오톡 찾기
        if not self.window_ctrl.find_kakao_window():
            self._emit_state(OrchestratorState.ERROR)
            self._emit_log("카카오톡이 실행되어 있지 않습니다!", "error")
            return {"success": False, "kakao_found": False,
                    "error": "카카오톡이 실행되어 있지 않습니다."}

        self._emit_log("카카오톡 발견!")

        # 활성화 및 위치 고정
        self.window_ctrl.activate_kakao()
        import time
        time.sleep(0.5)
        self.window_ctrl.calculate_kakao_position()
        positioned = self.window_ctrl.position_kakao_window()

        if not positioned:
            self._emit_log("카카오톡 창 위치 조정 실패. 수동으로 배치해주세요.", "warning")

        time.sleep(0.5)  # 창 이동 후 안정화 대기

        # 스크린샷 기반 캘리브레이션
        self._emit_log("카카오톡 창 캡처 중...")
        calibration = self.window_ctrl.calibrate(self.screen_capture)

        if not calibration.get("success"):
            self._emit_state(OrchestratorState.ERROR)
            self._emit_log(f"캘리브레이션 실패: {calibration.get('error', '')}", "error")
            return {"success": False, "error": calibration.get("error", "캘리브레이션 실패")}

        self.coordinates = calibration["coordinates"]
        self._emit_log(f"캘리브레이션 완료! 스크린샷: {calibration['screenshot_path']}")

        # 좌표 로그
        search_icon = self.coordinates.get("search_icon", {})
        self._emit_log(f"  돋보기 아이콘: ({search_icon.get('x')}, {search_icon.get('y')})")
        msg_input = self.coordinates.get("message_input", {})
        self._emit_log(f"  메시지 입력창: ({msg_input.get('x')}, {msg_input.get('y')})")

        return {
            "success": True,
            "calibration_pending": True,
            "screenshot_path": calibration["screenshot_path"],
            "screen": screen_info,
            "kakao_rect": self.window_ctrl.kakao_rect,
            "coordinates": self.coordinates
        }

    def confirm_calibration(self) -> dict:
        """
        캘리브레이션 확인 - 사용자가 스크린샷을 확인한 후 호출
        KakaoSender를 초기화하고 발송 준비 상태로 전환
        """
        if not self.coordinates:
            return {"success": False, "error": "먼저 초기화를 실행해주세요."}

        # KakaoSender 초기화
        self.sender = KakaoSender(self.coordinates, self.config)

        self._emit_state(OrchestratorState.READY)
        self._emit_log("셋팅 확인 완료! 발송 준비 완료.")

        return {"success": True}

    # -- 발송 준비 --

    def prepare_send_queue(self, category: str, template_content: str) -> list[dict]:
        """
        발송 큐 준비

        Args:
            category: 카테고리 ("friend", "family", "business", "all" 등)
            template_content: 메시지 템플릿 텍스트
        """
        contacts = self.contact_mgr.get_by_category(category)
        self.send_queue = []

        for contact in contacts:
            message = self.message_engine.substitute(
                template_content,
                contact.to_dict()
            )
            self.send_queue.append({
                "contact": contact,
                "message": message
            })

        self.total_count = len(self.send_queue)
        self.current_index = 0

        self._emit_log(f"발송 큐 준비 완료: {self.total_count}명")
        return [
            {
                "name": item["contact"].name,
                "category": item["contact"].category,
                "message": item["message"]
            }
            for item in self.send_queue
        ]

    def prepare_custom_queue(self, contacts: list[Contact], template_content: str,
                            image_path: str = None,
                            template_contents: list = None) -> list[dict]:
        """선택된 연락처로 발송 큐 준비
        template_contents: 여러 변형 메시지 리스트 (있으면 랜덤 선택)
        """
        self.send_queue = []

        for contact in contacts:
            if template_contents and len(template_contents) > 1:
                message = self.message_engine.substitute_random(
                    template_contents, contact.to_dict()
                )
            else:
                message = self.message_engine.substitute(
                    template_content,
                    contact.to_dict()
                )
            self.send_queue.append({
                "contact": contact,
                "message": message,
                "image_path": image_path
            })

        self.total_count = len(self.send_queue)
        self.current_index = 0

        self._emit_log(f"발송 큐 준비 완료: {self.total_count}명")
        return [
            {
                "name": item["contact"].name,
                "category": item["contact"].category,
                "message": item["message"]
            }
            for item in self.send_queue
        ]

    # -- 발송 실행 --

    def start_sending(self):
        """발송 시작 (별도 스레드)"""
        if self.state == OrchestratorState.SENDING:
            self._emit_log("이미 발송 중입니다.", "warning")
            return

        if not self.send_queue:
            self._emit_log("발송할 대상이 없습니다.", "warning")
            return

        if not self.sender:
            self._emit_log("먼저 초기화를 실행해주세요.", "error")
            return

        # 완료/에러 상태에서 다시 시작 → 처음부터
        if self.state in (OrchestratorState.COMPLETED, OrchestratorState.ERROR,
                          OrchestratorState.IDLE, OrchestratorState.READY):
            self.current_index = 0
            self._emit_log("발송 큐를 처음부터 다시 시작합니다.")

        self.sender.resume()
        self._send_thread = threading.Thread(target=self._send_loop, daemon=True)
        self._send_thread.start()

    def _send_loop(self):
        """발송 루프 (스레드에서 실행) - 안전 장치 포함"""
        self._emit_state(OrchestratorState.SENDING)
        self.report.start_session()
        self._emit_log("=" * 40)
        self._emit_log("발송 시작!")
        self._emit_log("=" * 40)

        # 안전 정지 콜백 등록
        def on_safety(msg):
            self._emit_log(f"🛑 {msg}", "error")

        self.sender.on_safety_stop(on_safety)

        while self.current_index < self.total_count:
            # 정지 체크 (일시정지 or 안전 정지)
            if self.sender._stop_flag:
                if self.sender._safety_error:
                    self._emit_state(OrchestratorState.ERROR)
                    self._emit_log("=" * 40)
                    self._emit_log(f"안전 정지로 발송 중단!", "error")
                    self._emit_log(f"원인: {self.sender._safety_error}", "error")
                    self._emit_log("=" * 40)
                    break
                else:
                    self._emit_state(OrchestratorState.PAUSED)
                    self._emit_log("발송 일시정지됨")
                    return

            # 일일 한도 체크
            if not self.sender.check_daily_limit():
                self._emit_log(f"일일 발송 한도({self.sender.daily_limit}명) 도달! 내일 다시 시도하세요.", "warning")
                break

            item = self.send_queue[self.current_index]
            contact = item["contact"]
            message = item["message"]
            image_path = item.get("image_path")

            self._emit_progress(self.current_index + 1, self.total_count, contact.name)
            self._emit_log(f"[{self.current_index + 1}/{self.total_count}] "
                           f"{contact.name}에게 발송 중..."
                           + (" (이미지 첨부)" if image_path else ""))

            # 발송
            from core.kakao_sender import _debug_log
            _debug_log(f"orchestrator: {contact.name}에게 send_to_contact 호출")
            result = self.sender.send_to_contact(contact.name, message, image_path=image_path)
            _debug_log(f"orchestrator: 결과={result.status} detail={result.detail}")
            result_dict = result.to_dict()
            self.report.add_result(result_dict)
            self._emit_result(result_dict)

            if result.status == SendResult.SUCCESS:
                self.contact_mgr.mark_sent(contact.id)
                self._emit_log(f"  {contact.name} 발송 성공!", "success")
            elif result.status == SendResult.FAILED_SAFETY:
                self._emit_log(f"  {contact.name}: {result.detail}", "error")
                break
            elif result.status == SendResult.FAILED_NOT_FOUND:
                self._emit_log(f"  {contact.name}: 검색 결과 없음 - 건너뜀", "warning")
            else:
                self._emit_log(f"  {contact.name}: {result.detail}", "error")

            self.current_index += 1

            # 안전 정지 확인
            if self.sender._stop_flag:
                break

            # 마지막이 아니면 딜레이
            if self.current_index < self.total_count:
                # N명마다 긴 휴식 (계정 보호)
                if self.sender.should_rest():
                    rest = self.sender.take_rest()
                    self._emit_log(f"  계정 보호 휴식: {rest:.0f}초 쉬는 중...", "info")

                delay = self.sender.random_delay()
                self._emit_log(f"  다음 발송까지 {delay:.0f}초 대기...")

        # 완료/중단 처리
        if self.sender._safety_error:
            self._emit_state(OrchestratorState.ERROR)
        elif self.current_index >= self.total_count:
            self._emit_state(OrchestratorState.COMPLETED)
        else:
            self._emit_state(OrchestratorState.PAUSED)

        self._emit_log("=" * 40)
        stats = self.report.get_statistics()
        self._emit_log(f"총 {stats['total']}명 | "
                       f"성공 {stats['success']}명 | "
                       f"실패 {stats['failed']}명 | "
                       f"성공률 {stats['success_rate']}%")
        self._emit_log("=" * 40)

        log_path = self.report.save_session_log()
        self._emit_log(f"로그 저장: {log_path}")

    def pause_sending(self):
        """발송 일시정지"""
        if self.sender:
            self.sender.stop()
            self._emit_log("발송 일시정지 요청됨")

    def resume_sending(self):
        """발송 재개"""
        if self.state == OrchestratorState.PAUSED:
            self.sender.resume()
            self._send_thread = threading.Thread(target=self._send_loop, daemon=True)
            self._send_thread.start()
            self._emit_log("발송 재개!")

    def stop_sending(self):
        """발송 완전 중지 - 상태 완전 초기화"""
        if self.sender:
            self.sender.stop()
            # 플래그 초기화 → 다음 발송 시 재시작 가능
            self.sender._stop_flag = False
            self.sender._safety_error = None
        self.current_index = 0
        self.send_queue = []
        self.state = OrchestratorState.IDLE
        self._emit_state(OrchestratorState.IDLE)
        self._emit_log("발송 중지됨. 큐가 초기화되었습니다.")

    # -- 리포트 --

    # -- 세종텔레콤 연동 --

    def init_sejong(self, config_dict: dict = None) -> dict:
        """세종텔레콤 DB 연결 초기화"""
        try:
            sejong_cfg = config_dict or self.config.get("sejong", {})
            sc = SejongConfig(sejong_cfg)
            self.sejong_sender = SejongSender(sc)
            result = self.sejong_sender.test_connection()
            if result["success"]:
                self._emit_log(f"세종텔레콤 DB 연결 성공: {result['message']}")
            else:
                self._emit_log(f"세종텔레콤 DB 연결 실패: {result['message']}", "error")
            return result
        except Exception as e:
            self._emit_log(f"세종텔레콤 초기화 실패: {e}", "error")
            return {"success": False, "message": str(e)}

    def start_sejong_sending(self):
        """세종텔레콤(알림톡/SMS) 발송 시작 (별도 스레드)"""
        if self.state == OrchestratorState.SENDING:
            self._emit_log("이미 발송 중입니다.", "warning")
            return
        if not self.send_queue:
            self._emit_log("발송할 대상이 없습니다.", "warning")
            return
        if not self.sejong_sender:
            self._emit_log("세종텔레콤 DB 연결이 필요합니다.", "error")
            return

        self.current_index = 0
        self._send_thread = threading.Thread(target=self._sejong_send_loop, daemon=True)
        self._send_thread.start()

    def _sejong_send_loop(self):
        """세종텔레콤 발송 루프"""
        self._emit_state(OrchestratorState.SENDING)
        self.report.start_session()
        self._emit_log("=" * 40)
        method_label = "알림톡" if self.send_method == "alimtalk" else "SMS/LMS"
        self._emit_log(f"{method_label} 발송 시작! (세종텔레콤)")
        self._emit_log("=" * 40)

        sejong_cfg = self.config.get("sejong", {})
        template_code = sejong_cfg.get("kakao", {}).get("template_code", "")

        while self.current_index < self.total_count:
            item = self.send_queue[self.current_index]
            contact = item["contact"]
            message = item["message"]

            self._emit_progress(self.current_index + 1, self.total_count, contact.name)
            self._emit_log(f"[{self.current_index + 1}/{self.total_count}] "
                           f"{contact.name}에게 {method_label} 발송 중...")

            phone = contact.phone
            if not phone:
                result_dict = {
                    "contact_name": contact.name, "phone": "",
                    "status": "failed", "detail": "전화번호 없음"
                }
                self.report.add_result(result_dict)
                self._emit_result(result_dict)
                self._emit_log(f"  {contact.name}: 전화번호 없음 - 건너뜀", "warning")
                self.current_index += 1
                continue

            # 발송 방법에 따라 분기
            if self.send_method == "alimtalk":
                result = self.sejong_sender.send_alimtalk(
                    phone=phone, message=message,
                    template_code=template_code,
                    contact_name=contact.name
                )
            else:  # sms
                result = self.sejong_sender.send_auto(
                    phone=phone, message=message,
                    contact_name=contact.name
                )

            result_dict = result.to_dict()
            # 통일된 status 키
            result_dict["status"] = "success" if result.status == SejongSendResult.SUCCESS else "failed"
            self.report.add_result(result_dict)
            self._emit_result(result_dict)

            if result.status == SejongSendResult.SUCCESS:
                self.contact_mgr.mark_sent(contact.id)
                self._emit_log(f"  {contact.name} 접수 성공! ({result.detail})", "success")
            else:
                self._emit_log(f"  {contact.name}: {result.detail}", "error")

            self.current_index += 1

        # 완료
        self._emit_state(OrchestratorState.COMPLETED)
        self._emit_log("=" * 40)
        stats = self.report.get_statistics()
        self._emit_log(f"총 {stats['total']}명 | "
                       f"성공 {stats['success']}명 | "
                       f"실패 {stats['failed']}명 | "
                       f"성공률 {stats['success_rate']}%")
        self._emit_log("=" * 40)
        log_path = self.report.save_session_log()
        self._emit_log(f"로그 저장: {log_path}")

    def get_current_stats(self) -> dict:
        """현재 발송 통계"""
        return self.report.get_statistics()

    def export_report(self, filepath: str = None) -> str:
        """리포트 엑셀 내보내기"""
        return self.report.export_report_excel(filepath)

    def get_send_history(self, limit: int = 10) -> list:
        """최근 발송 이력"""
        return self.report.get_history(limit)
