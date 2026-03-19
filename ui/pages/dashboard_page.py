"""
Dashboard Page - 메인 대시보드
SaaS 모드: 잔액/사용량/발송이력 + 매크로 컨트롤
로컬 모드: 매크로 컨트롤만
"""

import customtkinter as ctk
from ui.theme import AppTheme as T
from ui.components.widgets import StatCard, LogPanel


class DashboardPage(ctk.CTkFrame):
    """대시보드 페이지"""

    def __init__(self, parent, orchestrator=None, api_client=None, **kwargs):
        super().__init__(parent, fg_color=T.BG_DARK, **kwargs)
        self.orchestrator = orchestrator
        self.api_client = api_client
        self._build()

    def _build(self):
        # -- 페이지 제목 --
        header = ctk.CTkFrame(self, fg_color="transparent", height=50)
        header.pack(fill="x", padx=24, pady=(20, 16))
        header.pack_propagate(False)

        ctk.CTkLabel(
            header, text="📊 대시보드",
            font=(T.get_font_family(), T.FONT_SIZE_TITLE, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(side="left")

        if self.api_client:
            ctk.CTkButton(
                header, text="🔄 새로고침", width=90, height=30,
                font=(T.get_font_family(), T.FONT_SIZE_SMALL),
                fg_color=T.BG_HOVER, hover_color=T.BORDER,
                text_color=T.TEXT_PRIMARY, corner_radius=6,
                command=self.refresh_stats
            ).pack(side="right")

        # -- 통계 카드 영역 --
        stats_frame = ctk.CTkFrame(self, fg_color="transparent")
        stats_frame.pack(fill="x", padx=24, pady=(0, 16))

        if self.api_client:
            # SaaS 모드: 잔액 + 오늘발송 + 이번달발송 + 상태
            stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="stat")
            self.stat_balance = StatCard(stats_frame, "잔액", "0원", T.SUCCESS)
            self.stat_balance.grid(row=0, column=0, padx=(0, 8), sticky="ew")
            self.stat_today = StatCard(stats_frame, "오늘 발송", "0건", T.ACCENT)
            self.stat_today.grid(row=0, column=1, padx=4, sticky="ew")
            self.stat_month = StatCard(stats_frame, "이번 달", "0건", T.INFO)
            self.stat_month.grid(row=0, column=2, padx=4, sticky="ew")
            self.stat_status = StatCard(stats_frame, "상태", "대기", T.TEXT_SECONDARY)
            self.stat_status.grid(row=0, column=3, padx=(8, 0), sticky="ew")
        else:
            # 로컬 모드: 연락처 + 오늘발송 + 성공률 + 상태
            stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="stat")
            self.stat_contacts = StatCard(stats_frame, "총 연락처", "0명", T.INFO)
            self.stat_contacts.grid(row=0, column=0, padx=(0, 8), sticky="ew")
            self.stat_today = StatCard(stats_frame, "오늘 발송", "0건", T.ACCENT)
            self.stat_today.grid(row=0, column=1, padx=4, sticky="ew")
            self.stat_success = StatCard(stats_frame, "성공률", "0%", T.SUCCESS)
            self.stat_success.grid(row=0, column=2, padx=4, sticky="ew")
            self.stat_status = StatCard(stats_frame, "상태", "대기", T.TEXT_SECONDARY)
            self.stat_status.grid(row=0, column=3, padx=(8, 0), sticky="ew")

        # -- 빠른 실행 (매크로 컨트롤 - 양쪽 모두) --
        quick_frame = ctk.CTkFrame(self, fg_color=T.BG_CARD, corner_radius=T.CARD_RADIUS,
                                    border_width=1, border_color=T.BORDER)
        quick_frame.pack(fill="x", padx=24, pady=(0, 16))

        ctk.CTkLabel(
            quick_frame, text="⚡ 빠른 실행",
            font=(T.get_font_family(), T.FONT_SIZE_HEADER, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(padx=T.CARD_PADDING, pady=(T.CARD_PADDING, 8), anchor="w")

        btn_frame = ctk.CTkFrame(quick_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=T.CARD_PADDING, pady=(0, T.CARD_PADDING))

        ctk.CTkButton(
            btn_frame, text="🔧 카카오톡 초기화",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY,
            height=T.BUTTON_HEIGHT, corner_radius=6,
            command=self._on_initialize
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame, text="📐 카카오톡 자동 배치",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color="#1a5276", hover_color="#2471a3",
            text_color=T.TEXT_PRIMARY,
            height=T.BUTTON_HEIGHT, corner_radius=6,
            command=self._on_position_kakao
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame, text="📋 엑셀 가져오기",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY,
            height=T.BUTTON_HEIGHT, corner_radius=6,
            command=self._on_import_excel
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame, text="🚀 발송 시작",
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            fg_color=T.ACCENT, hover_color=T.ACCENT_HOVER,
            text_color=T.TEXT_ON_ACCENT,
            height=T.BUTTON_HEIGHT, corner_radius=6,
            command=self._on_quick_send
        ).pack(side="left")

        # -- 하단: 최근 이력 / 로그 --
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(fill="both", expand=True, padx=24, pady=(0, 16))
        bottom_frame.grid_columnconfigure((0, 1), weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        # 최근 발송
        history_card = ctk.CTkFrame(bottom_frame, fg_color=T.BG_CARD,
                                     corner_radius=T.CARD_RADIUS,
                                     border_width=1, border_color=T.BORDER)
        history_card.grid(row=0, column=0, padx=(0, 8), sticky="nsew")

        ctk.CTkLabel(
            history_card, text="📋 최근 발송 이력",
            font=(T.get_font_family(), T.FONT_SIZE_HEADER, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(padx=T.CARD_PADDING, pady=T.CARD_PADDING, anchor="w")

        self.history_list = ctk.CTkTextbox(
            history_card,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT,
            text_color=T.TEXT_SECONDARY,
            corner_radius=6,
            state="disabled"
        )
        self.history_list.pack(fill="both", expand=True,
                               padx=T.CARD_PADDING, pady=(0, T.CARD_PADDING))

        # 시스템 로그
        self.log_panel = LogPanel(bottom_frame, title="🔧 시스템 로그")
        self.log_panel.grid(row=0, column=1, padx=(8, 0), sticky="nsew")

        # 초기 로드
        self.refresh_stats()

    def refresh_stats(self):
        """통계 새로고침"""
        if self.api_client and self.api_client.is_logged_in:
            self._refresh_saas_stats()
        elif self.orchestrator:
            count = self.orchestrator.contact_mgr.get_count()
            self.stat_contacts.update_value(f"{count}명")

    def _refresh_saas_stats(self):
        """SaaS 모드 통계 갱신"""
        try:
            daily = self.api_client.get_daily_usage()
            self.stat_balance.update_value(f"{daily.get('balance', 0):,}원")
            self.stat_today.update_value(f"{daily.get('count', 0):,}건")

            monthly = self.api_client.get_monthly_usage()
            self.stat_month.update_value(f"{monthly.get('count', 0):,}건")

            # 최근 발송 이력
            history = self.api_client.get_send_history(page=1, size=15)
            self.history_list.configure(state="normal")
            self.history_list.delete("1.0", "end")
            if history:
                for item in history:
                    time_str = item.get("created_at", "")[:16]
                    name = item.get("contact_name", "-")
                    msg_type = item.get("msg_type", "")
                    status = item.get("status", "")
                    cost = item.get("cost", 0)
                    self.history_list.insert("end",
                        f"{time_str}  {name}  [{msg_type}]  {status}  {cost}원\n")
            else:
                self.history_list.insert("end", "발송 기록이 없습니다.")
            self.history_list.configure(state="disabled")

            self.log_panel.add_log("통계 갱신 완료", "success")
        except Exception as e:
            self.log_panel.add_log(f"통계 조회 실패: {e}", "error")

    def _on_initialize(self):
        if not self.orchestrator:
            return

        positions_path = self.orchestrator.base_dir / "config" / "learned_positions.json"
        if not positions_path.exists():
            self.log_panel.add_log("학습된 위치가 없습니다.", "warning")
            self.log_panel.add_log("설정 → '위치 학습 시작'을 먼저 실행해주세요.", "info")
            self.stat_status.update_value("미학습", T.WARNING)
            return

        import json
        with open(positions_path, "r", encoding="utf-8") as f:
            positions = json.load(f)

        self.log_panel.add_log("저장된 위치 로드 중...")
        self.orchestrator.coordinates = positions

        if not self.orchestrator.window_ctrl.find_kakao_window():
            self.log_panel.add_log("카카오톡이 실행되어 있지 않습니다!", "error")
            self.stat_status.update_value("오류", T.ERROR)
            return

        self.orchestrator.window_ctrl.activate_kakao()
        self.log_panel.add_log("카카오톡 발견!")

        import time
        time.sleep(0.3)
        self.orchestrator.window_ctrl.calculate_kakao_position()
        positioned = self.orchestrator.window_ctrl.position_kakao_window()
        if positioned:
            rect = self.orchestrator.window_ctrl.kakao_rect
            self.log_panel.add_log(
                f"카카오톡 자동 배치: ({rect['x']}, {rect['y']}) "
                f"{rect['width']}×{rect['height']}", "success"
            )
        else:
            self.log_panel.add_log("카카오톡 자동 배치 실패. 수동으로 배치해주세요.", "warning")

        result = self.orchestrator.confirm_calibration()
        if result.get("success"):
            self.log_panel.add_log("초기화 완료! 발송 준비 완료.", "success")
            self.stat_status.update_value("준비", T.SUCCESS)
            for key, pos in positions.items():
                if isinstance(pos, dict) and "x" in pos:
                    self.log_panel.add_log(
                        f"  {pos.get('description', key)}: ({pos['x']}, {pos['y']})",
                        "info"
                    )
        else:
            self.log_panel.add_log("초기화 실패", "error")
            self.stat_status.update_value("오류", T.ERROR)

    def auto_initialize(self):
        """앱 시작 시 자동 초기화"""
        if not self.orchestrator:
            return
        positions_path = self.orchestrator.base_dir / "config" / "learned_positions.json"
        if positions_path.exists():
            self.log_panel.add_log("저장된 설정 자동 로드 중...", "info")
            self._on_initialize()

    def _on_position_kakao(self):
        if not self.orchestrator:
            return
        if not self.orchestrator.window_ctrl.find_kakao_window():
            self.log_panel.add_log("카카오톡이 실행되어 있지 않습니다!", "error")
            return
        self.orchestrator.window_ctrl.activate_kakao()
        import time
        time.sleep(0.3)
        self.orchestrator.window_ctrl.calculate_kakao_position()
        positioned = self.orchestrator.window_ctrl.position_kakao_window()
        if positioned:
            rect = self.orchestrator.window_ctrl.kakao_rect
            self.log_panel.add_log(
                f"카카오톡 자동 배치 완료: ({rect['x']}, {rect['y']}) "
                f"{rect['width']}x{rect['height']}", "success"
            )
        else:
            self.log_panel.add_log("카카오톡 자동 배치 실패", "error")

    def _on_import_excel(self):
        from tkinter import filedialog, messagebox
        filepath = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not filepath:
            return

        if self.api_client and self.api_client.is_logged_in:
            try:
                result = self.api_client.import_contacts(filepath)
                self.log_panel.add_log(f"엑셀 가져오기: {result.get('added', 0)}명 추가", "success")
            except Exception as e:
                self.log_panel.add_log(f"엑셀 가져오기 실패: {e}", "error")
        elif self.orchestrator:
            result = self.orchestrator.contact_mgr.import_from_excel(filepath)
            self.log_panel.add_log(
                f"엑셀 가져오기: {result['success']}명 추가, {result['skipped']}명 건너뜀"
            )
        self.refresh_stats()

    def _on_quick_send(self):
        self.log_panel.add_log("발송 페이지로 이동하세요 (🚀 발송)")

    def add_log(self, message: str, level: str = "info"):
        """외부에서 로그 추가"""
        self.log_panel.add_log(message, level)
