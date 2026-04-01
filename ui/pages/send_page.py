"""
Send Page - 발송 실행 페이지 (LOCAL-ONLY)
대상 선택 (Treeview 다중선택) + 메시지 작성 + 발송 실행
카카오톡 봇 전용 로컬 모드
"""

import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from tkinter import messagebox
from ui.theme import AppTheme as T
from ui.components.widgets import ProgressCard, LogPanel


class SendPage(ctk.CTkFrame):
    """발송 실행 페이지 - 대상 선택 + 메시지 + 발송 (카카오톡 봇 전용)"""

    def __init__(self, parent, orchestrator=None, message_page=None, **kwargs):
        super().__init__(parent, fg_color=T.BG_DARK, **kwargs)
        self.orchestrator = orchestrator
        self.message_page = message_page
        self.selected_ids = set()  # 선택된 연락처 ID
        self.contact_checkboxes = {}  # {contact_id: checkbox}
        self._selected_template = None  # 선택된 템플릿 (변형 랜덤용)
        self._build()

    def _build(self):
        # -- 헤더 --
        header = ctk.CTkFrame(self, fg_color="transparent", height=50)
        header.pack(fill="x", padx=24, pady=(20, 8))
        header.pack_propagate(False)

        ctk.CTkLabel(
            header, text="🚀 메시지 발송",
            font=(T.get_font_family(), T.FONT_SIZE_TITLE, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(side="left")

        # -- 발송 방법 (카카오톡 봇 전용) --
        self.send_method_var = ctk.StringVar(value="카카오톡 봇")
        self._selected_method_key = "카카오톡 봇"

        method_frame = ctk.CTkFrame(self, fg_color="transparent", height=72)
        method_frame.pack(fill="x", padx=24, pady=(0, 4))
        method_frame.pack_propagate(False)

        self._method_cards = {}
        card_btn = ctk.CTkButton(
            method_frame, text="💬\n카카오톡 봇\n무료",
            width=140, height=60,
            font=(T.get_font_family(), 10, "bold"),
            fg_color=T.BG_HOVER, hover_color=T.BG_HOVER,
            text_color=T.TEXT_PRIMARY,
            border_width=2,
            border_color="#2ea043",
            corner_radius=10,
            command=lambda: None
        )
        card_btn.pack(side="left", padx=(0, 8))
        self._method_cards["카카오톡 봇"] = (card_btn, "#2ea043")

        # 발송 방법 상태 바
        self.method_status_label = ctk.CTkLabel(
            self, text="📨 카카오톡 봇으로 메시지 보내기 (무료)",
            font=(T.get_font_family(), 11, "bold"),
            text_color=T.TEXT_SECONDARY,
            fg_color=T.BG_CARD, corner_radius=6, height=28
        )
        self.method_status_label.pack(fill="x", padx=24, pady=(0, 8))

        # -- 상단: 대상 선택 (좌) + 메시지 작성 (우) --
        top_frame = ctk.CTkFrame(self, fg_color="transparent", height=350)
        top_frame.pack(fill="both", padx=24, pady=(0, 8))
        top_frame.pack_propagate(False)
        top_frame.grid_columnconfigure(0, weight=2)
        top_frame.grid_columnconfigure(1, weight=3)
        top_frame.grid_rowconfigure(0, weight=1)

        # ═══ 좌측: 대상 선택 ═══
        left_card = ctk.CTkFrame(top_frame, fg_color=T.BG_CARD,
                                  corner_radius=T.CARD_RADIUS,
                                  border_width=1, border_color=T.BORDER)
        left_card.grid(row=0, column=0, padx=(0, 6), sticky="nsew")

        # 제목
        left_header = ctk.CTkFrame(left_card, fg_color="transparent", height=36)
        left_header.pack(fill="x", padx=12, pady=(12, 8))
        left_header.pack_propagate(False)

        ctk.CTkLabel(
            left_header, text="👥 발송 대상",
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(side="left")

        self.selected_count_label = ctk.CTkLabel(
            left_header, text="0명 선택",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL, "bold"),
            text_color=T.ACCENT
        )
        self.selected_count_label.pack(side="right")

        # 검색
        self.contact_search = ctk.CTkEntry(
            left_card, placeholder_text="🔍 이름, 회사, 메모 검색...",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, border_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=32, corner_radius=6
        )
        self.contact_search.pack(fill="x", padx=12, pady=(0, 6))
        self.contact_search.bind("<KeyRelease>", lambda e: self._refresh_contact_list())

        # 카테고리 필터
        self.cat_filter_frame = ctk.CTkFrame(left_card, fg_color="transparent", height=28)
        self.cat_filter_frame.pack(fill="x", padx=12, pady=(0, 6))
        self.cat_filter_frame.pack_propagate(False)
        self.cat_filter_var = ctk.StringVar(value="all")

        # 전체선택 / 해제
        select_frame = ctk.CTkFrame(left_card, fg_color="transparent", height=28)
        select_frame.pack(fill="x", padx=12, pady=(0, 4))
        select_frame.pack_propagate(False)

        ctk.CTkButton(
            select_frame, text="전체 선택", width=70, height=24,
            font=(T.get_font_family(), 9),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_SECONDARY, corner_radius=4,
            command=self._select_all
        ).pack(side="left", padx=(0, 4))

        ctk.CTkButton(
            select_frame, text="전체 해제", width=70, height=24,
            font=(T.get_font_family(), 9),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_SECONDARY, corner_radius=4,
            command=self._deselect_all
        ).pack(side="left")

        # Treeview 스타일
        style = ttk.Style()
        style.configure("Send.Treeview",
                         background="#1c2333", foreground="#e6edf3",
                         fieldbackground="#1c2333", borderwidth=0,
                         font=(T.get_font_family(), 10), rowheight=28)
        style.configure("Send.Treeview.Heading",
                         background="#2d333b", foreground="#e6edf3",
                         font=(T.get_font_family(), 9, "bold"), borderwidth=0)
        style.map("Send.Treeview",
                   background=[("selected", "#2f81f7")],
                   foreground=[("selected", "#ffffff")])

        # 연락처 Treeview (체크박스 + 이름 + 카테고리 + 전화번호)
        tree_frame = tk.Frame(left_card, bg="#1c2333")
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        send_cols = ("check", "name", "category", "phone")
        self.send_tree = ttk.Treeview(
            tree_frame, columns=send_cols, show="headings",
            selectmode="none", style="Send.Treeview"
        )
        self.send_tree.heading("check", text="V", anchor="center")
        self.send_tree.heading("name", text="이름", anchor="w")
        self.send_tree.heading("category", text="카테고리", anchor="w")
        self.send_tree.heading("phone", text="전화번호", anchor="w")
        self.send_tree.column("check", width=30, minwidth=30, stretch=False)
        self.send_tree.column("name", width=80, minwidth=60)
        self.send_tree.column("category", width=60, minwidth=50)
        self.send_tree.column("phone", width=100, minwidth=80)

        # 체크된 항목 태그 (배경색 변경)
        self.send_tree.tag_configure("checked", background="#1a3a2a", foreground="#3fb950")
        self.send_tree.tag_configure("unchecked", background="#1c2333", foreground="#e6edf3")

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.send_tree.yview)
        self.send_tree.configure(yscrollcommand=sb.set)
        self.send_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 클릭 → 체크 토글
        self.send_tree.bind("<Button-1>", self._on_tree_click_toggle)
        self._send_tree_id_map = {}  # iid → contact_id

        # ═══ 우측: 메시지 작성 ═══
        right_card = ctk.CTkFrame(top_frame, fg_color=T.BG_CARD,
                                   corner_radius=T.CARD_RADIUS,
                                   border_width=1, border_color=T.BORDER)
        right_card.grid(row=0, column=1, padx=(6, 0), sticky="nsew")

        # 제목
        right_header = ctk.CTkFrame(right_card, fg_color="transparent", height=36)
        right_header.pack(fill="x", padx=12, pady=(12, 8))
        right_header.pack_propagate(False)

        ctk.CTkLabel(
            right_header, text="💬 메시지",
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(side="left")

        # 템플릿 불러오기
        tmpl_frame = ctk.CTkFrame(right_card, fg_color="transparent", height=32)
        tmpl_frame.pack(fill="x", padx=12, pady=(0, 6))
        tmpl_frame.pack_propagate(False)

        ctk.CTkLabel(
            tmpl_frame, text="템플릿:",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        ).pack(side="left", padx=(0, 4))

        self.template_var = ctk.StringVar(value="직접 입력")
        self.template_menu = ctk.CTkOptionMenu(
            tmpl_frame, values=["직접 입력"],
            variable=self.template_var,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER,
            text_color=T.TEXT_PRIMARY, height=28, corner_radius=4,
            command=self._on_template_select
        )
        self.template_menu.pack(side="left", fill="x", expand=True, padx=(0, 4))

        ctk.CTkButton(
            tmpl_frame, text="🔄", width=28, height=28,
            font=(T.get_font_family(), 12),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_SECONDARY, corner_radius=4,
            command=self._refresh_templates
        ).pack(side="right")

        # 템플릿 미리보기 스니펫 (항상 존재, 텍스트 비면 높이 0)
        self._tmpl_preview_frame = ctk.CTkFrame(right_card, fg_color="transparent", height=0)
        self._tmpl_preview_frame.pack(fill="x", padx=12, pady=0)
        self._tmpl_preview_frame.pack_propagate(False)

        self.template_preview_label = ctk.CTkLabel(
            self._tmpl_preview_frame, text="",
            font=(T.get_font_family(), 9),
            text_color=T.TEXT_MUTED, anchor="w",
            fg_color=T.BG_INPUT, corner_radius=4, height=20
        )
        self.template_preview_label.pack(fill="x")

        # 변수 버튼들
        var_frame = ctk.CTkFrame(right_card, fg_color="transparent", height=28)
        var_frame.pack(fill="x", padx=12, pady=(0, 4))
        var_frame.pack_propagate(False)

        ctk.CTkLabel(
            var_frame, text="변수:",
            font=(T.get_font_family(), 9),
            text_color=T.TEXT_MUTED
        ).pack(side="left", padx=(0, 4))

        for var in ["%이름%", "%카테고리%", "%회사%", "%직급%", "%생일%", "%기념일%", "%날짜%", "%요일%"]:
            ctk.CTkButton(
                var_frame, text=var, width=50, height=22,
                font=(T.get_font_family(), 9),
                fg_color=T.BG_HOVER, hover_color=T.BORDER,
                text_color=T.INFO, corner_radius=4,
                command=lambda v=var: self.msg_editor.insert("insert", v)
            ).pack(side="left", padx=1)

        # 메시지 에디터
        self.msg_editor = ctk.CTkTextbox(
            right_card,
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_INPUT, text_color=T.TEXT_PRIMARY,
            corner_radius=6, border_width=1, border_color=T.BORDER
        )
        self.msg_editor.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self.msg_editor.insert("1.0", "안녕하세요 %이름%님!\n\n")

        # 이미지 첨부
        img_frame = ctk.CTkFrame(right_card, fg_color="transparent", height=28)
        img_frame.pack(fill="x", padx=12, pady=(0, 4))
        img_frame.pack_propagate(False)

        self.image_path = None

        ctk.CTkButton(
            img_frame, text="📎 이미지 첨부", width=100, height=24,
            font=(T.get_font_family(), 9),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, corner_radius=4,
            command=self._select_image
        ).pack(side="left", padx=(0, 4))

        self.image_label = ctk.CTkLabel(
            img_frame, text="첨부 없음",
            font=(T.get_font_family(), 9),
            text_color=T.TEXT_MUTED
        )
        self.image_label.pack(side="left", padx=(0, 4))

        self.image_clear_btn = ctk.CTkButton(
            img_frame, text="✕", width=22, height=22,
            font=(T.get_font_family(), 10),
            fg_color="transparent", hover_color=T.ERROR,
            text_color=T.TEXT_MUTED, corner_radius=4,
            command=self._clear_image
        )

        # 미리보기
        preview_header = ctk.CTkFrame(right_card, fg_color="transparent", height=22)
        preview_header.pack(fill="x", padx=12)
        preview_header.pack_propagate(False)

        ctk.CTkLabel(
            preview_header, text="👁 미리보기",
            font=(T.get_font_family(), 9, "bold"),
            text_color=T.TEXT_SECONDARY
        ).pack(side="left")

        ctk.CTkButton(
            preview_header, text="새로고침", width=60, height=18,
            font=(T.get_font_family(), 9),
            fg_color="transparent", hover_color=T.BG_HOVER,
            text_color=T.TEXT_MUTED, corner_radius=4,
            command=self._update_preview
        ).pack(side="right")

        self.preview_box = ctk.CTkTextbox(
            right_card, height=70,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, text_color=T.SUCCESS,
            corner_radius=6, state="disabled"
        )
        self.preview_box.pack(fill="x", padx=12, pady=(2, 12))

        # -- 하단: 컨트롤 + 로그 --
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(fill="both", expand=True, padx=24, pady=(0, 12))

        # 컨트롤 바
        control_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent", height=50)
        control_frame.pack(fill="x", pady=(0, 8))
        control_frame.pack_propagate(False)

        self.start_btn = ctk.CTkButton(
            control_frame, text="▶  발송 시작", width=140,
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            fg_color="#2ea043", hover_color="#3fb950",
            text_color="#ffffff", height=42, corner_radius=8,
            command=self._start_send
        )
        self.start_btn.pack(side="left", padx=(0, 6))

        self.pause_btn = ctk.CTkButton(
            control_frame, text="⏸  일시정지", width=110,
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color="#b08800", hover_color="#d29922",
            text_color="#ffffff", height=42, corner_radius=8,
            command=self._pause_send
        )
        self.pause_btn.pack(side="left", padx=(0, 6))

        self.stop_btn = ctk.CTkButton(
            control_frame, text="⏹  중지", width=90,
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color="#b62324", hover_color="#f85149",
            text_color="#ffffff", height=42, corner_radius=8,
            command=self._stop_send
        )
        self.stop_btn.pack(side="left", padx=(0, 6))

        self.schedule_btn = ctk.CTkButton(
            control_frame, text="⏰ 예약", width=90,
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color="#2471a3", hover_color="#3498db",
            text_color="#ffffff", height=42, corner_radius=8,
            command=self._schedule_send
        )
        self.schedule_btn.pack(side="left", padx=(0, 12))

        # 딜레이 설정
        ctk.CTkLabel(
            control_frame, text="간격:",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        ).pack(side="left", padx=(0, 2))

        self.delay_min = ctk.CTkEntry(
            control_frame, width=40, height=28,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, border_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, corner_radius=4
        )
        self.delay_min.pack(side="left", padx=(0, 2))
        self.delay_min.insert(0, "3")

        ctk.CTkLabel(
            control_frame, text="~",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        ).pack(side="left", padx=2)

        self.delay_max = ctk.CTkEntry(
            control_frame, width=40, height=28,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, border_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, corner_radius=4
        )
        self.delay_max.pack(side="left", padx=(0, 2))
        self.delay_max.insert(0, "3")

        ctk.CTkLabel(
            control_frame, text="초",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        ).pack(side="left")

        # 진행 바
        self.progress_card = ProgressCard(control_frame)
        self.progress_card.pack(side="right")

        # 로그
        self.log_panel = LogPanel(bottom_frame, title="📡 발송 로그")
        self.log_panel.pack(fill="both", expand=True)

        # 초기 로드
        self._refresh_all()

    # ═══ 이미지 첨부 ═══

    def _select_image(self):
        from tkinter import filedialog
        filepath = filedialog.askopenfilename(
            title="이미지 선택",
            filetypes=[("이미지 파일", "*.png *.jpg *.jpeg *.bmp *.gif")]
        )
        if filepath:
            self.image_path = filepath
            filename = filepath.split("/")[-1].split("\\")[-1]
            self.image_label.configure(text=filename, text_color=T.ACCENT)
            self.image_clear_btn.pack(side="left")

    def _clear_image(self):
        self.image_path = None
        self.image_label.configure(text="첨부 없음", text_color=T.TEXT_MUTED)
        self.image_clear_btn.pack_forget()

    # ═══ 연락처 리스트 ═══

    def _refresh_all(self):
        """전체 새로고침"""
        self._refresh_category_filter()
        self._refresh_contact_list()
        self._refresh_templates()

    def _refresh_category_filter(self):
        """카테고리 필터 버튼 생성"""
        for w in self.cat_filter_frame.winfo_children():
            w.destroy()

        categories = [("all", "전체")]
        if self.orchestrator:
            for cat in self.orchestrator.contact_mgr.get_all_categories():
                label = {"friend": "친구", "family": "가족", "business": "사업체",
                         "vip": "VIP", "other": "기타"}.get(cat, cat)
                categories.append((cat, label))

        current = self.cat_filter_var.get()
        for cat_id, cat_name in categories:
            is_active = cat_id == current
            btn = ctk.CTkButton(
                self.cat_filter_frame, text=cat_name, height=24,
                font=(T.get_font_family(), 9),
                fg_color=T.ACCENT if is_active else T.BG_HOVER,
                hover_color=T.ACCENT_HOVER if is_active else T.BORDER,
                text_color=T.TEXT_ON_ACCENT if is_active else T.TEXT_SECONDARY,
                corner_radius=12,
                command=lambda cid=cat_id: self._on_cat_filter(cid)
            )
            btn.pack(side="left", padx=(0, 3))

    def _on_cat_filter(self, cat_id):
        self.cat_filter_var.set(cat_id)
        self._refresh_category_filter()
        self._refresh_contact_list()

    def _refresh_contact_list(self):
        """연락처 Treeview 갱신 (체크박스 + 즉시 로드)"""
        for item in self.send_tree.get_children():
            self.send_tree.delete(item)
        self._send_tree_id_map.clear()

        search = self.contact_search.get().strip().lower()
        category = self.cat_filter_var.get()

        cat_label_map = {
            "friend": "친구", "family": "가족", "business": "사업체",
            "vip": "VIP", "other": "기타", "미지정": "미지정"
        }

        if self.orchestrator:
            contacts = self.orchestrator.contact_mgr.get_by_category(category)
            if search:
                contacts = [
                    c for c in contacts
                    if search in c.name.lower()
                    or search in c.company.lower()
                    or search in c.memo.lower()
                    or search in c.category.lower()
                ]

            for contact in reversed(contacts):
                is_checked = contact.id in self.selected_ids
                check_mark = "V" if is_checked else ""
                tag = "checked" if is_checked else "unchecked"
                cat_text = cat_label_map.get(contact.category, contact.category)
                iid = self.send_tree.insert("", "end", values=(
                    check_mark, contact.name, cat_text, contact.phone or ""
                ), tags=(tag,))
                self._send_tree_id_map[iid] = contact.id

        self._update_selected_count()

    def _on_tree_click_toggle(self, event):
        """Treeview 행 클릭 → 체크 토글"""
        iid = self.send_tree.identify_row(event.y)
        if not iid:
            return
        cid = self._send_tree_id_map.get(iid)
        if not cid:
            return
        # 토글
        if cid in self.selected_ids:
            self.selected_ids.discard(cid)
            self.send_tree.set(iid, "check", "")
            self.send_tree.item(iid, tags=("unchecked",))
        else:
            self.selected_ids.add(cid)
            self.send_tree.set(iid, "check", "V")
            self.send_tree.item(iid, tags=("checked",))
        self._update_selected_count()

    def _select_all(self):
        """현재 보이는 목록 전체 선택"""
        for iid in self.send_tree.get_children():
            cid = self._send_tree_id_map.get(iid)
            if cid:
                self.selected_ids.add(cid)
            self.send_tree.set(iid, "check", "V")
            self.send_tree.item(iid, tags=("checked",))
        self._update_selected_count()

    def _deselect_all(self):
        """전체 해제"""
        for iid in self.send_tree.get_children():
            self.send_tree.set(iid, "check", "")
            self.send_tree.item(iid, tags=("unchecked",))
        self.selected_ids.clear()
        self._update_selected_count()

    def _update_selected_count(self):
        total = len(self.send_tree.get_children())
        selected = len(self.selected_ids)
        self.selected_count_label.configure(text=f"{selected}명 선택 / {total}명")

    # ═══ 메시지 / 템플릿 ═══

    def _refresh_templates(self):
        """템플릿 드롭다운 갱신"""
        names = ["직접 입력"]
        self._template_map = {}
        if self.orchestrator:
            for tmpl in self.orchestrator.message_engine.get_templates():
                label = tmpl.name
                if len(tmpl.contents) > 1:
                    label += f" ({len(tmpl.contents)}변형)"
                names.append(label)
                self._template_map[label] = tmpl
        self.template_menu.configure(values=names)

    def _on_template_select(self, name):
        if name == "직접 입력":
            self._selected_template = None
            self._tmpl_preview_frame.configure(height=0)
            self.template_preview_label.configure(text="")
            return
        tmpl = self._template_map.get(name)
        if tmpl:
            # 미리보기 스니펫 표시
            snippet = tmpl.content.replace("\n", " ")[:50]
            if len(tmpl.content) > 50:
                snippet += "..."
            self.template_preview_label.configure(text=f"  📄 {snippet}")
            self._tmpl_preview_frame.configure(height=24)
            self._selected_template = tmpl
            self.msg_editor.delete("1.0", "end")
            self.msg_editor.insert("1.0", tmpl.content)
            # 변형 개수 표시
            if len(tmpl.contents) > 1:
                self.log_panel.add_log(
                    f"템플릿 '{name}': {len(tmpl.contents)}개 변형 → 랜덤 발송",
                    "info"
                )
            # 템플릿에 이미지가 첨부되어 있으면 자동 로드
            if tmpl.image_path:
                import os
                self.image_path = tmpl.image_path
                filename = os.path.basename(tmpl.image_path)
                self.image_label.configure(text=filename, text_color=T.ACCENT)
                self.image_clear_btn.pack(side="left")
            self._update_preview()

    def _update_preview(self):
        content = self.msg_editor.get("1.0", "end").strip()
        if not content:
            return

        preview = content
        if self.orchestrator:
            contacts = self.orchestrator.contact_mgr.get_all()
            if contacts:
                preview = self.orchestrator.message_engine.substitute(
                    content, contacts[0].to_dict()
                )
            else:
                sample = {"name": "홍길동", "company": "ABC회사", "position": "대리"}
                preview = self.orchestrator.message_engine.substitute(content, sample)

        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", preview)
        self.preview_box.configure(state="disabled")

    def get_current_message(self) -> str:
        return self.msg_editor.get("1.0", "end").strip()

    # ═══ 발송 ═══

    def _start_send(self):
        # 디버그 로그
        if hasattr(self, 'log_panel'):
            self.log_panel.add_log(
                f"발송 시작 → 카카오톡 봇, 선택: {len(self.selected_ids)}명",
                "info"
            )
        # 선택 확인
        if not self.selected_ids:
            messagebox.showwarning("발송 불가", "발송 대상을 선택해주세요.")
            return

        message = self.get_current_message()
        if not message:
            messagebox.showwarning("발송 불가", "메시지를 입력해주세요.")
            return

        if not self.orchestrator:
            self.log_panel.add_log("orchestrator 없음", "error")
            return

        # sender가 없으면 자동 초기화 시도 (별도 스레드)
        if not self.orchestrator.sender:
            self.log_panel.add_log("카카오톡 봇 자동 초기화 중...", "info")
            self.start_btn.configure(state="disabled")

            import threading
            def _init_kakao():
                success = False
                msg = ""
                try:
                    from pathlib import Path
                    positions_path = self.orchestrator.base_dir / "config" / "learned_positions.json"
                    if not positions_path.exists():
                        msg = "위치 학습 데이터가 없습니다.\n설정 → '위치 학습 시작'을 먼저 실행하세요."
                    else:
                        import json
                        with open(positions_path, "r", encoding="utf-8") as f:
                            self.orchestrator.coordinates = json.load(f)
                        if self.orchestrator.window_ctrl.find_kakao_window():
                            self.orchestrator.window_ctrl.activate_kakao()
                            import time
                            time.sleep(0.3)
                            self.orchestrator.window_ctrl.calculate_kakao_position()
                            self.orchestrator.window_ctrl.position_kakao_window()
                            self.orchestrator.confirm_calibration()
                            if self.orchestrator.sender:
                                success = True
                            else:
                                msg = "카카오톡 초기화에 실패했습니다."
                        else:
                            msg = "카카오톡이 실행되어 있지 않습니다.\n카카오톡 PC를 먼저 실행하세요."
                except Exception as e:
                    msg = f"초기화 오류: {e}"

                def _done():
                    self.start_btn.configure(state="normal")
                    if success:
                        self.log_panel.add_log("카카오톡 봇 초기화 완료!", "success")
                        # 초기화 성공 → 발송 재시도
                        self._start_send()
                    else:
                        self.log_panel.add_log(f"초기화 실패: {msg}", "error")
                        messagebox.showwarning("카카오톡 봇 사용 불가", msg)
                self.after(0, _done)

            threading.Thread(target=_init_kakao, daemon=True).start()
            return  # 스레드에서 완료 후 _start_send 재호출

        # sender 있으면 정상 진행 - 연락처 가져오기 (화면 표시 순서 = 최신순)
        selected_contacts = [
            c for c in reversed(self.orchestrator.contact_mgr.get_all())
            if c.id in self.selected_ids
        ]
        if not selected_contacts:
            messagebox.showwarning("발송 불가", "선택된 연락처가 없습니다.")
            return

        tmpl = getattr(self, '_selected_template', None)
        template_contents = None
        if tmpl and len(tmpl.contents) > 1:
            template_contents = tmpl.contents

        self.log_panel.add_log(f"큐 준비: {len(selected_contacts)}명, 메시지: {message[:30]}...", "info")
        try:
            self.orchestrator.prepare_custom_queue(
                selected_contacts, message,
                image_path=self.image_path,
                template_contents=template_contents
            )
        except Exception as e:
            self.log_panel.add_log(f"큐 준비 오류: {e}", "error")
            return

        variation_info = f" ({len(template_contents)}개 변형 랜덤)" if template_contents else ""
        self.log_panel.add_log(f"[카카오톡 봇] 발송 큐: {len(self.orchestrator.send_queue)}건{variation_info}", "info")

        self.orchestrator.on_progress(self._on_progress)
        self.orchestrator.on_result(self._on_result)
        self.orchestrator.on_log(self._on_log)
        self.orchestrator.on_state_change(self._on_state_change)
        self.start_btn.configure(state="disabled")

        self.log_panel.add_log(f"카카오봇 발송 시작 (sender={self.orchestrator.sender is not None})", "info")
        try:
            d_min = int(self.delay_min.get())
            d_max = int(self.delay_max.get())
            self.orchestrator.sender.delay_min = d_min
            self.orchestrator.sender.delay_max = d_max
        except ValueError:
            pass
        self.orchestrator.start_sending()

    def _pause_send(self):
        if not self.orchestrator:
            return
        if self.orchestrator.state == "sending":
            self.orchestrator.pause_sending()
            self.pause_btn.configure(
                text="▶  재개",
                fg_color="#2ea043", hover_color="#3fb950"
            )
        elif self.orchestrator.state == "paused":
            self.orchestrator.resume_sending()
            self.pause_btn.configure(
                text="⏸  일시정지",
                fg_color="#b08800", hover_color="#d29922"
            )

    def _stop_send(self):
        if not self.orchestrator:
            return
        if self.orchestrator.state not in ("sending", "paused"):
            return
        if messagebox.askyesno("발송 중지", "정말 발송을 중지하시겠습니까?"):
            self.orchestrator.stop_sending()
            self.pause_btn.configure(
                text="⏸  일시정지",
                fg_color="#b08800", hover_color="#d29922"
            )
            self.log_panel.add_log("중지됨. 다시 시작할 수 있습니다.", "info")

    def _schedule_send(self):
        """예약 발송 다이얼로그"""
        if not self.orchestrator:
            return

        if not self.selected_ids:
            messagebox.showwarning("예약 불가", "발송 대상을 선택해주세요.")
            return

        message = self.get_current_message()
        if not message:
            messagebox.showwarning("예약 불가", "메시지를 입력해주세요.")
            return

        dialog = ScheduleDialog(self)
        self.wait_window(dialog)

        if dialog.result:
            job = self.orchestrator.scheduler.add_job(
                scheduled_time=dialog.result,
                contact_ids=list(self.selected_ids),
                template_content=message,
                image_path=self.image_path,
            )
            self.log_panel.add_log(
                f"예약 등록 완료: {job.display_time} ({len(self.selected_ids)}명)",
                "success"
            )

    # -- 콜백 (스레드 → 메인 스레드) --

    def _on_progress(self, current, total, name):
        self.after(0, lambda: self.progress_card.update_progress(current, total, name))

    def _on_result(self, result):
        def _update():
            emoji = "✅" if result["status"] == "success" else "❌"
            self.log_panel.add_log(
                f"{emoji} {result['contact_name']} - {result.get('detail', result['status'])}",
                "success" if result["status"] == "success" else "error"
            )
        self.after(0, _update)

    def _on_log(self, message, level):
        self.after(0, lambda: self.log_panel.add_log(message, level))

    def _on_state_change(self, state):
        def _update():
            if state in ("completed", "error", "idle"):
                self.start_btn.configure(state="normal")
                self.pause_btn.configure(
                    text="⏸  일시정지",
                    fg_color="#b08800", hover_color="#d29922"
                )
                if state == "completed":
                    self.log_panel.add_log("발송 완료!", "success")
                elif state == "error":
                    self.log_panel.add_log("발송 오류로 중단. 다시 시작할 수 있습니다.", "error")
        self.after(0, _update)


class ScheduleDialog(ctk.CTkToplevel):
    """예약 발송 시간 설정 다이얼로그"""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.title("예약 발송")
        self.geometry("350x280")
        self.configure(fg_color=T.BG_DARK)
        self.result = None
        self.transient(parent)
        self.grab_set()
        self._build()

    def _build(self):
        from datetime import datetime, timedelta

        ctk.CTkLabel(
            self, text="⏰ 예약 발송 시간 설정",
            font=(T.get_font_family(), T.FONT_SIZE_HEADER, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(padx=24, pady=(24, 16))

        # 날짜
        date_frame = ctk.CTkFrame(self, fg_color="transparent")
        date_frame.pack(fill="x", padx=24, pady=(0, 8))

        ctk.CTkLabel(
            date_frame, text="날짜:",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            text_color=T.TEXT_SECONDARY
        ).pack(side="left", padx=(0, 8))

        now = datetime.now()
        years = [str(y) for y in range(now.year, now.year + 2)]
        months = [f"{m:02d}" for m in range(1, 13)]
        days = [f"{d:02d}" for d in range(1, 32)]

        self.year_var = ctk.StringVar(value=str(now.year))
        ctk.CTkOptionMenu(
            date_frame, values=years, variable=self.year_var,
            width=80, height=30, font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER, text_color=T.TEXT_PRIMARY
        ).pack(side="left", padx=2)

        ctk.CTkLabel(date_frame, text="-", text_color=T.TEXT_MUTED).pack(side="left")

        self.month_var = ctk.StringVar(value=f"{now.month:02d}")
        ctk.CTkOptionMenu(
            date_frame, values=months, variable=self.month_var,
            width=60, height=30, font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER, text_color=T.TEXT_PRIMARY
        ).pack(side="left", padx=2)

        ctk.CTkLabel(date_frame, text="-", text_color=T.TEXT_MUTED).pack(side="left")

        self.day_var = ctk.StringVar(value=f"{now.day:02d}")
        ctk.CTkOptionMenu(
            date_frame, values=days, variable=self.day_var,
            width=60, height=30, font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER, text_color=T.TEXT_PRIMARY
        ).pack(side="left", padx=2)

        # 시간
        time_frame = ctk.CTkFrame(self, fg_color="transparent")
        time_frame.pack(fill="x", padx=24, pady=(0, 8))

        ctk.CTkLabel(
            time_frame, text="시간:",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            text_color=T.TEXT_SECONDARY
        ).pack(side="left", padx=(0, 8))

        hours = [f"{h:02d}" for h in range(24)]
        minutes = [f"{m:02d}" for m in range(0, 60, 5)]

        self.hour_var = ctk.StringVar(value=f"{now.hour:02d}")
        ctk.CTkOptionMenu(
            time_frame, values=hours, variable=self.hour_var,
            width=60, height=30, font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER, text_color=T.TEXT_PRIMARY
        ).pack(side="left", padx=2)

        ctk.CTkLabel(time_frame, text=":", text_color=T.TEXT_MUTED).pack(side="left")

        self.min_var = ctk.StringVar(value=f"{(now.minute // 5 * 5):02d}")
        ctk.CTkOptionMenu(
            time_frame, values=minutes, variable=self.min_var,
            width=60, height=30, font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER, text_color=T.TEXT_PRIMARY
        ).pack(side="left", padx=2)

        # 버튼
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=16)

        ctk.CTkButton(
            btn_frame, text="취소", width=100, height=36,
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, corner_radius=6,
            command=self.destroy
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame, text="⏰ 예약 등록", width=140, height=36,
            fg_color=T.ACCENT, hover_color=T.ACCENT_HOVER,
            text_color=T.TEXT_ON_ACCENT, corner_radius=6,
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            command=self._confirm
        ).pack(side="right")

    def _confirm(self):
        from datetime import datetime
        try:
            dt = datetime(
                int(self.year_var.get()),
                int(self.month_var.get()),
                int(self.day_var.get()),
                int(self.hour_var.get()),
                int(self.min_var.get())
            )
            if dt <= datetime.now():
                messagebox.showwarning("시간 오류", "현재 시간 이후로 설정해주세요.")
                return
            self.result = dt.isoformat()
            self.destroy()
        except ValueError as e:
            messagebox.showwarning("날짜 오류", f"올바른 날짜를 입력해주세요.\n{e}")
