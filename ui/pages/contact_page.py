"""
Contact Page - 연락처 관리 페이지
커스텀 카테고리, 엑셀 샘플 다운로드, 가져오기/내보내기
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from ui.theme import AppTheme as T


class ContactPage(ctk.CTkFrame):
    """연락처 관리 페이지"""

    def __init__(self, parent, orchestrator=None, api_client=None, **kwargs):
        super().__init__(parent, fg_color=T.BG_DARK, **kwargs)
        self.orchestrator = orchestrator
        self.api_client = api_client  # SaaS 모드
        self._contacts_cache = []  # API 연락처 캐시
        self._build()

    def _build(self):
        # -- 헤더 --
        header = ctk.CTkFrame(self, fg_color="transparent", height=50)
        header.pack(fill="x", padx=24, pady=(20, 16))
        header.pack_propagate(False)

        ctk.CTkLabel(
            header, text="👥 연락처 관리",
            font=(T.get_font_family(), T.FONT_SIZE_TITLE, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(side="left")

        # 버튼들
        ctk.CTkButton(
            header, text="📥 내보내기", width=90,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=30, corner_radius=6,
            command=self._export_excel
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            header, text="📤 가져오기", width=90,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=30, corner_radius=6,
            command=self._import_excel
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            header, text="📋 샘플 다운", width=90,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color="#1a5276", hover_color="#2471a3",
            text_color=T.TEXT_PRIMARY, height=30, corner_radius=6,
            command=self._download_sample
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            header, text="+ 추가", width=80,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL, "bold"),
            fg_color=T.ACCENT, hover_color=T.ACCENT_HOVER,
            text_color=T.TEXT_ON_ACCENT, height=30, corner_radius=6,
            command=self._add_contact
        ).pack(side="right")

        # -- 카테고리 필터 --
        filter_frame = ctk.CTkFrame(self, fg_color="transparent")
        filter_frame.pack(fill="x", padx=24, pady=(0, 8))

        self.category_var = ctk.StringVar(value="all")
        self.cat_btn_frame = ctk.CTkFrame(filter_frame, fg_color="transparent")
        self.cat_btn_frame.pack(side="left", fill="x", expand=True)

        # 카테고리 추가 버튼
        ctk.CTkButton(
            filter_frame, text="+ 카테고리", width=90, height=28,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.ACCENT, corner_radius=14,
            command=self._add_category
        ).pack(side="right")

        self._refresh_category_buttons()

        # -- 검색 --
        search_frame = ctk.CTkFrame(self, fg_color="transparent", height=40)
        search_frame.pack(fill="x", padx=24, pady=(0, 12))
        search_frame.pack_propagate(False)

        self.search_entry = ctk.CTkEntry(
            search_frame, placeholder_text="🔍 이름, 회사, 메모로 검색...",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_INPUT, border_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=36, corner_radius=6
        )
        self.search_entry.pack(fill="x")
        self.search_entry.bind("<KeyRelease>", lambda e: self._on_search())

        # -- 연락처 목록 --
        self.list_frame = ctk.CTkScrollableFrame(
            self, fg_color=T.BG_DARK,
            scrollbar_button_color=T.BG_HOVER,
            scrollbar_button_hover_color=T.BORDER
        )
        self.list_frame.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        # -- 하단 카운트 --
        self.count_label = ctk.CTkLabel(
            self, text="총 0명",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        )
        self.count_label.pack(padx=24, pady=(0, 12))

        self.refresh_list()

    def _refresh_category_buttons(self):
        """카테고리 버튼 동적 생성"""
        for w in self.cat_btn_frame.winfo_children():
            w.destroy()

        categories = [("all", "전체")]
        if self.orchestrator:
            for cat in self.orchestrator.contact_mgr.get_all_categories():
                cat_label = {
                    "friend": "친구", "family": "가족", "business": "사업체",
                    "vip": "VIP", "other": "기타"
                }.get(cat, cat)
                categories.append((cat, cat_label))
        else:
            for cat, label in [("friend", "친구"), ("family", "가족"),
                               ("business", "사업체"), ("vip", "VIP"), ("other", "기타")]:
                categories.append((cat, label))

        current = self.category_var.get()
        for cat_id, cat_name in categories:
            is_active = cat_id == current
            btn = ctk.CTkButton(
                self.cat_btn_frame, text=cat_name, width=70, height=28,
                font=(T.get_font_family(), T.FONT_SIZE_SMALL),
                fg_color=T.ACCENT if is_active else T.BG_HOVER,
                hover_color=T.ACCENT_HOVER if is_active else T.BORDER,
                text_color=T.TEXT_ON_ACCENT if is_active else T.TEXT_SECONDARY,
                corner_radius=14,
                command=lambda cid=cat_id: self._filter_category(cid)
            )
            btn.pack(side="left", padx=(0, 4))

    def refresh_list(self, category: str = "all", search: str = ""):
        """연락처 목록 새로고침"""
        for widget in self.list_frame.winfo_children():
            widget.destroy()

        if self.api_client and self.api_client.is_logged_in:
            # SaaS 모드: API에서 가져오기
            try:
                cat = category if category != "all" else None
                self._contacts_cache = self.api_client.get_contacts(
                    category=cat, search=search if search else None
                )
            except Exception:
                self._contacts_cache = []

            for c_data in self._contacts_cache:
                self._create_contact_row_api(c_data)
            self.count_label.configure(text=f"총 {len(self._contacts_cache)}명")
        elif self.orchestrator:
            # 로컬 모드
            contacts = self.orchestrator.contact_mgr.get_by_category(category)
            if search:
                search = search.lower()
                contacts = [
                    c for c in contacts
                    if search in c.name.lower()
                    or search in c.company.lower()
                    or search in c.memo.lower()
                ]
            for contact in contacts:
                self._create_contact_row(contact)
            self.count_label.configure(text=f"총 {len(contacts)}명")

    def _create_contact_row(self, contact):
        """연락처 행 생성"""
        row = ctk.CTkFrame(
            self.list_frame, fg_color=T.BG_CARD,
            corner_radius=6, height=56,
            border_width=1, border_color=T.BORDER
        )
        row.pack(fill="x", pady=3)
        row.pack_propagate(False)

        cat_color = T.CATEGORY_COLORS.get(contact.category, T.TEXT_MUTED)
        ctk.CTkLabel(
            row, text=f" {contact.category} ",
            font=(T.get_font_family(), 9, "bold"),
            fg_color=cat_color, text_color=T.BG_DARK,
            corner_radius=4, width=60
        ).pack(side="left", padx=(12, 8), pady=8)

        ctk.CTkLabel(
            row, text=contact.name,
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            text_color=T.TEXT_PRIMARY, width=120
        ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            row, text=contact.company or "-",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_SECONDARY, width=120
        ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            row, text=contact.phone or "-",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED, width=120
        ).pack(side="left", padx=(0, 12))

        # 생일/기념일 표시
        date_info = ""
        if contact.birthday:
            date_info += f"🎂{contact.birthday}"
        if contact.anniversary:
            date_info += f" 🎉{contact.anniversary}" if date_info else f"🎉{contact.anniversary}"
        ctk.CTkLabel(
            row, text=date_info or "-",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED, width=100
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            row, text=contact.memo[:20] + "..." if len(contact.memo) > 20 else contact.memo or "-",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        ).pack(side="left", padx=(0, 12), expand=True)

        ctk.CTkButton(
            row, text="🗑", width=30, height=26,
            font=(T.get_font_family(), 12),
            fg_color="transparent", hover_color=T.ERROR,
            text_color=T.TEXT_MUTED, corner_radius=4,
            command=lambda cid=contact.id: self._delete_contact(cid)
        ).pack(side="right", padx=(0, 8))

        ctk.CTkButton(
            row, text="✏️", width=30, height=26,
            font=(T.get_font_family(), 12),
            fg_color="transparent", hover_color=T.BG_HOVER,
            text_color=T.TEXT_MUTED, corner_radius=4,
            command=lambda c=contact: self._edit_contact(c)
        ).pack(side="right", padx=(0, 4))

        # 인라인 카테고리 변경 (한글 표시)
        cat_label_map = {"friend": "친구", "family": "가족", "business": "사업체", "vip": "VIP", "other": "기타"}
        if self.orchestrator:
            raw_cats = self.orchestrator.contact_mgr.get_all_categories()
            for c in raw_cats:
                if c not in cat_label_map:
                    cat_label_map[c] = c
        cat_value_map = {v: k for k, v in cat_label_map.items()}
        cat_labels = list(cat_label_map.values())
        current_label = cat_label_map.get(contact.category, contact.category)
        cat_var = ctk.StringVar(value=current_label)
        cat_menu = ctk.CTkOptionMenu(
            row, values=cat_labels, variable=cat_var,
            width=75, height=24,
            font=(T.get_font_family(), 9),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER,
            text_color=T.TEXT_PRIMARY, corner_radius=4,
            command=lambda val, _id=contact.id, _m=cat_value_map: self._quick_change_category(_id, _m.get(val, val))
        )
        cat_menu.pack(side="right", padx=(0, 4))

    def _create_contact_row_api(self, c_data: dict):
        """API 연락처 행 생성 (SaaS 모드)"""
        row = ctk.CTkFrame(
            self.list_frame, fg_color=T.BG_CARD,
            corner_radius=6, height=56,
            border_width=1, border_color=T.BORDER
        )
        row.pack(fill="x", pady=3)
        row.pack_propagate(False)

        cat = c_data.get("category", "other")
        cat_color = T.CATEGORY_COLORS.get(cat, T.TEXT_MUTED)
        ctk.CTkLabel(
            row, text=f" {cat} ",
            font=(T.get_font_family(), 9, "bold"),
            fg_color=cat_color, text_color=T.BG_DARK,
            corner_radius=4, width=60
        ).pack(side="left", padx=(12, 8), pady=8)

        ctk.CTkLabel(
            row, text=c_data.get("name", ""),
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            text_color=T.TEXT_PRIMARY, width=120
        ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            row, text=c_data.get("company", "") or "-",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_SECONDARY, width=120
        ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            row, text=c_data.get("phone", "") or "-",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED, width=120
        ).pack(side="left", padx=(0, 12))

        memo = c_data.get("memo", "") or ""
        ctk.CTkLabel(
            row, text=memo[:20] + "..." if len(memo) > 20 else memo or "-",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        ).pack(side="left", padx=(0, 12), expand=True)

        cid = c_data.get("id")
        ctk.CTkButton(
            row, text="🗑", width=30, height=26,
            font=(T.get_font_family(), 12),
            fg_color="transparent", hover_color=T.ERROR,
            text_color=T.TEXT_MUTED, corner_radius=4,
            command=lambda: self._delete_contact_api(cid)
        ).pack(side="right", padx=(0, 8))

        ctk.CTkButton(
            row, text="✏️", width=30, height=26,
            font=(T.get_font_family(), 12),
            fg_color="transparent", hover_color=T.BG_HOVER,
            text_color=T.TEXT_MUTED, corner_radius=4,
            command=lambda d=c_data: self._edit_contact_api(d)
        ).pack(side="right", padx=(0, 4))

        # 인라인 카테고리 변경 (한글 표시)
        cat_labels = ["고객", "친구", "가족", "사업체", "VIP", "기타"]
        cat_values = ["customer", "friend", "family", "business", "vip", "other"]
        cat_label_map = dict(zip(cat_values, cat_labels))
        cat_value_map = dict(zip(cat_labels, cat_values))
        current_label = cat_label_map.get(cat, cat)
        cat_var = ctk.StringVar(value=current_label)
        cat_menu = ctk.CTkOptionMenu(
            row, values=cat_labels, variable=cat_var,
            width=75, height=24,
            font=(T.get_font_family(), 9),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER,
            text_color=T.TEXT_PRIMARY, corner_radius=4,
            command=lambda val, _id=cid, _m=cat_value_map: self._quick_change_category_api(_id, _m.get(val, val))
        )
        cat_menu.pack(side="right", padx=(0, 4))

    def _quick_change_category_api(self, contact_id, new_category):
        """SaaS: 카테고리 즉시 변경"""
        try:
            self.api_client.update_contact(contact_id, {"category": new_category})
        except Exception:
            pass

    def _quick_change_category(self, contact_id, new_category):
        """로컬: 카테고리 즉시 변경"""
        if self.orchestrator:
            self.orchestrator.contact_mgr.update(contact_id, category=new_category)

    def _delete_contact_api(self, contact_id):
        if messagebox.askyesno("삭제 확인", "정말 삭제하시겠습니까?"):
            try:
                self.api_client.delete_contact(contact_id)
                self.refresh_list(category=self.category_var.get())
            except Exception as e:
                messagebox.showerror("오류", str(e))

    def _edit_contact_api(self, c_data):
        dialog = ContactDialogAPI(self, title="연락처 편집", contact_data=c_data)
        self.wait_window(dialog)
        if dialog.result:
            try:
                self.api_client.update_contact(c_data["id"], dialog.result)
                self.refresh_list(category=self.category_var.get())
            except Exception as e:
                messagebox.showerror("오류", str(e))

    def _filter_category(self, category: str):
        self.category_var.set(category)
        self._refresh_category_buttons()
        self.refresh_list(category=category, search=self.search_entry.get())

    def _on_search(self):
        self.refresh_list(
            category=self.category_var.get(),
            search=self.search_entry.get()
        )

    def _add_contact(self):
        if self.api_client and self.api_client.is_logged_in:
            dialog = ContactDialogAPI(self, title="연락처 추가")
            self.wait_window(dialog)
            if dialog.result:
                try:
                    self.api_client.create_contact(dialog.result)
                    self._refresh_category_buttons()
                    self.refresh_list(category=self.category_var.get())
                except Exception as e:
                    messagebox.showerror("오류", str(e))
            return

        dialog = ContactDialog(self, title="연락처 추가", orchestrator=self.orchestrator)
        self.wait_window(dialog)
        if dialog.result and self.orchestrator:
            from core.contact_manager import Contact
            contact = Contact(**dialog.result)
            if self.orchestrator.contact_mgr.add(contact):
                self._refresh_category_buttons()
                self.refresh_list(category=self.category_var.get())
            else:
                messagebox.showwarning("중복", "동일한 이름과 카테고리의 연락처가 이미 있습니다.")

    def _edit_contact(self, contact):
        dialog = ContactDialog(self, title="연락처 편집", contact=contact,
                               orchestrator=self.orchestrator)
        self.wait_window(dialog)
        if dialog.result and self.orchestrator:
            self.orchestrator.contact_mgr.update(contact.id, **dialog.result)
            self._refresh_category_buttons()
            self.refresh_list(category=self.category_var.get())

    def _delete_contact(self, contact_id):
        if messagebox.askyesno("삭제 확인", "정말 삭제하시겠습니까?"):
            if self.orchestrator:
                self.orchestrator.contact_mgr.delete(contact_id)
                self.refresh_list(category=self.category_var.get())

    def _add_category(self):
        dialog = CategoryDialog(self, orchestrator=self.orchestrator)
        self.wait_window(dialog)
        if dialog.result and self.orchestrator:
            if self.orchestrator.contact_mgr.add_category(dialog.result):
                self._refresh_category_buttons()
            else:
                messagebox.showwarning("중복", "이미 존재하는 카테고리입니다.")

    def _download_sample(self):
        filepath = filedialog.asksaveasfilename(
            title="샘플 엑셀 저장",
            defaultextension=".xlsx",
            initialfile="연락처_샘플.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filepath and self.orchestrator:
            try:
                self.orchestrator.contact_mgr.create_sample_excel(filepath)
                messagebox.showinfo("다운로드 완료",
                                    f"샘플 파일 저장 완료!\n{filepath}\n\n"
                                    "이 파일을 수정한 후 '가져오기'로 업로드하세요.")
            except Exception as e:
                messagebox.showerror("오류", str(e))

    def _import_excel(self):
        filepath = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not filepath:
            return

        if self.api_client and self.api_client.is_logged_in:
            try:
                result = self.api_client.import_contacts(filepath)
                messagebox.showinfo("가져오기 완료", f"추가: {result.get('added', 0)}명")
                self.refresh_list(category=self.category_var.get())
            except Exception as e:
                messagebox.showerror("오류", str(e))
        elif self.orchestrator:
            result = self.orchestrator.contact_mgr.import_from_excel(filepath)
            messagebox.showinfo(
                "가져오기 완료",
                f"추가: {result['success']}명\n건너뜀: {result['skipped']}명"
            )
            self._refresh_category_buttons()
            self.refresh_list(category=self.category_var.get())

    def _export_excel(self):
        filepath = filedialog.asksaveasfilename(
            title="엑셀 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not filepath:
            return

        if self.api_client and self.api_client.is_logged_in:
            try:
                self.api_client.export_contacts(filepath)
                messagebox.showinfo("내보내기 완료", f"저장 완료: {filepath}")
            except Exception as e:
                messagebox.showerror("오류", str(e))
        elif self.orchestrator:
            self.orchestrator.contact_mgr.export_to_excel(filepath)
            messagebox.showinfo("내보내기 완료", f"저장 완료: {filepath}")


class CategoryDialog(ctk.CTkToplevel):
    """카테고리 추가 다이얼로그"""

    def __init__(self, parent, orchestrator=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.title("카테고리 추가")
        self.geometry("350x200")
        self.configure(fg_color=T.BG_DARK)
        self.result = None
        self.orchestrator = orchestrator
        self.transient(parent)
        self.grab_set()
        self._build()

    def _build(self):
        ctk.CTkLabel(
            self, text="새 카테고리 이름",
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            text_color=T.TEXT_PRIMARY
        ).pack(padx=24, pady=(24, 8), anchor="w")

        self.entry = ctk.CTkEntry(
            self, placeholder_text="예: 동호회, 학교, 교회...",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_INPUT, border_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=40, corner_radius=6
        )
        self.entry.pack(fill="x", padx=24)
        self.entry.focus()

        if self.orchestrator:
            existing = self.orchestrator.contact_mgr.get_all_categories()
            ctk.CTkLabel(
                self, text=f"기존: {', '.join(existing)}",
                font=(T.get_font_family(), 9),
                text_color=T.TEXT_MUTED, wraplength=300
            ).pack(padx=24, pady=(4, 0), anchor="w")

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=16)

        ctk.CTkButton(
            btn_frame, text="취소", width=100, height=36,
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, corner_radius=6,
            command=self.destroy
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame, text="추가", width=100, height=36,
            fg_color=T.ACCENT, hover_color=T.ACCENT_HOVER,
            text_color=T.TEXT_ON_ACCENT, corner_radius=6,
            command=self._save
        ).pack(side="right")

    def _save(self):
        name = self.entry.get().strip()
        if not name:
            messagebox.showwarning("필수 입력", "카테고리 이름을 입력해주세요.")
            return
        self.result = name
        self.destroy()


class ContactDialog(ctk.CTkToplevel):
    """연락처 추가/편집 다이얼로그"""

    def __init__(self, parent, title="연락처", contact=None, orchestrator=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.title(title)
        self.geometry("420x640")
        self.configure(fg_color=T.BG_DARK)
        self.result = None
        self.contact = contact
        self.orchestrator = orchestrator

        self.transient(parent)
        self.grab_set()
        self._build()

    def _build(self):
        scroll = ctk.CTkScrollableFrame(
            self, fg_color="transparent",
            scrollbar_button_color=T.BG_HOVER
        )
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        fields = [
            ("이름 *", "name"),
            ("전화번호", "phone"),
            ("회사", "company"),
            ("직급", "position"),
            ("메모", "memo"),
        ]

        self.entries = {}

        for label_text, field_name in fields:
            ctk.CTkLabel(
                scroll, text=label_text,
                font=(T.get_font_family(), T.FONT_SIZE_SMALL),
                text_color=T.TEXT_SECONDARY
            ).pack(padx=24, pady=(12, 4), anchor="w")

            entry = ctk.CTkEntry(
                scroll, font=(T.get_font_family(), T.FONT_SIZE_BODY),
                fg_color=T.BG_INPUT, border_color=T.BORDER,
                text_color=T.TEXT_PRIMARY, height=36, corner_radius=6
            )
            entry.pack(fill="x", padx=24)
            self.entries[field_name] = entry

            if self.contact and hasattr(self.contact, field_name):
                entry.insert(0, getattr(self.contact, field_name) or "")

        # 생일 / 기념일
        date_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        date_frame.pack(fill="x", padx=24, pady=(12, 0))

        # 생일
        bd_frame = ctk.CTkFrame(date_frame, fg_color="transparent")
        bd_frame.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkLabel(
            bd_frame, text="생일 (MM-DD)",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_SECONDARY
        ).pack(anchor="w")
        bd_entry = ctk.CTkEntry(
            bd_frame, placeholder_text="예: 03-15",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_INPUT, border_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=36, corner_radius=6
        )
        bd_entry.pack(fill="x")
        self.entries["birthday"] = bd_entry
        if self.contact and self.contact.birthday:
            bd_entry.insert(0, self.contact.birthday)

        # 기념일
        an_frame = ctk.CTkFrame(date_frame, fg_color="transparent")
        an_frame.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            an_frame, text="기념일 (MM-DD)",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_SECONDARY
        ).pack(anchor="w")
        an_entry = ctk.CTkEntry(
            an_frame, placeholder_text="예: 05-10",
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_INPUT, border_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=36, corner_radius=6
        )
        an_entry.pack(fill="x")
        self.entries["anniversary"] = an_entry
        if self.contact and self.contact.anniversary:
            an_entry.insert(0, self.contact.anniversary)

        # 카테고리 선택 (동적 목록)
        ctk.CTkLabel(
            scroll, text="카테고리",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_SECONDARY
        ).pack(padx=24, pady=(12, 4), anchor="w")

        categories = ["friend", "family", "business", "vip", "other"]
        if self.orchestrator:
            categories = self.orchestrator.contact_mgr.get_all_categories()

        self.category_var = ctk.StringVar(
            value=self.contact.category if self.contact else "other"
        )
        self.category_menu = ctk.CTkOptionMenu(
            scroll, values=categories,
            variable=self.category_var,
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_INPUT, button_color=T.BG_HOVER,
            text_color=T.TEXT_PRIMARY, height=36, corner_radius=6
        )
        self.category_menu.pack(fill="x", padx=24)

        # 버튼
        btn_frame = ctk.CTkFrame(self, fg_color=T.BG_CARD, height=70)
        btn_frame.pack(fill="x", side="bottom")
        btn_frame.pack_propagate(False)

        ctk.CTkButton(
            btn_frame, text="취소", width=120,
            font=(T.get_font_family(), T.FONT_SIZE_BODY),
            fg_color=T.BG_HOVER, hover_color=T.BORDER,
            text_color=T.TEXT_PRIMARY, height=40, corner_radius=6,
            command=self.destroy
        ).pack(side="left", padx=(24, 8), pady=15)

        ctk.CTkButton(
            btn_frame, text="저장하기", width=120,
            font=(T.get_font_family(), T.FONT_SIZE_BODY, "bold"),
            fg_color=T.ACCENT, hover_color=T.ACCENT_HOVER,
            text_color=T.TEXT_ON_ACCENT, height=40, corner_radius=6,
            command=self._save
        ).pack(side="right", padx=(8, 24), pady=15)

    def _save(self):
        name = self.entries["name"].get().strip()
        if not name:
            messagebox.showwarning("필수 입력", "이름을 입력해주세요.")
            return

        self.result = {
            "name": name,
            "phone": self.entries["phone"].get().strip(),
            "company": self.entries["company"].get().strip(),
            "position": self.entries["position"].get().strip(),
            "memo": self.entries["memo"].get().strip(),
            "birthday": self.entries["birthday"].get().strip(),
            "anniversary": self.entries["anniversary"].get().strip(),
            "category": self.category_var.get()
        }
        self.destroy()


class ContactDialogAPI(ctk.CTkToplevel):
    """API용 연락처 추가/편집 다이얼로그 (SaaS 모드)"""

    def __init__(self, parent, title="연락처", contact_data=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.title(title)
        self.geometry("420x520")
        self.configure(fg_color=T.BG_DARK)
        self.result = None
        self.contact_data = contact_data or {}
        self.transient(parent)
        self.grab_set()
        self._build()

    def _build(self):
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True)

        fields = [
            ("이름 *", "name"), ("전화번호", "phone"), ("회사", "company"),
            ("직급", "position"), ("메모", "memo"), ("생일 (MM-DD)", "birthday"),
            ("기념일 (MM-DD)", "anniversary"),
        ]
        self.entries = {}

        for label_text, field_name in fields:
            ctk.CTkLabel(scroll, text=label_text,
                         font=(T.get_font_family(), T.FONT_SIZE_SMALL),
                         text_color=T.TEXT_SECONDARY).pack(padx=24, pady=(8, 2), anchor="w")
            entry = ctk.CTkEntry(scroll, font=(T.get_font_family(), T.FONT_SIZE_BODY),
                                 fg_color=T.BG_INPUT, border_color=T.BORDER,
                                 text_color=T.TEXT_PRIMARY, height=36, corner_radius=6)
            entry.pack(fill="x", padx=24)
            self.entries[field_name] = entry
            val = self.contact_data.get(field_name, "")
            if val:
                entry.insert(0, val)

        # 카테고리
        ctk.CTkLabel(scroll, text="카테고리",
                     font=(T.get_font_family(), T.FONT_SIZE_SMALL),
                     text_color=T.TEXT_SECONDARY).pack(padx=24, pady=(8, 2), anchor="w")
        categories = ["customer", "friend", "family", "business", "vip", "other"]
        self.category_var = ctk.StringVar(value=self.contact_data.get("category", "other"))
        ctk.CTkOptionMenu(scroll, values=categories, variable=self.category_var,
                          font=(T.get_font_family(), T.FONT_SIZE_BODY),
                          fg_color=T.BG_INPUT, text_color=T.TEXT_PRIMARY,
                          height=36, corner_radius=6).pack(fill="x", padx=24)

        # 버튼
        btn_frame = ctk.CTkFrame(self, fg_color=T.BG_CARD, height=60)
        btn_frame.pack(fill="x", side="bottom")
        btn_frame.pack_propagate(False)
        ctk.CTkButton(btn_frame, text="취소", width=100, fg_color=T.BG_HOVER,
                      hover_color=T.BORDER, text_color=T.TEXT_PRIMARY, height=36,
                      command=self.destroy).pack(side="left", padx=24, pady=12)
        ctk.CTkButton(btn_frame, text="저장", width=100, fg_color=T.ACCENT,
                      hover_color=T.ACCENT_HOVER, text_color=T.TEXT_ON_ACCENT, height=36,
                      command=self._save).pack(side="right", padx=24, pady=12)

    def _save(self):
        name = self.entries["name"].get().strip()
        if not name:
            messagebox.showwarning("필수 입력", "이름을 입력해주세요.")
            return
        self.result = {
            "name": name,
            "phone": self.entries["phone"].get().strip(),
            "company": self.entries["company"].get().strip(),
            "position": self.entries["position"].get().strip(),
            "memo": self.entries["memo"].get().strip(),
            "birthday": self.entries["birthday"].get().strip(),
            "anniversary": self.entries["anniversary"].get().strip(),
            "category": self.category_var.get()
        }
        self.destroy()
