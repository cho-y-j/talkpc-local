"""
Contact Page - 연락처 관리 페이지
Treeview 기반 고속 테이블 + 커스텀 카테고리
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
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

        # -- Treeview 스타일 (다크 테마) --
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Contact.Treeview",
                         background="#1c2333", foreground="#e6edf3",
                         fieldbackground="#1c2333", borderwidth=0,
                         font=(T.get_font_family(), 11),
                         rowheight=32)
        style.configure("Contact.Treeview.Heading",
                         background="#2d333b", foreground="#e6edf3",
                         font=(T.get_font_family(), 10, "bold"),
                         borderwidth=0)
        style.map("Contact.Treeview",
                   background=[("selected", "#2f81f7")],
                   foreground=[("selected", "#ffffff")])

        # -- Treeview 테이블 --
        tree_frame = ctk.CTkFrame(self, fg_color="#1c2333", corner_radius=8)
        tree_frame.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        columns = ("name", "category", "phone", "company", "memo")
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            selectmode="extended", style="Contact.Treeview"
        )
        self.tree.heading("name", text="이름", anchor="w")
        self.tree.heading("category", text="카테고리", anchor="w")
        self.tree.heading("phone", text="전화번호", anchor="w")
        self.tree.heading("company", text="회사", anchor="w")
        self.tree.heading("memo", text="메모", anchor="w")

        self.tree.column("name", width=100, minwidth=80)
        self.tree.column("category", width=80, minwidth=60)
        self.tree.column("phone", width=130, minwidth=100)
        self.tree.column("company", width=120, minwidth=80)
        self.tree.column("memo", width=200, minwidth=100)

        # 스크롤바
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 더블클릭 → 편집
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        # 우클릭 → 메뉴
        self.tree.bind("<Button-3>", self._on_tree_right_click)

        # contact id → tree item 매핑
        self._tree_id_map = {}

        # -- 하단: 카운트 + 선택 삭제 --
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=24, pady=(0, 12))

        self.count_label = ctk.CTkLabel(
            bottom, text="총 0명",
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            text_color=T.TEXT_MUTED
        )
        self.count_label.pack(side="left")

        ctk.CTkButton(
            bottom, text="선택 삭제", width=80, height=28,
            font=(T.get_font_family(), T.FONT_SIZE_SMALL),
            fg_color=T.BG_HOVER, hover_color="#f85149",
            text_color=T.TEXT_PRIMARY, corner_radius=6,
            command=self._delete_selected
        ).pack(side="right")

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
        """연락처 목록 새로고침 (Treeview - 즉시 로드)"""
        # 기존 항목 전부 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._tree_id_map.clear()

        cat_label_map = {
            "friend": "친구", "family": "가족", "business": "사업체",
            "vip": "VIP", "other": "기타", "미지정": "미지정"
        }

        if self.api_client and self.api_client.is_logged_in:
            try:
                cat = category if category != "all" else None
                self._contacts_cache = self.api_client.get_contacts(
                    category=cat, search=search if search else None
                )
            except Exception:
                self._contacts_cache = []
            for c in self._contacts_cache:
                cat_text = cat_label_map.get(c.get("category", ""), c.get("category", ""))
                iid = self.tree.insert("", 0, values=(
                    c.get("name", ""), cat_text,
                    c.get("phone", ""), c.get("company", ""),
                    c.get("memo", "")[:30]
                ))
                self._tree_id_map[iid] = c.get("id", "")
            self.count_label.configure(text=f"총 {len(self._contacts_cache)}명")
        elif self.orchestrator:
            contacts = self.orchestrator.contact_mgr.get_by_category(category)
            if search:
                search = search.lower()
                contacts = [
                    c for c in contacts
                    if search in c.name.lower()
                    or search in c.company.lower()
                    or search in c.memo.lower()
                ]
            # 최근 추가 순 (최신이 위) → insert at index 0이 아닌 "end"로 역순 삽입
            for contact in reversed(contacts):
                cat_text = cat_label_map.get(contact.category, contact.category)
                iid = self.tree.insert("", "end", values=(
                    contact.name, cat_text,
                    contact.phone or "", contact.company or "",
                    (contact.memo or "")[:30]
                ))
                self._tree_id_map[iid] = contact.id
            self.count_label.configure(text=f"총 {len(contacts)}명")

    # -- Treeview 이벤트 --

    def _get_contact_by_tree_item(self, iid):
        """Treeview item → Contact 객체"""
        contact_id = self._tree_id_map.get(iid)
        if contact_id and self.orchestrator:
            for c in self.orchestrator.contact_mgr.get_all():
                if c.id == contact_id:
                    return c
        return None

    def _on_tree_double_click(self, event):
        """더블클릭 → 편집"""
        sel = self.tree.selection()
        if not sel:
            return
        contact = self._get_contact_by_tree_item(sel[0])
        if contact:
            self._edit_contact(contact)

    def _on_tree_right_click(self, event):
        """우클릭 → 컨텍스트 메뉴"""
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        self.tree.selection_set(iid)
        contact = self._get_contact_by_tree_item(iid)
        if not contact:
            return

        menu = tk.Menu(self, tearoff=0, bg="#2d333b", fg="#e6edf3",
                       activebackground="#2f81f7", activeforeground="#fff")
        menu.add_command(label=f"편집: {contact.name}", command=lambda: self._edit_contact(contact))
        menu.add_separator()

        # 카테고리 변경 서브메뉴
        cat_menu = tk.Menu(menu, tearoff=0, bg="#2d333b", fg="#e6edf3",
                           activebackground="#2f81f7", activeforeground="#fff")
        all_cats = self.orchestrator.contact_mgr.get_all_categories() if self.orchestrator else []
        for cat in all_cats:
            cat_menu.add_command(
                label=cat,
                command=lambda c=contact, ca=cat: self._quick_change_category(c.id, ca)
            )
        menu.add_cascade(label="카테고리 변경", menu=cat_menu)
        menu.add_separator()
        menu.add_command(label="삭제", command=lambda: self._delete_contact(contact.id))
        menu.tk_popup(event.x_root, event.y_root)

    def _delete_selected(self):
        """선택된 연락처 일괄 삭제"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("알림", "삭제할 연락처를 선택하세요.")
            return
        if not messagebox.askyesno("삭제 확인", f"{len(sel)}명을 삭제하시겠습니까?"):
            return
        if self.orchestrator:
            for iid in sel:
                cid = self._tree_id_map.get(iid)
                if cid:
                    self.orchestrator.contact_mgr.delete(cid)
        self.refresh_list(category=self.category_var.get())

    def _quick_change_category(self, contact_id, new_category):
        """카테고리 즉시 변경"""
        if self.orchestrator:
            self.orchestrator.contact_mgr.update(contact_id, category=new_category)
            self._refresh_category_buttons()
            self.refresh_list(category=self.category_var.get())

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

        # 현재 보고 있는 카테고리를 기본값으로 전달
        current_cat = self.category_var.get()
        default_cat = current_cat if current_cat != "all" else "other"
        dialog = ContactDialog(self, title="연락처 추가", orchestrator=self.orchestrator,
                               default_category=default_cat)
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
            # 현재 카테고리 탭 → 카테고리 없는 항목에 자동 지정
            current_cat = self.category_var.get()
            default_cat = current_cat if current_cat != "all" else None
            result = self.orchestrator.contact_mgr.import_from_excel(
                filepath, default_category=default_cat
            )
            cat_msg = f"\n카테고리 미지정 → '{default_cat}' 자동 지정" if default_cat else ""
            messagebox.showinfo(
                "가져오기 완료",
                f"추가: {result['success']}명\n건너뜀: {result['skipped']}명{cat_msg}"
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

    def __init__(self, parent, title="연락처", contact=None, orchestrator=None,
                 default_category=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.title(title)
        self.geometry("420x640")
        self.configure(fg_color=T.BG_DARK)
        self.result = None
        self.contact = contact
        self.orchestrator = orchestrator
        self.default_category = default_category

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

        default_cat = self.contact.category if self.contact else (self.default_category or "other")
        self.category_var = ctk.StringVar(value=default_cat)
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
