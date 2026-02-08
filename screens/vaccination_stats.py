import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import requests
import time

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class VaccinationStatsScreen:
    def __init__(
        self,
        main_frame,
        build_header,
        go_back_callback,
        server_base_url,
        access_password,
        admin_pin,
        refresh_ms=30000,
    ):
        self.main_frame = main_frame
        self.build_header = build_header
        self.go_back_callback = go_back_callback
        self.server_base_url = server_base_url.rstrip("/")
        self.access_password = access_password
        self.admin_pin = admin_pin
        self.refresh_ms = refresh_ms

        self.after_id = None
        self.root = None

        # === СТАРОЕ (ОСТАВЛЕНО, НЕ УДАЛЯЮ) ===
        self.total_var = tk.StringVar(value="Всего: —")
        self.v_var = tk.StringVar(value="Привит: —")
        self.nv_var = tk.StringVar(value="Не привит: —")
        self.m_var = tk.StringVar(value="Медотвод: —")
        self.e_var = tk.StringVar(value="Пусто: —")

        # === НОВОЕ ДЛЯ КАРТОЧЕК ===
        self.total_num_var = tk.StringVar(value="—")
        self.v_num_var = tk.StringVar(value="—")
        self.nv_num_var = tk.StringVar(value="—")
        self.m_num_var = tk.StringVar(value="—")
        self.e_num_var = tk.StringVar(value="—")

        self.tree = None
        self.status_label = None

        # вкладки
        self.notebook = None
        self.tab_table = None
        self.tab_charts = None
        self.tab_search = None

        # график
        self.chart_holder = None
        self.chart_canvas = None

        # чекбоксы отделений (dep -> BooleanVar)
        self.dep_vars = {}

        # UI контейнер для чекбоксов
        self.dep_canvas = None
        self.dep_scrollbar = None
        self.dep_checks_frame = None

        # мини-вкладки графиков справа
        self.chart_tabs = None
        self.tab_vaccinated = None
        self.tab_not_vaccinated = None

        # последние данные
        # rows: (dep, total, vaccinated, pct_vaccinated, not_vaccinated, medical, empty)
        self._last_rows = []

        # ===== Поиск =====
        self.search_query_var = tk.StringVar(value="")
        self.search_status_var = tk.StringVar(value="Введите ФИО/полис и нажмите «Найти»")
        self.search_tree = None

        # кэш сотрудников (чтобы не опрашивать сервер каждый раз)
        self._employees_cache = []          # list[dict]
        self._employees_cache_ts = 0.0
        self._employees_cache_ttl = 180.0   # секунд

    # =====================
    # ENTRY
    # =====================
    def show(self):
        if not self._check_password():
            return
        self._build_ui()
        self.refresh()

    # =====================
    # PASSWORD
    # =====================
    def _check_password(self):
        pwd = simpledialog.askstring(
            "Доступ",
            "Введите пароль для просмотра статистики вакцинации:",
            show="*",
            parent=self.main_frame.winfo_toplevel(),
        )
        if pwd is None:
            return False
        if pwd != self.access_password:
            messagebox.showerror("Ошибка", "Неверный пароль")
            return False
        return True

    # =====================
    # UI
    # =====================
    def _build_ui(self):
        for w in self.main_frame.winfo_children():
            w.destroy()

        self.build_header(self.main_frame, back_callback=self._go_back)

        tk.Label(
            self.main_frame,
            text="Статистика вакцинации",
            font=("Segoe UI", 18, "bold"),
            bg="#f4f6f8",
        ).pack(pady=(20, 10))

        self.root = tk.Frame(self.main_frame, bg="#f4f6f8")
        self.root.pack(expand=True, fill="both", padx=20, pady=10)
        self.root.bind("<Destroy>", lambda e: self._stop_timer())

        # ==========================
        # ИТОГИ (красивые карточки)
        # ==========================
        totals = tk.LabelFrame(self.root, text="Итоги", bg="#f4f6f8", padx=12, pady=10)
        totals.pack(fill="x", pady=(0, 12))

        cards = tk.Frame(totals, bg="#f4f6f8")
        cards.pack(fill="x")

        def make_card(parent, title, value_var):
            card = tk.Frame(parent, bg="white", bd=1, relief="solid")
            card.pack(side="left", padx=8, pady=6, expand=True, fill="x")

            tk.Label(
                card,
                text=title,
                bg="white",
                fg="#6b7280",
                font=("Segoe UI", 9)
            ).pack(anchor="w", padx=12, pady=(8, 0))

            tk.Label(
                card,
                textvariable=value_var,
                bg="white",
                fg="#111827",
                font=("Segoe UI", 18, "bold")
            ).pack(anchor="w", padx=12, pady=(2, 10))

            return card

        make_card(cards, "Всего", self.total_num_var)
        make_card(cards, "Привит", self.v_num_var)
        make_card(cards, "Не привит", self.nv_num_var)
        make_card(cards, "Медотвод", self.m_num_var)
        make_card(cards, "Пусто", self.e_num_var)

        # ======================================================
        # NOTEBOOK
        # ======================================================
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")

        self.tab_table = tk.Frame(self.notebook, bg="#f4f6f8")
        self.tab_charts = tk.Frame(self.notebook, bg="#f4f6f8")
        self.tab_search = tk.Frame(self.notebook, bg="#f4f6f8")

        self.notebook.add(self.tab_table, text="Таблица")
        self.notebook.add(self.tab_charts, text="Графики")
        self.notebook.add(self.tab_search, text="Поиск")

        # ===== TAB: TABLE =====
        table_box = tk.LabelFrame(
            self.tab_table,
            text="По отделениям",
            bg="#f4f6f8",
            padx=10,
            pady=10
        )
        table_box.pack(expand=True, fill="both", padx=8, pady=8)

        cols = ("dep", "total", "vaccinated", "percent", "not_vaccinated", "medical", "empty")
        self.tree = ttk.Treeview(table_box, columns=cols, show="headings", height=14)

        headers = {
            "dep": "Отделение",
            "total": "Всего",
            "vaccinated": "Привит",
            "percent": "% привит",
            "not_vaccinated": "Не привит",
            "medical": "Медотвод",
            "empty": "Пусто",
        }

        for c in cols:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, anchor="center", width=90)

        self.tree.column("dep", anchor="w", width=320)

        vsb = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", expand=True, fill="both")
        vsb.pack(side="right", fill="y")

        self.tree.tag_configure("green", background="#d1fae5")
        self.tree.tag_configure("yellow", background="#fef9c3")
        self.tree.tag_configure("red", background="#fee2e2")

        # ✅ двойной клик открывает список сотрудников
        self.tree.bind("<Double-1>", self._on_department_double_click)

        # ===== TAB: CHARTS =====
        charts_wrap = tk.Frame(self.tab_charts, bg="#f4f6f8")
        charts_wrap.pack(expand=True, fill="both", padx=8, pady=8)

        # слева — выбор отделений + кнопки
        left_panel = tk.LabelFrame(charts_wrap, text="Выбор отделений", bg="#f4f6f8", padx=10, pady=10)
        left_panel.pack(side="left", fill="y", padx=(0, 10))

        tk.Label(
            left_panel,
            text="Отметь отделения галочками:",
            bg="#f4f6f8",
            font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(0, 6))

        btns = tk.Frame(left_panel, bg="#f4f6f8")
        btns.pack(fill="x", pady=(0, 10))

        ttk.Button(btns, text="Выбрать все", command=self._select_all_departments).pack(fill="x", pady=3)
        ttk.Button(btns, text="Очистить", command=self._clear_departments_selection).pack(fill="x", pady=3)
        ttk.Button(btns, text="Построить график", command=self._render_selected_chart).pack(fill="x", pady=(10, 3))

        self.dep_canvas = tk.Canvas(left_panel, highlightthickness=0, width=420, height=470)
        self.dep_canvas.pack(side="left", fill="y")

        self.dep_scrollbar = ttk.Scrollbar(left_panel, orient="vertical", command=self.dep_canvas.yview)
        self.dep_scrollbar.pack(side="right", fill="y")

        self.dep_canvas.configure(yscrollcommand=self.dep_scrollbar.set)

        self.dep_checks_frame = tk.Frame(self.dep_canvas, bg="#f4f6f8")
        self.dep_canvas.create_window((0, 0), window=self.dep_checks_frame, anchor="nw")

        self.dep_checks_frame.bind(
            "<Configure>",
            lambda e: self.dep_canvas.configure(scrollregion=self.dep_canvas.bbox("all"))
        )

        def _on_mousewheel(event):
            try:
                self.dep_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        self.dep_canvas.bind("<Enter>", lambda e: self.dep_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.dep_canvas.bind("<Leave>", lambda e: self.dep_canvas.unbind_all("<MouseWheel>"))

        # справа — вкладки графиков + график
        right_panel = tk.Frame(charts_wrap, bg="#f4f6f8")
        right_panel.pack(side="right", expand=True, fill="both")

        top_row = tk.Frame(right_panel, bg="#f4f6f8")
        top_row.pack(fill="x", pady=(0, 8))

        self.chart_tabs = ttk.Notebook(top_row)
        self.chart_tabs.pack(side="left")

        self.tab_vaccinated = tk.Frame(self.chart_tabs, bg="#f4f6f8")
        self.tab_not_vaccinated = tk.Frame(self.chart_tabs, bg="#f4f6f8")

        self.chart_tabs.add(self.tab_vaccinated, text="Процент привитых")
        self.chart_tabs.add(self.tab_not_vaccinated, text="Процент непривитых")

        chart_box = tk.LabelFrame(right_panel, text="", bg="#f4f6f8", padx=10, pady=10)
        chart_box.pack(expand=True, fill="both")

        self.chart_holder = tk.Frame(chart_box, bg="#f4f6f8")
        self.chart_holder.pack(expand=True, fill="both")

        self.chart_tabs.bind("<<NotebookTabChanged>>", lambda e: self._render_selected_chart(silent=True))

        def on_main_tab_changed(_event=None):
            try:
                tab_id = self.notebook.select()
                tab_text = self.notebook.tab(tab_id, "text")
            except Exception:
                return
            if tab_text == "Графики":
                self._populate_departments_checks()

        self.notebook.bind("<<NotebookTabChanged>>", on_main_tab_changed)

        # ===== TAB: SEARCH =====
        self._build_search_tab()

        # Нижняя панель
        bottom = tk.Frame(self.root, bg="#f4f6f8")
        bottom.pack(fill="x", pady=(10, 0))

        ttk.Button(bottom, text="🔄 Обновить", command=self.refresh).pack(side="left")

        self.status_label = tk.Label(
            bottom,
            text="",
            bg="#f4f6f8",
            fg="#6b7280",
            font=("Segoe UI", 9),
        )
        self.status_label.pack(side="left", padx=12)

    # =====================
    # SEARCH TAB UI
    # =====================
    def _build_search_tab(self):
        wrap = tk.Frame(self.tab_search, bg="#f4f6f8")
        wrap.pack(expand=True, fill="both", padx=8, pady=8)

        top = tk.LabelFrame(wrap, text="Поиск сотрудника", bg="#f4f6f8", padx=10, pady=10)
        top.pack(fill="x", pady=(0, 10))

        tk.Label(
            top,
            text="ФИО или номер полиса (можно часть):",
            bg="#f4f6f8",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        row = tk.Frame(top, bg="#f4f6f8")
        row.pack(fill="x", pady=(6, 0))

        entry = ttk.Entry(row, textvariable=self.search_query_var)
        entry.pack(side="left", fill="x", expand=True)

        def do_search(_e=None):
            self._search_employee()

        entry.bind("<Return>", do_search)

        ttk.Button(row, text="Найти", command=self._search_employee).pack(side="left", padx=8)
        ttk.Button(row, text="Сброс", command=self._clear_search).pack(side="left")

        tk.Label(
            top,
            textvariable=self.search_status_var,
            bg="#f4f6f8",
            fg="#6b7280",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(8, 0))

        box = tk.LabelFrame(wrap, text="Результаты", bg="#f4f6f8", padx=10, pady=10)
        box.pack(expand=True, fill="both")

        cols = ("dep", "fio", "position", "status", "policy", "date", "place")
        self.search_tree = ttk.Treeview(box, columns=cols, show="headings")

        headers = {
            "dep": "Отделение",
            "fio": "ФИО",
            "position": "Должность",
            "status": "Статус",
            "policy": "Полис",
            "date": "Дата",
            "place": "Место",
        }

        widths = {
            "dep": 260,
            "fio": 260,
            "position": 200,
            "status": 90,
            "policy": 140,
            "date": 90,
            "place": 220,
        }

        for c in cols:
            self.search_tree.heading(c, text=headers[c])
            self.search_tree.column(c, width=widths[c], anchor="w")

        self.search_tree.column("status", anchor="center")
        self.search_tree.column("date", anchor="center")

        vsb = ttk.Scrollbar(box, orient="vertical", command=self.search_tree.yview)
        self.search_tree.configure(yscrollcommand=vsb.set)

        self.search_tree.pack(side="left", expand=True, fill="both")
        vsb.pack(side="right", fill="y")

        self.search_tree.tag_configure("vacc", background="#d1fae5")
        self.search_tree.tag_configure("not", background="#fee2e2")
        self.search_tree.tag_configure("med", background="#fef9c3")
        self.search_tree.tag_configure("empty", background="#f3f4f6")

        self.search_tree.bind("<Double-1>", self._on_search_result_double_click)

    def _clear_search(self):
        self.search_query_var.set("")
        self.search_status_var.set("Введите ФИО/полис и нажмите «Найти»")
        if self.search_tree:
            for i in self.search_tree.get_children():
                self.search_tree.delete(i)

    def _on_search_result_double_click(self, _event=None):
        if not self.search_tree:
            return
        sel = self.search_tree.selection()
        if not sel:
            return
        vals = self.search_tree.item(sel[0], "values")
        if not vals:
            return
        dep = vals[0]
        if dep:
            self._open_department_details(dep)

    def _get_all_employees_cached(self):
        now = time.time()
        if self._employees_cache and (now - self._employees_cache_ts) < self._employees_cache_ttl:
            return self._employees_cache

        if not self._last_rows:
            self.refresh()

        deps = [r[0] for r in self._last_rows]
        employees = []

        for idx, dep in enumerate(deps, start=1):
            self.search_status_var.set(f"Сканирование отделений: {idx}/{len(deps)} …")
            try:
                self.main_frame.update_idletasks()
            except Exception:
                pass

            try:
                r = requests.get(
                    f"{self.server_base_url}/api/department",
                    params={"name": dep},
                    headers={"X-Admin-Pin": self.admin_pin},
                    timeout=10,
                )
                if r.status_code == 403:
                    raise RuntimeError("Доступ запрещён (неверный PIN).")

                r.raise_for_status()
                data = r.json()
                emps = data.get("employees", []) or []

                for e in emps:
                    employees.append({
                        "dep": dep,
                        "fio": (e.get("fio") or "").strip(),
                        "position": (e.get("position") or "").strip(),
                        "status": (e.get("status") or "").strip(),
                        "policy": (e.get("policy_number") or "").strip(),
                        "date": (e.get("vaccination_date") or "").strip(),
                        "place": (e.get("vaccination_place") or "").strip(),
                    })

            except Exception:
                # если одно отделение не отдалось — продолжаем
                continue

        self._employees_cache = employees
        self._employees_cache_ts = time.time()
        return employees

    def _search_employee(self):
        q = (self.search_query_var.get() or "").strip()
        if not q:
            messagebox.showinfo("Поиск", "Введите ФИО или номер полиса.")
            return

        if self.search_tree:
            for i in self.search_tree.get_children():
                self.search_tree.delete(i)

        try:
            all_emps = self._get_all_employees_cached()

            q_low = q.lower()
            found = []
            for e in all_emps:
                fio = (e.get("fio") or "").lower()
                policy = (e.get("policy") or "").lower()
                if q_low in fio or (q_low and q_low in policy):
                    found.append(e)

            found.sort(key=lambda x: (x.get("dep", ""), x.get("fio", "")))

            for e in found:
                st = e.get("status") or ""
                if st == "привит":
                    tag = "vacc"
                elif st == "не привит":
                    tag = "not"
                elif st == "медотвод":
                    tag = "med"
                else:
                    tag = "empty"

                self.search_tree.insert(
                    "",
                    "end",
                    values=(
                        e.get("dep", ""),
                        e.get("fio", ""),
                        e.get("position", ""),
                        e.get("status", ""),
                        e.get("policy", ""),
                        e.get("date", ""),
                        e.get("place", ""),
                    ),
                    tags=(tag,),
                )

            self.search_status_var.set(f"Найдено: {len(found)} (кэш {int(self._employees_cache_ttl)} сек.)")

        except Exception as ex:
            self.search_status_var.set("Ошибка поиска")
            messagebox.showerror("Ошибка", str(ex))

    # =====================
    # TABLE => DEPARTMENT DETAILS
    # =====================
    def _on_department_double_click(self, _event=None):
        if not self.tree:
            return
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        values = self.tree.item(item_id, "values")
        if not values:
            return
        dep = values[0]
        self._open_department_details(dep)

    def _open_department_details(self, department: str):
        win = tk.Toplevel(self.main_frame.winfo_toplevel())
        win.title(f"Вакцинация — {department}")
        win.geometry("1100x650")
        win.minsize(900, 500)

        top = tk.Frame(win)
        top.pack(fill="x", padx=12, pady=10)

        tk.Label(top, text=department, font=("Segoe UI", 14, "bold")).pack(side="left")

        status = tk.Label(top, text="Загрузка...", fg="#6b7280")
        status.pack(side="left", padx=12)

        body = tk.Frame(win)
        body.pack(expand=True, fill="both", padx=12, pady=(0, 12))

        cols = ("fio", "position", "status", "policy", "date", "place")
        tree = ttk.Treeview(body, columns=cols, show="headings")

        headers = {
            "fio": "ФИО",
            "position": "Должность",
            "status": "Статус",
            "policy": "Полис",
            "date": "Дата",
            "place": "Место",
        }

        widths = {
            "fio": 260,
            "position": 220,
            "status": 90,
            "policy": 140,
            "date": 90,
            "place": 220,
        }

        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=widths[c], anchor="w")

        tree.column("status", anchor="center")
        tree.column("date", anchor="center")

        vsb = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)

        tree.pack(side="left", expand=True, fill="both")
        vsb.pack(side="right", fill="y")

        tree.tag_configure("vacc", background="#d1fae5")   # привит
        tree.tag_configure("not", background="#fee2e2")    # не привит
        tree.tag_configure("med", background="#fef9c3")    # медотвод
        tree.tag_configure("empty", background="#f3f4f6")  # пусто

        def load():
            try:
                r = requests.get(
                    f"{self.server_base_url}/api/department",
                    params={"name": department},
                    headers={"X-Admin-Pin": self.admin_pin},
                    timeout=8,
                )
                if r.status_code == 403:
                    status.config(text="403 (PIN)", fg="#b91c1c")
                    messagebox.showerror("Доступ запрещён", "Неверный PIN на сервере статистики.")
                    return

                r.raise_for_status()
                data = r.json()
                employees = data.get("employees", []) or []

                for i in tree.get_children():
                    tree.delete(i)

                for e in employees:
                    fio = e.get("fio") or ""
                    position = e.get("position") or ""
                    st = e.get("status") or ""
                    policy = e.get("policy_number") or ""
                    d = e.get("vaccination_date") or ""
                    place = e.get("vaccination_place") or ""

                    if st == "привит":
                        tag = "vacc"
                    elif st == "не привит":
                        tag = "not"
                    elif st == "медотвод":
                        tag = "med"
                    else:
                        tag = "empty"

                    tree.insert("", "end", values=(fio, position, st, policy, d, place), tags=(tag,))

                status.config(text=f"Сотрудников: {len(employees)}", fg="#111827")

            except Exception as ex:
                status.config(text="Ошибка загрузки", fg="#b91c1c")
                messagebox.showerror("Ошибка", str(ex))

        ttk.Button(top, text="Обновить", command=load).pack(side="right")
        load()

    # =====================
    # CHECKBOX HELPERS
    # =====================
    def _populate_departments_checks(self):
        if self.dep_checks_frame is None:
            return

        current_deps = [r[0] for r in self._last_rows]

        for dep in list(self.dep_vars.keys()):
            if dep not in current_deps:
                del self.dep_vars[dep]

        for w in self.dep_checks_frame.winfo_children():
            w.destroy()

        for dep in current_deps:
            var = self.dep_vars.get(dep)
            if var is None:
                var = tk.BooleanVar(value=False)
                self.dep_vars[dep] = var

            cb = ttk.Checkbutton(self.dep_checks_frame, text=dep, variable=var)
            cb.pack(anchor="w", pady=2)

        try:
            self.dep_canvas.update_idletasks()
            self.dep_canvas.configure(scrollregion=self.dep_canvas.bbox("all"))
        except Exception:
            pass

    def _select_all_departments(self):
        for var in self.dep_vars.values():
            var.set(True)

    def _clear_departments_selection(self):
        for var in self.dep_vars.values():
            var.set(False)
        self._render_chart([])

    def _get_selected_rows(self):
        selected = [dep for dep, var in self.dep_vars.items() if var.get()]
        if not selected:
            return []
        s = set(selected)
        return [r for r in self._last_rows if r[0] in s]

    def _render_selected_chart(self, silent=False):
        rows = self._get_selected_rows()
        if not rows:
            if not silent:
                messagebox.showinfo("Выбор отделений", "Отметьте хотя бы одно отделение галочкой.")
            return

        try:
            tab_id = self.chart_tabs.select()
            tab_text = self.chart_tabs.tab(tab_id, "text")
        except Exception:
            tab_text = "Процент привитых"

        if tab_text == "Процент непривитых":
            rows2 = []
            for dep, t, v, pct_v, nv, m, e in rows:
                not_ok = int(nv) + int(m) + int(e)
                pct_not_ok = round((not_ok / t) * 100, 1) if t else 0.0
                rows2.append((dep, t, v, pct_v, nv, m, e, pct_not_ok))

            rows2.sort(key=lambda x: x[7], reverse=True)
            self._render_chart(rows2, mode="not_ok")
        else:
            rows.sort(key=lambda x: x[3], reverse=True)
            self._render_chart(rows, mode="vaccinated")

    # =====================
    # DATA
    # =====================
    def refresh(self):
        self._stop_timer()

        for i in self.tree.get_children():
            self.tree.delete(i)

        try:
            r = requests.get(
                f"{self.server_base_url}/api/stats",
                headers={"X-Admin-Pin": self.admin_pin},
                timeout=5,
            )
            if r.status_code == 403:
                messagebox.showerror("Доступ запрещён", "Неверный PIN на сервере статистики.")
                self.status_label.config(text="403 (PIN)")
                return

            r.raise_for_status()
            data = r.json()

            total = int(data.get("total", 0))
            vaccinated = int(data.get("vaccinated", 0))
            not_vaccinated = int(data.get("not_vaccinated", 0))
            medical = int(data.get("medical", 0))
            empty = int(data.get("empty", 0))

            # === НОВОЕ (карточки) ===
            self.total_num_var.set(str(total))
            self.v_num_var.set(str(vaccinated))
            self.nv_num_var.set(str(not_vaccinated))
            self.m_num_var.set(str(medical))
            self.e_num_var.set(str(empty))

            # === СТАРОЕ (ОСТАВЛЕНО, НЕ УДАЛЯЮ) ===
            self.total_var.set(f"Всего: {total}")
            self.v_var.set(f"Привит: {vaccinated}")
            self.nv_var.set(f"Не привит: {not_vaccinated}")
            self.m_var.set(f"Медотвод: {medical}")
            self.e_var.set(f"Пусто: {empty}")

            rows = []
            by_dep = data.get("by_department", {}) or {}

            for dep, st in by_dep.items():
                t = int(st.get("total", 0))
                v = int(st.get("vaccinated", 0))
                nv = int(st.get("not_vaccinated", 0))
                m = int(st.get("medical", 0))
                e = int(st.get("empty", 0))
                pct = round(v / t * 100, 1) if t else 0.0
                rows.append((dep, t, v, pct, nv, m, e))

            rows.sort(key=lambda x: (x[3], x[1]), reverse=True)

            for dep, t, v, pct, nv, m, e in rows:
                if pct >= 90:
                    tag = "green"
                elif pct >= 70:
                    tag = "yellow"
                else:
                    tag = "red"
                self.tree.insert("", "end", values=(dep, t, v, pct, nv, m, e), tags=(tag,))

            self._last_rows = rows
            self._populate_departments_checks()

            # сброс кэша поиска (чтобы видеть актуальные данные)
            self._employees_cache = []
            self._employees_cache_ts = 0.0

            self.status_label.config(text="Обновлено: " + time.strftime("%H:%M:%S"))

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            self.status_label.config(text="Ошибка обновления")

        self.after_id = self.main_frame.after(self.refresh_ms, self.refresh)

    # =====================
    # CHART
    # =====================
    def _render_chart(self, rows, mode="vaccinated"):
        if not self.chart_holder or not self.chart_holder.winfo_exists():
            return

        if self.chart_canvas is not None:
            try:
                self.chart_canvas.get_tk_widget().destroy()
            except Exception:
                pass
            self.chart_canvas = None

        if not rows:
            return

        if mode == "not_ok":
            deps = [r[0] for r in rows][::-1]
            pcts = [float(r[7]) for r in rows][::-1]
            title = "Процент непривитых по отделениям"
            xlabel = "% непривитых"
        else:
            deps = [r[0] for r in rows][::-1]
            pcts = [float(r[3]) for r in rows][::-1]
            title = "Процент привитых по отделениям"
            xlabel = "% привитых"

        fig = Figure(figsize=(12.5, 6.5), dpi=100)
        ax = fig.add_subplot(111)

        ax.barh(deps, pcts)
        ax.set_xlim(0, 100)
        ax.set_xlabel(xlabel)
        ax.set_title(title)

        for i, v in enumerate(pcts):
            ax.text(min(v + 1, 99), i, f"{v}%", va="center", fontsize=9)

        fig.subplots_adjust(left=0.38, right=0.98, top=0.92, bottom=0.10)

        canvas = FigureCanvasTkAgg(fig, master=self.chart_holder)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both")
        self.chart_canvas = canvas

    # =====================
    # EXIT
    # =====================
    def _stop_timer(self):
        if self.after_id:
            try:
                self.main_frame.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None

    def _go_back(self):
        self._stop_timer()
        self.go_back_callback()


def open_vaccination_stats(
    main_frame,
    build_header,
    go_back_callback,
    server_base_url,
    access_password,
    admin_pin,
    refresh_ms=30000,
):
    VaccinationStatsScreen(
        main_frame,
        build_header,
        go_back_callback,
        server_base_url,
        access_password,
        admin_pin,
        refresh_ms,
    ).show()
