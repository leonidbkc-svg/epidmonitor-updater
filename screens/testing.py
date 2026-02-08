# screens/testing.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import json
from urllib.parse import urlencode
from urllib.request import Request, urlopen

# ✅ новый импорт: экран тестирования ординаторов
from screens.ordinators_test import open_ordinators_test_screen


# -------------------------
# HTTP (без requests)
# -------------------------
def _http_get_json(url: str, headers: dict | None = None, timeout: int = 7):
    headers = headers or {}
    req = Request(url, headers=headers, method="GET")
    with urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
        return json.loads(raw)


def open_testing_screen(
    main_frame,
    build_header,
    go_back_callback,
    server_base_url: str,
    admin_pin: str,
):
    """
    Экран "Тестирование":
      - Верх: header (назад)
      - Внутри: Notebook
          1) Тестирование для мед. персонала (подтягивает БД с сервера)
          2) Тестирование для студентов
              - студенты Сеченовского университета
              - ординаторы (бывш. практиканты)
    """
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame, back_callback=go_back_callback)

    title = tk.Label(
        main_frame,
        text="Тестирование",
        font=("Segoe UI", 18, "bold"),
        bg="#f4f6f8",
        fg="#111827",
    )
    title.pack(pady=(20, 10))

    container = tk.Frame(main_frame, bg="#f4f6f8")
    container.pack(expand=True, fill="both", padx=12, pady=(0, 12))

    nb = ttk.Notebook(container)
    nb.pack(expand=True, fill="both")

    # -------------------------
    # TAB 1: Медперсонал
    # -------------------------
    tab_staff = tk.Frame(nb, bg="white")
    nb.add(tab_staff, text="Мед. персонал")

    _build_staff_tab(
        parent=tab_staff,
        server_base_url=server_base_url,
        admin_pin=admin_pin,
    )

    # -------------------------
    # TAB 2: Студенты (внутренний notebook)
    # -------------------------
    tab_students = tk.Frame(nb, bg="white")
    nb.add(tab_students, text="Студенты")

    _build_students_tab(
        parent=tab_students,
        main_frame=main_frame,
        build_header=build_header,
        go_back_callback=go_back_callback,
        server_base_url=server_base_url,
        admin_pin=admin_pin,
    )


def _build_students_tab(
    parent,
    main_frame,
    build_header,
    go_back_callback,
    server_base_url: str,
    admin_pin: str,
):
    wrap = tk.Frame(parent, bg="white")
    wrap.pack(expand=True, fill="both", padx=14, pady=14)

    info = tk.Label(
        wrap,
        text="Раздел для студентов и ординаторов.",
        font=("Segoe UI", 11),
        bg="white",
        fg="#374151",
    )
    info.pack(anchor="w", pady=(0, 10))

    nb2 = ttk.Notebook(wrap)
    nb2.pack(expand=True, fill="both")

    # -------------------------
    # TAB: Студенты Сеченовского
    # -------------------------
    tab_sechenov = tk.Frame(nb2, bg="white")
    nb2.add(tab_sechenov, text="Студенты Сеченовского")

    tk.Label(
        tab_sechenov,
        text="Здесь позже будут тесты для студентов Сеченовского университета.",
        font=("Segoe UI", 11),
        bg="white",
        fg="#111827",
        wraplength=900,
        justify="left",
    ).pack(anchor="w", padx=14, pady=14)

    # -------------------------
    # TAB: Ординаторы (бывш. Практиканты)
    # -------------------------
    tab_interns = tk.Frame(nb2, bg="white")
    nb2.add(tab_interns, text="Ординаторы")

    inner = tk.Frame(tab_interns, bg="white")
    inner.pack(expand=True, fill="both", padx=18, pady=18)

    tk.Label(
        inner,
        text="Входной тест по профилактике ИСМП",
        font=("Segoe UI", 16, "bold"),
        bg="white",
        fg="#111827",
        justify="left",
    ).pack(anchor="w", pady=(0, 10))

    body_text = (
        "В целях обеспечения эпидемиологической безопасности пациентов и медицинского персонала, "
        "а также оценки уровня базовых знаний по профилактике инфекций, связанных с оказанием "
        "медицинской помощи (ИСМП), всем студентам, прибывающим на практику в подразделения Центра, "
        "необходимо пройти входное тестирование.\n\n"
        "Тестирование проводится до начала практической деятельности и направлено на проверку знаний:\n"
        "• основных понятий ИСМП;\n"
        "• механизмов и факторов передачи инфекций;\n"
        "• правил гигиены рук;\n"
        "• использования средств индивидуальной защиты;\n"
        "• обращения с медицинскими отходами;\n"
        "• соблюдения санитарно-противоэпидемического режима в медицинских организациях.\n\n"
        "Результаты тестирования используются для определения готовности студента к работе в клинических "
        "подразделениях и носят обучающий и профилактический характер."
    )

    tk.Label(
        inner,
        text=body_text,
        font=("Segoe UI", 11),
        bg="white",
        fg="#374151",
        wraplength=900,
        justify="left",
    ).pack(anchor="w", pady=(0, 16))

    def _go_to_test():
        # ✅ Переход на отдельный экран тестовой платформы для ординаторов
        open_ordinators_test_screen(
            main_frame=main_frame,
            build_header=build_header,
            go_back_callback=lambda: open_testing_screen(
                main_frame=main_frame,
                build_header=build_header,
                go_back_callback=go_back_callback,
                server_base_url=server_base_url,
                admin_pin=admin_pin,
            ),
        )

    btn = ttk.Button(
        inner,
        text="➡️ Перейти к тестированию",
        style="Main.TButton",
        command=_go_to_test,
    )
    btn.pack(anchor="w")


def _build_staff_tab(parent, server_base_url: str, admin_pin: str):
    # UI state
    department_var = tk.StringVar()
    fio_var = tk.StringVar()

    departments_list = []
    employees_by_dep = {}  # dep -> [fio, fio...]

    # Layout
    wrap = tk.Frame(parent, bg="white")
    wrap.pack(expand=True, fill="both", padx=16, pady=16)

    tk.Label(
        wrap,
        text="Выберите отделение и сотрудника, затем можно приступать к тестированию.",
        font=("Segoe UI", 11),
        bg="white",
        fg="#374151",
        wraplength=950,
        justify="left",
    ).pack(anchor="w", pady=(0, 14))

    form = tk.Frame(wrap, bg="white")
    form.pack(anchor="w", fill="x")

    # Department
    tk.Label(form, text="Отделение", bg="white", fg="#111827").grid(
        row=0, column=0, sticky="w", pady=(0, 6)
    )
    dep_combo = ttk.Combobox(
        form,
        textvariable=department_var,
        values=[],
        state="readonly",
        width=45,
    )
    dep_combo.grid(row=1, column=0, sticky="w")

    # FIO
    tk.Label(form, text="Фамилия (ФИО)", bg="white", fg="#111827").grid(
        row=2, column=0, sticky="w", pady=(14, 6)
    )
    fio_combo = ttk.Combobox(
        form,
        textvariable=fio_var,
        values=[],
        state="disabled",
        width=45,
    )
    fio_combo.grid(row=3, column=0, sticky="w")

    # Status line
    status_var = tk.StringVar(value="Загрузка списка отделений…")
    status_lbl = tk.Label(
        wrap,
        textvariable=status_var,
        font=("Segoe UI", 10),
        bg="white",
        fg="#6b7280",
    )
    status_lbl.pack(anchor="w", pady=(14, 10))

    # Start button
    start_btn = ttk.Button(
        wrap,
        text="✅ Приступить к тестированию",
        style="Main.TButton",
        state="disabled",
    )
    start_btn.pack(anchor="w", pady=(6, 0))

    def _set_start_enabled():
        if department_var.get() and fio_var.get():
            start_btn.config(state="normal")
        else:
            start_btn.config(state="disabled")

    def _on_start():
        # тестов пока нет — заглушка
        messagebox.showinfo(
            "Тестирование",
            f"Ок: {department_var.get()}\n{fio_var.get()}\n\nТесты позже подключим.",
        )

    start_btn.config(command=_on_start)

    # Events
    def on_department_selected(_=None):
        fio_var.set("")
        _set_start_enabled()

        dep = department_var.get()
        if not dep:
            fio_combo.config(values=[], state="disabled")
            return

        # если уже загружали — просто подставляем
        if dep in employees_by_dep:
            values = employees_by_dep.get(dep, [])
            fio_combo.config(values=values, state="readonly" if values else "disabled")
            status_var.set("Выберите сотрудника.")
            return

        # иначе — тянем с сервера
        fio_combo.config(values=[], state="disabled")
        status_var.set("Загрузка сотрудников отдела…")

        def worker():
            try:
                # preferred endpoint: vaccination server
                url = f"{server_base_url.rstrip('/')}/api/department?{urlencode({'name': dep})}"
                data = _http_get_json(
                    url,
                    headers={"X-Admin-Pin": admin_pin},
                    timeout=10,
                )
                emps = data.get("employees", []) if isinstance(data, dict) else []
                if not emps:
                    # fallback (older server): /api/employees?department=
                    url = f"{server_base_url.rstrip('/')}/api/employees?{urlencode({'department': dep})}"
                    data = _http_get_json(
                        url,
                        headers={"X-Admin-Pin": admin_pin},
                        timeout=10,
                    )
                    emps = data if isinstance(data, list) else []

                fios = sorted(
                    [
                        (e.get("fio") or e.get("full_name") or "").strip()
                        for e in emps
                        if (e.get("fio") or e.get("full_name"))
                    ],
                    key=str.lower,
                )
            except Exception as e:
                fios = None
                err = str(e)

            def apply():
                if fios is None:
                    employees_by_dep[dep] = []
                    fio_combo.config(values=[], state="disabled")
                    status_var.set("Ошибка загрузки сотрудников. Проверь связь с сервером.")
                    messagebox.showerror("Ошибка", f"Не удалось загрузить сотрудников:\n{err}")
                    return

                employees_by_dep[dep] = fios
                fio_combo.config(values=fios, state="readonly" if fios else "disabled")
                status_var.set(
                    "Выберите сотрудника." if fios else "В этом отделении нет сотрудников."
                )

            parent.after(0, apply)

        threading.Thread(target=worker, daemon=True).start()

    def on_fio_selected(_=None):
        _set_start_enabled()

    dep_combo.bind("<<ComboboxSelected>>", on_department_selected)
    fio_combo.bind("<<ComboboxSelected>>", on_fio_selected)

    # Initial load departments (from /api/stats)
    def load_departments():
        def worker():
            try:
                url = f"{server_base_url.rstrip('/')}/api/stats"
                data = _http_get_json(
                    url,
                    headers={"X-Admin-Pin": admin_pin},
                    timeout=10,
                )
                by_dep = data.get("by_department", {}) or {}
                deps = sorted([d for d in by_dep.keys() if d and d != "—"], key=str.lower)
            except Exception as e:
                deps = None
                err = str(e)

            def apply():
                nonlocal departments_list
                if deps is None:
                    status_var.set("Ошибка загрузки отделений. Проверь сервер/пин.")
                    messagebox.showerror("Ошибка", f"Не удалось загрузить отделения:\n{err}")
                    dep_combo.config(values=[], state="disabled")
                    return

                departments_list = deps
                dep_combo.config(values=departments_list, state="readonly")
                status_var.set("Выберите отделение.")

            parent.after(0, apply)

        threading.Thread(target=worker, daemon=True).start()

    load_departments()
