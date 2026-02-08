from __future__ import annotations

import os
import re
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

import pandas as pd

from config.app_config import load_config, save_config
from analysis.swabs_journal import SwabsJournal, SHEETS_DEFAULT
from analysis.dep_mapper import load_aliases, save_aliases, try_match_department
from screens.dep_map_dialog import ask_user_map_unknowns
from services import webdav_sync

from analysis.report_builder import build_docx_report


# === ПУТЬ К ШАБЛОНУ DOCX (жёстко в коде, как изначально) ===
REPORT_TEMPLATE_PATH = r"\\192.168.137.17\c$\EpidArchive\documents\шаблон.docx"


DEFAULT_DEPARTMENTS = [
    "1АФО", "АО", "1ОАПБ", "2ОАПБ", "1ГО", "2ГО", "ГО", "ХО", "ООГ",
    "ОРИТН", "ОПННД", "ОАР",
    "ОПЕРБЛОК 2 ЭТАЖ",
    "ОПЕРБЛОК 3 ЭТАЖ НАДСТРОЙ",
    "ОПЕРБЛОК 4 ЭТАЖ КДЦ",
    "ОПЕРБЛОК 5 ЭТАЖ",
    "ОПЕРБЛОК 6 ЭТАЖ",
    "ОПЕРБЛОК 8 ЭТАЖ",
    "ОПЕРБЛОК 6 ЭТАЖ ФПЦ",
    "Клиническая генетика",
    "ОРИТН_ФПЦ", "ОН1", "ОН2_ФПЦ", "ТЕСТ",
    "СДП", "ОАиУ", "ОПМЖ", "ОХН (хир)",
    "2АФО", "ОАРИТН", "ОАР_ФПЦ",
    "2ОПННД", "ОПЛТ", "ОАР_КДЦ",
    "Столовая", "Пищеблок", "Аптека", "Участок по стирке и ремонту белья",
    "1ПАО", "2ПАО", "3ПАО", "ОТЭГ", "ОЭГиР",
    "ООХМЛ", "ОГРХ", "ОАР_Г", "ОРИТ_3"
]


def get_departments(cfg=None) -> list[str]:
    if cfg is None:
        cfg = load_config()
    raw = cfg.get("departments", []) or []
    deps = [str(d).strip() for d in raw if str(d).strip()]
    return deps or list(DEFAULT_DEPARTMENTS)


# ===== ESKAPE подсветка =====
ESKAPE_PATTERNS = [
    r"\benterococcus\s+faecium\b", r"\be\.?\s*faecium\b",
    r"\bstaphylococcus\s+aureus\b", r"\bs\.?\s*aureus\b",
    r"\bklebsiella\s+pneumoniae\b", r"\bk\.?\s*pneumoniae\b",
    r"\bacinetobacter\s+baumannii\b", r"\ba\.?\s*baumannii\b",
    r"\bpseudomonas\s+aeruginosa\b", r"\bp\.?\s*aeruginosa\b",
    r"\benterobacter\b",
]


def _norm_microbe(x: object) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if s in ("", "-", "—", "–"):
        return ""
    s = s.replace("\xa0", " ").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _is_eskape(microbe_text: str) -> bool:
    if not microbe_text:
        return False
    for pat in ESKAPE_PATTERNS:
        if re.search(pat, microbe_text, flags=re.IGNORECASE):
            return True
    return False


def build_swab_monitoring_screen(main_frame, build_header, go_back_callback):
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame)

    cfg = load_config()
    webdav_url = (cfg.get("webdav_url") or "").strip()
    if webdav_url:
        from microbio_app import DATA_ROOT
        webdav_sync.sync_down(DATA_ROOT)

    departments = get_departments(cfg)

    def _ensure_report_dir() -> str | None:
        path = (cfg.get("swabs_report_dir") or "").strip()
        if not path:
            messagebox.showinfo(
                "Папка отчётов",
                "Выберите папку, куда будут сохраняться сформированные отчёты."
            )
            path = filedialog.askdirectory(title="Выберите папку для отчётов")
            if not path:
                return None
            cfg["swabs_report_dir"] = path
            save_config(cfg)
        try:
            os.makedirs(path, exist_ok=True)
        except Exception:
            pass
        return path

    def _ensure_template_path() -> str | None:
        path = (cfg.get("swabs_template_path") or "").strip()
        if path and os.path.exists(path):
            return path

        if os.path.exists(REPORT_TEMPLATE_PATH):
            cfg["swabs_template_path"] = REPORT_TEMPLATE_PATH
            save_config(cfg)
            return REPORT_TEMPLATE_PATH

        # если есть WebDAV и шаблон уже синхронизирован — используем его
        if webdav_url:
            from microbio_app import DATA_ROOT
            candidate = os.path.join(DATA_ROOT, "documents", "шаблон.docx")
            if os.path.exists(candidate):
                cfg["swabs_template_path"] = candidate
                save_config(cfg)
                return candidate

        messagebox.showinfo(
            "Шаблон отчёта",
            "Выберите файл шаблона DOCX для формирования отчёта."
        )
        new_file = filedialog.askopenfilename(
            title="Выберите шаблон отчёта",
            filetypes=[("Word", "*.docx")]
        )
        if not new_file:
            return None

        if webdav_url:
            try:
                from microbio_app import DATA_ROOT
                target = os.path.join(DATA_ROOT, "documents", "шаблон.docx")
                os.makedirs(os.path.dirname(target), exist_ok=True)
                shutil.copy2(new_file, target)
                webdav_sync.upload_file(target, DATA_ROOT)
                new_file = target
            except Exception:
                pass

        cfg["swabs_template_path"] = new_file
        save_config(cfg)
        return new_file
    def _aliases_path() -> str:
        p = (cfg.get("swabs_aliases_json") or "").strip()
        if webdav_url:
            from microbio_app import DATA_ROOT
            return os.path.join(DATA_ROOT, "config", "swabs_dep_aliases.json")
        if not p:
            from microbio_app import DATA_ROOT
            p = os.path.join(DATA_ROOT, "config", "swabs_dep_aliases.json")
            cfg["swabs_aliases_json"] = p
            save_config(cfg)
        return p

    tk.Label(
        main_frame,
        text="Мониторинг (журнал)",
        font=("Segoe UI", 18, "bold"),
        bg="#f4f6f8"
    ).pack(pady=(18, 8))

    root = tk.Frame(main_frame, bg="#f4f6f8")
    root.pack(expand=True, fill="both", padx=18, pady=10)

    # ===== Верхняя панель =====
    top = tk.Frame(root, bg="#f4f6f8")
    top.pack(fill="x", pady=(0, 10))

    journal_lbl = tk.Label(
        top,
        text=f"Файл: {cfg.get('swabs_journal_xlsx', '')}",
        bg="#f4f6f8",
        fg="#444",
        anchor="w"
    )
    journal_lbl.pack(side="left", expand=True, fill="x", padx=(10, 10))

    def change_journal_path():
        new_file = filedialog.askopenfilename(
            title="Выберите Excel-журнал",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if not new_file:
            return
        cfg["swabs_journal_xlsx"] = new_file
        save_config(cfg)
        journal_lbl.config(text=f"Файл: {cfg.get('swabs_journal_xlsx', '')}")
        reload_journal()

    btn_journal = ttk.Button(top, text="Путь файла", command=change_journal_path)
    btn_journal.pack(side="right", padx=(6, 0))

    btn_refresh = ttk.Button(top, text="Обновить", command=lambda: reload_journal())
    btn_refresh.pack(side="right", padx=(6, 0))

    # ===== Основная область =====
    body = tk.Frame(root, bg="#f4f6f8")
    body.pack(expand=True, fill="both")

    # ===== ЛЕВАЯ ПАНЕЛЬ со вкладками =====
    left = tk.Frame(body, bg="#f4f6f8")
    left.pack(side="left", fill="y", padx=(0, 12))

    left_nb = ttk.Notebook(left)
    left_nb.pack(fill="both", expand=False)

    # --- Вкладка "По отделению" ---
    tab_by_dep = tk.Frame(left_nb, bg="#f4f6f8")
    left_nb.add(tab_by_dep, text="По отделению")

    tk.Label(tab_by_dep, text="Отделение", font=("Segoe UI", 10, "bold"), bg="#f4f6f8").pack(anchor="w")
    dep_var = tk.StringVar()
    dep_cb = ttk.Combobox(tab_by_dep, textvariable=dep_var, values=departments, state="readonly", width=26)
    dep_cb.pack(anchor="w", pady=(4, 6))

    # КНОПКА ОТЧЁТА (скрыта, пока нет выбранной даты)
    base_btn_color = "#2563eb"
    pulse_btn_color = "#60a5fa"

    btn_report = tk.Button(
        tab_by_dep,
        text="Сделать отчёт",
        font=("Segoe UI", 10, "bold"),
        bg=base_btn_color,
        fg="white",
        activebackground="#1d4ed8",
        activeforeground="white",
        relief="flat",
        padx=10,
        pady=6,
        cursor="hand2"
    )
    # НЕ pack сейчас — появится после выбора даты

    def _hex_to_rgb(hex_color: str):
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

    def _rgb_to_hex(rgb):
        return "#%02x%02x%02x" % rgb

    def _blend(c1: str, c2: str, t: float) -> str:
        r1, g1, b1 = _hex_to_rgb(c1)
        r2, g2, b2 = _hex_to_rgb(c2)
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        return _rgb_to_hex((r, g, b))

    def _animate_pulse(step: int, direction: int, steps: int = 12):
        if not btn_report.winfo_exists():
            return
        if not btn_report.winfo_ismapped():
            btn_report.after(15000, _pulse_report_button)
            return

        t = step / max(steps, 1)
        if direction < 0:
            t = 1.0 - t
        color = _blend(base_btn_color, pulse_btn_color, t)
        btn_report.configure(bg=color, activebackground=color)

        if step < steps:
            btn_report.after(50, _animate_pulse, step + 1, direction, steps)
        else:
            if direction > 0:
                btn_report.after(50, _animate_pulse, 0, -1, steps)

    def _pulse_report_button():
        if not btn_report.winfo_exists():
            return
        _animate_pulse(0, 1)
        btn_report.after(15000, _pulse_report_button)

    btn_report.after(15000, _pulse_report_button)

    tk.Label(tab_by_dep, text="Даты", font=("Segoe UI", 10, "bold"), bg="#f4f6f8").pack(anchor="w", pady=(10, 0))
    dates_list = tk.Listbox(tab_by_dep, width=18, height=20)
    dates_list.pack(anchor="w", pady=(4, 0))

    # --- Вкладка "По датам" ---
    tab_by_date = tk.Frame(left_nb, bg="#f4f6f8")
    left_nb.add(tab_by_date, text="По датам")

    tk.Label(tab_by_date, text="Даты", font=("Segoe UI", 10, "bold"), bg="#f4f6f8").pack(anchor="w")
    dates_all_list = tk.Listbox(tab_by_date, width=18, height=24)
    dates_all_list.pack(anchor="w", pady=(4, 0))

    # ===== ПРАВАЯ ЧАСТЬ =====
    right = tk.Frame(body, bg="#f4f6f8")
    right.pack(side="left", expand=True, fill="both")

    nb = ttk.Notebook(right)
    nb.pack(expand=True, fill="both")

    tree_by_sheet = {}

    # --- для "подсветки" вкладок наличием данных ---
    tab_by_sheet = {}
    base_tab_text = {}

    def make_tree(parent):
        frame = tk.Frame(parent, bg="#f4f6f8")
        frame.pack(expand=True, fill="both")
        tree = ttk.Treeview(frame, show="headings")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side="left", expand=True, fill="both")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        return tree

    for sh in SHEETS_DEFAULT:
        tab = tk.Frame(nb, bg="#f4f6f8")
        nb.add(tab, text=sh)
        tab_by_sheet[sh] = tab
        base_tab_text[sh] = sh
        tree_by_sheet[sh] = make_tree(tab)

    status = tk.Label(root, text="", bg="#f4f6f8", fg="#444", anchor="w")
    status.pack(fill="x", pady=(8, 0))

    journal = SwabsJournal(path=cfg.get("swabs_journal_xlsx", ""))

    # выбранная пара для отчёта
    selected_dep: str = ""
    selected_day: pd.Timestamp | None = None

    def set_status(msg: str):
        status.config(text=msg)

    def clear_tree(tree: ttk.Treeview):
        tree.delete(*tree.get_children())
        tree["columns"] = ()

    def fill_tree(tree: ttk.Treeview, df: pd.DataFrame):
        clear_tree(tree)
        if df is None or df.empty:
            return

        tree.tag_configure("eskape", background="#ffcccc")
        tree.tag_configure("detected", background="#fff2cc")

        cols = list(df.columns)
        tree["columns"] = cols

        for c in cols:
            title = str(c).strip().lower()
            w = 120 if "дата" in title else 160
            tree.heading(c, text=c)
            tree.column(c, width=w, anchor="w")

        culture_col = None
        for c in cols:
            if "выделенная культура" in str(c).strip().lower():
                culture_col = c
                break

        for _, row in df.iterrows():
            values = []
            for c in cols:
                v = row.get(c)
                if pd.isna(v):
                    values.append("")
                    continue
                if isinstance(v, pd.Timestamp):
                    values.append(v.strftime("%d.%m.%Y"))
                elif "дата" in str(c).lower():
                    dtv = pd.to_datetime(v, errors="coerce", dayfirst=True)
                    values.append(dtv.strftime("%d.%m.%Y") if not pd.isna(dtv) else str(v))
                else:
                    values.append(str(v))

            tag = None
            if culture_col:
                microbe = _norm_microbe(row.get(culture_col, ""))
                if microbe:
                    tag = "eskape" if _is_eskape(microbe) else "detected"

            tree.insert("", "end", values=values, tags=((tag,) if tag else ()))


    def update_tabs_presence(dep: str, day: pd.Timestamp | None):
        """
        Показываем индикатор 🟩 в названии вкладки, если в ней есть данные на выбранную дату.
        (В ttk.Notebook нельзя надёжно красить отдельные табы фоном на Windows,
         поэтому делаем индикатор в названии — работает стабильно.)
        """
        if not dep or day is None:
            for sh in SHEETS_DEFAULT:
                nb.tab(tab_by_sheet[sh], text=base_tab_text[sh])
            return

        for sh in SHEETS_DEFAULT:
            df = journal.filter_day(dep, day, sh)
            has_data = df is not None and not df.empty

            title = base_tab_text[sh]
            if has_data:
                title = f"{title}  🟩"
            nb.tab(tab_by_sheet[sh], text=title)


    # =============================
    # ✅ ВОССТАНОВЛЕНИЕ ALIASES.JSON
    # =============================
    def ensure_aliases_file():
        """
        Если aliases.json удалили — создаём пустые заново.
        """
        p = _aliases_path()
        if not p:
            return
        try:
            aliases = load_aliases(p)
            save_aliases(p, aliases)  # создаст файл, если его нет
            if webdav_url:
                from microbio_app import DATA_ROOT
                webdav_sync.upload_file(p, DATA_ROOT)
        except Exception as e:
            messagebox.showwarning("Привязки", f"Не удалось создать/прочитать файл привязок:\n{p}\n\n{e}")

    def build_mapping_with_dialog() -> dict:
        """
        mapping строим "как надо":
        - базово берём aliases.json
        - добираем точные совпадения (exact_norm)
        - спрашиваем только реально неизвестные (unknown_departments_for_mapping)
        """
        aliases_path = _aliases_path()
        aliases = load_aliases(aliases_path)
        mapping = dict(aliases)

        # добираем то, что можно сопоставить без окна (exact_norm)
        for raw in journal.unique_raw_departments():
            if raw in mapping:
                continue
            matched, _reason = try_match_department(raw, departments, aliases)
            if matched:
                mapping[raw] = matched

        unknown = journal.unknown_departments_for_mapping(mapping)
        if unknown:
            manual = ask_user_map_unknowns(main_frame, unknown, departments)
            if manual:
                manual = {k: v for k, v in manual.items() if isinstance(k, str) and isinstance(v, str)}
                aliases.update(manual)
                save_aliases(aliases_path, aliases)
                if webdav_url:
                    from microbio_app import DATA_ROOT
                    webdav_sync.upload_file(aliases_path, DATA_ROOT)
                mapping.update(manual)

        return mapping

    def hide_report_button():
        btn_report.pack_forget()

    def show_report_button():
        if not btn_report.winfo_ismapped():
            btn_report.pack(anchor="w", pady=(4, 0))

    def _show_wait_dialog(message: str):
        win = tk.Toplevel(main_frame)
        win.title("Подождите")
        win.transient(main_frame)
        win.grab_set()
        win.resizable(False, False)

        frame = tk.Frame(win, padx=20, pady=16)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=message, font=("Segoe UI", 10)).pack(pady=(0, 8))
        dots_var = tk.StringVar(value="...")
        dots_lbl = tk.Label(frame, textvariable=dots_var, font=("Segoe UI", 10))
        dots_lbl.pack(pady=(0, 8))

        pb = ttk.Progressbar(frame, mode="indeterminate", length=220)
        pb.pack()
        pb.start(10)

        win.update_idletasks()
        w = win.winfo_reqwidth()
        h = win.winfo_reqheight()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(0, int((sw - w) / 2))
        y = max(0, int((sh - h) / 2))
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.update()

        def animate_dots():
            cur = dots_var.get()
            dots_var.set("." if cur == "..." else cur + ".")
            if win.winfo_exists():
                win.after(350, animate_dots)

        animate_dots()
        win.configure(cursor="watch")

        return win, pb

    def do_make_report():
        nonlocal selected_dep, selected_day
        if not selected_dep or selected_day is None:
            return

        report_dir = _ensure_report_dir()
        if not report_dir:
            return
        template_path = _ensure_template_path()
        if not template_path:
            return

        df_swabs = journal.filter_day(selected_dep, selected_day, "Абиотические")
        df_air = journal.filter_day(selected_dep, selected_day, "Воздух")
        df_smears = journal.filter_day(selected_dep, selected_day, "Персонал")

        date_str = pd.Timestamp(selected_day).strftime("%d.%m.%Y")
        has_air = df_air is not None and not df_air.empty
        has_smears = df_smears is not None and not df_smears.empty

        file_name = f"Протокол результатов {selected_dep} {date_str}"
        if has_air:
            file_name += " + воздух"
        if has_smears:
            file_name += " + мазок"
        file_name += ".docx"

        temp_dir = Path.cwd() / "_tmp_reports"
        temp_dir.mkdir(parents=True, exist_ok=True)
        out_path = str(temp_dir / file_name)

        wait_win, wait_pb = _show_wait_dialog("Формируем отчет, пожалуйста подождите…")

        def _finish(ok, payload):
            try:
                wait_pb.stop()
                wait_win.destroy()
            except Exception:
                pass

            if not ok:
                messagebox.showerror("Ошибка", f"Не удалось сформировать отчет:\n{payload}")
                return

            saved_docx, auto_word_docx, auto_pdf = payload
            msg = "Отчет успешно сформирован и сохранен:\n\n"
            if auto_word_docx:
                msg += f"Word:\n{auto_word_docx}\n\n"
            else:
                msg += "Word:\n(не удалось сохранить в сетевую папку)\n\n"

            if auto_pdf:
                msg += f"PDF:\n{auto_pdf}"
            else:
                msg += "PDF:\n(не удалось сохранить в сетевую папку)"

            if webdav_url:
                try:
                    from microbio_app import DATA_ROOT
                    for p in (auto_word_docx, auto_pdf):
                        if p and os.path.abspath(p).startswith(os.path.abspath(DATA_ROOT)):
                            webdav_sync.upload_file(p, DATA_ROOT)
                except Exception:
                    pass

            messagebox.showinfo("Готово", msg)

        def _worker():
            try:
                import pythoncom  # type: ignore
                pythoncom.CoInitialize()
            except Exception:
                pythoncom = None
            try:
                result = build_docx_report(
                    template_path=template_path,
                    out_path=out_path,
                    dep=selected_dep,
                    day=selected_day,
                    df_swabs=df_swabs,
                    df_air=df_air,
                    df_smears=df_smears,
                    auto_word_dir=report_dir,
                    auto_pdf_dir=report_dir,
                )
            except Exception as e:
                main_frame.after(0, _finish, False, e)
                return
            finally:
                try:
                    if pythoncom is not None:
                        pythoncom.CoUninitialize()
                except Exception:
                    pass

            main_frame.after(0, _finish, True, result)

        import threading
        threading.Thread(target=_worker, daemon=True).start()

    btn_report.configure(command=do_make_report)

    def reload_journal():
        ensure_aliases_file()

        try:
            journal.path = cfg.get("swabs_journal_xlsx", "")
            journal.load()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать журнал:\n{e}")
            set_status("Ошибка чтения журнала")
            return

        try:
            mapping = build_mapping_with_dialog()
            journal.apply_department_mapping(mapping)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка привязки подразделений:\n{e}")

        set_status(f"Журнал загружен: {journal.path}")
        on_left_tab_changed()

    # ===== Режим "По отделению" =====
    def update_dates(event=None):
        nonlocal selected_dep, selected_day
        selected_dep = dep_var.get().strip()
        selected_day = None
        hide_report_button()

        dates_list.delete(0, tk.END)
        for sh in SHEETS_DEFAULT:
            clear_tree(tree_by_sheet[sh])

        update_tabs_presence(selected_dep, None)

        if not selected_dep:
            set_status("Выберите отделение")
            return

        dates = journal.list_dates_for_department(selected_dep)
        if not dates:
            set_status("Дат не найдено")
            return

        for d in dates:
            dates_list.insert(tk.END, d.strftime("%d.%m.%Y"))

        dates_list.selection_clear(0, tk.END)
        dates_list.selection_set(0)
        show_day()

    def show_day(event=None):
        nonlocal selected_dep, selected_day
        selected_dep = dep_var.get().strip()
        sel = dates_list.curselection()
        if not selected_dep or not sel:
            hide_report_button()
            return

        date_str = dates_list.get(sel[0])
        selected_day = pd.to_datetime(date_str, dayfirst=True).normalize()

        for sh in SHEETS_DEFAULT:
            df = journal.filter_day(selected_dep, selected_day, sh)
            fill_tree(tree_by_sheet[sh], df)

        update_tabs_presence(selected_dep, selected_day)

        set_status(f"{selected_dep} — {date_str}")
        show_report_button()

    # ===== Режим "По датам" =====
    def update_dates_all():
        hide_report_button()
        dates_all_list.delete(0, tk.END)
        for sh in SHEETS_DEFAULT:
            clear_tree(tree_by_sheet[sh])

        update_tabs_presence("", None)

        dates = journal.list_all_dates()
        if not dates:
            set_status("Дат не найдено")
            return

        for d in dates:
            dates_all_list.insert(tk.END, d.strftime("%d.%m.%Y"))

        dates_all_list.selection_clear(0, tk.END)
        dates_all_list.selection_set(0)
        show_day_all()

    def show_day_all(event=None):
        hide_report_button()
        sel = dates_all_list.curselection()
        if not sel:
            return

        date_str = dates_all_list.get(sel[0])
        day_ts = pd.to_datetime(date_str, dayfirst=True).normalize()

        for sh in SHEETS_DEFAULT:
            df = journal.filter_day_all(day_ts, sh)
            fill_tree(tree_by_sheet[sh], df)

        set_status(f"Все отделения — {date_str}")

    def on_left_tab_changed(event=None):
        tab_text = left_nb.tab(left_nb.select(), "text")
        if tab_text == "По датам":
            update_dates_all()
        else:
            if dep_var.get().strip():
                update_dates()
            else:
                hide_report_button()
                for sh in SHEETS_DEFAULT:
                    clear_tree(tree_by_sheet[sh])
                update_tabs_presence("", None)
                set_status("Выберите отделение")

    dep_cb.bind("<<ComboboxSelected>>", update_dates)
    dates_list.bind("<<ListboxSelect>>", show_day)
    dates_all_list.bind("<<ListboxSelect>>", show_day_all)
    left_nb.bind("<<NotebookTabChanged>>", on_left_tab_changed)

    reload_journal()
