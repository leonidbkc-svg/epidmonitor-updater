import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from analysis.microbes import analyze_microbes
from utils.charts import bar_with_value_labels, barh_with_value_labels, reset_canvas


def build_microbes_tab(parent, report_state, normalize_gram):
    container = tk.Frame(parent, bg="white")
    container.pack(expand=True, fill="both")

    # ---------- ЛЕВАЯ ПАНЕЛЬ ----------
    left = tk.Frame(container, bg="#f9fafb", width=380)
    left.pack(side="left", fill="y")
    left.pack_propagate(False)

    btn = ttk.Button(left, text="Загрузить Excel", style="Main.TButton")
    btn.pack(fill="x", padx=15, pady=(20, 10))

    text_box = tk.Text(left, font=("Segoe UI", 10))
    text_box.pack(expand=True, fill="both", padx=15, pady=10)

    # место под кнопку классификации (чтобы не плодилась)
    gram_btn_holder = tk.Frame(left, bg="#f9fafb")
    gram_btn_holder.pack(fill="x", padx=15, pady=(0, 10))

    # ---------- ПРАВАЯ ПАНЕЛЬ ----------
    right = tk.Frame(container, bg="white")
    right.pack(side="right", expand=True, fill="both")

    canvas_holder = tk.Frame(right, bg="white")
    canvas_holder.pack(expand=True, fill="both")

    current_canvas = {"obj": None}

    # ручные правки: { "Микроорганизм": "Gram+" | "Gram-" | "Грибы" }
    manual_gram = {}

    last_microbes = None
    last_result = None

    # --------------------------------------------------
    # ОКНО РУЧНОЙ КЛАССИФИКАЦИИ GRAM
    # --------------------------------------------------
    def open_gram_fix_window(items):
        win = tk.Toplevel(parent)
        win.title("Классификация Gram")
        win.geometry("520x560")
        win.minsize(480, 420)
        win.transient(parent)
        win.grab_set()

        tk.Label(
            win,
            text="Выберите категорию для микроорганизмов",
            font=("Segoe UI", 11, "bold")
        ).pack(pady=(12, 6))

        # ---------- ПРОКРУЧИВАЕМАЯ ОБЛАСТЬ ----------
        scroll_container = tk.Frame(win)
        scroll_container.pack(expand=True, fill="both", padx=12, pady=(0, 10))

        canvas = tk.Canvas(scroll_container, highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)

        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor="nw")

        def _on_configure(_):
            canvas.configure(scrollregion=canvas.bbox("all"))

        frame.bind("<Configure>", _on_configure)

        # колесо мыши
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        vars_map = {}

        # по умолчанию ставим то, что уже есть (ручное или авто)
        def _current_gram_for(name: str) -> str:
            if report_state.get("microbes"):
                for mm in report_state["microbes"].get("microbes", []):
                    if mm.get("microbe") == name:
                        g = (mm.get("gram") or "").strip()
                        if g in ("Gram+", "Gram-", "Грибы", "Не классифицировано"):
                            return g
                        return normalize_gram(g)
            return "Gram+"

        for name in items:
            row = tk.Frame(frame)
            row.pack(fill="x", padx=6, pady=4)

            tk.Label(row, text=name, anchor="w").pack(side="left", expand=True, fill="x")

            current_value = manual_gram.get(name, _current_gram_for(name))
            var = tk.StringVar(value=current_value)
            vars_map[name] = var

            ttk.Combobox(
                row,
                textvariable=var,
                values=["Gram+", "Gram-", "Грибы"],
                state="readonly",
                width=10
            ).pack(side="right")

        def apply():
            # сохраняем только выбранные (остальные не трогаем)
            for name, var in vars_map.items():
                manual_gram[name] = var.get()

            # синхронизация в report_state — МЕНЯЕМ ТОЛЬКО РУЧНЫЕ
            if report_state.get("microbes") and report_state["microbes"].get("microbes"):
                for mm in report_state["microbes"]["microbes"]:
                    nm = mm.get("microbe")
                    if nm in manual_gram:
                        mm["gram"] = manual_gram[nm]
                        mm["gram_source"] = "manual"

            win.destroy()
            render()

        ttk.Button(
            win,
            text="Применить",
            style="Main.TButton",
            command=apply
        ).pack(fill="x", padx=20, pady=(0, 14))

        def _on_close():
            canvas.unbind_all("<MouseWheel>")
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _on_close)

    # --------------------------------------------------
    # ОТРИСОВКА ТЕКСТА + ГРАФИКОВ
    # --------------------------------------------------
    def render():
        nonlocal last_microbes, last_result

        if not last_microbes or not last_result:
            return

        microbes = last_microbes
        result = last_result

        # очистка текста
        text_box.delete(1.0, tk.END)

        # убираем старую кнопку классификации (чтобы не плодилась)
        for w in gram_btn_holder.winfo_children():
            w.destroy()

        # ---------- ТЕКСТ: микроорганизмы ----------
        text_box.insert(
            tk.END,
            f"МИКРООРГАНИЗМЫ (всего: {result['total']})\n\n"
        )
        for i, m in enumerate(microbes, start=1):
            text_box.insert(
                tk.END,
                f"{i}. {m['microbe']} — {m['count']} ({m['percent']}%)\n"
            )

        # ---------- GRAM: расчёт для UI ----------
        text_box.insert(tk.END, "\nРАСПРЕДЕЛЕНИЕ ПО GRAM\n\n")

        gram_map = {
            "Gram+": {"count": 0, "items": []},
            "Gram-": {"count": 0, "items": []},
            "Грибы": {"count": 0, "items": []},
            "Не классифицировано": {"count": 0, "items": []},
        }

        for m in microbes:
            name = m["microbe"]

            # ВАЖНО:
            # 1) если есть ручная — берём её
            # 2) иначе берём уже ЗАФИКСИРОВАННЫЙ (авто) gram из данных (мы его проставляем в load_excel)
            gram = manual_gram.get(name, m.get("gram", "Не классифицировано"))

            if gram not in gram_map:
                gram = "Не классифицировано"

            gram_map[gram]["count"] += int(m.get("count", 0))
            gram_map[gram]["items"].append(name)

        total = int(result.get("total", 0)) or 0
        for gram, data in gram_map.items():
            if data["count"] > 0 and total > 0:
                percent = round(data["count"] / total * 100, 1)
                text_box.insert(tk.END, f"{gram}: {data['count']} ({percent}%)\n")

        # список неклассифицированных + кнопка
        if gram_map["Не классифицировано"]["items"]:
            text_box.insert(tk.END, "\nНЕ КЛАССИФИЦИРОВАННЫЕ:\n")
            for name in gram_map["Не классифицировано"]["items"]:
                text_box.insert(tk.END, f"  {name}\n")

            ttk.Button(
                gram_btn_holder,
                text="Классифицировать Gram",
                style="Secondary.TButton",
                command=lambda items=gram_map["Не классифицировано"]["items"][:]: open_gram_fix_window(items)
            ).pack(fill="x")

        # ---------- ГРАФИКИ ----------
        reset_canvas(current_canvas)

        fig = Figure(figsize=(11, 6), dpi=100)
        ax1 = fig.add_subplot(121)
        ax2 = fig.add_subplot(122)

        # --- МИКРООРГАНИЗМЫ ---
        labels = [m["microbe"] for m in microbes][::-1]
        values = [int(m.get("count", 0)) for m in microbes][::-1]

        barh_with_value_labels(
            ax1,
            labels,
            values,
            title="Микроорганизмы (количество)",
            xlabel="Количество"
        )

        # --- GRAM ---
        g_labels = [k for k, v in gram_map.items() if v["count"] > 0]
        g_values = [gram_map[k]["count"] for k in g_labels]

        color_map = {
            "Gram+": "#4E79A7",
            "Gram-": "#E15759",
            "Грибы": "#59A14F",
            "Не классифицировано": "#9D9D9D",
        }

        bar_with_value_labels(
            ax2,
            g_labels,
            g_values,
            title="Распределение по Gram",
            colors=[color_map.get(k, "#cccccc") for k in g_labels]
        )

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=canvas_holder)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both")
        current_canvas["obj"] = canvas

    def apply_result(result):
        nonlocal last_microbes, last_result

        manual_gram.clear()
        for m in result.get("microbes", []):
            if m.get("gram_source") == "manual":
                manual_gram[m.get("microbe")] = m.get("gram", "Gram+")
            else:
                from utils.gram import classify_gram
                gram_raw = (m.get("gram") or "").strip()
                if gram_raw in ("Gram+", "Gram-", "Грибы", "Не классифицировано"):
                    m["gram"] = gram_raw
                else:
                    m["gram"] = normalize_gram(classify_gram(m.get("microbe", "")))
                m["gram_source"] = m.get("gram_source") or "auto"

        microbes = sorted(
            result["microbes"],
            key=lambda x: x["count"],
            reverse=True
        )

        last_microbes = microbes
        last_result = result
        render()

    # --------------------------------------------------
    # ЗАГРУЗКА EXCEL
    # --------------------------------------------------
    def load_excel():
        nonlocal last_microbes, last_result

        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file:
            return

        try:
            result = analyze_microbes(file)
            report_state["microbes"] = result
            apply_result(result)
        except Exception as e:
            messagebox.showerror("Ошибка анализа микроорганизмов", str(e))

    btn.config(command=load_excel)

    if report_state.get("microbes"):
        apply_result(report_state["microbes"])
