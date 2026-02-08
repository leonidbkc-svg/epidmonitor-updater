import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from analysis.loci import analyze_loci, LOCUS_COLORS
from utils.charts import barh_with_value_labels, reset_canvas, stacked_barh


def build_loci_tab(parent, report_state):
    container = tk.Frame(parent, bg="white")
    container.pack(expand=True, fill="both")

    # ---------- ЛЕВАЯ ПАНЕЛЬ ----------
    left = tk.Frame(container, bg="#f9fafb", width=380)
    left.pack(side="left", fill="y")
    left.pack_propagate(False)

    btn = ttk.Button(left, text="Загрузить Excel", style="Main.TButton")
    btn.pack(fill="x", padx=15, pady=(20, 10))

    # переключатель режима
    mode = tk.StringVar(value="groups")

    ttk.Radiobutton(
        left, text="По группам",
        variable=mode, value="groups"
    ).pack(anchor="w", padx=20)

    ttk.Radiobutton(
        left, text="Детализация (локусы)",
        variable=mode, value="detail"
    ).pack(anchor="w", padx=20, pady=(0, 10))

    text_box = tk.Text(left, font=("Segoe UI", 10), wrap="word")
    text_box.pack(expand=True, fill="both", padx=15, pady=10)

    # ---------- ПРАВАЯ ПАНЕЛЬ ----------
    right = tk.Frame(container, bg="white")
    right.pack(side="right", expand=True, fill="both")

    canvas_holder = tk.Frame(right, bg="white")
    canvas_holder.pack(expand=True, fill="both")

    current_canvas = {"obj": None}
    last_result = {"data": None}

    def sort_groups_for_plot(groups):
        return sorted(groups, key=lambda g: g["count"], reverse=True)

    def loci_sorted_by_total(groups):
        totals = {}
        for g in groups:
            for it in g.get("items", []):
                name = it["name"]
                totals[name] = totals.get(name, 0) + int(it["count"])

        return sorted(totals.keys(), key=lambda k: totals[k], reverse=True)

    def render(result):
        reset_canvas(current_canvas)

        fig = Figure(figsize=(11, 6), dpi=100)
        ax = fig.add_subplot(111)

        groups = sort_groups_for_plot(result["groups"])
        group_names = [g["group"] for g in groups]

        text_box.delete(1.0, tk.END)
        text_box.insert(
            tk.END,
            "БИОМАТЕРИАЛЫ\n"
            "────────────\n\n"
        )

        for g in groups:
            text_box.insert(
                tk.END,
                f'{g["group"]} — {g["count"]} ({g["percent"]}%)\n'
            )
            if mode.get() == "detail":
                for i in g.get("items", []):
                    text_box.insert(
                        tk.END,
                        f'  • {i["name"]} — {i["count"]} ({i["percent"]}%)\n'
                    )
            text_box.insert(tk.END, "\n")

        if mode.get() == "groups":
            values = [g["count"] for g in groups]

            barh_with_value_labels(
                ax,
                group_names[::-1],
                values[::-1],
                title="Биоматериалы по группам",
                xlabel="Количество"
            )

        else:
            loci = loci_sorted_by_total(groups)
            group_totals = [int(g["count"]) for g in groups]
            max_total = max(group_totals) if group_totals else 0
            groups_items = [g.get("items", []) for g in groups]
            stacked_barh(
                ax,
                group_names,
                loci,
                groups_items,
                colors_map=LOCUS_COLORS
            )

            for idx, total_val in enumerate(group_totals):
                ax.text(
                    total_val + (max_total * 0.01 + 0.5 if max_total else 0.5),
                    idx,
                    str(total_val),
                    va="center",
                    fontsize=9,
                    fontweight="bold"
                )

            ax.set_title("Структура биоматериалов по локусам (стековая диаграмма)")
            ax.set_xlabel("Количество")

            if max_total:
                ax.set_xlim(0, max_total * 1.20)

            ax.legend(
                title="Локус",
                fontsize=8,
                bbox_to_anchor=(1.02, 1),
                loc="upper left",
                borderaxespad=0
            )

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=canvas_holder)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both")
        current_canvas["obj"] = canvas

    def apply_result(result):
        last_result["data"] = result
        render(result)

    def load_excel():
        file = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file:
            return

        try:
            result = analyze_loci(file)
            report_state["loci"] = result
            apply_result(result)
        except Exception as e:
            messagebox.showerror("Ошибка анализа локусов", str(e))

    def on_mode_change(*_):
        if last_result["data"]:
            render(last_result["data"])

    mode.trace_add("write", on_mode_change)
    btn.config(command=load_excel)

    if report_state.get("loci"):
        apply_result(report_state["loci"])
