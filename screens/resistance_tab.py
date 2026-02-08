import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from analysis.resistance import analyze_resistance
from utils.charts import barh_with_value_labels, reset_canvas


def build_resistance_tab(parent, report_state):
    container = tk.Frame(parent, bg="white")
    container.pack(expand=True, fill="both")

    # ---------- ЛЕВАЯ ПАНЕЛЬ ----------
    left = tk.Frame(container, bg="#f9fafb", width=380)
    left.pack(side="left", fill="y")
    left.pack_propagate(False)

    btn = ttk.Button(left, text="Загрузить Excel", style="Main.TButton")
    btn.pack(fill="x", padx=15, pady=(20, 10))

    text_box = tk.Text(
        left,
        font=("Segoe UI", 10),
        wrap="word"
    )
    text_box.pack(expand=True, fill="both", padx=15, pady=10)

    # ---------- ПРАВАЯ ПАНЕЛЬ ----------
    right = tk.Frame(container, bg="white")
    right.pack(side="right", expand=True, fill="both")

    canvas_holder = tk.Frame(right, bg="white")
    canvas_holder.pack(expand=True, fill="both")

    current_canvas = {"obj": None}

    def apply_result(result):
        microbes = sorted(
            result["microbes"],
            key=lambda x: x["r_percent"],
            reverse=True
        )

        antibiotics = sorted(
            result["antibiotics"],
            key=lambda x: x["r_percent"],
            reverse=True
        )

        text_box.delete(1.0, tk.END)

        text_box.insert(
            tk.END,
            "РЕЗИСТЕНТНОСТЬ ПО МИКРООРГАНИЗМАМ\n"
            "──────────────────────────────\n\n"
        )

        for m in microbes:
            text_box.insert(
                tk.END,
                f'{m["microbe"]:<30} '
                f'R {m["r_percent"]:>5}% '
                f'({m["r_count"]}/{m["total"]})\n'
            )

        text_box.insert(
            tk.END,
            "\nРЕЗИСТЕНТНОСТЬ ПО АНТИБИОТИКАМ\n"
            "──────────────────────────\n\n"
        )

        for a in antibiotics:
            text_box.insert(
                tk.END,
                f'{a["antibiotic"]:<30} '
                f'R {a["r_percent"]:>5}% '
                f'({a["r_count"]}/{a["total"]})\n'
            )

        if result["invalid"]:
            text_box.insert(
                tk.END,
                "\n⚠ НЕКОРРЕКТНЫЕ ДАННЫЕ\n"
                "────────────────────\n\n"
            )
            for i in result["invalid"]:
                text_box.insert(
                    tk.END,
                    f'{i["microbe"]} | {i["antibiotic"]} | {i["result"]}\n'
                )

        reset_canvas(current_canvas)

        # ДВА ГРАФИКА В ОДНОЙ FIGURE
        fig = Figure(figsize=(11, 6), dpi=100)

        # ===== 1. МИКРООРГАНИЗМЫ =====
        ax1 = fig.add_subplot(121)

        m_labels = [m["microbe"] for m in microbes]
        m_values = [float(m["r_percent"]) for m in microbes]

        m_colors = []
        for v in m_values:
            if v >= 20:
                m_colors.append("#E15759")
            elif v >= 10:
                m_colors.append("#F28E2B")
            else:
                m_colors.append("#59A14F")

        barh_with_value_labels(
            ax1,
            m_labels[::-1],
            m_values[::-1],
            title="Резистентность по микроорганизмам",
            xlabel="R (%)",
            colors=m_colors[::-1],
            value_fmt=lambda v: f"{v}%"
        )

        # ===== 2. АНТИБИОТИКИ =====
        ax2 = fig.add_subplot(122)

        a_labels = [a["antibiotic"] for a in antibiotics]
        a_values = [float(a["r_percent"]) for a in antibiotics]

        a_colors = []
        for v in a_values:
            if v >= 20:
                a_colors.append("#E15759")
            elif v >= 10:
                a_colors.append("#F28E2B")
            else:
                a_colors.append("#59A14F")

        barh_with_value_labels(
            ax2,
            a_labels[::-1],
            a_values[::-1],
            title="Резистентность по антибиотикам",
            xlabel="R (%)",
            colors=a_colors[::-1],
            value_fmt=lambda v: f"{v}%"
        )

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=canvas_holder)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both")
        current_canvas["obj"] = canvas

    def load_excel():
        file = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file:
            return

        try:
            result = analyze_resistance(file)
            report_state["resistance"] = result
            apply_result(result)
        except Exception as e:
            messagebox.showerror(
                "Ошибка анализа резистентности",
                str(e)
            )

    btn.config(command=load_excel)

    if report_state.get("resistance"):
        apply_result(report_state["resistance"])
