import pandas as pd
from matplotlib.figure import Figure
import mplcursors


GROUPS = {
    "Стерильные": [
        "Кровь венозная",
        "Дистальный конец ЦВК",
        "Жидкость амниотическая",
        "Аутопсийный материал кровь",
        "Аутопсийный материал легкое",
        "Аутопсийный материал печень",
        "Молоко грудное",
        "Эякулят"
    ],
    "Нестерильные": [
        "Аспират эндотрахеальный",
        "Аспират трахеобронхиальный",
        "Аспират из полости матки",
        "Отделяемое раны",
        "Отделяемое наружного уха",
        "Отделяемое слизистой уретры",
        "Отделяемое слизистой цервикального канала",
        "Мазок вагинальный",
        "Мазок вагино-ректальный",
        "Мазок слизистой миндалин",
        "Мазок конъюнктивы",
        "Мокрота",
        "Кал",
        "Моча",
        "Аутопсийный материал содержимое кишечника"
    ],
    "Скрининговые": [
        "Мазок ректальный",
        "Мазок слизистой ротоглотки и носоглотки",
        "Мазок со слизистой ротоглотки и носоглотки",
        "Отделяемое слизистой носа"
    ]
}


def classify_locus(locus: str) -> str:
    for group, items in GROUPS.items():
        if locus in items:
            return group
    return "Не определено"


def analyze_locus(file_path: str, output_func):
    df = pd.read_excel(file_path)

    if "Локус" not in df.columns or "COUNT(*)" not in df.columns:
        raise ValueError("Excel должен содержать 'Локус' и 'COUNT(*)'")

    df = df[["Локус", "COUNT(*)"]]
    df["Локус"] = df["Локус"].astype(str).str.strip()
    df = df[df["Локус"] != "Не указано"]

    df["Группа"] = df["Локус"].apply(classify_locus)

    pivot = df.pivot_table(
        index="Группа",
        columns="Локус",
        values="COUNT(*)",
        aggfunc="sum",
        fill_value=0
    )

    # ===== TEXT =====
    output_func("\n📌 СВОДКА ПО ЛОКУСАМ:\n")

    for group in pivot.index:
        total = pivot.loc[group].sum()
        output_func(f"▶ {group}: {total}")
        for locus, val in pivot.loc[group].items():
            if val > 0:
                output_func(f"   • {locus}: {val}")
        output_func("")

    # ===== GRAPH =====
    fig = Figure(figsize=(9, 5), dpi=100)
    ax = fig.add_subplot(111)

    bottom = [0] * len(pivot.index)
    bars = []

    for locus in pivot.columns:
        values = pivot[locus].values
        bar = ax.bar(pivot.index, values, bottom=bottom, label=locus)
        bars.extend(bar)
        bottom = [b + v for b, v in zip(bottom, values)]

    ax.set_title("Распределение локусов по группам")
    ax.set_ylabel("COUNT(*)")
    ax.legend(fontsize=8, bbox_to_anchor=(1.02, 1), loc="upper left")

    # ===== HOVER =====
    cursor = mplcursors.cursor(bars, hover=True)

    @cursor.connect("add")
    def on_add(sel):
        bar = sel.artist
        height = bar.get_height()
        label = bar.get_label()
        sel.annotation.set_text(f"{label}\nКоличество: {int(height)}")

    fig.tight_layout()
    return fig
