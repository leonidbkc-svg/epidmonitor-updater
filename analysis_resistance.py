import pandas as pd
from matplotlib.figure import Figure
import mplcursors


def analyze_resistance(file_path: str, output_func):
    df = pd.read_excel(file_path)

    # ===== АВТООПРЕДЕЛЕНИЕ КОЛОНОК =====
    cols = df.columns

    microbe_col = None
    result_col = None
    count_col = None
    antibiotic_col = None

    for c in cols:
        if df[c].astype(str).str.contains(
            "staphylococcus|klebsiella|enterococcus|pseudomonas|escherichia|streptococcus",
            case=False, na=False
        ).any():
            microbe_col = c

        if set(df[c].astype(str).str.upper()) & {"R", "S", "I"}:
            result_col = c

        if df[c].dtype in ["int64", "float64"]:
            count_col = c

        if df[c].astype(str).str.contains("линезолид|меропенем|ванкомицин", case=False, na=False).any():
            antibiotic_col = c

    if not all([microbe_col, result_col, count_col]):
        raise ValueError("Не удалось автоматически определить колонки")

    df = df[[microbe_col, result_col, count_col]].copy()
    df.columns = ["microbe", "result", "count"]

    df["result"] = df["result"].astype(str).str.upper()

    # ===== АГРЕГАЦИЯ =====
    stats = []

    for microbe, sub in df.groupby("microbe"):
        total = sub["count"].sum()
        r_cnt = sub[sub["result"] == "R"]["count"].sum()
        r_pct = (r_cnt / total * 100) if total else 0

        stats.append({
            "microbe": microbe,
            "R_count": r_cnt,
            "R_pct": round(r_pct, 1)
        })

    res_df = pd.DataFrame(stats).sort_values("R_pct", ascending=False)

    # ===== TEXT =====
    output_func("\n📌 РЕЗИСТЕНТНОСТЬ ПО МИКРООРГАНИЗМАМ:\n")
    for _, r in res_df.iterrows():
        output_func(f"{r['microbe']} — R: {r['R_count']} ({r['R_pct']}%)")

    # ===== GRAPH =====
    fig = Figure(figsize=(9, 5), dpi=100)
    ax = fig.add_subplot(111)

    bars = ax.bar(res_df["microbe"], res_df["R_pct"])
    ax.set_title("Резистентность по микроорганизмам (R %)")
    ax.set_ylabel("R (%)")
    ax.set_ylim(0, max(res_df["R_pct"]) * 1.3 if len(res_df) else 1)
    ax.grid(axis="y", alpha=0.3)
    ax.set_xticks(range(len(res_df)))
    ax.set_xticklabels(res_df["microbe"], rotation=75, ha="right")

    for bar, pct in zip(bars, res_df["R_pct"]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 1,
            f"{pct}%",
            ha="center",
            fontsize=9,
            fontweight="bold"
        )

    # ===== HOVER =====
    cursor = mplcursors.cursor(bars, hover=True)

    @cursor.connect("add")
    def _(sel):
        idx = sel.index
        row = res_df.iloc[idx]
        sel.annotation.set_text(
            f"{row['microbe']}\nR: {row['R_count']}\nR%: {row['R_pct']}"
        )

    fig.tight_layout()
    return fig
