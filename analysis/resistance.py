import pandas as pd


def analyze_resistance(file_path: str) -> dict:
    df = pd.read_excel(file_path)
    cols = df.columns

    # ---------- ПОИСК КОЛОНОК ----------
    antibiotic_col = None
    for c in cols:
        if "линезолид" in " ".join(df[c].astype(str).str.lower()):
            antibiotic_col = c
            break
    if antibiotic_col is None:
        antibiotic_col = cols[2]

    result_col = None
    for c in cols:
        if set(df[c].astype(str).str.upper()) & {"R", "S", "I"}:
            result_col = c
            break
    if result_col is None:
        raise ValueError("Не найдена колонка с результатами (R/S/I)")

    microbe_col = None
    for c in cols:
        if df[c].astype(str).str.contains(
            "staphylococcus|klebsiella|enterococcus|pseudomonas|bacillus|streptococcus",
            case=False,
            na=False
        ).any():
            microbe_col = c
            break
    if microbe_col is None:
        raise ValueError("Не найдена колонка с микроорганизмами")

    count_col = None
    for c in cols:
        if df[c].dtype in ["int64", "float64"]:
            count_col = c
            break
    if count_col is None:
        raise ValueError("Не найдена колонка с количеством")

    # ---------- ОЧИСТКА ----------
    df = df[[microbe_col, antibiotic_col, result_col, count_col]].copy()
    df.columns = ["microbe", "antibiotic", "result", "count"]
    df["result"] = df["result"].astype(str).str.upper()

    # ---------- ПО МИКРООРГАНИЗМАМ ----------
    microbes = []
    for microbe, sub in df.groupby("microbe"):
        total = sub["count"].sum()
        r_count = sub[sub["result"] == "R"]["count"].sum()
        r_percent = round((r_count / total * 100), 1) if total else 0

        microbes.append({
            "microbe": microbe,
            "r_percent": r_percent,
            "r_count": int(r_count),
            "total": int(total)
        })

    microbes.sort(key=lambda x: x["r_percent"], reverse=True)

    # ---------- ПО АНТИБИОТИКАМ ----------
    antibiotics = []
    for ab, sub in df.groupby("antibiotic"):
        total = sub["count"].sum()
        r_count = sub[sub["result"] == "R"]["count"].sum()
        r_percent = round((r_count / total * 100), 1) if total else 0

        antibiotics.append({
            "antibiotic": ab,
            "r_percent": r_percent,
            "r_count": int(r_count),
            "total": int(total)
        })

    antibiotics.sort(key=lambda x: x["r_percent"], reverse=True)

    return {
        "microbes": microbes,
        "antibiotics": antibiotics,
        "invalid": []
    }
