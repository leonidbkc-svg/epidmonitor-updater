# analysis/swabs.py
import pandas as pd


def analyze_swabs(filepath):
    df = pd.read_excel(filepath, sheet_name=0)

    REQUIRED_COLUMNS = [
        "Номер образца",
        "Место отбора образца",
        "БГКП",
        "SA",
        "Pseudomonas",
        "Примечания"
    ]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"В файле отсутствуют столбцы: {', '.join(missing)}"
        )

    df = df[REQUIRED_COLUMNS]

    samples = {}

    for _, row in df.iterrows():
        sample = row["Номер образца"]
        if pd.isna(sample):
            continue

        sample = str(sample).strip()
        place = str(row["Место отбора образца"]).strip()

        if sample not in samples:
            samples[sample] = {
                "place": place,
                "findings": [],
                "positive": False
            }

        found_in_main = False

        # ---------- ОСНОВНЫЕ СТОЛБЦЫ ----------
        for col in ["БГКП", "SA", "Pseudomonas"]:
            val = row[col]
            if pd.isna(val) or str(val).strip() in ("", "-"):
                continue

            samples[sample]["positive"] = True
            found_in_main = True

            samples[sample]["findings"].append({
                "source": "main",
                "value": str(val).strip()
            })

        # ---------- ПРИМЕЧАНИЯ ----------
        # используем ТОЛЬКО если в основных столбцах пусто
        note = row["Примечания"]
        if not found_in_main:
            if not pd.isna(note) and str(note).strip() not in ("", "-"):
                samples[sample]["positive"] = True
                samples[sample]["findings"].append({
                    "source": "note",
                    "value": str(note).strip()
                })

    total = len(samples)
    positive_samples = {k: v for k, v in samples.items() if v["positive"]}

    positive = len(positive_samples)
    negative = total - positive
    percent = round((positive / total * 100), 1) if total else 0

    details = []
    for sample, data in positive_samples.items():
        for f in data["findings"]:
            details.append({
                "sample": sample,
                "place": data["place"],
                "source": f["source"],   # main / note
                "value": f["value"]
            })

    return {
        "total": total,
        "positive": positive,
        "negative": negative,
        "percent": percent,
        "details": details
    }
