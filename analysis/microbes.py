import pandas as pd
from utils.gram import classify_gram


def analyze_microbes(excel_path: str) -> dict:
    """
    Анализ микроорганизмов.
    Возвращает чистую структуру данных без GUI.
    """

    df = pd.read_excel(excel_path)

    required_cols = ["Обнар. микроорг.", "COUNT(*)"]
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"В Excel отсутствует колонка: {c}")

    df = df[required_cols]
    df.columns = ["microbe", "count"]
    df = df.sort_values("count", ascending=False)

    total = int(df["count"].sum())

    microbes = []
    gram_counter = {}
    unclassified = []

    for _, row in df.iterrows():
        microbe = str(row["microbe"])
        count = int(row["count"])
        gram = classify_gram(microbe)
        percent = round(count / total * 100, 1) if total else 0

        microbes.append({
            "microbe": microbe,
            "count": count,
            "percent": percent,
            "gram": gram
        })

        gram_counter.setdefault(gram, 0)
        gram_counter[gram] += count

        if gram == "Не классифицировано":
            unclassified.append({
                "microbe": microbe,
                "count": count,
                "percent": percent
            })

    gram_summary = []
    for gram, count in gram_counter.items():
        gram_summary.append({
            "gram": gram,
            "count": count,
            "percent": round(count / total * 100, 1) if total else 0
        })

    return {
        "total": total,
        "microbes": microbes,
        "gram_summary": gram_summary,
        "unclassified": unclassified
    }
