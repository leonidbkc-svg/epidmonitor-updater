import pandas as pd
from matplotlib.figure import Figure


def classify_gram(microbe: str) -> str:
    m = str(microbe).lower()

    gram_positive = [
        "staphylococcus", "streptococcus", "enterococcus",
        "bacillus", "lactobacillus", "lactiplantibacillus",
        "gemella", "rothia", "micrococcus", "corynebacterium"
    ]

    gram_negative = [
        "burkholderia", "klebsiella", "escherichia", "acinetobacter",
        "enterobacter", "serratia", "proteus", "morganella",
        "pseudomonas", "haemophilus", "gardnerella"
    ]

    fungi = ["candida", "malassezia"]

    if any(k in m for k in gram_positive):
        return "Грамположительные"
    if any(k in m for k in gram_negative):
        return "Грамотрицательные"
    if any(k in m for k in fungi):
        return "Грибы"

    return "Не классифицировано"


def analyze_microbes(file_path: str, output_func):
    df = pd.read_excel(file_path)

    df = df[['Обнар. микроорг.', 'COUNT(*)']]
    df.columns = ['microbe', 'count']
    df = df.sort_values("count", ascending=False)

    total = df["count"].sum()

    # ===== ТЕКСТ =====
    output_func("\n📊 РАНЖИРОВАННОЕ КОЛИЧЕСТВО:\n")
    for _, row in df.iterrows():
        p = row["count"] / total * 100
        output_func(f"{row['microbe']} — {row['count']} ({p:.1f}%)")

    # ===== ГРАМ =====
    df["Gram"] = df["microbe"].apply(classify_gram)
    gram_summary = df.groupby("Gram")["count"].sum()

    output_func("\n📌 СОСТАВ ПО ГРАМ-ОКРАСКЕ:\n")
    for g, c in gram_summary.items():
        output_func(f"{g}: {c} ({c/total*100:.1f}%)")

    return df, gram_summary
