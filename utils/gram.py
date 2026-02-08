def classify_gram(microbe: str) -> str:
    m = str(microbe).lower()

    gram_positive = [
        "staphylococcus", "streptococcus", "enterococcus",
        "bacillus", "lactobacillus", "lactiplantibacillus",
        "gemella", "rothia", "micrococcus", "corynebacterium"
    ]

    gram_negative = [
        "escherichia", "klebsiella", "enterobacter", "serratia",
        "proteus", "morganella", "pseudomonas",
        "acinetobacter", "haemophilus", "neisseria"
    ]

    fungi = ["candida", "malassezia"]

    if any(k in m for k in gram_positive):
        return "Грамположительные"
    if any(k in m for k in gram_negative):
        return "Грамотрицательные"
    if any(k in m for k in fungi):
        return "Грибы"

    return "Не классифицировано"
