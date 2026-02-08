from __future__ import annotations

import json
import os
import re
from typing import Dict, List, Optional, Tuple


def _norm(s: str) -> str:
    """
    Нормализация для "точного" сопоставления без процентов:
    - upper
    - убираем пробелы/подчёркивания/слэши/дефисы/точки/скобки/табуляции и т.п.
    """
    s = str(s or "").strip().upper()
    return re.sub(r"[ \t/_\-.()]+", "", s)


def load_aliases(path: str) -> Dict[str, str]:
    """
    Загружает алиасы raw->canonical. Если файла нет/битый — возвращает {}.
    """
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}
        # гарантируем dict[str,str]
        out: Dict[str, str] = {}
        for k, v in (data or {}).items():
            if isinstance(k, str) and isinstance(v, str):
                out[k] = v
        return out
    except Exception:
        return {}


def save_aliases(path: str, aliases: Dict[str, str]) -> None:
    """
    Сохраняет алиасы. Если папки нет — создаёт.
    ВАЖНО: если файл удалили, он будет создан заново.
    """
    if not path:
        return

    dirpath = os.path.dirname(path)
    if dirpath:
        os.makedirs(dirpath, exist_ok=True)

    safe: Dict[str, str] = {}
    for k, v in (aliases or {}).items():
        if isinstance(k, str) and isinstance(v, str):
            safe[k] = v

    with open(path, "w", encoding="utf-8") as f:
        json.dump(safe, f, ensure_ascii=False, indent=2)


def try_match_department(raw_dep: str, departments: List[str], aliases: Dict[str, str]) -> Tuple[Optional[str], str]:
    """
    Безопасное сопоставление:
    - алиас по исходной строке (точно) -> alias_exact
    - или точное совпадение после нормализации -> exact_norm
    Никаких "процентов похожести" и авто-угадываний.
    """
    raw = str(raw_dep or "").strip()
    if not raw:
        return None, "empty"

    # 1) алиас по исходной строке (как в Excel)
    if raw in (aliases or {}):
        v = aliases.get(raw)
        if isinstance(v, str) and v.strip():
            return v, "alias_exact"

    # 2) точное совпадение после нормализации
    raw_n = _norm(raw)
    for d in (departments or []):
        if raw_n == _norm(d):
            return d, "exact_norm"

    return None, "unknown"
