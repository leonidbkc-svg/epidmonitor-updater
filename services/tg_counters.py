# services/tg_counters.py
import json
import os
from typing import Tuple, Optional


def _status_file_path() -> str:
    """
    Файл, который будет обновлять бот.
    Храним рядом с DATA_ROOT, чтобы всем было доступно и без плясок.
    """
    from microbio_app import DATA_ROOT  # импорт внутри, чтобы не ловить циклы
    return os.path.join(DATA_ROOT, "tg_status.json")


def read_latest_numbers() -> Tuple[Optional[int], Optional[int], Optional[int]]:
    """
    Возвращает:
      (last_swab_number, last_air_number, last_stroke_number)

    Если данных нет/файл не создан — вернёт (None, None, None).
    Формат JSON:
      {
        "last_swab": 123,
        "last_air": 45,
        "last_stroke": 67
      }
    """
    path = _status_file_path()
    if not os.path.exists(path):
        return None, None, None

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}

        last_swab = data.get("last_swab", None)
        last_air = data.get("last_air", None)
        last_stroke = data.get("last_stroke", None)

        last_swab = int(last_swab) if last_swab is not None else None
        last_air = int(last_air) if last_air is not None else None
        last_stroke = int(last_stroke) if last_stroke is not None else None

        return last_swab, last_air, last_stroke
    except Exception:
        return None, None, None
