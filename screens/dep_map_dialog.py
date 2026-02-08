from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Optional


def ask_user_map_unknowns(parent, unknown: List[str], departments: List[str], preset: Optional[Dict[str, str]] = None) -> Optional[Dict[str, str]]:
    """
    Показывает окно сопоставления неизвестных значений (raw из Excel) с каноническими отделениями.

    ВАЖНО: возвращает строго Dict[str, str]
      { raw_value: selected_department }
    Если пользователь нажал "Отмена" -> None
    """

    unknown = [u for u in (unknown or []) if str(u).strip()]
    if not unknown:
        return {}

    top = tk.Toplevel(parent)
    top.title("Неизвестные подразделения")
    top.transient(parent)
    top.grab_set()

    top.geometry("900x520")

    info = tk.Label(
        top,
        text=(
            "Система не смогла однозначно сопоставить значения из столбца 'Подразделение'.\n"
            "Выберите отделение для каждого значения и нажмите «Сохранить»."
        ),
        justify="left"
    )
    info.pack(anchor="w", padx=12, pady=(10, 8))

    # Поиск
    search_var = tk.StringVar()
    search_row = tk.Frame(top)
    search_row.pack(fill="x", padx=12, pady=(0, 8))
    tk.Label(search_row, text="Поиск:").pack(side="left")
    search_ent = ttk.Entry(search_row, textvariable=search_var, width=40)
    search_ent.pack(side="left", padx=(6, 0), fill="x", expand=True)

    # Скролл-контейнер
    outer = tk.Frame(top)
    outer.pack(expand=True, fill="both", padx=12, pady=(0, 10))

    canvas = tk.Canvas(outer, highlightthickness=0)
    vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)

    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", expand=True, fill="both")

    inner = tk.Frame(canvas)
    win = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_inner_config(_evt=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _on_canvas_config(evt):
        canvas.itemconfigure(win, width=evt.width)

    inner.bind("<Configure>", _on_inner_config)
    canvas.bind("<Configure>", _on_canvas_config)

    # Список строк: raw -> combobox
    # Храним переменные отдельно, чтобы собрать результат
    rows: List[tuple[str, tk.StringVar, ttk.Combobox, tk.Label]] = []

    # Заголовок таблицы
    hdr = tk.Frame(inner)
    hdr.pack(fill="x", pady=(0, 6))
    tk.Label(hdr, text="Значение из Excel", width=50, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
    tk.Label(hdr, text="Отделение", width=40, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(10, 0))

    preset = preset or {}

    for raw in unknown:
        r = tk.Frame(inner)
        r.pack(fill="x", pady=2)

        raw_lbl = tk.Label(r, text=str(raw), width=50, anchor="w")
        raw_lbl.pack(side="left")

        preset_val = preset.get(str(raw))
        if preset_val in departments:
            var = tk.StringVar(value=preset_val)
        else:
            var = tk.StringVar(value="")
        cb = ttk.Combobox(r, textvariable=var, values=departments, state="readonly", width=38)
        cb.pack(side="left", padx=(10, 0))

        rows.append((str(raw), var, cb, raw_lbl))

    # Фильтр по поиску
    def apply_filter(*_):
        q = (search_var.get() or "").strip().lower()
        for raw, var, cb, raw_lbl in rows:
            show = (q in raw.lower()) if q else True
            cb.master.pack_forget()
            if show:
                cb.master.pack(fill="x", pady=2)

    search_var.trace_add("write", apply_filter)

    # Кнопки
    btns = tk.Frame(top)
    btns.pack(fill="x", padx=12, pady=(0, 12))

    result: Dict[str, str] = {}
    cancelled = {"flag": False}

    def on_cancel():
        cancelled["flag"] = True
        top.destroy()

    def on_save():
        out: Dict[str, str] = {}
        for raw, var, _cb, _lbl in rows:
            val = (var.get() or "").strip()
            if val:
                # ВАЖНО: только строка
                out[str(raw)] = str(val)
        result.update(out)
        top.destroy()

    ttk.Button(btns, text="Отмена", command=on_cancel).pack(side="right")
    ttk.Button(btns, text="Сохранить", command=on_save).pack(side="right", padx=(0, 8))

    top.wait_window()

    if cancelled["flag"]:
        return None
    return result
