# screens/ordinators_test.py
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict

from data.question_bank import get_ordinators_3686_questions


def open_ordinators_test_screen(
    main_frame,
    build_header,
    go_back_callback,
):
    """
    Экран тестирования для ординаторов:
      - показывает вопросы по одному
      - single: радиокнопки
      - multi: чекбоксы
      - подсчёт результата
    """
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame, back_callback=go_back_callback)

    questions = get_ordinators_3686_questions()
    if not questions:
        messagebox.showwarning("Тестирование", "Вопросы не найдены.")
        go_back_callback()
        return

    # --- верх ---
    tk.Label(
        main_frame,
        text="🧪 Тестирование (Ординаторы) — ИСМП / СанПиН 3.3686",
        font=("Segoe UI", 16, "bold"),
        bg="#f4f6f8",
        fg="#111827",
        wraplength=980,
        justify="left",
    ).pack(pady=(18, 10), padx=18, anchor="w")

    body = tk.Frame(main_frame, bg="#f4f6f8")
    body.pack(expand=True, fill="both", padx=18, pady=(0, 18))

    card = tk.Frame(body, bg="white", bd=1, relief="solid")
    card.pack(expand=True, fill="both")

    # состояние теста
    idx_var = tk.IntVar(value=0)
    chosen_single = tk.StringVar(value="")           # для single
    chosen_multi: Dict[str, tk.BooleanVar] = {}      # для multi
    user_answers: Dict[str, List[str]] = {}          # qid -> ["A","C"]

    # UI элементы
    progress_lbl = tk.Label(card, bg="white", fg="#6b7280", font=("Segoe UI", 10))
    progress_lbl.pack(anchor="w", padx=16, pady=(12, 4))

    q_lbl = tk.Label(
        card,
        text="",
        font=("Segoe UI", 13, "bold"),
        bg="white",
        fg="#111827",
        wraplength=960,
        justify="left",
    )
    q_lbl.pack(anchor="w", padx=16, pady=(0, 10))

    options_box = tk.Frame(card, bg="white")
    options_box.pack(anchor="w", padx=16, pady=(0, 12), fill="x")

    sep = ttk.Separator(card)
    sep.pack(fill="x", padx=16, pady=(6, 10))

    bottom = tk.Frame(card, bg="white")
    bottom.pack(fill="x", padx=16, pady=(0, 14))

    btn_prev = ttk.Button(bottom, text="⬅️ Назад")
    btn_next = ttk.Button(bottom, text="Далее ➡️")
    btn_finish = ttk.Button(bottom, text="✅ Завершить", style="Main.TButton")

    btn_prev.pack(side="left")
    btn_finish.pack(side="right")
    btn_next.pack(side="right", padx=(0, 10))

    def _clear_options_box():
        for w in options_box.winfo_children():
            w.destroy()

    def _get_current_question():
        return questions[idx_var.get()]

    def _load_saved_answer(qid: str, qtype: str):
        chosen_single.set("")
        for k in list(chosen_multi.keys()):
            del chosen_multi[k]

        saved = user_answers.get(qid, [])
        if qtype == "single":
            chosen_single.set(saved[0] if saved else "")
        else:
            for letter in ["A", "B", "C", "D", "E", "F"]:
                chosen_multi[letter] = tk.BooleanVar(value=(letter in saved))

    def _render_question():
        _clear_options_box()
        q = _get_current_question()

        progress_lbl.config(text=f"Вопрос {idx_var.get()+1} из {len(questions)} • {q['id']}")
        q_lbl.config(text=q["question"])

        qtype = q["type"]
        qid = q["id"]
        opts = q["options"]

        _load_saved_answer(qid, qtype)

        # рисуем варианты
        for letter, text in opts.items():
            line = tk.Frame(options_box, bg="white")
            line.pack(anchor="w", fill="x", pady=4)

            if qtype == "single":
                rb = ttk.Radiobutton(
                    line,
                    text=f"{letter}) {text}",
                    value=letter,
                    variable=chosen_single,
                )
                rb.pack(anchor="w")
            else:
                if letter not in chosen_multi:
                    chosen_multi[letter] = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(
                    line,
                    text=f"{letter}) {text}",
                    variable=chosen_multi[letter],
                )
                cb.pack(anchor="w")

        # кнопки
        btn_prev.config(state="normal" if idx_var.get() > 0 else "disabled")
        btn_next.config(state="normal" if idx_var.get() < len(questions) - 1 else "disabled")

    def _save_current_answer():
        q = _get_current_question()
        qid = q["id"]
        if q["type"] == "single":
            val = chosen_single.get().strip()
            user_answers[qid] = [val] if val else []
        else:
            selected = [k for k, v in chosen_multi.items() if v.get()]
            # сохранить только те буквы, которые реально есть в options
            selected = [x for x in selected if x in q["options"]]
            user_answers[qid] = selected

    def _on_prev():
        _save_current_answer()
        idx_var.set(idx_var.get() - 1)
        _render_question()

    def _on_next():
        _save_current_answer()
        idx_var.set(idx_var.get() + 1)
        _render_question()

    def _compute_score():
        correct = 0
        total = len(questions)
        details = []

        for q in questions:
            qid = q["id"]
            right = sorted(q["answer"])
            got = sorted([x for x in user_answers.get(qid, []) if x])

            if q["type"] == "single":
                ok = (len(got) == 1 and got == right)
            else:
                ok = (got == right)

            if ok:
                correct += 1
            else:
                details.append((qid, right, got))

        return correct, total, details

    def _on_finish():
        _save_current_answer()

        # проверим пропуски
        unanswered = []
        for q in questions:
            if not user_answers.get(q["id"]):
                unanswered.append(q["id"])

        if unanswered:
            if not messagebox.askyesno(
                "Есть пропуски",
                "Вы ответили не на все вопросы.\n\n"
                f"Пропущены: {', '.join(unanswered[:8])}"
                + ("…" if len(unanswered) > 8 else "")
                + "\n\nЗавершить тест всё равно?",
            ):
                return

        correct, total, details = _compute_score()
        percent = round((correct / total) * 100)

        msg = f"Результат: {correct}/{total} ({percent}%)"
        if details:
            # покажем первые несколько ошибок
            lines = []
            for qid, right, got in details[:6]:
                lines.append(f"{qid}: верно {','.join(right)} / ваш ответ {(','.join(got) if got else '—')}")
            msg += "\n\nОшибки (первые):\n" + "\n".join(lines)
            if len(details) > 6:
                msg += "\n…"
        messagebox.showinfo("Итоги тестирования", msg)

        # после завершения можно вернуть назад
        go_back_callback()

    btn_prev.config(command=_on_prev)
    btn_next.config(command=_on_next)
    btn_finish.config(command=_on_finish)

    _render_question()
