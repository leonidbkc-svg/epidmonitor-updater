# screens/question_bank.py
import tkinter as tk
from tkinter import ttk, messagebox


def open_question_bank_screen(
    main_frame,
    build_header,
    go_back_callback,
):
    """
    Экран "Банк вопросов":
      - header (назад)
      - заголовок
      - вкладки (например: ИСМП 3.3686, Отходы 3.3684, и т.д.)
      - пока заглушки, дальше подключишь CRUD/импорт/просмотр
    """
    # очистка экрана
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame, back_callback=go_back_callback)

    tk.Label(
        main_frame,
        text="🧠 Банк вопросов",
        font=("Segoe UI", 20, "bold"),
        bg="#f4f6f8",
    ).pack(pady=(24, 10))

    body = tk.Frame(main_frame, bg="#f4f6f8")
    body.pack(expand=True, fill="both", padx=24, pady=18)

    card = tk.Frame(body, bg="white", bd=1, relief="solid")
    card.pack(expand=True, fill="both")

    top = tk.Frame(card, bg="white")
    top.pack(fill="x", padx=16, pady=(14, 10))

    tk.Label(
        top,
        text="Здесь будет храниться и редактироваться банк вопросов для тестирования.",
        font=("Segoe UI", 11),
        bg="white",
        fg="#374151",
        wraplength=980,
        justify="left",
    ).pack(anchor="w")

    # Tabs
    nb = ttk.Notebook(card)
    nb.pack(expand=True, fill="both", padx=12, pady=12)

    # --- TAB: СанПиН 3.3686 (ИСМП)
    tab_3686 = tk.Frame(nb, bg="white")
    nb.add(tab_3686, text="СанПиН 3.3686 (ИСМП)")

    tk.Label(
        tab_3686,
        text=(
            "Раздел для вопросов по профилактике ИСМП (СанПиН 3.3686).\n\n"
            "План:\n"
            "• просмотр списка вопросов\n"
            "• поиск / фильтры\n"
            "• добавление / редактирование\n"
            "• импорт из файла (JSON/CSV)\n"
        ),
        font=("Segoe UI", 11),
        bg="white",
        fg="#111827",
        wraplength=980,
        justify="left",
    ).pack(anchor="w", padx=16, pady=16)

    ttk.Button(
        tab_3686,
        text="➕ Добавить вопрос (заглушка)",
        style="Main.TButton",
        command=lambda: messagebox.showinfo("Скоро", "Добавление вопроса будет подключено позже."),
    ).pack(anchor="w", padx=16, pady=(0, 16))

    # --- TAB: СанПиН 3.3684 (Отходы)
    tab_3684 = tk.Frame(nb, bg="white")
    nb.add(tab_3684, text="СанПиН 3.3684 (Отходы)")

    tk.Label(
        tab_3684,
        text=(
            "Раздел для вопросов по обращению с медицинскими отходами (СанПиН 3.3684).\n\n"
            "План:\n"
            "• классы отходов и маркировка\n"
            "• тара и правила заполнения\n"
            "• обеззараживание и транспортирование\n"
            "• типовые ошибки\n"
        ),
        font=("Segoe UI", 11),
        bg="white",
        fg="#111827",
        wraplength=980,
        justify="left",
    ).pack(anchor="w", padx=16, pady=16)

    ttk.Button(
        tab_3684,
        text="➕ Добавить вопрос (заглушка)",
        style="Main.TButton",
        command=lambda: messagebox.showinfo("Скоро", "Добавление вопроса будет подключено позже."),
    ).pack(anchor="w", padx=16, pady=(0, 16))

    # --- TAB: Прочее/настройки
    tab_other = tk.Frame(nb, bg="white")
    nb.add(tab_other, text="Настройки")

    tk.Label(
        tab_other,
        text=(
            "Настройки банка вопросов:\n"
            "• источник хранения (локально/сервер)\n"
            "• импорт/экспорт\n"
            "• версии тестов\n"
            "• уровни сложности\n"
        ),
        font=("Segoe UI", 11),
        bg="white",
        fg="#111827",
        wraplength=980,
        justify="left",
    ).pack(anchor="w", padx=16, pady=16)

    ttk.Button(
        tab_other,
        text="📦 Экспорт (заглушка)",
        style="Main.TButton",
        command=lambda: messagebox.showinfo("Скоро", "Экспорт будет подключён позже."),
    ).pack(anchor="w", padx=16, pady=(0, 16))
