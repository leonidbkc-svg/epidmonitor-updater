# screens/testing_menu.py
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

from screens.testing import open_testing_screen
from screens.tg_exam_stats import open_tg_exam_stats
from screens.question_bank import open_question_bank_screen


def open_testing_menu(
    main_frame,
    build_header,
    go_back_callback,
    data_root: str,
    base_url: str,
    report_api_key: str,
    test_server_base: str,
    test_admin_pin: str,
    vacc_server_base: str,
    vacc_admin_pin: str,
    qr_image_path: str,
):
    # очистим экран
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame, back_callback=go_back_callback)

    # ---------- заголовок ----------
    tk.Label(
        main_frame,
        text="🧪 Тестирование",
        font=("Segoe UI", 20, "bold"),
        bg="#f4f6f8",
        fg="#111827",
    ).pack(pady=(24, 10))

    # подзаголовок
    tk.Label(
        main_frame,
        text="Выберите нужный раздел тестирования.",
        font=("Segoe UI", 11),
        bg="#f4f6f8",
        fg="#6b7280",
    ).pack(pady=(0, 14))

    body = tk.Frame(main_frame, bg="#f4f6f8")
    body.pack(expand=True, fill="both", padx=30, pady=20)

    left = tk.Frame(body, bg="#f4f6f8")
    left.pack(side="left", fill="y")

    right = tk.Frame(body, bg="#f4f6f8")
    right.pack(side="right", fill="both", expand=True)

    # ---------- карточка слева: кнопки ----------
    card = tk.Frame(left, bg="white", bd=1, relief="solid")
    card.pack(anchor="nw")

    tk.Label(
        card,
        text="Основные разделы",
        font=("Segoe UI", 12, "bold"),
        bg="white",
        fg="#111827",
    ).pack(anchor="w", padx=14, pady=(12, 2))

    tk.Label(
        card,
        text="Выберите действие:",
        font=("Segoe UI", 10),
        bg="white",
        fg="#6b7280",
    ).pack(anchor="w", padx=14, pady=(0, 10))

    btns = tk.Frame(card, bg="white")
    btns.pack(fill="x", padx=14, pady=(0, 14))

    def _back_to_menu():
        open_testing_menu(
            main_frame,
            build_header,
            go_back_callback,
            data_root,
            base_url,
            report_api_key,
            test_server_base,
            test_admin_pin,
            vacc_server_base,
            vacc_admin_pin,
            qr_image_path,
        )

    ttk.Button(
        btns,
        text="🧑‍⚕️ Запуск тестирования",
        style="Main.TButton",
        command=lambda: open_testing_screen(
            main_frame=main_frame,
            build_header=build_header,
            go_back_callback=_back_to_menu,
            server_base_url=vacc_server_base,
            admin_pin=vacc_admin_pin,
        ),
    ).pack(fill="x", pady=(0, 10))

    ttk.Button(
        btns,
        text="📊 Статистика tg-exam",
        style="Main.TButton",
        command=lambda: open_tg_exam_stats(
            main_frame=main_frame,
            build_header=build_header,
            go_back_callback=_back_to_menu,
            data_root=data_root,
            base_url=base_url,
            report_api_key=report_api_key,
        ),
    ).pack(fill="x", pady=(0, 10))

    ttk.Button(
        btns,
        text="🗂 Банк вопросов",
        style="Main.TButton",
        command=lambda: open_question_bank_screen(
            main_frame=main_frame,
            build_header=build_header,
            go_back_callback=_back_to_menu,
        ),
    ).pack(fill="x")

    # ---------- правая карточка: QR ----------
    qr_box = tk.Frame(right, bg="white", bd=1, relief="solid")
    qr_box.pack(anchor="e", padx=10, pady=10)

    tk.Label(
        qr_box,
        text="QR для запуска теста",
        bg="white",
        fg="#111827",
        font=("Segoe UI", 12, "bold"),
    ).pack(anchor="w", padx=14, pady=(12, 4))

    tk.Label(
        qr_box,
        text="Отсканируйте, чтобы открыть\nтестирование в Telegram.",
        bg="white",
        fg="#6b7280",
        font=("Segoe UI", 10),
        justify="left",
    ).pack(anchor="w", padx=14, pady=(0, 10))

    try:
        img = Image.open(qr_image_path).resize((260, 260), Image.LANCZOS)
        qr_img = ImageTk.PhotoImage(img)

        lbl = tk.Label(qr_box, image=qr_img, bg="white")
        lbl.image = qr_img
        lbl.pack(padx=14, pady=(4, 8))
    except Exception:
        tk.Label(
            qr_box,
            text="(QR не найден)",
            bg="white",
            fg="#6b7280",
            font=("Segoe UI", 10),
        ).pack(padx=14, pady=(12, 12))

    tk.Label(
        qr_box,
        text="@EPID_TEST_BOT",
        bg="white",
        fg="#9ca3af",
        font=("Segoe UI", 9),
        justify="left",
    ).pack(anchor="w", padx=14, pady=(0, 12))
