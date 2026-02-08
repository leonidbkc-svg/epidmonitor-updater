# screens/documents.py
import os
import tkinter as tk
from tkinter import ttk, messagebox
from services import webdav_sync

# Office preview uses COM
try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None


def _get_documents_root() -> str:
    """
    Берём единый путь из microbio_app.py (чтобы совпадало с сетью/фолбэком).
    Импорт внутри — чтобы не было циклических импортов.
    """
    from microbio_app import DOCUMENTS_DIR, DATA_ROOT
    webdav_sync.sync_down(DATA_ROOT)
    os.makedirs(DOCUMENTS_DIR, exist_ok=True)
    return DOCUMENTS_DIR


def _copy_to_clipboard(widget: tk.Widget, text: str) -> None:
    try:
        widget.clipboard_clear()
        widget.clipboard_append(text)
        widget.update()  # чтобы точно записалось
    except Exception:
        pass


def _open_file(path: str) -> None:
    if not os.path.exists(path):
        messagebox.showerror("Ошибка", "Файл не найден")
        return
    try:
        os.startfile(path)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")


def _print_preview(path: str) -> None:
    """
    Печать через предпросмотр:
    - DOC/DOCX -> Word Print Preview
    - XLS/XLSX -> Excel Print Preview
    - остальное -> просто открыть файл (универсального preview нет)
    """
    if not os.path.exists(path):
        messagebox.showerror("Ошибка", "Файл не найден")
        return

    ext = os.path.splitext(path)[1].lower()

    # ---- WORD ----
    if ext in (".doc", ".docx"):
        if win32com is None:
            messagebox.showwarning(
                "Печать",
                "Не найден модуль win32com.\n"
                "Открою файл, печать сделайте из приложения."
            )
            _open_file(path)
            return

        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)

            # Включаем предпросмотр печати
            # Самый совместимый способ:
            doc.ActiveWindow.View.Type = 3  # wdPrintPreview = 3
            # Альтернатива (не всегда доступна):
            # doc.PrintPreview()

        except Exception as e:
            messagebox.showwarning(
                "Печать",
                f"Не удалось открыть предпросмотр в Word:\n{e}\n\nОткрою файл обычным способом."
            )
            _open_file(path)
        return

    # ---- EXCEL ----
    if ext in (".xls", ".xlsx"):
        if win32com is None:
            messagebox.showwarning(
                "Печать",
                "Не найден модуль win32com.\n"
                "Открою файл, печать сделайте из приложения."
            )
            _open_file(path)
            return

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            wb = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)

            # Предпросмотр печати (Excel сам показывает окно preview)
            wb.PrintPreview()

        except Exception as e:
            messagebox.showwarning(
                "Печать",
                f"Не удалось открыть предпросмотр в Excel:\n{e}\n\nОткрою файл обычным способом."
            )
            _open_file(path)
        return

    # ---- OTHER (PDF, images, etc.) ----
    # Универсально открыть print-preview невозможно (зависит от программы по умолчанию).
    # Поэтому аккуратно открываем файл, а пользователь печатает из приложения.
    messagebox.showinfo(
        "Печать",
        "Для этого типа файла предпросмотр зависит от программы по умолчанию.\n"
        "Открою документ — нажмите Печать в приложении."
    )
    _open_file(path)


def build_documents_screen(main_frame, build_header, go_back_callback):
    # очистка экрана
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame, back_callback=go_back_callback)

    tk.Label(
        main_frame,
        text="Документы",
        font=("Segoe UI", 18, "bold"),
        bg="#f4f6f8"
    ).pack(pady=(25, 10))

    root_dir = _get_documents_root()

    top_bar = tk.Frame(main_frame, bg="#f4f6f8")
    top_bar.pack(fill="x", padx=20, pady=(0, 10))

    ttk.Button(
        top_bar,
        text="🔄 Обновить",
        style="Secondary.TButton",
        command=lambda: refresh()
    ).pack(side="left")

    hint = tk.Label(
        top_bar,
        text=f"Папка: {root_dir}",
        bg="#f4f6f8",
        fg="#6b7280",
        font=("Segoe UI", 9)
    )
    hint.pack(side="right")

    container = tk.Frame(main_frame, bg="#f4f6f8")
    container.pack(expand=True, fill="both", padx=20, pady=10)

    canvas = tk.Canvas(container, bg="#f4f6f8", highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg="#f4f6f8")

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def clear_list():
        for w in scroll_frame.winfo_children():
            w.destroy()

    def refresh():
        clear_list()
        from microbio_app import DATA_ROOT
        webdav_sync.sync_down(DATA_ROOT)

        if not os.path.exists(root_dir):
            tk.Label(
                scroll_frame,
                text="Папка документов недоступна",
                font=("Segoe UI", 12),
                bg="#f4f6f8"
            ).pack(pady=40)
            return

        files = sorted(
            f for f in os.listdir(root_dir)
            if os.path.isfile(os.path.join(root_dir, f))
        )

        if not files:
            tk.Label(
                scroll_frame,
                text="Документы отсутствуют",
                font=("Segoe UI", 12),
                bg="#f4f6f8"
            ).pack(pady=40)
            return

        for fname in files:
            full_path = os.path.join(root_dir, fname)

            row = tk.Frame(scroll_frame, bg="#f4f6f8", padx=8, pady=6)
            row.pack(fill="x", padx=10, pady=4)

            label = tk.Label(
                row,
                text="📄 " + fname,
                bg="#f4f6f8",
                anchor="w",
                font=("Segoe UI", 11),
                cursor="hand2"
            )
            label.pack(fill="x", expand=True)

            # подсветка
            def on_enter(e, r=row):
                r.configure(bg="#e5e7eb")
                for c in r.winfo_children():
                    c.configure(bg="#e5e7eb")

            def on_leave(e, r=row):
                r.configure(bg="#f4f6f8")
                for c in r.winfo_children():
                    c.configure(bg="#f4f6f8")

            row.bind("<Enter>", on_enter)
            row.bind("<Leave>", on_leave)
            label.bind("<Enter>", on_enter)
            label.bind("<Leave>", on_leave)

            # двойной клик — открыть
            row.bind("<Double-Button-1>", lambda e, p=full_path: _open_file(p))
            label.bind("<Double-Button-1>", lambda e, p=full_path: _open_file(p))

            # ПКМ меню
            menu = tk.Menu(row, tearoff=0)
            menu.add_command(label="📂 Открыть", command=lambda p=full_path: _open_file(p))
            menu.add_command(label="📋 Копировать путь", command=lambda p=full_path: _copy_to_clipboard(main_frame, p))
            menu.add_separator()
            menu.add_command(label="🖨 Печать (предпросмотр)", command=lambda p=full_path: _print_preview(p))
            menu.add_separator()
            menu.add_command(
                label="🗑 Удалить",
                command=lambda p=full_path: delete_file(p)
            )

            def show_menu(event, m=menu):
                m.tk_popup(event.x_root, event.y_root)

            row.bind("<Button-3>", show_menu)
            label.bind("<Button-3>", show_menu)

    def delete_file(path: str):
        if not os.path.exists(path):
            refresh()
            return

        name = os.path.basename(path)
        if not messagebox.askyesno("Удалить документ", f"Удалить файл?\n\n{name}"):
            return

        try:
            os.remove(path)
            from microbio_app import DATA_ROOT
            webdav_sync.delete_path(path, DATA_ROOT)
        except Exception as e:
            messagebox.showerror("Ошибка удаления", str(e))
            return

        refresh()

    # первая загрузка
    refresh()
