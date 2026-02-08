# screens/photo_rounds.py
import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from services import webdav_sync
from datetime import datetime
from PIL import Image, ImageTk

SUPPORTED_EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp")


def _get_photo_root() -> str:
    """
    Берём единый путь архива из microbio_app.py
    (импорт внутри функции, чтобы не ловить циклические импорты).
    """
    from microbio_app import PHOTO_ROUNDS_DIR, DATA_ROOT
    webdav_sync.ensure_synced(DATA_ROOT)
    os.makedirs(PHOTO_ROUNDS_DIR, exist_ok=True)
    return PHOTO_ROUNDS_DIR


def _safe_name(name: str) -> str:
    name = (name or "").strip()
    name = name.replace("/", "_").replace("\\", "_")
    while "  " in name:
        name = name.replace("  ", " ")
    return name


def _delete_photo(full_path: str) -> bool:
    """Физически удаляет файл фото с диска (шары/ПК) + сообщает об ошибках."""
    if not full_path or not os.path.exists(full_path):
        messagebox.showwarning("Не найдено", "Файл не найден.")
        return False

    if not messagebox.askyesno("Удалить фото", "Удалить фотографию?\n\nФайл будет удалён с диска."):
        return False

    try:
        os.remove(full_path)
        from microbio_app import DATA_ROOT
        webdav_sync.delete_path(full_path, DATA_ROOT)
        return True
    except Exception as e:
        messagebox.showerror("Ошибка удаления", str(e))
        return False


def _delete_folder(folder_path: str) -> bool:
    """Удаляет папку дня целиком (со всем содержимым)."""
    if not folder_path or not os.path.isdir(folder_path):
        messagebox.showwarning("Не найдено", "Папка не найдена.")
        return False

    if not messagebox.askyesno("Удалить папку", "Удалить папку дня целиком?\n\nБудут удалены все фото внутри."):
        return False

    try:
        shutil.rmtree(folder_path)
        from microbio_app import DATA_ROOT
        webdav_sync.delete_path(folder_path, DATA_ROOT)
        return True
    except Exception as e:
        messagebox.showerror("Ошибка удаления", str(e))
        return False


def _save_photo_as(src_path: str) -> None:
    if not src_path or not os.path.exists(src_path):
        messagebox.showwarning("Не найдено", "Файл не найден.")
        return

    base = os.path.basename(src_path)
    ext = os.path.splitext(base)[1].lower()
    dst = filedialog.asksaveasfilename(
        title="Сохранить фото как",
        initialfile=base,
        defaultextension=ext,
        filetypes=[
            ("Изображения", "*.jpg;*.jpeg;*.png;*.webp;*.bmp"),
            ("Все файлы", "*.*"),
        ],
    )
    if not dst:
        return

    try:
        shutil.copy2(src_path, dst)
        messagebox.showinfo("Готово", f"Сохранено:\n{dst}")
    except Exception as e:
        messagebox.showerror("Ошибка сохранения", str(e))


def _parse_date_any(name: str) -> datetime | None:
    """
    Пытаемся понять дату папки в разных форматах.
    Поддержим:
      - dd.mm.yyyy
      - dd-mm-yyyy
      - yyyy-mm-dd
      - yyyy.mm.dd
      - dd_mm_yyyy
      - yyyy_mm_dd
    """
    s = (name or "").strip()
    if not s:
        return None

    candidates = [
        ("%d.%m.%Y", s),
        ("%d-%m-%Y", s),
        ("%Y-%m-%d", s),
        ("%Y.%m.%d", s),
        ("%d_%m_%Y", s),
        ("%Y_%m_%d", s),
    ]
    for fmt, val in candidates:
        try:
            return datetime.strptime(val, fmt)
        except Exception:
            pass
    return None


def _to_ddmmyyyy_folder_name(name: str) -> str | None:
    """Возвращает имя в dd.mm.yyyy если дата распарсилась, иначе None."""
    dt = _parse_date_any(name)
    if not dt:
        return None
    return dt.strftime("%d.%m.%Y")


def _day_folder_has_supported_images(day_path: str) -> bool:
    """Есть ли в папке дня хоть одна поддерживаемая картинка (рекурсивно не нужно)."""
    try:
        for fn in os.listdir(day_path):
            if fn.lower().endswith(SUPPORTED_EXT) and os.path.isfile(os.path.join(day_path, fn)):
                return True
    except Exception:
        return False
    return False


def build_photo_rounds_screen(main_frame, build_header, go_back_callback):
    # очистка экрана
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame, back_callback=go_back_callback)

    tk.Label(
        main_frame,
        text="Фотоотчёт обходов",
        font=("Segoe UI", 18, "bold"),
        bg="#f4f6f8"
    ).pack(pady=(20, 10))

    root_dir = _get_photo_root()

    top_bar = tk.Frame(main_frame, bg="#f4f6f8")
    top_bar.pack(fill="x", padx=20, pady=(0, 10))

    ttk.Button(
        top_bar,
        text="🔄 Обновить",
        style="Secondary.TButton",
        command=lambda: refresh_all()
    ).pack(side="left")

    save_selected_btn = ttk.Button(
        top_bar,
        text="Сохранить выбранное...",
        style="Secondary.TButton",
        state="disabled",
        command=lambda: _save_selected(),
    )
    save_selected_btn.pack(side="left", padx=(8, 0))

    hint = tk.Label(
        top_bar,
        text=f"Папка архива: {root_dir}",
        bg="#f4f6f8",
        fg="#6b7280",
        font=("Segoe UI", 9)
    )
    hint.pack(side="right")

    # основной контейнер 3 колонки
    container = tk.Frame(main_frame, bg="#f4f6f8")
    container.pack(expand=True, fill="both", padx=20, pady=10)

    # -------------------------
    # ЛЕВАЯ: отделения
    # -------------------------
    left = tk.Frame(container, bg="#f4f6f8")
    left.pack(side="left", fill="y")

    tk.Label(left, text="Отделения", bg="#f4f6f8", font=("Segoe UI", 11, "bold")).pack(anchor="w")

    dep_list = tk.Listbox(left, width=26, height=24)
    dep_list.pack(fill="y", pady=6)

    # -------------------------
    # СРЕДНЯЯ: даты (папки дней)
    # -------------------------
    mid = tk.Frame(container, bg="#f4f6f8")
    mid.pack(side="left", fill="y", padx=(15, 0))

    tk.Label(mid, text="Дни", bg="#f4f6f8", font=("Segoe UI", 11, "bold")).pack(anchor="w")

    day_list = tk.Listbox(mid, width=18, height=24)
    day_list.pack(fill="y", pady=6)

    # -------------------------
    # ПРАВАЯ: фото (прокрутка)
    # -------------------------
    right = tk.Frame(container, bg="#f4f6f8")
    right.pack(side="left", expand=True, fill="both", padx=(15, 0))

    tk.Label(right, text="Фотографии", bg="#f4f6f8", font=("Segoe UI", 11, "bold")).pack(anchor="w")

    scroll_container = tk.Frame(right, bg="#f4f6f8")
    scroll_container.pack(expand=True, fill="both", pady=6)

    canvas = tk.Canvas(scroll_container, bg="#f4f6f8", highlightthickness=0)
    vbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vbar.set)

    vbar.pack(side="right", fill="y")
    canvas.pack(side="left", expand=True, fill="both")

    grid_frame = tk.Frame(canvas, bg="#f4f6f8")
    canvas.create_window((0, 0), window=grid_frame, anchor="nw")

    def _on_frame_configure(_):
        canvas.configure(scrollregion=canvas.bbox("all"))

    grid_frame.bind("<Configure>", _on_frame_configure)

    # Храним ссылки на картинки, иначе Tk их "съест"
    thumbs_cache = []
    current_dep = {"name": None}
    current_day = {"name": None}
    selection = {
        "paths": set(),      # set[str]
        "widgets": {},       # dict[str, tuple[tk.Widget, tk.Widget, tk.Widget, tk.Widget]]
        "vars": {},          # dict[str, tk.BooleanVar]
        "order": [],         # list[str]
    }

    # Маппинг отображаемого дня -> реальная папка (на случай автопереименования / другого формата)
    day_display_to_folder = {}

    def _apply_selected_style(path: str, is_selected: bool) -> None:
        w = selection["widgets"].get(path)
        if not w:
            return
        card, img_lbl, name_lbl, chk = w

        # Более мягкое выделение: фон + тонкая рамка.
        bg = "#e8f0fe" if is_selected else "white"

        try:
            card.configure(bg=bg)
        except Exception:
            pass
        try:
            img_lbl.configure(bg=bg)
        except Exception:
            pass
        try:
            name_lbl.configure(bg=bg)
        except Exception:
            pass
        try:
            chk.configure(bg=bg, activebackground=bg)
        except Exception:
            pass
        try:
            card.configure(
                highlightthickness=(2 if is_selected else 0),
                highlightbackground="#2563eb",
                highlightcolor="#2563eb",
            )
        except Exception:
            pass

    def _update_save_button() -> None:
        try:
            cnt = len(selection["paths"])
            save_selected_btn.configure(state=("normal" if cnt else "disabled"))
            if cnt <= 1:
                save_selected_btn.configure(text="Сохранить выбранное...")
            else:
                save_selected_btn.configure(text=f"Сохранить выбранные ({cnt})...")
        except Exception:
            pass

    def _clear_selection() -> None:
        # Сначала снимем галочки (вызов command на Checkbutton при set() не происходит).
        for var in list(selection["vars"].values()):
            try:
                var.set(False)
            except Exception:
                pass
        for p in list(selection["paths"]):
            _apply_selected_style(p, False)
        selection["paths"].clear()
        _update_save_button()

    def _add_to_selection(p: str) -> None:
        if p in selection["paths"]:
            return
        selection["paths"].add(p)
        try:
            v = selection["vars"].get(p)
            if v is not None:
                v.set(True)
        except Exception:
            pass
        _apply_selected_style(p, True)
        _update_save_button()

    def _remove_from_selection(p: str) -> None:
        if p not in selection["paths"]:
            return
        selection["paths"].discard(p)
        try:
            v = selection["vars"].get(p)
            if v is not None:
                v.set(False)
        except Exception:
            pass
        _apply_selected_style(p, False)
        _update_save_button()

    def _toggle_selection(p: str) -> None:
        if p in selection["paths"]:
            _remove_from_selection(p)
        else:
            _add_to_selection(p)

    def _save_selected() -> None:
        paths = sorted(selection["paths"])
        if not paths:
            return

        if len(paths) == 1:
            _save_photo_as(paths[0])
            return

        dst_dir = filedialog.askdirectory(title="Куда сохранить выбранные фото")
        if not dst_dir:
            return

        saved = 0
        for src in paths:
            try:
                base = os.path.basename(src)
                name, ext = os.path.splitext(base)
                dst = os.path.join(dst_dir, base)
                # Не перезаписываем: добавим суффикс.
                n = 1
                while os.path.exists(dst):
                    dst = os.path.join(dst_dir, f"{name} ({n}){ext}")
                    n += 1
                shutil.copy2(src, dst)
                saved += 1
            except Exception:
                # пропускаем проблемные файлы, но продолжаем
                pass

        messagebox.showinfo("Готово", f"Сохранено фото: {saved}\nПапка:\n{dst_dir}")

    def list_departments():
        dep_list.delete(0, tk.END)
        if not os.path.isdir(root_dir):
            return
        deps = sorted([d for d in os.listdir(root_dir) if os.path.isdir(os.path.join(root_dir, d))])
        for d in deps:
            dep_list.insert(tk.END, d)

    def clear_photos():
        nonlocal thumbs_cache
        thumbs_cache = []
        selection["widgets"].clear()
        selection["vars"].clear()
        selection["order"].clear()
        _clear_selection()
        for w in grid_frame.winfo_children():
            w.destroy()
        canvas.yview_moveto(0)

    def list_days(dep_name: str):
        """
        Требование:
          - формат папки даты: dd.mm.yyyy
        Мы:
          1) пытаемся распарсить имя папки как дату
          2) если распарсилось и формат не dd.mm.yyyy — пробуем переименовать папку на dd.mm.yyyy
          3) в списке показываем dd.mm.yyyy
        """
        day_list.delete(0, tk.END)
        day_display_to_folder.clear()
        current_day["name"] = None
        clear_photos()

        if not dep_name:
            return

        dep_path = os.path.join(root_dir, dep_name)
        if not os.path.isdir(dep_path):
            return

        dirs = [d for d in os.listdir(dep_path) if os.path.isdir(os.path.join(dep_path, d))]

        # Сначала нормализуем имена папок (best-effort)
        normalized = []
        for d in dirs:
            old_path = os.path.join(dep_path, d)
            target = _to_ddmmyyyy_folder_name(d)

            if target and target != d:
                new_path = os.path.join(dep_path, target)
                # чтобы не затирать существующую папку при коллизии
                if not os.path.exists(new_path):
                    try:
                        os.rename(old_path, new_path)
                        d = target
                    except Exception:
                        # если не смогли переименовать — оставляем как есть
                        pass

            # Для сортировки: если дата распарсилась — используем её, иначе по имени
            dt = _parse_date_any(d)
            normalized.append((d, dt))

        # Сортировка по дате убыванию, неизвестные — вниз
        normalized.sort(key=lambda x: (x[1] is None, x[1] if x[1] else datetime.min), reverse=True)

        for folder_name, _dt in normalized:
            # Показываем dd.mm.yyyy если распарсилось, иначе как есть
            display = _to_ddmmyyyy_folder_name(folder_name) or folder_name
            day_list.insert(tk.END, display)
            day_display_to_folder[display] = folder_name

    def _after_possible_empty_day_cleanup(dep_name: str, day_folder_name: str):
        """
        Если после удаления фото папка дня пуста (нет ни одной поддерживаемой картинки) —
        удаляем папку дня, обновляем список дней и чистим сетку.
        """
        if not dep_name or not day_folder_name:
            return

        day_path = os.path.join(root_dir, dep_name, day_folder_name)
        if not os.path.isdir(day_path):
            # уже удалили
            list_days(dep_name)
            return

        if not _day_folder_has_supported_images(day_path):
            try:
                shutil.rmtree(day_path)
            except Exception:
                # если не вышло — оставим, но всё равно обновим интерфейс
                pass
            list_days(dep_name)
            clear_photos()
            current_day["name"] = None

    def list_photos(dep_name: str, day_display: str):
        clear_photos()

        if not dep_name or not day_display:
            return

        day_folder = day_display_to_folder.get(day_display, day_display)
        day_path = os.path.join(root_dir, dep_name, day_folder)
        if not os.path.isdir(day_path):
            return

        files = sorted([f for f in os.listdir(day_path) if f.lower().endswith(SUPPORTED_EXT)])

        if not files:
            # если в папке реально пусто — удалим папку (чтобы не копить пустые дни)
            _after_possible_empty_day_cleanup(dep_name, day_folder)
            tk.Label(
                grid_frame,
                text="Нет фотографий в этой папке",
                bg="#f4f6f8",
                fg="#6b7280",
                font=("Segoe UI", 11)
            ).pack(pady=30)
            return

        # сетка: 4 колонки
        cols = 4
        thumb_size = (220, 160)

        for i, fn in enumerate(files):
            full_path = os.path.join(day_path, fn)
            selection["order"].append(full_path)

            card = tk.Frame(
                grid_frame,
                bg="white",
                bd=1,
                relief="solid",
                highlightthickness=0,
            )
            r = i // cols
            c = i % cols
            card.grid(row=r, column=c, padx=8, pady=8, sticky="n")

            # превью
            try:
                img = Image.open(full_path)
                img.thumbnail(thumb_size, Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
            except Exception:
                photo = None

            thumbs_cache.append(photo)

            if photo:
                img_lbl = tk.Label(card, image=photo, bg="white", cursor="hand2")
                img_lbl.pack(padx=6, pady=6)
            else:
                img_lbl = tk.Label(card, text="(не удалось открыть)", bg="white", fg="red")
                img_lbl.pack(padx=20, pady=30)

            name_lbl = tk.Label(
                card,
                text=fn,
                bg="white",
                fg="#374151",
                font=("Segoe UI", 9),
                wraplength=220,
                justify="center"
            )
            name_lbl.pack(padx=6, pady=(0, 6))

            # Чекбокс выбора (без Ctrl/Shift): можно отметить несколько фото.
            var = tk.BooleanVar(value=False)
            selection["vars"][full_path] = var

            def _on_check(p=full_path):
                v = selection["vars"].get(p)
                if v is None:
                    return
                if v.get():
                    _add_to_selection(p)
                else:
                    _remove_from_selection(p)

            chk = tk.Checkbutton(
                card,
                variable=var,
                command=_on_check,
                text="",
                bg="white",
                activebackground="white",
                highlightthickness=0,
                bd=0,
                cursor="hand2",
            )
            chk.place(x=6, y=6)

            selection["widgets"][full_path] = (card, img_lbl, name_lbl, chk)

            def _open(_e=None, p=full_path):
                open_viewer(p)

            def _toggle_check(p=full_path):
                v = selection["vars"].get(p)
                if v is None:
                    return
                v.set(not v.get())
                _on_check(p)

            # Клик по карточке/фото/имени просто переключает галочку.
            card.bind("<Button-1>", lambda _e, p=full_path: _toggle_check(p))
            img_lbl.bind("<Button-1>", lambda _e, p=full_path: _toggle_check(p))
            name_lbl.bind("<Button-1>", lambda _e, p=full_path: _toggle_check(p))

            card.bind("<Double-Button-1>", _open)
            img_lbl.bind("<Double-Button-1>", _open)
            name_lbl.bind("<Double-Button-1>", _open)

            def _delete_and_refresh(p=full_path, dep=dep_name, day=day_folder, disp=day_display):
                ok = _delete_photo(p)
                if not ok:
                    return
                list_photos(dep, disp)  # обновим сетку
                _after_possible_empty_day_cleanup(dep, day)  # и если стало пусто — удалим день

            # ---------- ПКМ: сохранить / удалить ----------
            menu = tk.Menu(card, tearoff=0)
            menu.add_command(
                label="Сохранить как...",
                command=lambda p=full_path: _save_photo_as(p),
            )
            menu.add_separator()
            menu.add_command(
                label="🗑 Удалить фото",
                command=_delete_and_refresh,
            )

            def _show_menu(event, m=menu, card_ref=card, p=full_path):
                # ПКМ по фото: если оно не выбрано — просто отметим его галочкой.
                if p not in selection["paths"]:
                    v = selection["vars"].get(p)
                    if v is not None:
                        v.set(True)
                    _add_to_selection(p)
                m.tk_popup(event.x_root, event.y_root)

            card.bind("<Button-3>", _show_menu)
            img_lbl.bind("<Button-3>", _show_menu)
            name_lbl.bind("<Button-3>", _show_menu)

    def open_viewer(image_path: str):
        """Просмотр: по умолчанию вписать целиком; есть скроллы/перетаскивание; колесо — зум; ПКМ — удалить фото."""
        if not os.path.exists(image_path):
            messagebox.showerror("Ошибка", "Файл не найден")
            return

        win = tk.Toplevel(main_frame)
        win.title(os.path.basename(image_path))
        win.geometry("1100x750")
        win.minsize(700, 450)

        # верхняя панель
        top = tk.Frame(win)
        top.pack(fill="x")

        ttk.Label(top, text=image_path).pack(side="left", padx=10, pady=8)

        btns = tk.Frame(top)
        btns.pack(side="right", padx=10)

        # контейнер со скроллами
        container = tk.Frame(win)
        container.pack(expand=True, fill="both")

        view_canvas = tk.Canvas(container, bg="black", highlightthickness=0)
        vbar = ttk.Scrollbar(container, orient="vertical", command=view_canvas.yview)
        hbar = ttk.Scrollbar(container, orient="horizontal", command=view_canvas.xview)

        view_canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")
        view_canvas.pack(side="left", expand=True, fill="both")

        # загрузим исходник
        try:
            orig = Image.open(image_path).convert("RGB")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть изображение:\n{e}")
            win.destroy()
            return

        state = {"scale": 1.0, "imgtk": None, "base": orig}

        def redraw():
            cw = view_canvas.winfo_width()
            ch = view_canvas.winfo_height()
            if cw <= 10 or ch <= 10:
                return

            iw, ih = state["base"].size
            new_w = max(1, int(iw * state["scale"]))
            new_h = max(1, int(ih * state["scale"]))

            scaled = state["base"].resize((new_w, new_h), Image.LANCZOS)
            imgtk = ImageTk.PhotoImage(scaled)
            state["imgtk"] = imgtk

            view_canvas.delete("all")
            view_canvas.create_image(0, 0, image=imgtk, anchor="nw")
            view_canvas.config(scrollregion=(0, 0, new_w, new_h))

        def fit_to_window():
            cw = view_canvas.winfo_width()
            ch = view_canvas.winfo_height()
            if cw <= 10 or ch <= 10:
                return
            iw, ih = state["base"].size
            if iw <= 0 or ih <= 0:
                return
            state["scale"] = min(cw / iw, ch / ih)
            redraw()

        def zoom_100():
            state["scale"] = 1.0
            redraw()

        ttk.Button(btns, text="Вписать", command=fit_to_window).pack(side="left", padx=4)
        ttk.Button(btns, text="100%", command=zoom_100).pack(side="left", padx=4)
        ttk.Button(btns, text="Сохранить как...", command=lambda p=image_path: _save_photo_as(p)).pack(side="left", padx=4)

        # zoom колесом
        def on_wheel(event):
            if event.delta > 0:
                state["scale"] = min(8.0, state["scale"] * 1.15)
            else:
                state["scale"] = max(0.1, state["scale"] / 1.15)
            redraw()

        view_canvas.bind("<MouseWheel>", on_wheel)

        # панорамирование мышью
        def start_pan(event):
            view_canvas.scan_mark(event.x, event.y)

        def do_pan(event):
            view_canvas.scan_dragto(event.x, event.y, gain=1)

        view_canvas.bind("<ButtonPress-1>", start_pan)
        view_canvas.bind("<B1-Motion>", do_pan)

        # ---------- ПКМ В VIEWER: удалить фото ----------
        menu = tk.Menu(win, tearoff=0)
        menu.add_command(
            label="Сохранить как...",
            command=lambda p=image_path: _save_photo_as(p),
        )
        menu.add_separator()
        menu.add_command(
            label="🗑 Удалить фото",
            command=lambda p=image_path: _delete_from_viewer(p)
        )

        def show_menu(event, m=menu):
            m.tk_popup(event.x_root, event.y_root)

        view_canvas.bind("<Button-3>", show_menu)
        win.bind("<Button-3>", show_menu)

        def _delete_from_viewer(p: str):
            # определим день/отделение из пути (так надёжнее, чем хранить отдельно)
            dep = current_dep["name"]
            disp = current_day["name"]
            day_folder = day_display_to_folder.get(disp, disp) if disp else None

            ok = _delete_photo(p)
            if not ok:
                return

            # закрываем viewer
            try:
                win.destroy()
            except Exception:
                pass

            # обновляем правую сетку и при необходимости удаляем папку дня
            if dep and disp:
                list_photos(dep, disp)
            if dep and day_folder:
                _after_possible_empty_day_cleanup(dep, day_folder)

        # первая отрисовка: вписать
        win.after(80, fit_to_window)

    def on_dep_select(_=None):
        sel = dep_list.curselection()
        if not sel:
            return
        dep_name = dep_list.get(sel[0])
        current_dep["name"] = dep_name
        list_days(dep_name)

    def on_day_select(_=None):
        sel = day_list.curselection()
        if not sel:
            return
        day_display = day_list.get(sel[0])
        current_day["name"] = day_display
        list_photos(current_dep["name"], day_display)

    def refresh_all():
        from microbio_app import DATA_ROOT
        webdav_sync.sync_down(DATA_ROOT)
        list_departments()
        day_list.delete(0, tk.END)
        day_display_to_folder.clear()
        clear_photos()
        current_dep["name"] = None
        current_day["name"] = None

    # ---------- ПКМ по списку дней: удалить папку дня ----------
    day_menu = tk.Menu(day_list, tearoff=0)

    def _delete_selected_day(day_display: str):
        dep = current_dep["name"]
        if not dep or not day_display:
            return

        day_folder = day_display_to_folder.get(day_display, day_display)
        day_path = os.path.join(root_dir, dep, day_folder)

        if _delete_folder(day_path):
            # если удалили текущий день — очистим фото
            if current_day["name"] == day_display:
                current_day["name"] = None
                clear_photos()
            list_days(dep)

    # Обновляем команду динамически при открытии меню
    def _show_day_menu(event):
        dep = current_dep["name"]
        if not dep:
            return

        idx = day_list.nearest(event.y)
        if idx is None:
            return

        # выделяем элемент под курсором
        try:
            day_list.selection_clear(0, tk.END)
            day_list.selection_set(idx)
        except Exception:
            pass

        day_display = day_list.get(idx)

        day_menu.delete(0, tk.END)
        day_menu.add_command(
            label="🗑 Удалить папку дня",
            command=lambda d=day_display: _delete_selected_day(d)
        )
        day_menu.tk_popup(event.x_root, event.y_root)

    day_list.bind("<Button-3>", _show_day_menu)

    dep_list.bind("<<ListboxSelect>>", on_dep_select)
    day_list.bind("<<ListboxSelect>>", on_day_select)

    # первичная загрузка
    refresh_all()
