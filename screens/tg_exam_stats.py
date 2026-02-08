import os
import json
import hashlib
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from services.tg_exam_stats import fetch_new_results, compute_stats


def open_tg_exam_stats(
    main_frame,
    build_header,
    go_back_callback,
    data_root: str,
    base_url: str,
    report_api_key: str,
    refresh_ms: int = 60000,
):
    # --------------------------
    # helpers: files
    # --------------------------
    cache_path = os.path.join(data_root, "tg_exam_cache.jsonl")
    deleted_path = os.path.join(data_root, "tg_exam_deleted.json")

    def _ensure_data_root():
        os.makedirs(data_root, exist_ok=True)

    def _load_deleted_hashes() -> set[str]:
        if not os.path.exists(deleted_path):
            return set()
        try:
            with open(deleted_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return set(str(x) for x in data)
            if isinstance(data, dict) and isinstance(data.get("hashes"), list):
                return set(str(x) for x in data["hashes"])
        except Exception:
            pass
        return set()

    def _save_deleted_hashes(hashes: set[str]) -> None:
        tmp = deleted_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"hashes": sorted(hashes)}, f, ensure_ascii=False, indent=2)
        os.replace(tmp, deleted_path)

    def _stable_hash(obj: dict) -> str:
        """
        Стабильный хеш записи (не зависит от порядка ключей).
        Если сервер после обновления снова отдаст ту же запись — хеш совпадет,
        и мы ее не покажем/не оставим в кеше.
        """
        try:
            s = json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        except Exception:
            s = str(obj)
        return hashlib.sha1(s.encode("utf-8")).hexdigest()

    def _read_cache_lines():
        """
        Возвращает список строк исходного jsonl: [(line_idx, raw_line, obj, hash), ...]
        """
        if not os.path.exists(cache_path):
            return []
        out = []
        with open(cache_path, "r", encoding="utf-8") as f:
            for idx, line in enumerate(f):
                raw = line.rstrip("\n")
                if not raw.strip():
                    continue
                try:
                    obj = json.loads(raw)
                except Exception:
                    continue
                h = _stable_hash(obj)
                out.append((idx, raw, obj, h))
        return out

    def _rewrite_cache_keep_hashes(keep_hashes: set[str]) -> None:
        """
        Переписывает кеш, оставляя только записи, чьи hash входят в keep_hashes.
        """
        lines = _read_cache_lines()
        keep_raw = [raw + "\n" for (_, raw, _, h) in lines if h in keep_hashes]

        tmp = cache_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.writelines(keep_raw)
        os.replace(tmp, cache_path)

    def _purge_deleted_from_cache(deleted_hashes: set[str]) -> None:
        """
        Гарантирует, что удаленные записи физически удалены из jsonl,
        чтобы после перезапуска/обновления они не всплывали.
        """
        lines = _read_cache_lines()
        keep = [raw + "\n" for (_, raw, _, h) in lines if h not in deleted_hashes]

        tmp = cache_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.writelines(keep)
        os.replace(tmp, cache_path)

    # --------------------------
    # UI: clear + header
    # --------------------------
    for w in main_frame.winfo_children():
        w.destroy()

    build_header(main_frame, back_callback=go_back_callback)

    title = tk.Label(main_frame, text="📊 Статистика tg-exam", font=("Segoe UI", 16, "bold"), bg="#f4f6f8")
    title.pack(pady=(16, 6))

    top = tk.Frame(main_frame, bg="#f4f6f8")
    top.pack(fill="x", padx=16)

    status_var = tk.StringVar(value="Готово")
    tk.Label(top, textvariable=status_var, bg="#f4f6f8", fg="#374151").pack(side="left")

    # кнопки справа
    btn_refresh = ttk.Button(top, text="🔄 Обновить", style="Secondary.TButton")
    btn_refresh.pack(side="right")

    btn_delete = ttk.Button(top, text="🗑 Удалить отмеченные", style="Secondary.TButton")
    btn_delete.pack(side="right", padx=(0, 8))

    btn_clear_marks = ttk.Button(top, text="☐ Снять отметки", style="Secondary.TButton")
    btn_clear_marks.pack(side="right", padx=(0, 8))

    # --------------------------
    # KPI row
    # --------------------------
    kpi = tk.Frame(main_frame, bg="#f4f6f8")
    kpi.pack(fill="x", padx=16, pady=(10, 0))

    k_total = tk.StringVar(value="—")
    k_passed = tk.StringVar(value="—")

    def _kpi_card(parent, label, var):
        box = tk.Frame(parent, bg="white", bd=1, relief="solid")
        box.pack(side="left", padx=6, pady=6, ipadx=10, ipady=6)
        tk.Label(box, text=label, bg="white", fg="#6b7280", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(box, textvariable=var, bg="white", fg="#111827", font=("Segoe UI", 14, "bold")).pack(anchor="w")

    _kpi_card(kpi, "Попыток", k_total)
    _kpi_card(kpi, "Сдано", k_passed)

    # --------------------------
    # layout: table + chart
    # --------------------------
    body = tk.Frame(main_frame, bg="#f4f6f8")
    body.pack(expand=True, fill="both", padx=16, pady=12)

    left = tk.Frame(body, bg="#f4f6f8")
    left.pack(side="left", fill="both", expand=True)

    # --------------------------
    # Tree: чекбоксы через символы
    # --------------------------
    # первая колонка = "точечки/галочки"
    cols = ("mark", "fio", "score", "percent", "passed", "reason", "leave", "dur")
    tree = ttk.Treeview(left, columns=cols, show="headings", height=18)
    tree.pack(fill="both", expand=True)

    heads = {
        "mark": "✓",
        "fio": "ФИО",
        "score": "Баллы",
        "percent": "%",
        "passed": "Сдан",
        "reason": "Причина",
        "leave": "Уходы",
        "dur": "Сек",
    }

    tree.heading("mark", text=heads["mark"])
    tree.column("mark", width=40, anchor="center")

    tree.heading("fio", text=heads["fio"])
    tree.column("fio", width=240, anchor="w")

    for c in ("score", "percent", "passed", "reason", "leave", "dur"):
        tree.heading(c, text=heads[c])
        if c == "reason":
            tree.column(c, width=160, anchor="w")
        elif c in ("score",):
            tree.column(c, width=90, anchor="w")
        else:
            tree.column(c, width=70, anchor="center")

    # --------------------------
    # state: marks + deleted
    # --------------------------
    _ensure_data_root()
    deleted_hashes = _load_deleted_hashes()
    marked_hashes: set[str] = set()  # что пользователь "точечками" отметил на удаление

    # --------------------------
    # reading visible rows
    # --------------------------
    def _read_last_n_visible(n=30):
        """
        Возвращает последние n записей из кеша, исключая удаленные (по deleted_hashes).
        Формат: [(obj, hash), ...]
        """
        lines = _read_cache_lines()
        visible = [(obj, h) for (_, _, obj, h) in lines if h not in deleted_hashes]
        return visible[-n:]

    def _fmt_score(score, mx):
        if score is None and mx is None:
            return ""
        if mx is None:
            return str(score)
        return f"{score}/{mx}"

    # --------------------------
    # render
    # --------------------------
    def render():
        # KPI считаем из compute_stats (оно читает кеш),
        # но т.к. мы чистим кеш от удаленных — цифры будут согласованы.
        st = compute_stats(data_root)

        k_total.set(str(st.get("total_attempts", "—")))
        k_passed.set(str(st.get("passed_count", "—")))

        # table
        for i in tree.get_children():
            tree.delete(i)

        rows = _read_last_n_visible(30)
        # показываем последние сверху
        for obj, h in rows[::-1]:
            meta = obj.get("meta") or {}
            fio = meta.get("fio") or ""
            score = obj.get("score")
            mx = obj.get("max_score")
            pct = obj.get("percent")
            passed = "✅" if obj.get("passed") else "❌"
            reason = (meta.get("reason") or obj.get("meta", {}).get("reason") or "").strip()
            leave = meta.get("leaveCount") or 0
            dur = obj.get("duration_sec") or ""

            mark_symbol = "☑" if h in marked_hashes else "☐"

            # iid = hash, чтобы легко удалять/отмечать
            tree.insert(
                "", "end",
                iid=h,
                values=(mark_symbol, fio, _fmt_score(score, mx), pct, passed, reason, leave, dur)
            )

        # chart removed

    # --------------------------
    # details on double-click
    # --------------------------
    def _find_record_by_hash(target_hash: str):
        for obj, h in _read_last_n_visible(300):
            if h == target_hash:
                return obj
        lines = _read_cache_lines()
        for (_, _, obj, h) in lines:
            if h == target_hash:
                return obj
        return None

    def _format_ts(ts):
        try:
            dt = datetime.fromtimestamp(int(ts) / 1000.0)
            return dt.strftime("%H:%M:%S %d.%m.%Y")
        except Exception:
            return ""

    def _format_duration(sec):
        try:
            total = int(sec)
        except Exception:
            return ""
        if total < 0:
            total = 0
        minutes = total // 60
        seconds = total % 60
        return f"{minutes} мин {seconds} сек"

    def _build_details_text(obj: dict) -> str:
        meta = obj.get("meta") or {}
        lines = []
        lines.append("Результат тестирования")
        lines.append("")
        lines.append(f"ФИО: {meta.get('fio') or ''}")
        lines.append(f"Дата: {_format_ts(obj.get('ts')) or obj.get('date_iso','')}")
        lines.append(f"Тест: {obj.get('exam_title') or obj.get('exam_id') or ''}")
        lines.append(f"Баллы: {obj.get('score')}/{obj.get('max_score')}")
        lines.append(f"Процент: {obj.get('percent')}%")
        lines.append(f"Сдан: {'Да' if obj.get('passed') else 'Нет'}")
        lines.append(f"Длительность: {_format_duration(obj.get('duration_sec'))}")
        lines.append(f"Нарушения: {meta.get('leaveCount',0)} выход(-ов)")
        if meta.get("reason"):
            lines.append(f"Причина: {meta.get('reason')}")
        lines.append("")
        lines.append("Ответы:")
        answers = obj.get("answers") or []
        if not answers:
            lines.append("— нет данных")
        else:
            for i, a in enumerate(answers, 1):
                q = a.get("question") or a.get("question_text") or ""
                options = a.get("options") or []
                selected_ids = a.get("selected_ids") or []
                selected_texts = a.get("selected_texts") or []

                lines.append(f"{i}. {q}")
                if options:
                    lines.append("   Варианты:")
                    for idx, opt in enumerate(options, 1):
                        if isinstance(opt, dict):
                            text = opt.get("text") or opt.get("label") or ""
                            oid = opt.get("id") or opt.get("key") or ""
                            mark = " [выбран]" if (oid and oid in selected_ids) or (text and text in selected_texts) else ""
                            lines.append(f"     {idx}) {text}{mark}")
                        else:
                            text = str(opt)
                            mark = " [выбран]" if text in selected_texts else ""
                            lines.append(f"     {idx}) {text}{mark}")
                if selected_texts:
                    lines.append(f"   Выбрано: {', '.join([str(x) for x in selected_texts])}")
                elif selected_ids:
                    lines.append(f"   Выбрано (id): {', '.join([str(x) for x in selected_ids])}")
                else:
                    lines.append("   Выбрано: нет ответа")
                lines.append("")
        return "\n".join(lines)

    def _open_details_window(obj: dict):
        win = tk.Toplevel(main_frame)
        win.title("Ответы")
        win.geometry("900x700")
        win.configure(bg="white")

        txt = tk.Text(win, wrap="word", bg="white", fg="#111827", font=("Consolas", 10))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", _build_details_text(obj))
        txt.configure(state="disabled")

    def on_double_click(event):
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        obj = _find_record_by_hash(row_id)
        if not obj:
            messagebox.showinfo("Ответы", "Запись не найдена.")
            return
        _open_details_window(obj)

    tree.bind("<Double-1>", on_double_click, add="+")

    # --------------------------
    # mark toggle on click
    # --------------------------
    def toggle_mark(event):
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)  # '#1' - первая колонка
        if not row_id:
            return
        if col_id != "#1":
            return  # отмечаем только кликом по колонке "✓"

        h = row_id  # iid = hash
        if h in marked_hashes:
            marked_hashes.remove(h)
            tree.set(h, "mark", "☐")
        else:
            marked_hashes.add(h)
            tree.set(h, "mark", "☑")

    tree.bind("<Button-1>", toggle_mark, add="+")

    # --------------------------
    # actions
    # --------------------------
    def clear_marks():
        marked_hashes.clear()
        # обновим видимые символы
        for iid in tree.get_children():
            tree.set(iid, "mark", "☐")

    def delete_marked():
        nonlocal deleted_hashes

        if not marked_hashes:
            messagebox.showinfo("Удаление", "Ничего не отмечено.")
            return

        cnt = len(marked_hashes)
        if not messagebox.askyesno(
            "Удаление результатов",
            f"Удалить отмеченные записи: {cnt} шт.?\n\nОперация необратима."
        ):
            return

        # 1) добавляем в deleted список
        deleted_hashes = set(deleted_hashes) | set(marked_hashes)
        _save_deleted_hashes(deleted_hashes)

        # 2) чистим кеш физически (чтобы после refresh/перезапуска не всплывало)
        try:
            _purge_deleted_from_cache(deleted_hashes)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить кеш:\n{e}")
            return

        # 3) сбрасываем отметки и перерисовываем
        marked_hashes.clear()
        status_var.set(f"Удалено: {cnt}")
        render()

    btn_clear_marks.config(command=clear_marks)
    btn_delete.config(command=delete_marked)

    # --------------------------
    # refresh (network) + auto refresh
    # --------------------------
    def refresh():
        btn_refresh.config(state="disabled")
        status_var.set("Обновляем...")

        try:
            added, _cur = fetch_new_results(
                base_url=base_url,
                report_api_key=report_api_key,
                data_root=data_root
            )

            # важно: даже если сервер снова прислал старые записи,
            # мы их выкинем по deleted_hashes (и из файла, и из отображения).
            try:
                _purge_deleted_from_cache(deleted_hashes)
            except Exception:
                pass

            status_var.set(f"Готово. Новых: {added}.")
            render()
        except Exception as e:
            status_var.set("Ошибка обновления")
            messagebox.showerror("tg-exam", str(e))
        finally:
            btn_refresh.config(state="normal")

        # автообновление
        main_frame.after(refresh_ms, refresh)

    btn_refresh.config(command=refresh)

    # --------------------------
    # first render (offline) + immediate refresh
    # --------------------------
    try:
        # на всякий: при открытии экрана чистим кеш от уже удаленных
        _purge_deleted_from_cache(deleted_hashes)
    except Exception:
        pass

    try:
        render()
    except Exception:
        pass

    refresh()
