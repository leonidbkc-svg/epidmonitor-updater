import os
import sys
import json
import shutil
import hashlib
import zipfile
import tempfile
import subprocess
import urllib.request
import threading
import queue
import time
from pathlib import Path

OWNER = "leonidbkc-svg"
REPO = "epidmonitor-updater"


BASE_URL = f"https://github.com/{OWNER}/{REPO}/releases/latest/download/"
MANIFEST_URL = BASE_URL + "manifest.json"

APP_ROOT = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "EpidMonitor"
APP_DIR = APP_ROOT / "app"
VERSION_FILE = APP_ROOT / "version.txt"


def _req(url: str):
    return urllib.request.Request(url, headers={"User-Agent": "EpidLauncher/1.0"})


def download(url: str, dest: Path, on_progress=None):
    """
    on_progress: callable(read_bytes: int, total_bytes: int | None) -> None
    """
    with urllib.request.urlopen(_req(url), timeout=60) as r, open(dest, "wb") as f:
        total = None
        try:
            cl = r.headers.get("Content-Length")
            total = int(cl) if cl else None
        except Exception:
            total = None

        read_bytes = 0
        while True:
            chunk = r.read(1024 * 256)
            if not chunk:
                break
            f.write(chunk)
            read_bytes += len(chunk)
            if on_progress:
                on_progress(read_bytes, total)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_local_version() -> str | None:
    if VERSION_FILE.exists():
        return VERSION_FILE.read_text(encoding="utf-8").strip()
    return None


def write_local_version(ver: str):
    APP_ROOT.mkdir(parents=True, exist_ok=True)
    VERSION_FILE.write_text(ver, encoding="utf-8")


def run_exe(exe_path: Path, cwd: Path):
    return subprocess.Popen([str(exe_path)], cwd=str(cwd))

def _find_local_exe():
    try:
        local_ver = read_local_version()
    except Exception:
        local_ver = None

    if local_ver:
        target = APP_DIR / local_ver
        if target.exists():
            exes = list(target.rglob("*.exe"))
            if exes:
                return exes[0], target

    if APP_DIR.exists():
        try:
            dirs = sorted(
                [p for p in APP_DIR.iterdir() if p.is_dir()],
                key=lambda p: p.stat().st_mtime,
                reverse=True
            )
        except Exception:
            dirs = []

        for d in dirs:
            exes = list(d.rglob("*.exe"))
            if exes:
                return exes[0], d

    return None, None


def _show_error(title: str, message: str):
    try:
        import tkinter as tk
        from tkinter import messagebox as mb

        root = tk.Tk()
        root.withdraw()
        try:
            mb.showerror(title, message, parent=root)
        finally:
            root.destroy()
    except Exception:
        print(f"{title}: {message}", file=sys.stderr)


class SplashUI:
    def __init__(self, title: str = "EpidMonitor — запуск"):
        try:
            import tkinter as tk
            from tkinter import ttk

            self._tk = tk
            self._ttk = ttk
        except Exception:
            self.enabled = False
            return

        self.enabled = True
        self._queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self._close_after_ms: int | None = None
        self._indeterminate_running = False

        self.root = self._tk.Tk()
        self.root.title(title)
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        pad = 14
        frm = self._ttk.Frame(self.root, padding=pad)
        frm.grid(row=0, column=0, sticky="nsew")

        self._status_var = self._tk.StringVar(value="Запуск...")
        lbl = self._ttk.Label(frm, textvariable=self._status_var, anchor="w")
        lbl.grid(row=0, column=0, sticky="ew")

        self._progress = self._ttk.Progressbar(frm, mode="indeterminate", length=360)
        self._progress.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        frm.columnconfigure(0, weight=1)

        self._center(420, 120)
        self._progress.start(12)
        self._indeterminate_running = True

        self.root.after(50, self._drain_queue)

    def _center(self, w: int, h: int):
        try:
            self.root.update_idletasks()
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = int((sw - w) / 2)
            y = int((sh - h) / 2)
            self.root.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def post_status(self, text: str):
        if self.enabled:
            self._queue.put(("status", text))

    def post_progress(self, read_bytes: int, total_bytes: int | None):
        if self.enabled:
            self._queue.put(("progress", (read_bytes, total_bytes)))

    def request_close(self, after_ms: int = 0):
        if self.enabled:
            self._queue.put(("close", after_ms))

    def _set_indeterminate(self):
        if not self._indeterminate_running:
            self._progress.config(mode="indeterminate")
            self._progress.start(12)
            self._indeterminate_running = True

    def _set_determinate(self, value: int, maximum: int):
        if self._indeterminate_running:
            self._progress.stop()
            self._indeterminate_running = False
        self._progress.config(mode="determinate", maximum=max(1, maximum), value=value)

    def _drain_queue(self):
        if not self.enabled:
            return

        while True:
            try:
                kind, payload = self._queue.get_nowait()
            except queue.Empty:
                break

            if kind == "status":
                self._status_var.set(str(payload))
            elif kind == "progress":
                read_bytes, total_bytes = payload  # type: ignore[misc]
                if total_bytes:
                    self._set_determinate(int(read_bytes), int(total_bytes))
                else:
                    self._set_indeterminate()
            elif kind == "close":
                self._close_after_ms = int(payload)  # type: ignore[arg-type]
                if self._close_after_ms <= 0:
                    try:
                        self.root.destroy()
                    except Exception:
                        pass
                    return

        if self._close_after_ms is not None:
            self._close_after_ms -= 50
            if self._close_after_ms <= 0:
                try:
                    self.root.destroy()
                except Exception:
                    pass
                return

        self.root.after(50, self._drain_queue)

    def mainloop(self):
        if self.enabled:
            self.root.mainloop()


def _ensure_updated(status, progress):
    status("Проверка обновлений...")
    with urllib.request.urlopen(_req(MANIFEST_URL), timeout=30) as r:
        manifest = json.load(r)

    remote_ver = str(manifest["version"]).strip()
    zip_name = str(manifest["zip"]).strip()
    exe_name = str(manifest["exe_relpath"]).strip()
    expected_hash = str(manifest["sha256"]).strip().lower()

    local_ver = read_local_version()
    target_dir = APP_DIR / remote_ver
    exe_path = target_dir / exe_name

    need_update = (local_ver != remote_ver) or (not exe_path.exists())
    if not need_update:
        return exe_path, target_dir, remote_ver

    status("Загрузка обновления...")
    tmpdir = Path(tempfile.mkdtemp(prefix="epid_launcher_"))
    zip_path = tmpdir / zip_name

    def _format_eta(seconds: float | None) -> str:
        if seconds is None or seconds < 0 or seconds != seconds:
            return "--:--:--"
        sec = int(max(0, seconds))
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    last_t = None
    last_b = None
    rate_ema = None
    last_ui_t = 0.0

    def _on_dl(read_bytes: int, total_bytes: int | None):
        nonlocal last_t, last_b, rate_ema, last_ui_t
        now = time.monotonic()
        if last_t is not None and last_b is not None:
            dt = now - last_t
            db = read_bytes - last_b
            if dt > 0 and db > 0:
                inst_rate = db / dt
                if rate_ema is None:
                    rate_ema = inst_rate
                else:
                    rate_ema = rate_ema * 0.8 + inst_rate * 0.2
        last_t = now
        last_b = read_bytes

        if total_bytes and rate_ema:
            remaining = total_bytes - read_bytes
            eta = remaining / max(rate_ema, 1e-6)
            if now - last_ui_t >= 0.25:
                status("\u0417\u0430\u0433\u0440\u0443\u0437\u043a\u0430 \u043e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u044f... \u043e\u0441\u0442\u0430\u043b\u043e\u0441\u044c " + _format_eta(eta))
                last_ui_t = now
        elif now - last_ui_t >= 0.5:
            status("\u0417\u0430\u0433\u0440\u0443\u0437\u043a\u0430 \u043e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u044f... \u043e\u0441\u0442\u0430\u043b\u043e\u0441\u044c --:--:--")
            last_ui_t = now
        progress(read_bytes, total_bytes)

    try:
        download(BASE_URL + zip_name, zip_path, on_progress=_on_dl)

        progress(0, None)
        status("Проверка целостности...")
        real_hash = sha256(zip_path).lower()
        if real_hash != expected_hash:
            raise RuntimeError(f"SHA256 не совпадает.\nОжидалось: {expected_hash}\nФакт: {real_hash}")

        progress(0, None)
        status("Установка обновления...")
        if target_dir.exists():
            shutil.rmtree(target_dir, ignore_errors=True)
        target_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(target_dir)

        if not exe_path.exists():
            raise RuntimeError(f"После распаковки не найден {exe_name} в {target_dir}")

        write_local_version(remote_ver)
        return exe_path, target_dir, remote_ver
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main():
    splash = SplashUI()

    if not splash.enabled:
        try:
            exe_path, target_dir, _ = _ensure_updated(lambda _: None, lambda *_: None)
            run_exe(exe_path, target_dir)
        except Exception as e:
            fallback_exe, fallback_dir = _find_local_exe()
            if fallback_exe and fallback_dir:
                run_exe(fallback_exe, fallback_dir)
                return
            _show_error("EpidLauncher — ошибка", str(e))
            sys.exit(1)
        return

    error_text: list[str] = []

    def worker():
        try:
            exe_path, target_dir, remote_ver = _ensure_updated(splash.post_status, splash.post_progress)
            splash.post_status(f"Запуск EpidMonitor {remote_ver}...")
            run_exe(exe_path, target_dir)
            splash.request_close(after_ms=1500)
        except Exception as e:
            fallback_exe, fallback_dir = _find_local_exe()
            if fallback_exe and fallback_dir:
                splash.post_status("Не удалось обновить. Запуск локальной версии...")
                run_exe(fallback_exe, fallback_dir)
                splash.request_close(after_ms=1500)
                return
            error_text.append(str(e))
            splash.request_close(after_ms=0)

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    splash.mainloop()

    if error_text:
        _show_error("EpidLauncher — ошибка", error_text[0])
        sys.exit(1)


if __name__ == "__main__":
    main()
