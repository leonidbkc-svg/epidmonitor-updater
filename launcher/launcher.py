import os
import sys
import json
import shutil
import hashlib
import zipfile
import tempfile
import subprocess
import urllib.request
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


def download(url: str, dest: Path):
    with urllib.request.urlopen(_req(url), timeout=60) as r, open(dest, "wb") as f:
        shutil.copyfileobj(r, f)


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
    subprocess.Popen([str(exe_path)], cwd=str(cwd))


def main():
    try:
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

        if need_update:
            tmpdir = Path(tempfile.mkdtemp(prefix="epid_launcher_"))
            zip_path = tmpdir / zip_name

            download(BASE_URL + zip_name, zip_path)

            real_hash = sha256(zip_path).lower()
            if real_hash != expected_hash:
                raise RuntimeError(f"SHA256 не совпадает.\nОжидалось: {expected_hash}\nФакт: {real_hash}")

            if target_dir.exists():
                shutil.rmtree(target_dir, ignore_errors=True)
            target_dir.mkdir(parents=True, exist_ok=True)

            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(target_dir)

            if not exe_path.exists():
                raise RuntimeError(f"После распаковки не найден {exe_name} в {target_dir}")

            write_local_version(remote_ver)

        run_exe(exe_path, target_dir)

    except Exception as e:
        try:
            import tkinter.messagebox as mb
            mb.showerror("EpidLauncher — ошибка", str(e))
        except Exception:
            print("EpidLauncher error:", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
