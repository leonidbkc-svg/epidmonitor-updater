import os
import time
import requests
import urllib.parse
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime

from config.app_config import load_config

_SESSION = None
_BASE_URL = None
_AUTH = None
_LAST_ERROR = None
_SYNC_DONE = False


def get_default_local_root() -> str:
    return os.path.join(
        os.environ.get("APPDATA", os.path.expanduser("~")),
        "EpidMonitor"
    )

def _init():
    global _SESSION, _BASE_URL, _AUTH
    if _SESSION is not None:
        return
    cfg = load_config()
    url = (cfg.get("webdav_url") or "").strip()
    user = (cfg.get("webdav_user") or "").strip()
    pwd = (cfg.get("webdav_password") or "")
    if not url:
        _SESSION = False
        return
    if not url.endswith("/"):
        url += "/"
    _BASE_URL = url
    _AUTH = (user, pwd) if user else None
    s = requests.Session()
    if _AUTH:
        s.auth = _AUTH
    _SESSION = s


def _url(path: str) -> str:
    base = _BASE_URL or ""
    p = (path or "").lstrip("/")
    return urllib.parse.urljoin(base, p)


def _propfind(path: str, depth: int = 0):
    _init()
    if _SESSION is False:
        raise RuntimeError("WebDAV is not configured")
    headers = {
        "Depth": str(depth),
        "Content-Type": "text/xml; charset=utf-8",
    }
    body = (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        "<D:propfind xmlns:D=\"DAV:\">"
        "<D:prop>"
        "<D:resourcetype/>"
        "<D:getlastmodified/>"
        "<D:getcontentlength/>"
        "</D:prop>"
        "</D:propfind>"
    )
    r = _SESSION.request("PROPFIND", _url(path), headers=headers, data=body, timeout=30)
    if r.status_code == 404:
        return []
    if r.status_code not in (207, 200):
        raise RuntimeError(f"PROPFIND {path}: {r.status_code} {r.text[:200]}")
    try:
        root = ET.fromstring(r.text)
    except Exception:
        return []
    ns = {"D": "DAV:"}
    items = []
    for resp in root.findall("D:response", ns):
        href = resp.findtext("D:href", default="", namespaces=ns)
        href = urllib.parse.unquote(urllib.parse.urlparse(href).path or "")
        if href.startswith("/"):
            href = href[1:]
        # remove base path if any
        if _BASE_URL:
            base_path = urllib.parse.urlparse(_BASE_URL).path.lstrip("/")
            if base_path and href.startswith(base_path):
                href = href[len(base_path):].lstrip("/")

        res_type = resp.find(".//D:resourcetype", ns)
        is_dir = res_type is not None and res_type.find("D:collection", ns) is not None

        lastmod = resp.findtext(".//D:getlastmodified", default="", namespaces=ns)
        size = resp.findtext(".//D:getcontentlength", default="", namespaces=ns)

        items.append({
            "path": href.rstrip("/"),
            "is_dir": is_dir,
            "lastmod": lastmod,
            "size": int(size) if str(size).isdigit() else None,
        })
    return items


def _ensure_remote_dir(path: str):
    _init()
    if _SESSION is False:
        return
    parts = [p for p in (path or "").split("/") if p]
    cur = ""
    for p in parts:
        cur = f"{cur}/{p}" if cur else p
        r = _SESSION.request("MKCOL", _url(cur), timeout=20)
        if r.status_code in (201, 405, 301, 302, 204):
            continue


def _download(path: str, local_path: str):
    _init()
    if _SESSION is False:
        return
    r = _SESSION.get(_url(path), stream=True, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"GET {path}: {r.status_code}")
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as f:
        for chunk in r.iter_content(1024 * 1024):
            if chunk:
                f.write(chunk)


def _upload(local_path: str, remote_path: str):
    _init()
    if _SESSION is False:
        return
    _ensure_remote_dir(os.path.dirname(remote_path))
    with open(local_path, "rb") as f:
        r = _SESSION.put(_url(remote_path), data=f, timeout=60)
    if r.status_code not in (200, 201, 204):
        raise RuntimeError(f"PUT {remote_path}: {r.status_code} {r.text[:200]}")


def _delete(remote_path: str):
    _init()
    if _SESSION is False:
        return
    r = _SESSION.request("DELETE", _url(remote_path), timeout=30)
    if r.status_code not in (200, 204, 404):
        raise RuntimeError(f"DELETE {remote_path}: {r.status_code} {r.text[:200]}")


def _relpath(local_path: str, local_root: str) -> str:
    rel = os.path.relpath(local_path, local_root)
    return rel.replace("\\", "/")


def get_last_error() -> str | None:
    return _LAST_ERROR


def sync_down(local_root: str) -> bool:
    global _LAST_ERROR, _SYNC_DONE
    _init()
    if _SESSION is False:
        return False
    try:
        os.makedirs(local_root, exist_ok=True)
        items = _propfind("", depth=1)
        # build list of dirs/files recursively
        queue = [i for i in items if i.get("is_dir")]
        files = [i for i in items if not i.get("is_dir")]

        while queue:
            d = queue.pop(0)
            sub = _propfind(d["path"], depth=1)
            for it in sub:
                if it["path"] == d["path"]:
                    continue
                if it.get("is_dir"):
                    queue.append(it)
                else:
                    files.append(it)

        # ensure directories
        for d in [i for i in items if i.get("is_dir")] + [i for i in queue if i.get("is_dir")]:
            if not d["path"]:
                continue
            os.makedirs(os.path.join(local_root, d["path"]), exist_ok=True)

        # download files if missing or older
        for f in files:
            rel = f["path"]
            if not rel:
                continue
            local_path = os.path.join(local_root, rel)
            remote_ts = None
            if f.get("lastmod"):
                try:
                    remote_ts = parsedate_to_datetime(f["lastmod"]).timestamp()
                except Exception:
                    remote_ts = None
            need = True
            if os.path.exists(local_path) and remote_ts:
                local_ts = os.path.getmtime(local_path)
                if local_ts >= remote_ts - 1:
                    need = False
            if need:
                _download(rel, local_path)
                if remote_ts:
                    os.utime(local_path, (remote_ts, remote_ts))

        _LAST_ERROR = None
        _SYNC_DONE = True
        return True
    except Exception as e:
        _LAST_ERROR = str(e)
        return False


def ensure_synced(local_root: str) -> bool:
    global _SYNC_DONE
    if _SYNC_DONE:
        return True
    return sync_down(local_root)


def upload_file(local_path: str, local_root: str) -> None:
    _init()
    if _SESSION is False:
        return
    try:
        rel = _relpath(local_path, local_root)
        _upload(local_path, rel)
    except Exception as e:
        global _LAST_ERROR
        _LAST_ERROR = str(e)


def delete_path(local_path: str, local_root: str) -> None:
    _init()
    if _SESSION is False:
        return
    try:
        rel = _relpath(local_path, local_root)
        _delete(rel)
    except Exception as e:
        global _LAST_ERROR
        _LAST_ERROR = str(e)
