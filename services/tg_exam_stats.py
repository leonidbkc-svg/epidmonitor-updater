import json
import os
import time
from typing import Dict, Any, List, Tuple
import requests

def _read_json(path: str, default: Any):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def _write_json(path: str, obj: Any):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def _append_jsonl(path: str, rows: List[Dict[str, Any]]):
    with open(path, "a", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

def fetch_new_results(
    base_url: str,
    report_api_key: str,
    data_root: str,
    limit: int = 1000,
    timeout: int = 20,
) -> Tuple[int, int]:
    """
    Забирает только новые результаты с tg-exam.
    Возвращает (added_count, new_cursor).
    """
    os.makedirs(data_root, exist_ok=True)
    cursor_path = os.path.join(data_root, "tg_exam_cursor.json")
    cache_path = os.path.join(data_root, "tg_exam_cache.jsonl")

    cur = _read_json(cursor_path, {"from": 0})
    cursor = int(cur.get("from", 0) or 0)

    added_total = 0
    session = requests.Session()
    session.headers.update({"X-API-Key": report_api_key})

    while True:
        params = {"from": cursor, "limit": limit}
        url = base_url.rstrip("/") + "/api/admin/results"

        r = session.get(url, params=params, timeout=timeout)
        if r.status_code != 200:
            raise RuntimeError(f"API error {r.status_code}: {r.text[:300]}")

        payload = r.json()
        if not payload.get("ok"):
            raise RuntimeError(f"API not ok: {payload}")

        rows = payload.get("results") or []
        if not rows:
            break

        _append_jsonl(cache_path, rows)
        added_total += len(rows)

        # если сервер отдаёт next_from — используем, иначе считаем сами
        next_from = payload.get("next_from")
        if isinstance(next_from, (int, float)):
            cursor = int(next_from)
        else:
            last_ts = int(rows[-1].get("ts") or 0)
            cursor = last_ts + 1

        # safety: если внезапно сервер не поддерживает limit и вернул всё
        if len(rows) < limit:
            break

        time.sleep(0.05)

    _write_json(cursor_path, {"from": cursor, "updated_at": int(time.time())})
    return added_total, cursor

def _iter_cache(cache_path: str):
    if not os.path.exists(cache_path):
        return
    with open(cache_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                yield json.loads(line)
            except Exception:
                continue

def compute_stats(data_root: str) -> Dict[str, Any]:
    """
    Считает агрегаты по кэшу tg_exam_cache.jsonl.
    """
    cache_path = os.path.join(data_root, "tg_exam_cache.jsonl")

    total = 0
    passed = 0
    sum_percent = 0.0

    sum_score = 0
    sum_max = 0

    dur_n = 0
    sum_dur = 0

    # античит
    sum_blur = 0
    sum_hidden = 0
    sum_leave = 0

    any_violation = 0
    auto_finish = 0

    last_ts = 0

    for r in _iter_cache(cache_path):
        total += 1
        ts = int(r.get("ts") or 0)
        last_ts = max(last_ts, ts)

        if r.get("passed") is True:
            passed += 1

        pct = r.get("percent")
        if isinstance(pct, (int, float)):
            sum_percent += float(pct)

        sc = r.get("score")
        mx = r.get("max_score")
        if isinstance(sc, (int, float)) and isinstance(mx, (int, float)):
            sum_score += int(sc)
            sum_max += int(mx)

        dur = r.get("duration_sec")
        if isinstance(dur, (int, float)) and dur >= 0:
            dur_n += 1
            sum_dur += int(dur)

        meta = r.get("meta") or {}
        bc = int(meta.get("blurCount") or 0)
        hc = int(meta.get("hiddenCount") or 0)
        lc = int(meta.get("leaveCount") or 0)
        reason = str(meta.get("reason") or "")

        sum_blur += bc
        sum_hidden += hc
        sum_leave += lc

        if (bc + hc + lc) > 0:
            any_violation += 1

        if reason == "too_many_violations":
            auto_finish += 1

    avg_percent = round(sum_percent / total, 2) if total else 0.0
    pass_rate = round((passed / total) * 100, 2) if total else 0.0

    avg_dur = round(sum_dur / dur_n, 1) if dur_n else 0.0
    avg_blur = round(sum_blur / total, 2) if total else 0.0
    avg_hidden = round(sum_hidden / total, 2) if total else 0.0
    avg_leave = round(sum_leave / total, 2) if total else 0.0

    viol_rate = round((any_violation / total) * 100, 2) if total else 0.0
    auto_rate = round((auto_finish / total) * 100, 2) if total else 0.0

    return {
        "total_attempts": total,
        "passed_count": passed,
        "pass_rate_pct": pass_rate,
        "avg_percent": avg_percent,
        "avg_score": round(sum_score / total, 2) if total else 0.0,
        "avg_max": round(sum_max / total, 2) if total else 0.0,
        "avg_duration_sec": avg_dur,
        "violations_pct": viol_rate,
        "auto_finish_pct": auto_rate,
        "avg_blur": avg_blur,
        "avg_hidden": avg_hidden,
        "avg_leave": avg_leave,
        "last_ts": last_ts,
    }
