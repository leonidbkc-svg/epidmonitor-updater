from __future__ import annotations

import json
from typing import Iterable, List, Dict, Generator, Optional

import requests


def _chat_url(base_url: str) -> str:
    base = (base_url or "").rstrip("/")
    if base.endswith("/v1"):
        return base + "/chat/completions"
    return base + "/v1/chat/completions"


def stream_chat(
    *,
    base_url: str,
    api_key: str,
    agent_id: str,
    messages: List[Dict[str, str]],
    timeout_s: int = 120,
) -> Generator[str, None, None]:
    """
    Yields chunks of assistant text. Falls back to non-streaming if stream isn't supported.
    """
    url = _chat_url(base_url)
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": "timeweb-agent",
        "messages": messages,
        "stream": True,
        "agent_id": agent_id,
    }

    try:
        with requests.post(url, headers=headers, json=payload, stream=True, timeout=timeout_s) as r:
            r.encoding = "utf-8"
            r.raise_for_status()
            got_any = False
            for raw in r.iter_lines(decode_unicode=True):
                if not raw:
                    continue
                if raw.startswith("data:"):
                    data = raw[len("data:"):].strip()
                else:
                    data = raw.strip()

                if data == "[DONE]":
                    return
                try:
                    obj = json.loads(data)
                except Exception:
                    continue

                choices = obj.get("choices") or []
                if not choices:
                    continue
                delta = choices[0].get("delta") or {}
                content = delta.get("content")
                if content:
                    got_any = True
                    yield content

            if got_any:
                return
    except Exception:
        pass

    # Fallback: non-streaming
    payload["stream"] = False
    r = requests.post(url, headers=headers, json=payload, timeout=timeout_s)
    r.encoding = "utf-8"
    r.raise_for_status()
    obj = r.json()
    text = ""
    try:
        text = obj["choices"][0]["message"]["content"]
    except Exception:
        text = json.dumps(obj, ensure_ascii=False)

    if text:
        # yield in chunks for "typing" feel
        step = 40
        for i in range(0, len(text), step):
            yield text[i:i + step]
