"""Discover API endpoints used by JavaScript-rendered cost information sites."""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlparse

from .browser_capture import (
    _CdpSession,
    _find_available_port,
    _get_browser_target,
    _stop_browser_process,
    find_chromium_browsers,
)
from .config import config


SENSITIVE_KEYWORDS = ("password", "passwd", "token", "secret", "cookie", "authorization", "api_key", "apikey")
COST_ENDPOINT_KEYWORDS = (
    "hyxx", "xxj", "clj", "price", "cost", "material", "zaojia", "download",
    "file", "query", "search", "list", "info", "api", "信息价", "材料价", "造价",
)


def _is_sensitive_key(key: str) -> bool:
    lowered = key.lower()
    return any(keyword in lowered for keyword in SENSITIVE_KEYWORDS)


def _redact_value(value, key: str = ""):
    if key and _is_sensitive_key(key):
        return "[REDACTED]"
    if isinstance(value, dict):
        return {item_key: _redact_value(item_value, str(item_key)) for item_key, item_value in value.items()}
    if isinstance(value, list):
        return [_redact_value(item) for item in value[:100]]
    return value


def sanitize_post_data(post_data: str) -> str:
    """Keep request shape useful for adapters without persisting credentials."""
    text = (post_data or "").strip()
    if not text:
        return ""
    try:
        parsed = json.loads(text)
        return json.dumps(_redact_value(parsed), ensure_ascii=False)[:8000]
    except (TypeError, ValueError, json.JSONDecodeError):
        pass
    if "=" in text:
        try:
            fields = [
                (key, "[REDACTED]" if _is_sensitive_key(key) else value)
                for key, value in parse_qsl(text, keep_blank_values=True)
            ]
            if fields:
                return urlencode(fields, doseq=True)[:8000]
        except ValueError:
            pass
    redacted = re.sub(
        r"(?i)(password|passwd|token|secret|cookie|authorization|api[_-]?key)(\s*[:=]\s*)[^&\s,;}]+",
        r"\1\2[REDACTED]",
        text,
    )
    return redacted[:8000]


def _endpoint_score(endpoint: dict) -> int:
    searchable = " ".join([
        str(endpoint.get("url", "")),
        str(endpoint.get("post_data", "")),
        str(endpoint.get("mime_type", "")),
    ]).lower()
    score = sum(2 for keyword in COST_ENDPOINT_KEYWORDS if keyword.lower() in searchable)
    if endpoint.get("method") == "POST":
        score += 1
    if any(token in searchable for token in ("json", "excel", "spreadsheet", "octet-stream")):
        score += 1
    return score


def _deduplicate(endpoints: list[dict]) -> list[dict]:
    result = []
    seen = set()
    for endpoint in endpoints:
        key = (endpoint.get("method"), endpoint.get("url"), endpoint.get("post_data"))
        if key in seen:
            continue
        seen.add(key)
        result.append(endpoint)
    return result


def sniff_apis(url: str, timeout: int = 30, progress_callback=None, should_stop=None) -> dict:
    """Launch a local headless Chromium browser and capture XHR/fetch requests."""
    parsed = urlparse(url.strip())
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        return {"error": "请输入完整的 http/https 网站地址"}
    browsers = find_chromium_browsers()
    if not browsers:
        return {"error": "未找到 Microsoft Edge 或 Google Chrome"}

    errors = []
    for browser in browsers:
        result = _sniff_with_browser(browser, url, timeout, progress_callback, should_stop)
        if not result.get("error"):
            return result
        errors.append(f"{browser.name}: {result['error']}")
        if should_stop and should_stop():
            break
    return {"error": "；".join(errors)[:500] or "API 嗅探已取消"}


def _sniff_with_browser(browser: Path, url: str, timeout: int, progress_callback, should_stop) -> dict:
    cache_root = config.USER_DATA_DIR / "browser_cache"
    cache_root.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="api-sniff-", dir=cache_root))
    process = None
    socket = None
    try:
        if progress_callback:
            progress_callback("正在启动浏览器")
        port = _find_available_port()
        creation_flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        process = subprocess.Popen(
            [
                str(browser),
                "--headless=new",
                "--disable-gpu",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-background-networking",
                "--disable-component-update",
                "--remote-allow-origins=*",
                f"--remote-debugging-port={port}",
                f"--user-data-dir={profile}",
                "about:blank",
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creation_flags,
        )
        target = _get_browser_target(port, process, min(12, timeout))

        from websocket import create_connection

        socket = create_connection(
            target["webSocketDebuggerUrl"].replace("localhost", "127.0.0.1"),
            timeout=5,
            http_proxy_host=None,
        )
        cdp = _CdpSession(socket)
        cdp.command("Page.enable")
        cdp.command("Runtime.enable")
        cdp.command("Network.enable", {"maxTotalBufferSize": 50 * 1024 * 1024})
        cdp.command("Page.navigate", {"url": url})
        if progress_callback:
            progress_callback("网页已打开，正在拦截接口请求")

        captured_by_id = {}
        load_seen_at = None
        last_activity = time.monotonic()
        started_at = time.monotonic()
        deadline = started_at + max(8, timeout)
        while time.monotonic() < deadline:
            if should_stop and should_stop():
                return {"error": "API 嗅探已取消"}
            event = cdp.next_event(min(0.5, deadline - time.monotonic()))
            if event is None:
                if load_seen_at and time.monotonic() - last_activity >= 5 and time.monotonic() - started_at >= 8:
                    break
                continue
            method = event.get("method", "")
            params = event.get("params", {})
            if method == "Page.loadEventFired":
                load_seen_at = time.monotonic()
            elif method == "Network.requestWillBeSent" and params.get("type") in {"XHR", "Fetch"}:
                request = params.get("request", {})
                headers = request.get("headers", {})
                captured_by_id[params.get("requestId", "")] = {
                    "method": request.get("method", "GET"),
                    "url": request.get("url", ""),
                    "resource_type": params.get("type", ""),
                    "post_data": sanitize_post_data(request.get("postData", "")),
                    "content_type": headers.get("Content-Type", headers.get("content-type", "")),
                    "status": 0,
                    "mime_type": "",
                }
                last_activity = time.monotonic()
            elif method == "Network.responseReceived":
                request_id = params.get("requestId", "")
                endpoint = captured_by_id.get(request_id)
                if endpoint is not None:
                    response = params.get("response", {})
                    endpoint["status"] = int(response.get("status", 0) or 0)
                    endpoint["mime_type"] = response.get("mimeType", "")
                    last_activity = time.monotonic()

        final_url = cdp.command(
            "Runtime.evaluate",
            {"expression": "location.href", "returnByValue": True},
        ).get("result", {}).get("value", url)
        captured = _deduplicate([endpoint for endpoint in captured_by_id.values() if endpoint.get("url")])
        relevant = sorted(
            (endpoint for endpoint in captured if _endpoint_score(endpoint) >= 2),
            key=_endpoint_score,
            reverse=True,
        )
        return {
            "total_captured": len(captured),
            "relevant": relevant,
            "all": captured,
            "browser": browser.name,
            "page_url": final_url,
            "captured_at": datetime.now().isoformat(timespec="seconds"),
        }
    except Exception as error:
        return {"error": str(error)[:500], "browser": browser.name}
    finally:
        if socket is not None:
            try:
                socket.close()
            except Exception:
                pass
        if process is not None:
            _stop_browser_process(process)
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    import sys

    target_url = sys.argv[1] if len(sys.argv) > 1 else "https://ciac.zjw.sh.gov.cn/JGBXMGCZJInterWeb/pc/#/HyxxHynr?bmCode=003002"
    print(json.dumps(sniff_apis(target_url, timeout=20), ensure_ascii=False, indent=2))
