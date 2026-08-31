"""Capture JavaScript-rendered official pages through a local Chromium browser."""
from __future__ import annotations

import json
import os
import shutil
import socket as network_socket
import subprocess
import tempfile
import time
from collections import deque
from pathlib import Path
from urllib.request import ProxyHandler, build_opener

from .config import config


MAX_CAPTURED_BODY_BYTES = 25 * 1024 * 1024
MAX_CAPTURED_TOTAL_BYTES = 100 * 1024 * 1024


def find_chromium_browsers() -> list[Path]:
    configured = os.getenv("DASHUO_BROWSER_PATH", "").strip()
    if configured:
        path = Path(configured)
        return [path] if path.is_file() else []
    candidates = []
    if os.name == "nt":
        try:
            import winreg

            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                for executable in ("chrome.exe", "msedge.exe"):
                    try:
                        key = winreg.OpenKey(
                            hive,
                            rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{executable}",
                        )
                        candidates.append(Path(winreg.QueryValue(key, None)))
                    except OSError:
                        pass
        except ImportError:
            pass
        candidates.extend([
            Path(os.getenv("PROGRAMFILES", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.getenv("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.getenv("PROGRAMFILES(X86)", "")) / "Microsoft/Edge/Application/msedge.exe",
            Path(os.getenv("PROGRAMFILES", "")) / "Microsoft/Edge/Application/msedge.exe",
        ])
    for command in ("msedge", "google-chrome", "chromium", "chrome"):
        located = shutil.which(command)
        if located:
            candidates.append(Path(located))
    result = []
    seen = set()
    for path in candidates:
        normalized = str(path).lower()
        if path and path.is_file() and normalized not in seen:
            seen.add(normalized)
            result.append(path)
    return result


def find_chromium_browser() -> Path | None:
    browsers = find_chromium_browsers()
    return browsers[0] if browsers else None


class _CdpSession:
    def __init__(self, socket):
        self.socket = socket
        self.next_id = 1
        self.events = deque()

    def command(self, method: str, params: dict | None = None, timeout: float = 10.0) -> dict:
        command_id = self.next_id
        self.next_id += 1
        self.socket.send(json.dumps({"id": command_id, "method": method, "params": params or {}}))
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            message = self._receive(deadline - time.monotonic())
            if not message:
                continue
            if message.get("id") == command_id:
                if "error" in message:
                    raise RuntimeError(message["error"].get("message", f"CDP {method} failed"))
                return message.get("result", {})
            if "method" in message:
                self.events.append(message)
        raise TimeoutError(f"浏览器命令超时: {method}")

    def next_event(self, timeout: float = 1.0) -> dict | None:
        if self.events:
            return self.events.popleft()
        return self._receive(timeout)

    def _receive(self, timeout: float) -> dict | None:
        from websocket import WebSocketTimeoutException

        self.socket.settimeout(max(0.05, timeout))
        try:
            return json.loads(self.socket.recv())
        except WebSocketTimeoutException:
            return None


def capture_dynamic_page(
    url: str,
    timeout: int = 35,
    auto_scroll: bool = False,
    cookie_header: str = "",
) -> dict:
    """Return rendered HTML and selected network responses for one public page."""
    browsers = find_chromium_browsers()
    if not browsers:
        return {"success": False, "error": "未找到 Microsoft Edge 或 Google Chrome，无法加载动态官网"}

    errors = []
    for browser in browsers:
        result = _capture_with_browser(
            browser, url, timeout,
            auto_scroll=auto_scroll,
            cookie_header=cookie_header,
        )
        if result.get("success"):
            return result
        errors.append(f"{browser.name}: {result.get('error', '启动失败')}")
        rendered = _capture_dump_dom(browser, url, timeout)
        if rendered.get("success"):
            return rendered
        errors.append(f"{browser.name} 渲染回退: {rendered.get('error', '失败')}")
    return {"success": False, "error": "；".join(errors)[:500]}


def _capture_with_browser(
    browser: Path,
    url: str,
    timeout: int,
    *,
    auto_scroll: bool = False,
    cookie_header: str = "",
) -> dict:

    cache_root = config.USER_DATA_DIR / "browser_cache"
    cache_root.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="capture-", dir=cache_root))
    process = None
    socket = None
    try:
        creation_flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        port = _find_available_port()
        process = subprocess.Popen(
            [
                str(browser),
                "--headless=new",
                "--disable-gpu",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-background-networking",
                "--disable-component-update",
                f"--remote-debugging-port={port}",
                "--remote-allow-origins=*",
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
        cdp.command("Network.enable", {"maxTotalBufferSize": MAX_CAPTURED_TOTAL_BYTES})
        if cookie_header:
            for part in cookie_header.split(";"):
                name, separator, value = part.strip().partition("=")
                if separator and name:
                    cdp.command(
                        "Network.setCookie",
                        {"name": name, "value": value, "url": url},
                    )
        cdp.command("Page.navigate", {"url": url})

        responses = {}
        load_seen_at = None
        last_activity = time.monotonic()
        next_scroll_at = time.monotonic() + 2.0
        stable_scrolls = 0
        previous_height = 0
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if auto_scroll and time.monotonic() >= next_scroll_at:
                try:
                    scroll_result = cdp.command(
                        "Runtime.evaluate",
                        {
                            "expression": """
                                (() => {
                                    const elements = [document.scrollingElement].concat(
                                        Array.from(document.querySelectorAll('*'))
                                    );
                                    const scrollables = elements.filter(
                                        element => element && element.scrollHeight > element.clientHeight + 20
                                    );
                                    for (const element of scrollables) {
                                        element.scrollTop = element.scrollHeight;
                                    }
                                    return scrollables.reduce(
                                        (height, element) => Math.max(height, element.scrollHeight),
                                        document.documentElement ? document.documentElement.scrollHeight : 0
                                    );
                                })()
                            """,
                            "returnByValue": True,
                        },
                        timeout=3,
                    )
                    height = int(scroll_result.get("result", {}).get("value", 0) or 0)
                    try:
                        cdp.command(
                            "Input.dispatchMouseEvent",
                            {"type": "mouseWheel", "x": 360, "y": 640, "deltaY": 1200},
                            timeout=2,
                        )
                    except Exception:
                        pass
                    stable_scrolls = stable_scrolls + 1 if height == previous_height else 0
                    previous_height = height
                except Exception:
                    stable_scrolls += 1
                next_scroll_at = time.monotonic() + 1.0
            event = cdp.next_event(min(0.5, deadline - time.monotonic()))
            if event is None:
                if load_seen_at and time.monotonic() - max(load_seen_at, last_activity) >= 4 and (
                    not auto_scroll or stable_scrolls >= 5
                ):
                    break
                continue
            method = event.get("method", "")
            params = event.get("params", {})
            if method == "Page.loadEventFired":
                load_seen_at = time.monotonic()
            elif method == "Network.responseReceived":
                response = params.get("response", {})
                if _should_capture_response(params.get("type", ""), response):
                    responses[params.get("requestId", "")] = {
                        "url": response.get("url", ""),
                        "mime_type": response.get("mimeType", ""),
                        "status": response.get("status", 0),
                    }
                last_activity = time.monotonic()
            elif method in {"Network.loadingFinished", "Network.loadingFailed"}:
                last_activity = time.monotonic()

        rendered = cdp.command(
            "Runtime.evaluate",
            {
                "expression": "document.documentElement ? document.documentElement.outerHTML : ''",
                "returnByValue": True,
            },
        ).get("result", {}).get("value", "")
        final_url = cdp.command(
            "Runtime.evaluate",
            {"expression": "location.href", "returnByValue": True},
        ).get("result", {}).get("value", url)

        captured = []
        total_size = 0
        for request_id, metadata in list(responses.items())[:120]:
            try:
                body_result = cdp.command("Network.getResponseBody", {"requestId": request_id}, timeout=4)
            except Exception:
                continue
            body = body_result.get("body", "")
            estimated_size = len(body) * (3 if body_result.get("base64Encoded") else 1)
            if estimated_size > MAX_CAPTURED_BODY_BYTES or total_size + estimated_size > MAX_CAPTURED_TOTAL_BYTES:
                continue
            total_size += estimated_size
            metadata.update({
                "body": body,
                "base64_encoded": bool(body_result.get("base64Encoded")),
            })
            captured.append(metadata)

        return {
            "success": True,
            "browser": browser.name,
            "url": final_url,
            "html": rendered,
            "resources": captured,
        }
    except Exception as error:
        return {"success": False, "error": str(error)[:500], "browser": browser.name}
    finally:
        if socket is not None:
            try:
                socket.close()
            except Exception:
                pass
        if process is not None:
            _stop_browser_process(process)
        shutil.rmtree(profile, ignore_errors=True)


def _capture_dump_dom(browser: Path, url: str, timeout: int) -> dict:
    """Fallback for machines whose security software blocks CDP WebSockets."""
    profile = Path(tempfile.mkdtemp(prefix="dom-capture-", dir=config.USER_DATA_DIR / "browser_cache"))
    try:
        creation_flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        completed = subprocess.run(
            [
                str(browser),
                "--headless=new",
                "--disable-gpu",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-background-networking",
                "--disable-component-update",
                f"--user-data-dir={profile}",
                "--virtual-time-budget=15000",
                "--dump-dom",
                url,
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=max(20, timeout),
            creationflags=creation_flags,
        )
        html = completed.stdout.decode("utf-8", errors="replace")
        if "<html" not in html.lower():
            return {"success": False, "error": "浏览器未返回网页内容"}
        return {
            "success": True,
            "browser": browser.name,
            "url": url,
            "html": html,
            "resources": [],
            "capture_mode": "dump_dom",
        }
    except subprocess.TimeoutExpired:
        return {"success": False, "error": "浏览器页面加载超时"}
    except Exception as error:
        return {"success": False, "error": str(error)[:300]}
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def _stop_browser_process(process) -> None:
    if process.poll() is not None:
        return
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
            check=False,
        )
        return
    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()


def _find_available_port() -> int:
    with network_socket.socket(network_socket.AF_INET, network_socket.SOCK_STREAM) as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


def _get_browser_target(port: int, process, timeout: int = 5) -> dict:
    opener = build_opener(ProxyHandler({}))
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError("浏览器启动失败")
        try:
            with opener.open(f"http://127.0.0.1:{port}/json/list", timeout=1) as response:
                targets = json.loads(response.read().decode("utf-8"))
            target = next((item for item in targets if item.get("type") == "page"), None)
            if target:
                return target
        except Exception:
            time.sleep(0.1)
    raise TimeoutError("浏览器调试端口启动超时")


def _should_capture_response(resource_type: str, response: dict) -> bool:
    if int(response.get("status", 0) or 0) >= 400:
        return False
    mime_type = str(response.get("mimeType", "")).lower()
    url = str(response.get("url", "")).lower()
    if resource_type in {"XHR", "Fetch"}:
        return True
    return any(token in mime_type for token in ("json", "excel", "spreadsheet", "csv", "pdf")) or any(
        token in url for token in (".xlsx", ".xls", ".csv", ".pdf", "/api/", "/query", "/list")
    )
