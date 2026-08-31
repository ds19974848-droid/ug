"""????????????Chrome??????????Cookie?DPAPI?????"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

from .browser_capture import (
    _CdpSession,
    _find_available_port,
    _get_browser_target,
    _stop_browser_process,
    find_chromium_browsers,
)
from .config import config
from .source_credentials import save_source_credentials, load_source_credentials


def launch_visible_browser(url, *, timeout_seconds=300, progress_callback=None):
    parsed = urlparse(url.strip())
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        return {"success": False, "error": "?????? http/https ??"}
    browsers = find_chromium_browsers()
    if not browsers:
        return {"success": False, "error": "??? Microsoft Edge ? Google Chrome"}
    errors = []
    for browser in browsers[:1]:
        try:
            result = _launch_one_browser(browser, url, timeout_seconds, progress_callback)
            if result.get("success"):
                return result
            errors.append(f"{browser.name}: {result.get('error', '????')}")
        except Exception as exc:
            errors.append(f"{browser.name}: {exc}")
    return {"success": False, "error": ";".join(errors)[:500]}


def _launch_one_browser(browser, url, timeout_seconds, progress_callback=None):
    cache_root = config.USER_DATA_DIR / "browser_cache"
    cache_root.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="browser-login-", dir=cache_root))
    process = None
    try:
        if progress_callback:
            progress_callback("???????...")
        port = _find_available_port()
        creation_flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        cmd = [
            str(browser),
            f"--remote-debugging-port={port}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-networking",
            "--disable-component-update",
            "--remote-allow-origins=*",
            f"--user-data-dir={profile}",
            url,
        ]
        process = subprocess.Popen(
            cmd,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creation_flags,
            close_fds=True,
        )
        if progress_callback:
            progress_callback("???????????...")
        target = _get_browser_target(port, process, min(15, timeout_seconds))
        from websocket import create_connection
        socket = create_connection(
            target["webSocketDebuggerUrl"].replace("localhost", "127.0.0.1"),
            timeout=8,
            http_proxy_host=None,
        )
        return {
            "success": True,
            "browser_name": browser.name,
            "cdp_port": port,
            "debugger_url": target["webSocketDebuggerUrl"],
            "process": process,
            "socket": socket,
            "profile_dir": profile,
        }
    except Exception as exc:
        if process is not None and process.poll() is None:
            _stop_browser_process(process)
        shutil.rmtree(profile, ignore_errors=True)
        return {"success": False, "error": str(exc)[:300]}


def capture_cookies_from_cdp(cdp_session):
    result = cdp_session.command("Network.getAllCookies")
    cookies = result.get("cookies", []) if result else []
    return [
        {
            "name": c.get("name", ""),
            "value": c.get("value", ""),
            "domain": c.get("domain", ""),
            "path": c.get("path", "/"),
            "expires": c.get("expires", 0),
            "httpOnly": c.get("httpOnly", False),
            "secure": c.get("secure", False),
        }
        for c in cookies
    ]


def cookies_to_header_string(cookies):
    parts = []
    for c in cookies:
        if c.get("name"):
            parts.append(f"{c['name']}={c['value']}")
    return "; ".join(parts)


def cookies_from_header_string(header):
    results = []
    for part in header.split(";"):
        part = part.strip()
        if "=" in part:
            name, _, value = part.partition("=")
            results.append({"name": name.strip(), "value": value.strip()})
    return results


def save_browser_cookies(source_id, cookie_header):
    save_source_credentials(source_id, auth_type="cookie", username="browser_login", secret=cookie_header)


def get_source_cookie_header(source_id):
    credential = load_source_credentials(source_id, include_secret=True)
    if credential.get("auth_type") != "cookie":
        return None
    secret = credential.get("secret", "")
    return secret or None


def has_valid_cookies(source_id):
    return bool(get_source_cookie_header(source_id))


def wait_for_user_login_and_capture(cdp_session, *, login_detection_url=None,
                                     timeout_seconds=300, progress_callback=None,
                                     should_stop=None, capture_requested=None):
    started_at = time.monotonic()
    deadline = started_at + timeout_seconds
    load_fired = False
    page_url = ""
    cdp_session.command("Page.enable")
    cdp_session.command("Network.enable")
    cdp_session.command("Runtime.enable")
    last_status = ""

    def capture_now(reason=""):
        try:
            cookies = capture_cookies_from_cdp(cdp_session)
            header = cookies_to_header_string(cookies)
            if not cookies:
                return {"success": False, "error": "当前页面没有可保存的登录 Cookie，请确认是在软件打开的浏览器中登录"}
            return {
                "success": True,
                "cookies": cookies,
                "cookie_header": header,
                "final_url": page_url,
                "cookie_count": len(cookies),
                "warning": reason,
            }
        except Exception as exc:
            return {"success": False, "error": f"Cookie捕获失败：{exc}"}

    while time.monotonic() < deadline:
        if should_stop and should_stop():
            return {"success": False, "error": "???????"}
        if capture_requested and capture_requested():
            if progress_callback:
                progress_callback("正在确认登录状态并保存 Cookie...")
            return capture_now()
        event = cdp_session.next_event(min(1.0, deadline - time.monotonic()))
        if event is None:
            elapsed = int(time.monotonic() - started_at)
            status = f"????? ({elapsed}s / {timeout_seconds}s)"
            if status != last_status and progress_callback:
                progress_callback(status)
                last_status = status
            if load_fired:
                try:
                    auth_text = cdp_session.command(
                        "Runtime.evaluate",
                        {
                            "expression": "document.body ? document.body.innerText.slice(0,12000) : ''",
                            "returnByValue": True,
                        },
                    ).get("result", {}).get("value", "")
                    if any(marker in auth_text for marker in (
                        "退出登录", "安全退出", "注销登录", "个人中心", "我的账户",
                    )):
                        return capture_now("已根据页面登录标识自动确认")
                except Exception:
                    pass
            continue
        method = event.get("method", "")
        if method == "Page.loadEventFired":
            load_fired = True
        if method == "Page.frameNavigated":
            frame = event.get("params", {}).get("frame", {})
            page_url = frame.get("url", page_url)
        if load_fired:
            try:
                eval_r = cdp_session.command(
                    "Runtime.evaluate",
                    {"expression": "location.href", "returnByValue": True},
                )
                current_url = eval_r.get("result", {}).get("value", "") or ""
                if current_url:
                    page_url = current_url
            except Exception:
                pass
            if login_detection_url and login_detection_url in (page_url or ""):
                if progress_callback:
                    progress_callback("???????????? Cookie...")
                return capture_now()
    try:
        cookies = capture_cookies_from_cdp(cdp_session)
        header = cookies_to_header_string(cookies)
        if cookies:
            return {"success": True, "cookies": cookies, "cookie_header": header,
                    "final_url": page_url, "cookie_count": len(cookies),
                    "warning": "??????????Cookie????????????"}
    except Exception:
        pass
    return {"success": False, "error": f"?????? ({timeout_seconds}s)???????????????????"}


def close_browser_login(session_info):
    if "socket" in session_info and session_info["socket"]:
        try:
            session_info["socket"].close()
        except Exception:
            pass
    if "process" in session_info:
        _stop_browser_process(session_info["process"])
    if "profile_dir" in session_info:
        shutil.rmtree(session_info["profile_dir"], ignore_errors=True)
