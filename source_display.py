"""User-facing source labels and redaction for internal retrieval evidence."""
from __future__ import annotations

import re
from copy import deepcopy


_PUBLIC_MARKET_RE = re.compile(r"造价\s*HOME|zaojiahome(?:\.com)?", re.IGNORECASE)
_URL_RE = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)


def is_public_market_source(name: str = "", url: str = "", source_type: str = "") -> bool:
    text = f"{name} {url} {source_type}".lower()
    return "zaojiahome" in text or "造价home" in text.replace(" ", "")


def display_source_name(name: str = "", url: str = "", source_type: str = "") -> str:
    """Return a neutral label without exposing the internal market source."""
    if is_public_market_source(name, url, source_type):
        return "公开市场参考数据"
    if source_type == "user_url":
        return "用户输入来源"
    return _redact_text(name) or "信息价来源"


def display_source_url(url: str = "", name: str = "", source_type: str = "") -> str:
    if is_public_market_source(name, url, source_type):
        return "已隐藏公开来源地址"
    return url or ""


def display_file_name(file_name: str = "", source_name: str = "", source_url: str = "", source_type: str = "") -> str:
    """Return a safe file label without exposing internal source naming."""
    if is_public_market_source(source_name, source_url, source_type):
        return "公开市场价格文件"
    return _redact_text(file_name) or "来源文件"


def _redact_text(value) -> str:
    text = str(value or "")
    text = _URL_RE.sub("来源地址已隐藏", text)
    return _PUBLIC_MARKET_RE.sub("公开市场参考数据", text)


def redact_research_evidence(records: list[dict] | None) -> list[dict]:
    """Strip URLs and identifying market-site labels before evidence is shown or persisted."""
    redacted = []
    for original in records or []:
        if not isinstance(original, dict):
            continue
        item = deepcopy(original)
        for key in ("url", "host", "attachment_candidates", "api_candidates"):
            item.pop(key, None)
        for key in ("title", "snippet", "excerpt", "query", "source_name", "name"):
            if key in item:
                item[key] = _redact_text(item[key])
        item["source_display"] = display_source_name(
            item.get("title") or item.get("name") or "",
            "",
            item.get("source_type") or "",
        )
        redacted.append(item)
    return redacted


def redact_text(value) -> str:
    """Redact URLs and the internal market source name from a user-facing message."""
    return _redact_text(value)
