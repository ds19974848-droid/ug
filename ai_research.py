"""Web research helpers for AI quota composition fallback."""
from __future__ import annotations

import logging
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

from .config import config
from .quota_service import (
    _clause_material_terms,
    extract_quota_key_terms,
    split_boq_work_items,
)
from .source_display import redact_research_evidence


logger = logging.getLogger(__name__)
_SEARCH_CACHE: dict[tuple[str, int], tuple[float, list[dict]]] = {}
_PAGE_CACHE: dict[str, tuple[float, str]] = {}
_CACHE_TTL_SECONDS = 30 * 60


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _search_headers() -> dict[str, str]:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.6",
    }


def _is_official_result(url: str, title: str = "") -> bool:
    value = f"{url} {title}".lower()
    if ".gov.cn" in value or "zjj." in value or "zjz." in value:
        return True
    return any(marker in value for marker in config.OFFICIAL_DOMAIN_PATTERNS)


def _relevance_score(record: dict) -> int:
    text = f"{record.get('title', '')} {record.get('snippet', '')}".lower()
    keywords = (
        "信息价", "造价", "定额", "人材机", "人工费", "材料费", "机械费",
        "单价", "价格", "报价", "套", "台班", "d400", "φ", "φ",
    )
    score = sum(1 for keyword in keywords if keyword in text)
    if record.get("official"):
        score += 12
    return score


def _query_terms(boq: dict) -> list[str]:
    """Keep only specific BOQ anchors so search does not drift by generic words."""
    terms = extract_quota_key_terms(boq.get("name", ""), boq.get("feature", ""))
    material_terms = _clause_material_terms(
        f"{_text(boq.get('name'))} {_text(boq.get('feature'))}"
    )
    values = [
        _text(boq.get("name")),
        *[str(value) for value in terms.get("crafts") or []],
        *[str(value) for value in terms.get("materials") or []],
        *[str(value) for value in terms.get("road_scope_groups") or []],
        *[str(value) for value in material_terms],
    ]
    compact = []
    for value in values:
        value = re.sub(r"\s+", " ", value).strip()
        if value and value not in compact:
            compact.append(value)
    return compact[:5]


def _evidence_score(record: dict, anchors: list[str], region: str, period: str) -> int:
    text = " ".join(_text(record.get(key)) for key in ("title", "snippet", "excerpt"))
    score = _relevance_score(record)
    matched = sum(1 for anchor in anchors if len(anchor) >= 2 and anchor in text)
    score += matched * 8
    if region and region in text:
        score += 8
    if period and (period in text or period[:4] in text):
        score += 5
    if any(term in text for term in ("元", "单价", "信息价", "定额", "人材机", "台班")):
        score += 6
    return score


def _search_sogou(query: str, limit: int) -> list[dict]:
    """Sogou is a useful free fallback for Chinese commercial price pages."""
    try:
        response = requests.get(
            "https://www.sogou.com/web",
            params={"query": query, "ie": "utf8"},
            headers=_search_headers(),
            timeout=config.AI_WEB_SEARCH_TIMEOUT,
        )
        response.raise_for_status()
    except Exception as error:
        logger.info("Sogou search failed for %s: %s", query, error)
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    results = []
    seen = set()
    for item in soup.select("div.vrwrap, div.rb, div.results > div.vrwrap, div.results > div.rb"):
        anchor = item.select_one("h3 a, a")
        snippet = item.select_one(
            "div.space-txt, div.text-layout, .fz-mid, p",
        )
        title = anchor.get_text(" ", strip=True) if anchor else ""
        href = anchor.get("href") if anchor else ""
        url = urljoin(str(response.url), href) if href else ""
        text = snippet.get_text(" ", strip=True) if snippet else ""
        key = url or title
        if not title or not key or key in seen:
            continue
        seen.add(key)
        results.append({
            "query": query,
            "title": title,
            "url": url,
            "host": urlparse(url).hostname or "",
            "snippet": text,
            "official": _is_official_result(url, title),
            "engine": "sogou",
        })
        if len(results) >= int(limit):
            break
    return results


def search_web(query: str, limit: int = 6) -> list[dict]:
    """Return bounded search results from Bing and Sogou."""
    query = _text(query)
    if not query:
        return []
    cache_key = (query, int(limit))
    cached = _SEARCH_CACHE.get(cache_key)
    if cached and time.monotonic() - cached[0] < _CACHE_TTL_SECONDS:
        return [dict(value) for value in cached[1]]
    try:
        response = requests.get(
            "https://www.bing.com/search",
            params={
                "q": query,
                "count": max(3, min(int(limit), 10)),
                "mkt": "zh-CN",
                "setlang": "zh-hans",
            },
            headers=_search_headers(),
            timeout=config.AI_WEB_SEARCH_TIMEOUT,
        )
        response.raise_for_status()
    except Exception as error:
        logger.warning("AI web search failed for %s: %s", query, error)
        bing_results = []
    else:
        soup = BeautifulSoup(response.text, "html.parser")
        bing_results = []
        seen_urls = set()
        for item in soup.select("li.b_algo"):
            anchor = item.select_one("h2 a")
            snippet = item.select_one(".b_caption p, .b_lineclamp2, .b_lineclamp3, .b_lineclamp4")
            title = anchor.get_text(" ", strip=True) if anchor else ""
            url = anchor.get("href") if anchor else ""
            text = snippet.get_text(" ", strip=True) if snippet else ""
            if not title or not url or url in seen_urls:
                continue
            seen_urls.add(url)
            hostname = urlparse(url).hostname or ""
            bing_results.append({
                "query": query, "title": title, "url": url, "host": hostname,
                "snippet": text, "official": _is_official_result(url, title), "engine": "bing",
            })
            if len(bing_results) >= int(limit):
                break
    combined = bing_results + _search_sogou(query, limit)
    deduped = []
    seen_urls = set()
    for result in combined:
        url = result.get("url")
        if not url or url in seen_urls:
            continue
        seen_urls.add(url)
        deduped.append(result)
    deduped.sort(
        key=lambda value: (
            not value.get("official"),
            -_relevance_score(value),
        )
    )
    result = deduped[: max(1, int(limit))]
    _SEARCH_CACHE[cache_key] = (time.monotonic(), [dict(value) for value in result])
    return result


def fetch_page_text(url: str, max_chars: int = 8000) -> str:
    """Fetch a short HTML excerpt from a search result."""
    url = _text(url)
    if not url.startswith(("http://", "https://")):
        return ""
    cached = _PAGE_CACHE.get(url)
    if cached and time.monotonic() - cached[0] < _CACHE_TTL_SECONDS:
        return cached[1]
    try:
        response = requests.get(
            url,
            headers=_search_headers(),
            timeout=config.AI_WEB_SEARCH_TIMEOUT,
            allow_redirects=True,
        )
        response.raise_for_status()
    except Exception as error:
        logger.info("AI research page fetch failed for %s: %s", url, error)
        return ""
    content_type = response.headers.get("Content-Type", "").lower()
    if "html" not in content_type and "text" not in content_type:
        return ""
    soup = BeautifulSoup(response.text, "html.parser")
    for element in soup(["script", "style", "noscript", "svg"]):
        element.decompose()
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    result = text[:max_chars]
    _PAGE_CACHE[url] = (time.monotonic(), result)
    return result


def _research_queries(boq: dict, context: dict, work_items: list[str]) -> list[str]:
    project = context.get("project") or {}
    name = _text(boq.get("name"))
    region_values = []
    for value in (
        _text(project.get("city")),
        _text(project.get("province")),
        _text(project.get("location")),
    ):
        if value and value not in region_values:
            region_values.append(value)
    region = " ".join(region_values)
    period = _text(project.get("pricing_date") or project.get("price_year"))
    base_queries = []
    anchors = _query_terms(boq)
    clause_anchors = [
        re.sub(r"[，,、:：;；]+", " ", _text(value)).strip()
        for value in work_items[:3]
        if _text(value)
    ]
    anchors = list(dict.fromkeys([*anchors, *clause_anchors]))[:8]
    focus = " ".join(anchors[:3]) or name
    clauses = [
        re.sub(r"[，,、:：;；]+", " ", _text(value)).strip()
        for value in work_items[:3]
        if _text(value)
    ]
    if clauses:
        base_queries.append(f"{name or focus} {' '.join(clauses[:2])} 定额 人材机")
    if focus:
        base_queries.append(f"{focus} 定额 人材机")
        base_queries.append(f"{focus} 施工含量 单价")
    if region and focus:
        base_queries.append(f"{focus} {region} {period} 信息价")
        base_queries.append(f"{focus} {region} 造价 site:gov.cn")
    return list(dict.fromkeys(value for value in base_queries if value))


def research_boq_evidence(
    boq: dict,
    context: dict,
    *,
    max_queries: int = 4,
    max_results_per_query: int = 3,
    fetch_pages: int = 3,
) -> list[dict]:
    """Search and fetch bounded evidence for one unmatched BOQ row."""
    if not config.AI_WEB_SEARCH_ENABLED:
        return []
    work_items = split_boq_work_items(boq.get("name", ""), boq.get("feature", ""))
    queries = _research_queries(boq, context, work_items)[: max(1, int(max_queries))]
    project = context.get("project") or {}
    region = " ".join(_text(project.get(key)) for key in ("city", "province") if _text(project.get(key)))
    period = _text(project.get("effective_pricing_period") or project.get("pricing_date"))
    province = _text(project.get("province") or project.get("pricing_province"))
    anchors = _query_terms(boq)
    records = []
    seen_urls = set()
    # The public structured 造价HOME workbook/PDF is a direct evidence source
    # for every city that has been verified in its province catalog. Other
    # regions continue through the national search engines and public pages.
    try:
        from .api_fetcher import search_zaojiahome_market_reference
        records.extend(search_zaojiahome_market_reference(
            _text(project.get("city")),
            period,
            keywords=anchors,
            unit=_text(boq.get("unit")),
            limit=max_results_per_query * 2,
            province=province,
        ))
    except Exception as error:
        logger.info("Public market reference search failed: %s", error)
    for query in queries:
        for result in search_web(query, limit=max_results_per_query):
            url = _text(result.get("url"))
            if not url or url in seen_urls:
                continue
            seen_urls.add(url)
            records.append({
                **result,
                "evidence_score": _evidence_score(result, anchors, region, period),
            })
        if len(records) >= int(max_queries) * int(max_results_per_query):
            break
    records.sort(key=lambda value: (-value.get("evidence_score", 0), not value.get("official")))
    # Fetch only the highest-quality candidates and do it concurrently.
    fetch_targets = records[: max(0, int(fetch_pages))]
    with ThreadPoolExecutor(max_workers=min(3, len(fetch_targets) or 1)) as executor:
        futures = {executor.submit(fetch_page_text, record["url"]): record for record in fetch_targets}
        for future in as_completed(futures):
            record = futures[future]
            try:
                record["excerpt"] = future.result()
            except Exception:
                record["excerpt"] = ""
    accepted = []
    for record in records:
        evidence_text = " ".join(_text(record.get(key)) for key in ("title", "snippet", "excerpt"))
        matched_anchors = sum(1 for anchor in anchors if len(anchor) >= 2 and anchor in evidence_text)
        record["matched_anchors"] = matched_anchors
        record["trusted"] = bool(record.get("official")) and bool(record.get("excerpt"))
        # Do not let an irrelevant SEO page become AI pricing evidence.
        if matched_anchors or record.get("official") or record.get("evidence_score", 0) >= 14:
            accepted.append(record)
    return redact_research_evidence(
        accepted[: max(1, int(max_queries) * int(max_results_per_query))]
    )
