"""Evidence-based discovery of official engineering price sources."""
from __future__ import annotations

import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse

import requests
from bs4 import BeautifulSoup

from .ai_research import search_web
from .config import config
from .utils import is_official_domain


_PRICE_MARKERS = (
    "信息价", "造价信息", "工程造价", "材料价格", "材料信息", "人材机",
    "价格信息", "市场价格", "建设工程价格", "工程价格",
)
_LOGIN_MARKERS = ("登录", "会员", "统一身份认证", "用户中心", "验证码")
_INSTITUTION_MARKERS = (
    "住房和城乡建设局", "住房城乡建设局", "住房和城乡建设厅", "住房城乡建设厅",
    "住建局", "住建委", "建设工程造价管理站", "工程造价管理站", "造价总站",
    "标准定额站", "定额站", "建设工程价格信息",
)
_TRACKING_QUERY_KEYS = {"from", "source", "spm", "utm_source", "utm_medium", "utm_campaign"}


def canonical_source_url(url: str) -> str:
    """Normalize a source URL for comparison without losing functional query parameters."""
    value = (url or "").strip()
    parsed = urlparse(value)
    host = (parsed.hostname or "").lower().rstrip(".")
    if not host:
        return ""
    port = parsed.port
    netloc = host if not port or (parsed.scheme == "https" and port == 443) or (parsed.scheme == "http" and port == 80) else f"{host}:{port}"
    path = re.sub(r"/{2,}", "/", parsed.path or "/").rstrip("/") or "/"
    query = urlencode(
        sorted(
            (key, value)
            for key, value in parse_qsl(parsed.query, keep_blank_values=True)
            if key.lower() not in _TRACKING_QUERY_KEYS
        )
    )
    return urlunparse(((parsed.scheme or "https").lower(), netloc, path, "", query, ""))


def build_official_source_queries(
    city: str,
    period: str = "",
    specialty: str = "",
    province: str = "",
) -> list[str]:
    city = (city or "").strip()
    period = (period or "").strip()
    specialty = (specialty or "").strip()
    province = (province or "").strip()
    context = " ".join(value for value in (period, specialty) if value and value != "全部专业")
    suffix = f" {context}" if context else ""
    queries = [
        f'"{city}" 住房和城乡建设局 工程造价 信息价{suffix}',
        f'"{city}" 住房和城乡建设委员会 造价信息{suffix}',
        f'"{city}" 工程造价管理站 材料信息价{suffix}',
        f'"{city}" 标准定额站 建设工程 信息价{suffix}',
        f'"{city}" 建设工程价格信息 月刊{suffix}',
        f'"{city}" 人材机 信息价 查询{suffix}',
        f'site:gov.cn "{city}" "信息价" 建设工程{suffix}',
        f'site:gov.cn "{city}" 材料价格 造价信息 PDF Excel{suffix}',
    ]
    if province and province != city:
        queries.extend([
            f'"{province}" 住房和城乡建设厅 "{city}" 信息价{suffix}',
            f'"{province}" 建设工程造价总站 "{city}" 材料价格{suffix}',
        ])
    return queries


def _response_excerpt(response, max_bytes: int = 240_000) -> tuple[str, str]:
    chunks = []
    size = 0
    for chunk in response.iter_content(chunk_size=32_768):
        if not chunk:
            continue
        remaining = max_bytes - size
        chunks.append(chunk[:remaining])
        size += min(len(chunk), remaining)
        if size >= max_bytes:
            break
    raw = b"".join(chunks)
    encoding = response.encoding or "utf-8"
    try:
        html = raw.decode(encoding, errors="replace")
    except LookupError:
        html = raw.decode("utf-8", errors="replace")
    soup = BeautifulSoup(html, "html.parser")
    for element in soup(["script", "style", "noscript", "svg"]):
        element.decompose()
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))[:6000]
    return html, text


def _allowed_official_url(url: str, trusted_hosts: set[str] | None = None) -> bool:
    host = (urlparse(url).hostname or "").lower().rstrip(".")
    return is_official_domain(url) or bool(host and host in (trusted_hosts or set()))


def _page_clues(
    base_url: str,
    html: str,
    trusted_hosts: set[str] | None = None,
) -> tuple[list[str], list[str]]:
    api_urls = []
    attachments = []
    if not html:
        return api_urls, attachments
    soup = BeautifulSoup(html, "html.parser")
    for element in soup.select("a[href], form[action], script[src]"):
        target = element.get("href") or element.get("action") or element.get("src") or ""
        absolute = urljoin(base_url, target.strip())
        if not _allowed_official_url(absolute, trusted_hosts):
            continue
        lower = absolute.lower()
        if re.search(r"\.(?:pdf|xlsx?|csv|zip)(?:$|[?#])", lower):
            if absolute not in attachments:
                attachments.append(absolute)
        if any(marker in lower for marker in ("/api/", "/interface/", "query", "search", "price", "list", "xxj", "jgb")):
            if absolute not in api_urls:
                api_urls.append(absolute)
    endpoint_pattern = re.compile(
        r"[\"']((?:https?://[^\"']+|/[^\"']{3,220})(?:api|interface|query|search|price|material|xxj|jgb)[^\"']*)[\"']",
        re.IGNORECASE,
    )
    for match in endpoint_pattern.finditer(html):
        absolute = urljoin(base_url, match.group(1).replace("\\/", "/"))
        if _allowed_official_url(absolute, trusted_hosts) and absolute not in api_urls:
            api_urls.append(absolute)
    return api_urls[:5], attachments[:5]


def _institution_name(text: str) -> str:
    for marker in _INSTITUTION_MARKERS:
        if marker in text:
            return marker
    return "政府官网"


def _jurisdiction(text: str, city: str, province: str, catalog_seed: bool) -> str:
    if city and city in text:
        return "目标城市"
    if province and province in text:
        return "省级来源"
    return "现有目录来源" if catalog_seed else "归属待核验"


def verify_official_search_result(
    record: dict,
    request_get=requests.get,
    *,
    city: str = "",
    province: str = "",
    trusted_hosts: set[str] | None = None,
) -> tuple[dict | None, str]:
    """Verify one search result and return only official, relevant evidence."""
    source_url = str(record.get("url") or "").strip()
    catalog_seed = bool(record.get("catalog_seed"))
    if not source_url or not _allowed_official_url(source_url, trusted_hosts):
        return None, "搜索结果不是已认可的官方域名"
    try:
        response = request_get(
            source_url,
            headers={"User-Agent": config.USER_AGENT, "Accept-Language": "zh-CN,zh;q=0.9"},
            timeout=max(8, int(config.AI_WEB_SEARCH_TIMEOUT)),
            allow_redirects=True,
            stream=True,
        )
    except Exception as error:
        return None, str(error)[:240]

    try:
        final_url = str(getattr(response, "url", source_url) or source_url)
        if not _allowed_official_url(final_url, trusted_hosts):
            return None, "最终跳转地址不是已认可的官方域名"
        status_code = int(getattr(response, "status_code", 0) or 0)
        if status_code in {404, 410} or status_code >= 500 or not status_code:
            return None, f"官方页面返回 HTTP {status_code or '未知'}"
        content_type = str(response.headers.get("Content-Type", "")).lower()
        html = ""
        excerpt = ""
        if status_code not in {401, 403} and any(value in content_type for value in ("html", "text", "json", "xml", "javascript")):
            html, excerpt = _response_excerpt(response)
        combined = " ".join(
            str(record.get(key) or "") for key in ("title", "snippet")
        ) + " " + excerpt
        if status_code in {401, 403}:
            return {
                "name": str(record.get("title") or "官方信息价来源").strip()[:300],
                "url": final_url,
                "source_kind": "api" if "/api/" in final_url.lower() else "web",
                "description": str(record.get("snippet") or "").strip()[:500],
                "excerpt": excerpt,
                "http_status": status_code,
                "content_type": content_type[:120],
                "verified": True,
                "verification_status": "需登录/受限",
                "login_required": True,
                "api_hint": "官方地址需要登录或会员授权，登录后再验证价格数据",
                "api_candidates": [],
                "attachment_candidates": [],
                "institution": _institution_name(combined),
                "jurisdiction": _jurisdiction(combined, city, province, catalog_seed),
                "relevance_score": 65,
                "catalog_seed": catalog_seed,
                "search_query": str(record.get("query") or ""),
                "search_engine": str(record.get("engine") or ""),
            }, ""
        if not any(marker in combined for marker in _PRICE_MARKERS):
            return None, "页面可访问，但未发现工程造价信息价内容证据"
        if city and city not in combined and (not province or province not in combined) and not catalog_seed:
            return None, f"页面存在价格文字，但无法确认属于{city}或{province or '对应省份'}"

        lower_url = final_url.lower()
        if "json" in content_type or "/api/" in lower_url:
            source_kind = "api"
        elif "pdf" in content_type or re.search(r"\.pdf(?:$|[?#])", lower_url):
            source_kind = "pdf"
        elif any(value in content_type for value in ("spreadsheet", "excel", "csv")) or re.search(r"\.(?:xlsx?|csv)(?:$|[?#])", lower_url):
            source_kind = "excel"
        else:
            source_kind = "web"
        api_urls, attachments = _page_clues(final_url, html, trusted_hosts)
        login_required = status_code in {401, 403} or any(marker in combined for marker in _LOGIN_MARKERS)
        if source_kind == "api":
            api_hint = "该地址返回接口型内容，仍需验证字段、分页和期数参数"
        elif api_urls:
            api_hint = f"页面发现 {len(api_urls)} 条接口/查询线索，需继续API嗅探"
        else:
            api_hint = "未发现可直接确认的接口，保存后可在来源配置中继续API嗅探"
        institution = _institution_name(combined)
        jurisdiction = _jurisdiction(combined, city, province, catalog_seed)
        relevance_score = 30
        relevance_score += 30 if jurisdiction == "目标城市" else 15 if jurisdiction == "省级来源" else 10
        relevance_score += 15 if institution != "政府官网" else 0
        relevance_score += 15 if source_kind == "api" else 10 if api_urls or attachments else 0
        return {
            "name": str(record.get("title") or "官方信息价来源").strip()[:300],
            "url": final_url,
            "source_kind": source_kind,
            "description": str(record.get("snippet") or "").strip()[:500],
            "excerpt": excerpt[:1200],
            "http_status": status_code,
            "content_type": content_type[:120],
            "verified": True,
            "verification_status": "需登录/受限" if status_code in {401, 403} else "官方域名可访问",
            "login_required": login_required,
            "api_hint": api_hint,
            "api_candidates": api_urls,
            "attachment_candidates": attachments,
            "institution": institution,
            "jurisdiction": jurisdiction,
            "relevance_score": min(relevance_score, 100),
            "catalog_seed": catalog_seed,
            "search_query": str(record.get("query") or ""),
            "search_engine": str(record.get("engine") or ""),
        }, ""
    finally:
        close = getattr(response, "close", None)
        if callable(close):
            close()


def discover_official_sources(
    city: str,
    period: str = "",
    specialty: str = "",
    province: str = "",
    *,
    limit: int = 8,
    seed_sources: list[dict] | None = None,
    search_fn=search_web,
    request_get=requests.get,
) -> dict:
    """Search and verify official source evidence; no result is persisted here."""
    queries = build_official_source_queries(city, period, specialty, province)
    raw_results = []
    trusted_hosts = set()
    for source in seed_sources or []:
        url = str(source.get("url") or "").strip()
        host = (urlparse(url).hostname or "").lower().rstrip(".")
        if not url or not host:
            continue
        trusted_hosts.add(host)
        raw_results.append({
            "query": "现有官方目录",
            "title": source.get("name") or f"{city}现有官方来源",
            "url": url,
            "snippet": source.get("note") or "工程造价信息价官方来源",
            "engine": "official_catalog",
            "catalog_seed": True,
        })
    def run_queries(query_values: list[str], result_limit: int = 8) -> list[dict]:
        found = []
        with ThreadPoolExecutor(max_workers=min(5, max(1, len(query_values)))) as executor:
            futures = {
                executor.submit(search_fn, query, result_limit): query
                for query in query_values
            }
            for future in as_completed(futures):
                query = futures[future]
                try:
                    values = future.result() or []
                except Exception:
                    values = []
                for result in values:
                    candidate = dict(result)
                    candidate.setdefault("query", query)
                    found.append(candidate)
        return found

    raw_results.extend(run_queries(queries))

    # Reuse successful-city experience: first identify the correct official host, then search inside it
    # for price columns instead of treating unrelated government pages as candidates.
    host_candidates = []
    for record in raw_results:
        url = str(record.get("url") or "")
        host = (urlparse(url).hostname or "").lower().rstrip(".")
        text = f"{record.get('title', '')} {record.get('snippet', '')}"
        if host and _allowed_official_url(url, trusted_hosts) and (city in text or record.get("catalog_seed")):
            if host not in host_candidates:
                host_candidates.append(host)
    host_queries = []
    for host in host_candidates[:6]:
        host_queries.extend([
            f'site:{host} "信息价" (材料价格 OR 工程造价 OR 标准定额)',
            f'site:{host} (价格查询 OR 人材机 OR 信息价月刊) (API OR Excel OR PDF)',
        ])
    targeted_results = run_queries(host_queries)
    for candidate in targeted_results:
        candidate["host_targeted"] = True
        raw_results.append(candidate)
    queries = [*queries, *host_queries]

    deduped = []
    seen = set()
    for record in raw_results:
        key = canonical_source_url(str(record.get("url") or ""))
        if not key or key in seen:
            continue
        seen.add(key)
        deduped.append(record)

    def search_record_score(record: dict) -> int:
        text = f"{record.get('title', '')} {record.get('snippet', '')}"
        score = 50 if record.get("catalog_seed") else 0
        score += 25 if city and city in text else 0
        score += 12 if province and province in text else 0
        score += sum(5 for marker in _PRICE_MARKERS if marker in text)
        score += 15 if any(marker in text for marker in _INSTITUTION_MARKERS) else 0
        score += 12 if record.get("host_targeted") else 0
        return score

    deduped.sort(key=search_record_score, reverse=True)
    verification_pool = deduped[:48]

    sources = []
    rejected = []
    with ThreadPoolExecutor(max_workers=min(6, max(1, len(verification_pool)))) as executor:
        futures = {
            executor.submit(
                verify_official_search_result,
                record,
                request_get,
                city=city,
                province=province,
                trusted_hosts=trusted_hosts,
            ): record
            for record in verification_pool
        }
        for future in as_completed(futures):
            record = futures[future]
            try:
                verified, reason = future.result()
            except Exception as error:
                verified, reason = None, str(error)[:240]
            if verified:
                sources.append(verified)
            else:
                rejected.append({"url": str(record.get("url") or ""), "reason": reason})
    sources.sort(
        key=lambda item: (
            item.get("jurisdiction") != "目标城市",
            -int(item.get("relevance_score") or 0),
            item.get("source_kind") != "api",
        )
    )
    sources = sources[: max(1, int(limit))]
    for index, source in enumerate(sources, 1):
        source["evidence_id"] = f"E{index}"
    return {
        "city": city,
        "province": province,
        "queries": queries,
        "strategy_steps": [
            "现有已核查官方目录",
            "市住建局/住建委",
            "市造价管理站/标准定额站",
            "省住建厅/省造价机构",
            "已发现官方网站内二次检索",
            "页面接口与价格附件线索识别",
        ],
        "sources": sources,
        "searched_count": len(raw_results),
        "candidate_count": len(deduped),
        "verified_attempt_count": len(verification_pool),
        "rejected": rejected[:20],
    }
