"""AI 服务抽象层：封装 DeepSeek API 调用，提供各模块 AI 能力"""
import json
import logging
from datetime import datetime
from typing import Optional

import requests

from .config import config
from .db import get_session
from .source_display import redact_research_evidence, redact_text

logger = logging.getLogger(__name__)


class AIService:
    """AI 服务单例，封装 DeepSeek API，无 Key 时所有方法返回安全默认值"""

    _instance: Optional["AIService"] = None
    _session: Optional[requests.Session] = None
    _available: Optional[bool] = None
    _connection_state = "not_configured"
    _connection_message = "未配置 API Key"

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    @property
    def is_available(self) -> bool:
        if self._available is None:
            self._check_availability()
        return self._available

    def _check_availability(self):
        if not config.DEEPSEEK_API_KEY:
            self._available = False
            self._connection_state = "not_configured"
            self._connection_message = "未配置 API Key"
            return
        base_url = str(config.DEEPSEEK_BASE_URL or "").strip()
        if not base_url.startswith(("https://", "http://")):
            self._available = False
            self._connection_state = "failed"
            self._connection_message = "AI 服务地址必须以 http:// 或 https:// 开头"
            return
        self._session = requests.Session()
        self._session.headers.update({
            "Authorization": f"Bearer {config.DEEPSEEK_API_KEY}",
            "Content-Type": "application/json",
            "User-Agent": f"DashuoCostCloud/{config.VERSION}",
        })
        self._available = True
        self._connection_state = "not_tested"
        self._connection_message = "已配置，尚未测试连接"

    @property
    def connection_state(self) -> str:
        if self._available is None:
            self._check_availability()
        return self._connection_state

    @property
    def connection_message(self) -> str:
        if self._available is None:
            self._check_availability()
        return self._connection_message

    def reload(self):
        """重新检测 API 可用性（用户更新配置后调用）"""
        self._available = None
        if self._session is not None:
            self._session.close()
        self._session = None
        self._check_availability()

    def _chat_completion(
        self,
        messages: list[dict],
        temperature: float,
        max_tokens: int,
        response_format: dict | None = None,
    ) -> str:
        if self._session is None:
            raise RuntimeError("AI 服务尚未初始化")
        endpoint = f"{str(config.DEEPSEEK_BASE_URL).rstrip('/')}/chat/completions"
        payload = {
            "model": config.DEEPSEEK_MODEL,
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        if response_format:
            payload["response_format"] = response_format
        response = self._session.post(
            endpoint,
            json=payload,
            timeout=60,
        )
        try:
            data = response.json()
        except ValueError as error:
            raise RuntimeError(f"AI 服务返回了无法解析的响应（HTTP {response.status_code}）") from error
        if not response.ok:
            error_data = data.get("error", {}) if isinstance(data, dict) else {}
            detail = error_data.get("message") if isinstance(error_data, dict) else ""
            raise RuntimeError(f"AI 服务请求失败（HTTP {response.status_code}）：{detail or '请检查地址、API Key 和网络'}")
        try:
            content = data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError) as error:
            raise RuntimeError("AI 服务未返回有效内容") from error
        if not isinstance(content, str) or not content.strip():
            raise RuntimeError("AI 服务未返回有效内容")
        return content.strip()

    def test_connection(self) -> tuple[bool, str]:
        """发起一次最小真实请求，不能仅凭 API Key 判定已连接。"""
        if not self.is_available or self._session is None:
            return False, self.connection_message
        try:
            self._chat_completion(
                [{"role": "user", "content": "只回复 OK"}],
                temperature=0,
                max_tokens=4,
            )
            self._connection_state = "connected"
            self._connection_message = f"已连接 {config.DEEPSEEK_MODEL}"
            return True, self._connection_message
        except Exception as error:
            self._connection_state = "failed"
            self._connection_message = str(error)
            return False, self._connection_message

    def _call(
        self,
        system_prompt: str,
        user_prompt: str,
        temperature: float = 0.3,
        max_tokens: int = 2048,
        response_format: dict | None = None,
    ) -> Optional[str]:
        """底层 API 调用"""
        if not self.is_available or self._session is None:
            return None
        try:
            content = self._chat_completion(
                [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=temperature,
                max_tokens=max_tokens,
                response_format=response_format,
            )
            self._connection_state = "connected"
            self._connection_message = f"已连接 {config.DEEPSEEK_MODEL}"
            return content
        except Exception as e:
            self._connection_state = "failed"
            self._connection_message = str(e)
            logger.warning(f"AI call failed: {e}")
            return None

    def chat(self, messages: list[dict], temperature: float = 0.3, max_tokens: int = 4096) -> tuple[bool, str]:
        """Run a multi-turn chat request and preserve the configured connection state."""
        if not self.is_available or self._session is None:
            return False, self.connection_message
        if not messages:
            return False, "消息不能为空"
        try:
            content = self._chat_completion(
                messages,
                temperature=temperature,
                max_tokens=max_tokens,
            )
            self._connection_state = "connected"
            self._connection_message = f"已连接 {config.DEEPSEEK_MODEL}"
            return True, content
        except Exception as error:
            self._connection_state = "failed"
            self._connection_message = str(error)
            logger.warning("AI chat failed: %s", error)
            return False, str(error)

    def _call_json(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        max_tokens: int = 2048,
        temperature: float = 0.2,
        response_format: dict | None = None,
        repair_prompt: str | None = None,
    ) -> Optional[dict]:
        """调用 API 并解析 JSON 响应"""
        text = self._call(
            system_prompt,
            user_prompt,
            temperature=temperature,
            max_tokens=max_tokens,
            response_format=response_format,
        )
        if text is None:
            return None
        def parse(text: str) -> Optional[dict]:
            text = text.strip()
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0]
            elif "```" in text:
                text = text.split("```")[1].split("```")[0]
            try:
                value = json.loads(text)
                return value if isinstance(value, dict) else None
            except json.JSONDecodeError:
                return None

        result = parse(text)
        if result is not None:
            return result
        logger.warning("AI response is not valid JSON; requesting one repair")
        if not repair_prompt:
            return None
        repaired = self._call(
            system_prompt,
            repair_prompt,
            temperature=0,
            max_tokens=max_tokens,
            response_format=response_format,
        )
        return parse(repaired or "") if repaired is not None else None

    def standardize_material_name(self, raw_name: str) -> dict:
        """标准化材料名称，解决同物不同名"""
        if not self.is_available:
            return {"standard_name": raw_name, "confidence": 0.0, "aliases": [], "category": ""}

        prompt = f"""你是工程造价材料专家。请将以下材料名称标准化为规范名称，并识别其分类和别名。
材料原始名称：{raw_name}

返回 JSON 格式：
{{
    "standard_name": "规范名称",
    "confidence": 0.0-1.0,
    "aliases": ["别名1", "别名2"],
    "category": "土建材料/装饰材料/安装材料/市政材料/园林材料/人工/机械/其他"
}}"""
        result = self._call_json("你是工程造价领域的材料标准化专家。只返回 JSON，不要解释。", prompt)
        if result is None:
            return {"standard_name": raw_name, "confidence": 0.0, "aliases": [], "category": ""}
        return {
            "standard_name": result.get("standard_name", raw_name),
            "confidence": result.get("confidence", 0.0),
            "aliases": result.get("aliases", []),
            "category": result.get("category", ""),
        }

    def suggest_prices(self, items: list[dict], region: str) -> list[dict]:
        """对清单项建议综合单价"""
        if not self.is_available or not items:
            return [{"seq_no": i.get("seq_no", idx+1), "suggested_price": None, "reference": "", "confidence": 0.0}
                    for idx, i in enumerate(items)]

        items_text = "\n".join(
            f"{i.get('seq_no', idx+1)}. {i.get('item_name', '')} | {i.get('spec', '')} | {i.get('unit', '')} | 工程量{i.get('quantity', '')}"
            for idx, i in enumerate(items)
        )
        prompt = f"""你是工程造价专家。请根据以下清单项和地区，给出综合单价建议和参考依据。

地区：{region}
清单项：
{items_text}

返回 JSON 数组：
[
    {{"seq_no": 序号, "suggested_price": 建议综合单价(数字), "unit": "单位", "reference": "参考依据说明", "confidence": 0.0-1.0}}
]
如果你不确定某项目，suggested_price 设为 null，confidence 设为 0。
不要凭空编造价格，基于常见工程造价数据给出参考。"""

        result = self._call_json("你是工程造价专家。基于实际工程数据给出合理参考价。只返回 JSON 数组。", prompt)
        if result is None:
            return [{"seq_no": i.get("seq_no", idx+1), "suggested_price": None, "reference": "AI 服务暂不可用", "confidence": 0.0}
                    for idx, i in enumerate(items)]
        if isinstance(result, dict):
            result = [result]
        return result

    def validate_quote(self, items: list[dict], region: str) -> dict:
        """验证报价合理性"""
        if not self.is_available:
            return {"overall_score": 0, "risk_level": "unknown", "warnings": [], "suggestions": ["未接入 AI，无法自动验证"]}

        items_text = "\n".join(
            f"{i.get('item_name','')} | {i.get('spec','')} | 数量{i.get('quantity','')} | 单价{i.get('unit_price','')} | 合价{i.get('total_price','')}"
            for i in items
        )
        prompt = f"""你是工程造价审计专家。请分析以下报价的合理性。

地区：{region}
报价清单：
{items_text}

返回 JSON：
{{
    "overall_score": 0-100,
    "risk_level": "low/medium/high",
    "warnings": ["警告1", "警告2"],
    "suggestions": ["建议1", "建议2"],
    "item_analysis": [
        {{"item_name": "名称", "price_reasonable": true/false, "comment": "评价", "expected_range": "合理区间"}}
    ]
}}"""

        result = self._call_json("你是工程造价审计专家。分析报价合理性并给出具体建议。只返回 JSON。", prompt)
        if result is None:
            return {"overall_score": 0, "risk_level": "unknown", "warnings": [], "suggestions": ["AI 调用失败"]}
        return result

    def search_official_sources(
        self,
        region: str,
        period: str = "",
        specialty: str = "",
        research_evidence: list[dict] | None = None,
    ) -> dict:
        """Classify verified source evidence without allowing AI-generated URLs."""
        evidence = [dict(item) for item in (research_evidence or []) if item.get("verified")]
        if not evidence:
            return {
                "sources": [],
                "search_note": "没有获得经搜索和官方域名验证的候选来源，未调用 AI 猜测网址。",
            }
        if not self.is_available:
            return {
                "sources": evidence,
                "search_note": "未配置 AI；已保留搜索引擎发现并验证的官方候选，可人工确认。",
                "ai_available": False,
            }

        period_text = period or datetime.now().strftime("%Y年%m月")
        specialty_text = specialty if specialty and specialty != "全部专业" else "各专业"
        compact_evidence = []
        for item in evidence[:10]:
            compact_evidence.append({
                "evidence_id": item.get("evidence_id"),
                "title": item.get("name", ""),
                "source_kind": item.get("source_kind", "web"),
                "http_status": item.get("http_status"),
                "content_type": item.get("content_type", ""),
                "description": item.get("description", "")[:400],
                "excerpt": item.get("excerpt", "")[:800],
                "api_hint": item.get("api_hint", ""),
                "login_required": bool(item.get("login_required")),
                "jurisdiction": item.get("jurisdiction", ""),
                "institution": item.get("institution", ""),
                "relevance_score": item.get("relevance_score", 0),
            })
        prompt = f"""请审核以下已经由程序实际搜索并完成官方域名验证的候选证据，判断其是否适合作为工程造价信息价来源。

地区：{region}
期数：{period_text}
专业：{specialty_text}
候选证据：
{json.dumps(compact_evidence, ensure_ascii=False)}

严格规则：
- 只能引用上面已有的 evidence_id，禁止输出网址，禁止增加候选
- 区分信息价网页、PDF/Excel 附件、公开查询接口和登录/会员页面
- 优先目标城市住建部门、造价站、定额站；其次才是明确覆盖该城市的省级主管部门
- 仅有招投标公告、新闻或政策说明而没有价格数据入口的，不推荐
- API 只能标为“接口候选”，除非证据已经显示 JSON/API 内容，不能声称已经可抓取

返回 JSON：
{{
  "sources": [
    {{
      "evidence_id": "E1",
      "recommended": true,
      "name": "来源名称",
      "source_kind": "api/web/pdf/excel",
      "login_required": false,
      "confidence": 0,
      "description": "判断依据"
    }}
  ],
  "search_note": "审核总结"
}}"""
        result = self._call_json(
            "你是工程造价官方信息源审核专家。只根据给定证据分类，不得生成网址。只返回 JSON。",
            prompt,
        )
        if not isinstance(result, dict):
            return {
                "sources": evidence,
                "search_note": "AI 审核失败；已保留程序验证过的候选供人工确认。",
                "ai_available": True,
            }

        classifications = {
            str(item.get("evidence_id") or ""): item
            for item in result.get("sources", [])
            if isinstance(item, dict) and item.get("evidence_id")
        }
        allowed_types = {"api", "web", "pdf", "excel"}
        reviewed = []
        for source in evidence:
            merged = dict(source)
            classification = classifications.get(str(source.get("evidence_id") or ""), {})
            if classification:
                merged["recommended"] = bool(classification.get("recommended", False))
                merged["ai_confidence"] = max(0, min(100, int(classification.get("confidence") or 0)))
                merged["ai_description"] = str(classification.get("description") or "")[:500]
                proposed_kind = str(classification.get("source_kind") or "").lower()
                if proposed_kind in allowed_types:
                    merged["source_kind"] = proposed_kind
                merged["login_required"] = bool(
                    source.get("login_required") or classification.get("login_required")
                )
                proposed_name = str(classification.get("name") or "").strip()
                if proposed_name:
                    merged["name"] = proposed_name[:300]
            else:
                merged.setdefault("recommended", False)
                merged["ai_description"] = "AI 未返回该候选的判断，保留程序验证结果。"
            reviewed.append(merged)
        return {
            "sources": reviewed,
            "search_note": str(result.get("search_note") or "AI 已按搜索证据完成分类。")[:1000],
            "ai_available": True,
        }

    def classify_material_category(self, name: str, spec: str = "") -> str:
        """辅助判断材料分类"""
        if not self.is_available:
            return ""

        prompt = f"""请判断以下材料属于哪个大类。
材料名称：{name}
规格：{spec}

可选分类：土建材料、装饰材料、安装材料、市政材料、园林材料、人工、机械、其他
返回 JSON：{{"category": "分类名", "reason": "判断理由"}}"""

        result = self._call_json("你是工程材料专家。只返回 JSON。", prompt)
        if result is None:
            return ""
        return result.get("category", "")

    def standardize_batch(self, names: list[str]) -> list[dict]:
        """批量标准化材料名称"""
        if not self.is_available or not names:
            return [{"original": n, "standard_name": n, "confidence": 0.0} for n in names]

        names_text = "\n".join(f"{i+1}. {n}" for i, n in enumerate(names))
        prompt = f"""请将以下材料名称统一标准化，合并同物不同名的情况。

材料列表：
{names_text}

返回 JSON 数组：
[
    {{"original": "原始名称", "standard_name": "标准名称", "confidence": 0.0-1.0, "note": "说明"}}
]"""

        result = self._call_json("你是工程材料标准化专家。只返回 JSON 数组。", prompt)
        if result is None:
            return [{"original": n, "standard_name": n, "confidence": 0.0} for n in names]
        return result

    def parse_unstructured_table(self, text: str) -> list[dict]:
        """从非结构化文本中提取材料价格表"""
        if not self.is_available:
            return []

        prompt = f"""请从以下文本中提取材料价格信息，整理为结构化表格。

文本内容：
{text[:3000]}

返回 JSON 数组：
[
    {{"material_name": "材料名称", "spec": "规格", "unit": "单位", "price": 价格数字, "period": "期数"}}
]
如果没有价格数据，返回空数组 []。"""

        result = self._call_json("你是数据提取专家。只返回 JSON 数组，不添加额外文字。", prompt)
        if result is None:
            return []
        return result if isinstance(result, list) else []

    def match_quota_candidates(self, boq: dict, candidates: list[dict]) -> dict:
        """Ask AI to review a bounded set of real local quota candidates."""
        if not candidates:
            return {
                "success": True,
                "matches": [],
                "summary": "本地硬性逻辑筛选后没有可供 AI 复核的候选定额。",
            }
        if not self.is_available:
            return {"success": False, "matches": [], "summary": self.connection_message}

        from src.quota_service import extract_quota_key_terms, _core_object_groups

        required_terms = extract_quota_key_terms(
            boq.get("name", ""),
            boq.get("feature", ""),
        )
        # A candidate rejected by the deterministic safety layer is evidence
        # for the prompt, never an option the model may select. Keeping it out
        # of both the prompt and the allow-list prevents a retry from reviving
        # a cross-major, object-conflicting, or specification-conflicting
        # quota.
        candidates = [
            candidate for candidate in candidates
            if candidate.get("logic_allowed", True)
        ]
        if not candidates:
            return {
                "success": True,
                "matches": [],
                "summary": "本地硬性逻辑筛选后没有可供 AI 复核的候选定额。",
            }
        allowed_ids = {
            int(candidate["quota_id"])
            for candidate in candidates
            if candidate.get("quota_id") is not None
        }
        safe_candidates = [
            {
                key: candidate.get(key)
                for key in (
                    "quota_id", "major", "code", "name", "feature", "unit",
                    "local_score", "name_score", "context_score", "code_score",
                    "logic_reasons", "logic_allowed", "logic_warnings", "components",
                )
            }
            for candidate in candidates
        ]
        payload = {
            "boq": {
                "code": str(boq.get("code") or ""),
                "name": str(boq.get("name") or ""),
                "feature": str(boq.get("feature") or ""),
                "unit": str(boq.get("unit") or ""),
                "quantity": boq.get("quantity"),
                "major": str(boq.get("major") or ""),
            },
            "required_terms": required_terms,
            "core_object_groups": sorted(_core_object_groups(
                f"{boq.get('name', '')} {boq.get('feature', '')}"
            )),
            "candidates": safe_candidates,
        }
        prompt = f"""请依据中国工程造价套定额逻辑，对一条清单与本地候选定额进行复核匹配。

必须遵守：
1. 只能从 candidates 中选择 quota_id，禁止编造定额、编码、价格或工料机。
2. 综合判断清单名称、编码、项目特征及工作内容、单位、施工部位、材料、工艺、规格；候选定额只看其类别、编码、名称、单位和价格组成明细。required_terms 是清单中辨识度最高的关键工艺或道路部位词，候选定额名称或组成必须覆盖这些词，或给出明确的同类工艺/部位换算理由，不能只靠“道路”“单位一致”或整体文字相似度选择。
2a. core_object_groups 是清单明确的核心工程对象。它们必须在候选定额名称或人材机组成中得到对应证据；例如“铝板景墙”不得选择“大理石墙面”，“预埋铁件”不得选择“幕墙预埋件”。颜色、厚度、尺寸、“综合考虑”等修饰词不能替代核心对象证据。
3. 一条清单最终只选择一个最准确的定额；该定额内部的人工、材料、辅材、机械等组成可以有多条，但不得并列套用多个定额。
4. 单位量纲、部位、拆除/新建、材料体系或明确规格冲突时不得选择。
5. 没有合理候选时 matches 返回空数组，不要勉强匹配。
6. confidence 是 AI 对造价逻辑适用性的置信度，不是文字相似度。

单位换算专项规则：清单单位与组价分量单位不同，必须核对 unitConversion、理论单位含量、损耗率和 totalQty；例如 m² 钢材应按面积×厚度×密度换算为 t/m²。若缺少完整物理参数，但已给出基于定额消耗或施工工效的正数单位含量，可保留为黄色可计算结果并说明依据；只有含量无依据、非正数或明显违反量纲/工程常识时才判定为高风险错误。

输入数据：
{json.dumps(payload, ensure_ascii=False, default=str)}

只返回以下 JSON：
{{
  "matches": [
    {{
      "quota_id": 真实候选ID,
      "role": "主体"或"补充工序",
      "confidence": 0到1,
      "reason": "说明名称、工作内容、单位、材料、工艺和规格的判断依据",
      "source_clause": "对应的清单工作内容"
    }}
  ],
  "summary": "总体判断；没有合适候选时说明缺少哪类定额"
}}"""
        result = self._call_json(
            "你是严谨的中国工程造价专业人员。只复核给定候选，不生成价格，只返回 JSON。",
            prompt,
            temperature=0.1,
            response_format={"type": "json_object"},
            repair_prompt=f"""上一次输出不是合法 JSON。请重新只返回一个 JSON 对象，不得使用 Markdown。
清单：{json.dumps(payload['boq'], ensure_ascii=False, default=str)}
候选定额ID：{sorted(allowed_ids)}""",
        )
        if not isinstance(result, dict):
            return {"success": False, "matches": [], "summary": "AI 返回内容无法解析。"}

        matches = []
        seen = set()
        raw_matches = result.get("matches")
        if not isinstance(raw_matches, list):
            raw_matches = []
        # The model is asked for one result, but legacy responses can contain
        # several. Read all of them so an invalid first item does not hide a
        # later valid candidate; the caller still receives one best match.
        for raw in raw_matches:
            if not isinstance(raw, dict):
                continue
            try:
                quota_id = int(raw.get("quota_id"))
                confidence = min(max(float(raw.get("confidence") or 0), 0.0), 1.0)
            except (TypeError, ValueError):
                continue
            if quota_id not in allowed_ids or quota_id in seen:
                continue
            seen.add(quota_id)
            matches.append({
                "quota_id": quota_id,
                "confidence": confidence,
                "reason": str(raw.get("reason") or "AI 未提供详细理由").strip(),
                "source_clause": str(raw.get("source_clause") or "AI 复核").strip(),
            })
        matches.sort(key=lambda value: float(value.get("confidence") or 0), reverse=True)
        if matches:
            matches = [{**matches[0], "role": "主体"}]
        return {
            "success": True,
            "matches": matches,
            "summary": str(result.get("summary") or "AI 已完成候选复核。").strip(),
        }

    def match_or_generate_quota(
        self,
        boq: dict,
        context: dict,
        repair_hints: list[str] | None = None,
    ) -> dict:
        """Review local candidates or generate a project-only market estimate."""
        if not self.is_available:
            return {
                "success": False,
                "decision": "none",
                "matches": [],
                "summary": self.connection_message,
            }
        candidates = [
            candidate for candidate in (context.get("candidates") or [])
            if candidate.get("logic_allowed", True)
        ]
        market_prices = context.get("market_prices") or []
        # Only logic-approved local quotas are legal existing matches. The
        # model can still generate a project-only quota when this set is empty.
        allowed_ids = {
            int(value["quota_id"])
            for value in candidates
            if value.get("quota_id") is not None
            and value.get("logic_allowed", True)
        }
        allowed_price_ids = {
            int(value["price_id"])
            for value in market_prices
            if value.get("price_id") is not None
        }
        payload = {
            "project": context.get("project") or {},
            "work_items": context.get("work_items") or [],
            "historical_project_references": context.get("historical_project_references") or [],
            "similar_project_results": context.get("similar_project_results") or [],
            "boq": {
                "code": str(boq.get("code") or ""),
                "name": str(boq.get("name") or ""),
                "feature": str(boq.get("feature") or ""),
                "unit": str(boq.get("unit") or ""),
                "quantity": boq.get("quantity"),
                "major": str(boq.get("major") or ""),
            },
            "required_terms": context.get("required_terms") or {},
            "material_requirements": context.get("material_requirements") or [],
            "resource_requirements": context.get("resource_requirements") or {},
            "core_object_groups": context.get("core_object_groups") or [],
            "core_material_families": context.get("core_material_families") or [],
            "candidates": candidates,
            "market_prices": market_prices,
            "current_composition": context.get("current_composition") or {},
        }
        prompt = f"""你是中国工程造价套定额和市场询价专家。请对一条清单做逐行复核。
目标是在造价逻辑允许范围内给出可复核结果，不要因文字不完全相同直接返回 none。
当 current_composition 存在时，这是一次“复核/补缺”任务：当前清单最终只能保留一个最准确的定额或一个 AI 补充定额；只纠正乱匹配、重复匹配、单位含量或计算错误，并补充该唯一结果中明确需要的人工、材料、机械、辅材、主材或专业分包组成。不得返回主体定额加多个补充定额，也不得为了补齐类别而添加与清单无关的分量。

必须同时依据：工程名称、项目特征及工作内容、清单编码、单位、工程量、work_items、页面选择的信息来源地区/信息价最新期、候选定额的类别/编码/名称/单位/价格明细、required_terms、已确认市场价。定额库历史“工作内容”字段不作为候选依据。required_terms 中的关键工艺或道路部位词必须由候选定额名称/组成或生成定额名称/组成覆盖，不能因为“道路”“单位一致”或整体文字相似度而选错工序。
core_object_groups 是更高优先级的工程对象证据。先核对核心对象/材料体系，再看工艺、规格、部位和价格；清单出现铝板、景墙、方钢、大理石、混凝土、路床等明确对象时，候选缺少对应对象或出现相冲突材料体系，不得写入 matches，也不得用长文本相似度放行。修饰词只用于规格和含量校验，不能主导定额选择。
            historical_project_references 是“企业参考定额表”中的历史案例，仅作为同类项目的辅助证据。只有工程对象、工艺、规格、单位和地区/计价期逻辑一致时才能参考其组成或价格；不得把历史综合单价直接当作正式定额，也不得隐去来源项目、地区和月份。历史案例单位不一致时仍必须做单位换算。
similar_project_results 是当前项目中已经形成组价的接近清单，最多3条，用于交叉分析人材机组成。先逐条执行 must_compare：只借鉴多条参考中与当前工程对象、工艺和做法一致的共同组成；参考之间不一致的分量不能按多数直接采用，必须回到当前清单特征判断。必须替换当前主材名称/规格，并按当前厚度、配合比、单位、地区和计价期重新计算含量与价格。不得照抄综合单价；qty 是理论含量，loss 单独计算，禁止把已含损耗的 qty 再乘损耗率。

地区与价格规则：
- 信息来源地区以 project.source_region 为唯一来源。已选择城市时只能使用该城市/省份的已确认信息价；未选择时使用信息价库最新期，不能读取项目城市、项目地址或项目计价日期来替换来源地区。
- 跨城市、跨计价期或模型估算的价格必须在 notes 和 calculation_basis 中写明来源地区、计价期与估算性质，并降低置信度；不能因为有一个价格就认为定额匹配成立。
- 信息来源地区为空不等于不能匹配：定额和人材机逻辑可继续处理，价格使用信息价库最新期或市场模型估算，并明确标注价格来源和复核风险。

规格硬约束：
- 清单明确 C25、C30、M10、DN100、厚度、HRB/Q 等规格时，候选必须含同一规格；C25 绝不能替换为 C30，规格缺失也不能当作一致。
- 外墙、内墙、室外、室内、楼地面、屋面等专用施工范围属于硬约束。清单没有明确“外墙”时，禁止选择名称或组成明确为“外墙”的定额；不得因为都包含“砖”“铺设”等泛词而放行。
- 规格冲突或候选缺少清单明确规格时，必须写入 coverage.uncovered，不能写入 matches；需要替代时只能生成项目级补充定额并说明换算和人工复核风险。

判断顺序：
1. 先逐条检查 work_items：每一条必须明确写入 matches.source_clause，或写入 coverage 中的 uncovered；不能用一个主体定额的总分数代替逐条核查。
2. existing：candidates 中只选择一个与清单对象、工艺、材料体系和单位口径最一致的定额，且候选 logic_allowed=true，或组成能明确覆盖清单；必须覆盖 required_terms 的约束；matches 数组最多1条，只能使用真实 quota_id。置信度低于0.40的候选不得写入 matches。
   当本专业没有可靠候选时，cross_major_candidate=true 的候选是系统经过全库召回后提供的复核证据；只有其对象、部位、工艺、规格、单位和组成全部适用时才能选用，并在 reason 中明确说明跨专业引用依据。专业名称本身不能掩盖逻辑冲突，也不能阻止一个实际适用的通用施工定额被复核采用。
3. generate：候选定额不适用或单位需要换算，但清单名称和单位能识别明确工程对象。即使 feature 为空、厚度/重量/做法未给出，也要按该对象最常见的安装组成生成项目级“AI补充定额/市场估算”，不要返回 none。
4. none：只有名称和单位都无法识别工程对象，或在一次受约束生成后仍无法给出任何正数量、正价格组成时才能返回。

生成规则：
- generated_quota 只是项目级参考，不得写入官方定额库。
- generated_quota.name 和 feature 必须保留清单工程对象核心词以及 required_terms 的关键工艺/部位词，严禁把“路床整形碾压”替换成“道路标线”等另一类工序。
- feature 只复述清单已明确内容；未明确内容写入 assumptions，不要伪装成清单事实。
- unit 可以与清单单位不同，但必须给出 unit_conversion，说明含量如何折算到一份清单工程量。
- components 只允许人工费、材料费、辅材费、主材费、机械费、专业分包。
- 每个 component.qty 表示每 1 份清单工程量需要多少 component.unit。
- 允许组成单位与清单单位不同，例如清单单位“块”，主材单位“块”，安装人工单位“工日”。
- 当组成单位与清单单位不同时，qty 必须明确表示“每1个清单计量单位的组成消耗量”。有清单明确尺寸时优先输出物理换算公式；没有完整尺寸、但能按定额消耗水平或施工工效推算出正数含量时，仍应输出可计算结果，并在 unitConversion/calculation_basis 写明“每1清单单位计取多少组成单位”的估算依据及复核风险，不要因此返回 none。不能把不同单位的 qty 无依据默认写成1。
- 简单换算可按工程常用理论公式计算；涉及钢构件组合、异形件、复杂损耗、多个规格或无法从清单确定换算参数时，必须标记“需要 AI 计算/人工确认”，不得猜测厚度、重量或含量。
- 主材优先对应清单中明确对象；不能增加与清单对象无关的主材。
- 可以推断厚度、单块重量、安装辅材和损耗，但必须写入 assumptions，说明是行业常规假设，confidence 相应降低。
- 对预埋铁件、钢板、钢构件、钢支架等钢材对象，如果清单没有明确钢材牌号、规格、长宽厚、单件重量或设计详图，允许按项目地区和最新月份采用市场常规假设（例如Q235B、常用板厚、合理损耗）形成暂估价格，但 assumptions、notes 和 calculation_basis 必须逐项写明假设、地区、月份、价格来源和适用范围，并降低置信度、标记人工复核。清单单位为 t 时，材料单位含量 1 表示每1吨清单的材料基数，是合理计量口径；只有单位不一致或没有换算公式时，才禁止使用含量1。
- 预埋铁件、钢盖板、幕墙埋件、钢支架等是不同工程对象；清单写“预埋铁件”时不得用“钢盖板”或“幕墙预埋件”替代，清单未写幕墙时不得增加幕墙施工范围。对象、部位或施工用途冲突时必须返回 none 或待确认，不得仅因都含“铁件/钢材”而匹配。
- 市场价只能引用 market_prices 中真实 price_id。没有匹配价格时，允许按地区同期常见市场价估算，source_evidence_ids=[], calculation_basis 以“市场模型估算：”开头，confidence 不高于 0.65。
- 含税单价=除税单价*(1+税率)；禁止把总工序价格当作组成价格；qty、loss、noTaxPrice 必须大于0。

输入数据：
{json.dumps(payload, ensure_ascii=False, default=str)}

只返回 JSON，结构如下：
{{
  "decision": "existing或generate或none",
  "matches": [{{"quota_id": 真实候选ID, "role": "唯一最佳定额", "confidence": 0.0, "reason": "依据", "source_clause": "对应清单条款"}}],
  "coverage": [{{"source_clause": "每一条 work_items", "status": "covered或uncovered", "reason": "覆盖依据或缺失原因"}}],
  "generated_quota": {{
    "major": "专业", "code": "AI补充编码", "name": "补充定额名称",
    "feature": "清单已明确的工作内容", "unit": "补充定额单位",
    "unit_conversion": "清单单位与定额单位的换算说明",
    "assumptions": ["未明确的厚度/重量/做法等假设"],
    "category": "AI补充定额/市场估算", "confidence": 0.0,
    "notes": "地区、计价期、边界和不确定性",
    "components": [{{
      "cat": "人工费等允许类别", "code": "", "name": "组成名称",
      "feature": "对应清单工作内容", "unit": "组成单位", "qty": 0,
      "unitConversion": "不同单位时填写可复核换算公式；相同单位填写空",
      "loss": 0, "noTaxPrice": 0, "taxRate": 0, "taxPrice": 0,
      "calculation_basis": "来源或估算方法", "source_evidence_ids": [], "note": ""
    }}]
  }},
  "summary": "选择或生成理由"
}}
禁止输出 Markdown 或解释文字。"""
        if repair_hints:
            prompt += "\n\n修复要求：\n" + "\n".join(
                f"- {value}" for value in repair_hints
            )
        result = self._call_json(
            "你是严谨的中国工程造价专业人员。优先给出可复核结果，再把缺失信息、换算和估算风险明确标出，只返回 JSON。",
            prompt,
            max_tokens=8192,
            temperature=0,
            response_format={"type": "json_object"},
            repair_prompt=f"""上一次输出不是合法 JSON。请严格只返回一个 JSON 对象，不得使用 Markdown。
重新复核清单：{json.dumps(boq, ensure_ascii=False, default=str)}
可用候选定额ID：{sorted(allowed_ids)}
如果候选不适用，必须生成带 components 的 generated_quota，并把不确定内容写入 assumptions。""",
        )
        if not isinstance(result, dict):
            detail = (
                self.connection_message
                if self.connection_state == "failed"
                else "AI 返回内容无法解析。"
            )
            return {
                "success": False,
                "decision": "none",
                "matches": [],
                "summary": detail,
            }

        matches = []
        seen = set()
        candidate_by_id = {
            int(value["quota_id"]): value
            for value in candidates
            if value.get("quota_id") is not None
        }
        allowed_apply_ids = {
            quota_id
            for quota_id, candidate in candidate_by_id.items()
            if candidate.get("logic_allowed", True)
        }
        raw_matches = result.get("matches") if isinstance(result.get("matches"), list) else []
        # Parse every returned candidate before applying the single-winner
        # contract. This keeps fallback matching usable when the provider
        # places a rejected candidate first.
        for raw in raw_matches:
            if not isinstance(raw, dict):
                continue
            try:
                quota_id = int(raw.get("quota_id"))
                confidence = min(max(float(raw.get("confidence") or 0), 0.0), 1.0)
            except (TypeError, ValueError):
                continue
            if quota_id not in allowed_apply_ids or quota_id in seen or confidence < 0.40:
                continue
            seen.add(quota_id)
            matches.append({
                "quota_id": quota_id,
                "confidence": confidence,
                "reason": str(raw.get("reason") or "AI 未提供详细理由").strip(),
                "source_clause": str(raw.get("source_clause") or "AI 复核").strip(),
            })

        matches.sort(key=lambda value: float(value.get("confidence") or 0), reverse=True)
        if matches:
            matches = [{**matches[0], "role": "主体"}]

        generated = result.get("generated_quota")
        if not isinstance(generated, dict):
            generated = None
        if generated:
            components = generated.get("components")
            if not isinstance(components, list):
                generated["components"] = []
            for component in generated.get("components") or []:
                if not isinstance(component, dict):
                    continue
                source_ids = component.get("source_evidence_ids")
                if not isinstance(source_ids, list):
                    component["source_evidence_ids"] = []
                    continue
                component["source_evidence_ids"] = [
                    int(value)
                    for value in source_ids
                    if str(value).lstrip("-").isdigit() and int(value) in allowed_price_ids
                ]
        decision = str(result.get("decision") or "none").strip().lower()
        if decision not in {"existing", "generate", "none"}:
            decision = "none"
        if decision == "existing" and not matches and generated and generated.get("components"):
            decision = "generate"
        if generated and generated.get("components") and decision == "none":
            decision = "generate"
        if decision == "generate" and not generated:
            decision = "none"
        return {
            "success": True,
            "decision": decision,
            "matches": matches,
            "coverage": [
                value for value in (result.get("coverage") or [])
                if isinstance(value, dict)
            ],
            "generated_quota": generated,
            "summary": str(result.get("summary") or "AI 未提供复核说明").strip(),
        }

    def research_quota(
        self,
        boq: dict,
        context: dict,
        research_evidence: list[dict],
    ) -> dict:
        """Use bounded web evidence to search for labor, material, and machinery."""
        if not self.is_available:
            return {
                "success": False,
                "decision": "none",
                "matches": [],
                "summary": self.connection_message,
            }
        if not research_evidence:
            return {
                "success": True,
                "decision": "none",
                "matches": [],
                "summary": "未检索到可用于本条清单的公开造价资料。",
            }

        evidence = []
        for record in redact_research_evidence(research_evidence):
            evidence.append({
                "evidence_id": str(record.get("evidence_id") or ""),
                "query": str(record.get("query") or ""),
                "title": str(record.get("title") or ""),
                "snippet": str(record.get("snippet") or "")[:1200],
                "excerpt": str(record.get("excerpt") or "")[:3500],
                "official": bool(record.get("official")),
                "trusted": bool(record.get("trusted")),
                "source_type": str(record.get("source_type") or ""),
                "material_name": str(record.get("material_name") or ""),
                "spec": str(record.get("spec") or ""),
                "unit": str(record.get("unit") or ""),
                "period": str(record.get("period") or ""),
                "market_price": record.get("market_price"),
                "market_tax_price": record.get("market_tax_price"),
                "price_basis": str(record.get("price_basis") or ""),
            })
        payload = {
            "project": context.get("project") or {},
            "work_items": context.get("work_items") or [],
            "boq": {
                "code": str(boq.get("code") or ""),
                "name": str(boq.get("name") or ""),
                "feature": str(boq.get("feature") or ""),
                "unit": str(boq.get("unit") or ""),
                "quantity": boq.get("quantity"),
                "major": str(boq.get("major") or ""),
            },
            "required_terms": context.get("required_terms") or {},
            "research_evidence": evidence,
        }
        prompt = f"""你正在为一条无法直接套定额的工程量清单查找人材机组成。只使用给出的检索证据，不能把模型记忆中的价格伪装成已检索到的价格。

判断要求：
1. 清单工程名称、项目特征及工作内容必须逐条复核，保留原工程对象核心词以及 required_terms 的关键工艺/部位词，不能替换成其他工程对象或工序。
2. 优先从证据的标题、摘要或正文中找出人工、材料、机械、专业分包的名称、单位、规格、价格和施工做法。
3. 证据中确实出现数字价格时，才能写入 noTaxPrice/taxPrice；不能只看到材料名称就编造价格。输出不得包含网址。
4. 证据没有完整价格时，可以生成市场模型估算，但 source_urls 必须为空，calculation_basis 以“市场模型估算：”开头，confidence 不得高于 0.55。
5. 网页检索资料不是官方确认信息价，只能作为项目级待复核证据；正式报价仍需人工确认地区、期数和税价口径。
6. unit 与清单单位不一致时必须给出 unit_conversion 和对应含量；面积转体积只能使用清单明确厚度或检索证据中的换算依据。复杂组合或缺少换算参数时，必须标记需要 AI 计算/人工确认，不能把 qty 默认写成 1。
7. components 只允许人工费、材料费、辅材费、主材费、机械费、专业分包；每个 qty、loss、noTaxPrice 必须大于等于 0 且合理。
8. source_urls 必须为空；来源只通过 evidence_id、地区、期数和来源类型说明。

输入数据：
{json.dumps(payload, ensure_ascii=False, default=str)}

只返回 JSON，结构如下：
{{
  "decision": "generate或none",
  "generated_quota": {{
    "major": "专业",
    "code": "AI检索补充编码",
    "name": "必须含 required_terms 关键工艺/部位的补充定额名称",
    "feature": "必须含 required_terms 关键工艺/部位的清单已明确内容",
    "unit": "补充定额单位",
    "unit_conversion": "单位换算说明",
    "assumptions": ["无法确认的规格、损耗或含量假设"],
    "category": "AI联网检索补充定额",
    "confidence": 0.0,
    "notes": "检索边界、来源和风险",
    "components": [{{
      "cat": "人工费等允许类别",
      "code": "",
      "name": "组成名称",
      "feature": "对应清单工作内容",
      "unit": "组成单位",
      "qty": 0,
      "unitConversion": "不同单位时填写可复核换算公式；相同单位填写空",
      "loss": 0,
      "noTaxPrice": 0,
      "taxRate": 0,
      "taxPrice": 0,
      "calculation_basis": "价格检索依据或市场模型估算，不得包含网址",
      "source_urls": [],
      "evidence_excerpt": "支持该分量的证据片段",
      "note": ""
    }}]
  }},
  "summary": "检索和生成理由"
}}
无法形成至少一个可计算分量时 decision 为 none。禁止输出 Markdown 或解释文字。"""
        result = self._call_json(
            "你是严谨的中国工程造价专业人员。只根据给定检索证据识别可追溯的人材机，不编造官方价格，只返回 JSON。",
            prompt,
            max_tokens=8192,
            temperature=0,
            response_format={"type": "json_object"},
        )
        if not isinstance(result, dict):
            return {
                "success": False,
                "decision": "none",
                "matches": [],
                "summary": "AI 检索结果无法解析。",
            }
        generated = result.get("generated_quota")
        if not isinstance(generated, dict):
            return {
                "success": True,
                "decision": "none",
                "matches": [],
                "summary": str(result.get("summary") or "检索资料无法形成可计算组成。"),
            }
        allowed_urls = set()
        components = generated.get("components")
        if not isinstance(components, list):
            components = []
        for component in components:
            if not isinstance(component, dict):
                continue
            source_urls = component.get("source_urls")
            if not isinstance(source_urls, list):
                source_urls = []
            component["source_urls"] = [str(value) for value in source_urls if str(value) in allowed_urls]
            component["calculation_basis"] = redact_text(
                component.get("calculation_basis") or ""
            )
        generated["components"] = components
        generated["source_type"] = "ai_web_research"
        decision = str(result.get("decision") or "none").strip().lower()
        if decision != "generate" or not components:
            decision = "none"
        return {
            "success": True,
            "decision": decision,
            "matches": [],
            "generated_quota": generated,
            "summary": str(result.get("summary") or "AI 已完成联网检索分析。").strip(),
            "research_sources": redact_research_evidence(research_evidence),
        }

    def audit_quota_composition(
        self,
        boq: dict,
        context: dict,
        current: dict,
    ) -> dict:
        """Review an existing quota composition for logic errors and omissions."""
        if not self.is_available:
            return {
                "success": False,
                "findings": [],
                "needs_correction": False,
                "summary": self.connection_message,
            }
        from src.quota_service import _core_object_groups
        payload = {
            "project": context.get("project") or {},
            "boq": {
                "code": str(boq.get("code") or ""),
                "name": str(boq.get("name") or ""),
                "feature": str(boq.get("feature") or ""),
                "unit": str(boq.get("unit") or ""),
                "quantity": boq.get("quantity"),
                "major": str(boq.get("major") or ""),
            },
            "work_items": context.get("work_items") or [],
            "required_terms": context.get("required_terms") or {},
            "candidates": context.get("candidates") or [],
            "current_quota_matches": [
                {
                    "role": str(value.get("role") or ""),
                    "quota_id": value.get("quota_id"),
                    "quota_code": str(value.get("quota_code") or ""),
                    "quota_name": str(value.get("quota_name") or ""),
                    "score": value.get("score"),
                    "evidence_level": str(value.get("evidence_level") or ""),
                    "source_type": str(value.get("source_type") or ""),
                    "reasons": value.get("reasons") or [],
                }
                for value in current.get("quota_matches") or []
            ],
            "current_components": [
                {
                    key: value.get(key)
                    for key in (
                        "cat", "category", "code", "name", "feature", "unit",
                        "qty", "loss", "loss_rate", "noTaxPrice", "no_tax_price",
                        "taxRate", "tax_rate", "taxPrice", "tax_price",
                        "boqUnit", "engineeringUnit", "unitConversion", "unit_conversion",
                        "totalQty", "engineeringQuantity", "engineeringNoTaxTotal",
                        "quotaRole", "quotaName", "evidenceLevel", "note",
                    )
                }
                for value in current.get("compositions") or []
            ],
            "local_audit_issues": current.get("audit_issues") or [],
            "local_audit_report": current.get("local_audit_report") or {},
            "current_evidence_level": str(current.get("evidence_level") or ""),
            "current_score": current.get("score"),
            "current_generated": current.get("generated_quota") or {},
            "core_object_groups": sorted(_core_object_groups(
                f"{boq.get('name', '')} {boq.get('feature', '')}"
            )),
        }
        prompt = f"""请复核一条已经套好定额的工程量清单。必须按固定三类核查机制输出，不能只给模糊结论。

第一类：人材机完整性
1. 项目特征和工作内容中逐条出现的人工、材料、机械、专业分包必须全部找到对应分量或定额依据。
2. 明确区分“漏匹配”和“合理合并”，例如一条工作内容可以由一条主体定额覆盖，不要求机械地拆成无意义的多条。
3. 同一份定额、同一个工料机、同一道工序重复计入必须报告为“重复匹配”。

第二类：乱匹配
1. 清单工程对象不能被替换，例如检查井盖不能套成检查井，灌木不能套成石材养护，围栏不能套成电子围栏。
2. 先核对输入中的 core_object_groups。核心对象/材料体系不一致时，必须判定为乱匹配；颜色、厚度、尺寸和“综合考虑”等修饰词不能掩盖对象冲突。
3. 核对材料体系、施工部位、拆除/新建、规格、厚度、单位、工艺和施工条件；required_terms 中的关键工艺或道路部位词未覆盖时，必须报告为“乱匹配”。
4. 低相似度替补、AI市场估算、无本地价格证据必须明确标出。

第三类：计算正确性
1. 每个分量按 含量 × 调整系数 × (1+损耗率) × 单价 验证除税和含税合计。
2. 清单单位与组价分量单位不同，优先核对 unitConversion、理论单位含量、损耗率和 totalQty；有明确尺寸时必须按物理公式换算。缺少完整尺寸但存在正数且符合定额消耗/施工工效的单位含量时，标黄并保留计算；只有无依据、非正数或明显违反工程常识时才判定高风险，不能把单位含量无依据默认为1。
3. 核对含税单价=除税单价×(1+税率)，税率、损耗率、含量和单价必须合理。
4. 核对人工费、材料费、辅材费、主材费、机械费、专业分包、管理费、利润和综合单价的汇总关系。
5. 检查重复计费、漏计管理费/利润、税价口径混用、负值/零值和总量不一致。

第四类：综合价合理性
1. 在组成和计算正确的基础上，判断套定额综合单价是否明显偏离原清单或历史案例参考价；参考价只作预警证据，不能因为偏差自动改价。
2. 核对综合单价×工程量是否等于清单合价，识别单位工程量、税价口径或小数位错误。
3. 检查管理费/利润占直接费比例异常、综合单价远高于直接费、有效组成却综合价为0等情况。
4. 综合价异常时必须进一步回溯定额对象、重复工序、单位含量、材料价格地区和计价月份，不能只按价格高低判定正确或错误。

输入数据：
{json.dumps(payload, ensure_ascii=False, default=str)}

只返回 JSON：
{{
  "findings": [
    {{
      "severity": "high/medium/low",
      "category": "人材机遗漏/重复匹配/乱匹配/计算错误/其他",
      "type": "更具体的问题类型",
      "message": "问题说明",
      "affected_component": "涉及的分量或定额",
      "expected_calculation": "应有计算口径",
      "actual_calculation": "当前实际值",
      "recommended_action": "纠正建议"
    }}
  ],
  "needs_correction": true或false,
  "reason": "总体判断",
  "summary": "复核结论"
}}
没有可靠问题时可返回空 findings 和 needs_correction=false。禁止输出 Markdown 或解释文字。"""
        result = self._call_json(
            "你是严谨的中国工程造价审计专业人员。只根据清单、当前组价和证据判断，不凭空修改价格。只返回 JSON。",
            prompt,
            max_tokens=4096,
            temperature=0,
            response_format={"type": "json_object"},
        )
        if not isinstance(result, dict):
            return {
                "success": False,
                "findings": [],
                "needs_correction": False,
                "summary": "AI 核查结果无法解析。",
            }
        findings = result.get("findings")
        if not isinstance(findings, list):
            findings = []
        return {
            "success": True,
            "findings": findings,
            "needs_correction": bool(result.get("needs_correction")),
            "reason": str(result.get("reason") or "").strip(),
            "summary": str(result.get("summary") or "AI 已完成当前组价核查。").strip(),
        }


ai = AIService()
