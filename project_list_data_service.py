"""Import and search historical project BOQ rows used by quota matching."""
from __future__ import annotations

import csv
import io
import json
import re
import zipfile
from pathlib import Path

from rapidfuzz import fuzz
from sqlalchemy import or_
from sqlalchemy.orm import Session

from .models import ProjectListData
from .excel_compat import read_excel_sheets
from .bill_list_extractor import extract_specialized_bill_list


HEADERS = [
    "序号", "工程名称", "项目特征及工作内容", "单位", "工程量", "综合单价分析",
    "综合单价", "合价", "人工费", "材料费", "机械费", "管理费", "利润", "备注",
]
ENTERPRISE_REFERENCE_TITLE = "企业参考定额表"
ENTERPRISE_MAIN_SHEET = "企业定额参考价总表"
ENTERPRISE_DETAIL_SHEET = "定额子目组成明细"
ALIASES = {
    "seq_no": ("序号", "编号", "顺序号"),
    "item_code": ("项目编码", "清单编码", "定额编码", "企业定额编码", "企业编码", "编码"),
    "item_name": ("工程名称", "项目名称", "清单名称", "清单项目", "项目特征名称", "定额名称", "定额子目名称", "设备名称", "产品名称", "型号", "名称"),
    "feature": ("项目特征及工作内容", "项目特征描述", "项目特征", "项目描述", "产品描述", "工作内容", "特征"),
    "work_content": ("工作内容", "施工内容", "工作内容描述", "施工工艺"),
    "unit": ("单位", "计量单位"),
    "quantity": ("工程量", "数量", "清单工程量", "定额工程量", "主材采购量", "采购量"),
    "analysis": ("综合单价分析", "组价分析", "价格分析"),
    "comprehensive_price": ("综合单价（元）", "综合单位（元）", "综合单价元", "综合单位元", "综合单价", "综合单位", "全费用综合单价", "子目单价(元)", "子目单价", "单价（元）", "单价(元)", "单价", "主材单价"),
    "total_price": ("合价（元）", "合价元", "合价", "总价", "综合合价", "金额（元）", "金额元", "金额", "不含税综合合价（元）", "不含税综合合价(元)"),
    "labor_cost": ("人工费", "人工费(元)"),
    "material_cost": ("材料费", "材料费(元)"),
    "machinery_cost": ("机械费", "机械费(元)"),
    "management_cost": ("管理费", "管理费(元)"),
    "profit": ("利润", "利润(元)"),
    "note": ("备注", "使用说明", "说明", "复核说明"),
}


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _number(value):
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", "").replace("%", "").strip())
    except (TypeError, ValueError):
        return None


_UNIT_PRICE_COST_FIELDS = (
    "labor_cost",
    "material_cost",
    "machinery_cost",
    "management_cost",
    "profit",
)


def _relative_difference(left, right) -> float:
    """Compare two amounts without making zero/blank values look equivalent."""
    if left is None or right is None:
        return float("inf")
    left = abs(float(left))
    right = abs(float(right))
    denominator = max(left, right, 0.000001)
    return abs(left - right) / denominator


def _normalize_unit_prices(
    *,
    quantity,
    comprehensive_price,
    total_price,
    costs: dict,
) -> tuple[dict, dict]:
    """Convert imported aggregate amounts to prices for one BOQ unit.

    Excel exports are inconsistent: ``综合单价`` is normally already a unit
    price, while cost breakdown columns may be either unit amounts or row
    totals. Conversion is therefore evidence-based. A breakdown is divided
    only when it agrees with the row total or with the explicit unit price
    after division. The original values remain in the caller's raw record.
    """
    normalized = {
        field: _number(costs.get(field))
        for field in _UNIT_PRICE_COST_FIELDS
    }
    quantity = _number(quantity)
    comprehensive_price = _number(comprehensive_price)
    total_price = _number(total_price)
    explicit_unit_price = comprehensive_price if comprehensive_price and comprehensive_price > 0 else None
    derived_unit_price = (
        total_price / quantity
        if total_price is not None and total_price > 0 and quantity and quantity > 0
        else None
    )
    unit_price = explicit_unit_price or derived_unit_price
    component_sum = sum(value for value in normalized.values() if value is not None and value > 0)
    conversion = False
    conversion_basis = ""

    if quantity and quantity > 0 and abs(quantity - 1.0) > 0.000001 and component_sum > 0:
        # Exact row-total evidence is strongest and also works for quantities
        # below one, such as a 0.85 t steel item.
        if total_price is not None and _relative_difference(component_sum, total_price) <= 0.25:
            conversion = True
            conversion_basis = f"费用组成合计 {component_sum:g} ÷ 工程量 {quantity:g}"
        # Some exports omit 合价 but contain 综合单价 and aggregate cost
        # columns. Allow a wider tolerance because fees/tax may be omitted.
        elif unit_price is not None and _relative_difference(component_sum / quantity, unit_price) <= 0.35:
            conversion = True
            conversion_basis = f"费用组成合计 {component_sum:g} ÷ 工程量 {quantity:g} ≈ 单位综合单价 {unit_price:g}"

    if conversion:
        normalized = {
            field: value / quantity if value is not None else None
            for field, value in normalized.items()
        }

    if comprehensive_price is None or comprehensive_price <= 0:
        comprehensive_price = derived_unit_price
    metadata = {
        "unit_price_basis": "per_boq_unit",
        "source_quantity": quantity,
        "source_comprehensive_price": explicit_unit_price,
        "source_total_price": total_price,
        "component_totals_converted": conversion,
        "conversion_formula": conversion_basis,
        "note": (
            f"已按每1个清单单位计价：{conversion_basis}"
            if conversion_basis
            else "按每1个清单单位保存单位价格；原工程量和总价仅作追溯"
        ),
    }
    result = {"comprehensive_price": comprehensive_price, **normalized}
    return result, metadata
def _header_key(value) -> str:
    return re.sub(r"[\s\n\r（）()【】\[\]：:、,，]+", "", _text(value)).lower()


def _find_mapping(headers, aliases=None):
    aliases = aliases or ALIASES
    normalized = {_header_key(value): index for index, value in enumerate(headers) if _text(value)}
    mapping = {}
    for field, field_aliases in aliases.items():
        for alias in field_aliases:
            if _header_key(alias) in normalized:
                mapping[field] = normalized[_header_key(alias)]
                break
    return mapping


def _cell(row, mapping, field):
    index = mapping.get(field)
    return row[index] if index is not None and index < len(row) else ""


def _combined_text(row, mapping, *fields):
    values = []
    for field in fields:
        value = _text(_cell(row, mapping, field))
        if value and value not in values:
            values.append(value)
    return "\n".join(values)


def _read_text_table(data: bytes):
    """Read text/CSV exports that were given an Excel extension by another tool."""
    encodings = []
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        encodings.extend(("utf-16", "utf-16-le", "utf-16-be"))
    encodings.extend(("utf-8-sig", "gb18030", "utf-16-le", "utf-16-be"))
    seen = set()
    for encoding in encodings:
        if encoding in seen:
            continue
        seen.add(encoding)
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        if not text.strip():
            continue
        delimiter = "\t" if "\t" in text else ","
        rows = [tuple(row) for row in csv.reader(io.StringIO(text), delimiter=delimiter)]
        if any(
            _text(cell) in {alias for values in ALIASES.values() for alias in values}
            for row in rows[:60]
            for cell in row
        ):
            return [("文本表格", rows)]
    return None


def _read_sheets(file_path: Path):
    """企业参考定额和项目清单共用统一 Excel 读取器。"""
    return read_excel_sheets(file_path)


def _find_header(rows, required_fields, aliases=None, scan_limit=60):
    for index, row in enumerate(rows[:scan_limit]):
        mapping = _find_mapping(row, aliases)
        if all(field in mapping for field in required_fields):
            return index, mapping
    return None, {}


_HEADER_SIGNAL_TOKENS = (
    "序号", "项目编码", "项目名称", "清单名称", "项目特征", "工程内容",
    "工作内容", "规格型号", "计量单位", "单位", "工程量", "数量",
    "综合单价", "合价", "总价", "金额", "人工费", "主材费",
    "材料费", "辅材费", "机械费", "管理费", "利润", "税率",
    "备注", "说明",
)


def _flatten_header_rows(header_rows):
    """Flatten merged/multi-level headers while preserving column positions."""
    width = max((len(row) for row in header_rows), default=0)
    flattened = []
    for column in range(width):
        labels = []
        for row in header_rows:
            value = _text(row[column]) if column < len(row) else ""
            if value and value not in labels:
                labels.append(value)
        flattened.append("".join(labels))
    return flattened


def _standard_header_mapping(header_rows):
    """Map common BOQ headers, including two-level Glodon-style headers."""
    columns = _flatten_header_rows(header_rows)
    normalized = [_header_key(value) for value in columns]
    mapping = {}

    def exact(*values):
        wanted = {_header_key(value) for value in values}
        return next((index for index, value in enumerate(normalized) if value in wanted), None)

    def row_exact(*values):
        wanted = {_header_key(value) for value in values}
        for row in reversed(header_rows):
            for index, value in enumerate(row):
                if _header_key(value) in wanted:
                    return index
        return None

    def first_index(*values):
        index = row_exact(*values)
        return index if index is not None else exact(*values)

    def contains(predicate):
        return next((index for index, value in enumerate(columns) if predicate(value, _header_key(value))), None)

    mapping["seq_no"] = first_index("序号", "编号", "顺序号")
    mapping["item_code"] = first_index(
        "项目编码", "清单编码", "定额编码", "企业定额编码", "企业编码", "编码",
    )
    mapping["item_name"] = first_index(
        "工程名称", "项目名称", "清单名称", "清单项目", "定额名称", "名称",
    )
    mapping["feature"] = first_index(
        "项目特征及工作内容", "项目特征描述", "项目特征", "项目描述", "特征",
    )
    mapping["work_content"] = first_index(
        "工程内容", "工作内容", "施工内容", "工作内容描述", "施工工艺",
    )
    mapping["unit"] = first_index("计量单位", "单位")
    mapping["quantity"] = first_index("工程量", "清单工程量", "数量", "工程数量", "清单数量")
    mapping["analysis"] = first_index("综合单价分析", "组价分析", "价格分析")

    # In a two-level header, the parent label may say "综合单价组价明细"
    # while a child column is the actual pretax unit price. Prefer the
    # explicit pretax column and never mistake the labor/material child
    # columns for the comprehensive price.
    mapping["comprehensive_price"] = None
    for index, value in enumerate(columns):
        key = _header_key(value)
        # A merged parent like 税前综合单价组价明细 is not the unit-price
        # column. The actual value is the child column 税前综合单价（元）.
        if key in {"税前综合单价元", "综合单价元"}:
            mapping["comprehensive_price"] = index
            break
    if mapping["comprehensive_price"] is None:
        for index, row in enumerate(header_rows):
            for column, value in enumerate(row):
                key = _header_key(value)
                if key in {"税前综合单价元", "综合单价元"}:
                    mapping["comprehensive_price"] = column
                    break
            if mapping["comprehensive_price"] is not None:
                break
    if mapping["comprehensive_price"] is None:
        mapping["comprehensive_price"] = contains(
        lambda text, key: "综合单价" in text
        and "组价明细" not in text
        and "税后" not in text
        and "含税" not in text
        and ("税前" in text or "元" in text)
        )
    if mapping["comprehensive_price"] is not None:
        # Ignore the merged parent "税前综合单价组价明细". When the
        # workbook has a child row, its explicit numeric label wins.
        current_label = _header_key(columns[mapping["comprehensive_price"]])
        if "组价明细" in current_label:
            mapping["comprehensive_price"] = None
            for index, row in enumerate(header_rows):
                for column, value in enumerate(row):
                    if _header_key(value) in {"税前综合单价元", "综合单价元"}:
                        mapping["comprehensive_price"] = column
                        break
                if mapping["comprehensive_price"] is not None:
                    break
    if mapping["comprehensive_price"] is None:
        mapping["comprehensive_price"] = contains(
            lambda text, key: "综合单价" in text
            and "组价明细" not in text
            and "税后" not in text
            and "含税" not in text
        )
    mapping["total_price"] = first_index(
        "含税总价", "合价（元）", "合价", "清单合价", "工程合价",
        "总价", "金额（元）", "金额",
    )
    if mapping["total_price"] is None:
        mapping["total_price"] = contains(
            lambda text, key: any(token in text for token in ("合价", "总价", "金额"))
        )
    mapping["main_material_loss"] = first_index("主材损耗率", "损耗率")
    mapping["composite_fee_rate"] = first_index("综合费率", "综合费用率")
    if mapping["composite_fee_rate"] is None:
        mapping["composite_fee_rate"] = contains(
            lambda text, key: "综合费率" in key or "综合费用率" in key
        )
    mapping["tax_rate"] = first_index("税率")
    # "税后综合单价" is a parent header placed one column before the
    # numeric child column. Prefer the explicit child "综合单价（元）"
    # after the tax-rate column; otherwise a tax rate is stored as the price.
    tax_rate_index = mapping.get("tax_rate")
    tax_inclusive_index = None
    if tax_rate_index is not None:
        for index in range(tax_rate_index + 1, len(columns)):
            child = _header_key(
                header_rows[-2][index]
                if len(header_rows) >= 2 and index < len(header_rows[-2])
                else ""
            )
            parent = _header_key(
                header_rows[0][index]
                if header_rows and index < len(header_rows[0])
                else ""
            )
            if child == "综合单价元" or parent in {"含税综合单价", "税后综合单价"}:
                tax_inclusive_index = index
                break
    mapping["tax_inclusive_price"] = (
        tax_inclusive_index
        if tax_inclusive_index is not None
        else next(
            (index for index, value in enumerate(columns)
             if _header_key(value) in {"综合单价元", "含税综合单价", "税后综合单价"}
             and (tax_rate_index is None or index > tax_rate_index)),
            None,
        )
    )

    # Parent labels such as "综合费率（含管理费、规费、利润）" are not
    # fee amounts. Only exact child labels may populate these columns.
    field_aliases = {
        "labor_cost": ("人工费",),
        "material_cost": ("材料费",),
        "main_material_cost": ("主材费",),
        "auxiliary_material_cost": ("辅材费",),
        "machinery_cost": ("机械费",),
        "management_cost": ("管理费",),
        "profit": ("利润",),
        "note": ("备注", "说明"),
    }
    for field, aliases in field_aliases.items():
        mapping[field] = first_index(*aliases)
    return mapping


def _is_header_continuation(row):
    text = "".join(_text(value) for value in row if _text(value))
    return bool(text) and any(token in text for token in _HEADER_SIGNAL_TOKENS)


def _find_standard_header(rows):
    def has_price(mapping):
        return any(mapping.get(field) is not None for field in (
            "comprehensive_price", "total_price", "labor_cost", "material_cost",
            "main_material_cost", "auxiliary_material_cost", "machinery_cost",
            "management_cost", "profit",
        ))

    for index in range(min(len(rows), 60)):
        depth = 1
        while depth < 3 and index + depth < len(rows) and _is_header_continuation(rows[index + depth]):
            depth += 1
        mapping = _standard_header_mapping(rows[index:index + depth])
        if mapping.get("item_name") is not None and mapping.get("unit") is not None and mapping.get("quantity") is not None and has_price(mapping):
            mapping["_header_depth"] = depth
            mapping["_format"] = (
                "项目清单参考价（含组价明细）"
                if mapping.get("main_material_cost") is not None
                or mapping.get("auxiliary_material_cost") is not None
                else "普通项目清单"
            )
            return index, mapping, False

    header, mapping = _find_header(rows, ("item_name", "unit", "quantity"))
    if header is not None and has_price(mapping):
        return header, mapping, False
    # Some reference sheets provide a unit price but intentionally omit a
    # project quantity; those rows are normalized to one reference unit.
    for index, row in enumerate(rows[:60]):
        mapping = _standard_header_mapping([row])
        if mapping.get("item_name") is not None and mapping.get("unit") is not None and has_price(mapping):
            mapping["_header_depth"] = 1
            mapping["_format"] = "普通项目清单"
            return index, mapping, True
    return None, {}, False


_SUMMARY_NAME_MARKERS = (
    "合计", "小计", "汇总", "总计", "措施项目", "其他项目", "规费", "税金",
)
_WORK_TYPE_MARKERS = (
    "找平", "铺贴", "瓷砖", "石材", "防水", "美缝", "吊顶", "乳胶漆",
    "涂料", "抹灰", "门窗", "水电", "灯具", "基层", "安装", "运输",
)


def _is_aggregate_reference_row(name: str, unit: str, feature: str, mapping: dict, row: tuple) -> bool:
    """Reject section totals and multi-process lump-sum rows, keep real items."""
    compact = re.sub(r"\s+", "", _text(name))
    if any(marker in compact for marker in _SUMMARY_NAME_MARKERS):
        return True
    normalized_unit = re.sub(r"[㎡m²^2]", "m2", _text(unit).lower())
    if normalized_unit not in {"项", "项次", "item"}:
        return False
    # A standalone item with a price is allowed. A parenthesized list of
    # several trades/processes is a package price, not a reusable quota row.
    marker_count = sum(1 for marker in _WORK_TYPE_MARKERS if marker in compact or marker in _text(feature))
    has_list = any(token in compact for token in ("、", "/", "，", ",", "等", "及"))
    has_package_word = any(token in compact for token in ("工程", "装修", "装饰", "综合", "包干", "一项"))
    return marker_count >= 2 and has_list and has_package_word


def _has_reference_price(row: tuple, mapping: dict) -> bool:
    fields = (
        "comprehensive_price", "total_price", "labor_cost", "material_cost",
        "main_material_cost", "auxiliary_material_cost", "machinery_cost",
        "management_cost", "profit",
    )
    return any(_number(_cell(row, mapping, field)) is not None for field in fields)


def _reference_analysis(row: tuple, mapping: dict) -> str:
    """Preserve the source's grouped pricing columns as readable evidence."""
    labels = (
        ("人工费", "labor_cost"),
        ("主材费", "main_material_cost"),
        ("主材损耗率", "main_material_loss"),
        ("辅材费", "auxiliary_material_cost"),
        ("机械费", "machinery_cost"),
        ("综合费率", "composite_fee_rate"),
        ("税前综合单价", "comprehensive_price"),
        ("税率", "tax_rate"),
        ("含税综合单价", "tax_inclusive_price"),
        ("含税总价", "total_price"),
    )
    values = []
    for label, field in labels:
        value = _cell(row, mapping, field)
        if value not in (None, ""):
            values.append(f"{label}：{_text(value)}")
    return "；".join(values)


_REGION_HINTS = (
    "北京", "天津", "上海", "重庆", "河北", "山西", "辽宁", "吉林", "黑龙江",
    "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南",
    "广东", "广西", "海南", "四川", "贵州", "云南", "西藏", "陕西", "甘肃",
    "青海", "宁夏", "新疆", "内蒙古", "成都", "绵阳", "德阳", "乐山", "泸州",
    "贵阳", "昆明", "西安", "兰州", "杭州", "宁波", "南京", "苏州", "无锡",
    "上海", "武汉", "郑州", "长沙", "广州", "深圳", "济南", "青岛", "合肥",
)


def _infer_region_period(sheets, text_hint: str = ""):
    values = []
    for _sheet_name, rows in sheets:
        for row in rows[:30]:
            values.extend(_text(value) for value in row if _text(value))
    text = " ".join(values) + " " + _text(text_hint)
    region = next((hint for hint in _REGION_HINTS if hint in text), "")
    match = re.search(
        r"(20\d{2})\s*(?:年\s*|[-/.])\s*(1[0-2]|0?[1-9])\s*(?:月|[-/.]|$)",
        text,
    )
    period = f"{match.group(1)}-{int(match.group(2)):02d}" if match else ""
    return region, period


def _infer_source_project(sheets, text_hint: str = "") -> str:
    """Find a human project title without treating section names as projects."""
    candidates = []
    for sheet_name, rows in sheets:
        for row in rows[:8]:
            for value in row[:8]:
                text = _text(value)
                if not text:
                    continue
                match = re.search(r"(?:工程名称|项目名称)\s*[：:]\s*(.+)", text)
                if match:
                    candidates.append(match.group(1).strip())
                elif any(token in text for token in ("工程量清单", "报价清单")):
                    cleaned = re.sub(r"(?:工程量清单|报价清单).*$", "", text).strip(" ：:－-_")
                    if len(cleaned) >= 4:
                        candidates.append(cleaned)
    if candidates:
        return max(candidates, key=len)
    return _text(text_hint)


def _enterprise_format(sheets):
    """Recognize enterprise reference sheets by headers, not filename or tab names."""
    main_aliases = {
        "item_code": ("企业定额编码", "企业编码"),
        "item_name": ("项目名称",),
        "feature": ("项目特征描述", "项目特征及工作内容", "项目特征", "工作内容", "施工内容"),
        "work_content": ("工作内容", "施工内容", "工作内容描述", "施工工艺"),
        "unit": ("单位",),
        "comprehensive_price": ("全费用综合单价",),
    }
    detail_aliases = {
        "item_code": ("企业编码", "企业定额编码"),
        "item_name": ("定额子目名称",),
        "unit": ("单位",),
        "quantity": ("工程量",),
        "comprehensive_price": ("子目单价(元)", "子目单价"),
    }
    main_result = None
    detail_result = None
    for sheet_name, rows in sheets:
        if main_result is None:
            header, mapping = _find_header(rows, ("item_code", "item_name", "feature", "unit", "comprehensive_price"), main_aliases)
            if header is not None:
                main_result = (sheet_name, rows, header, mapping)
        if detail_result is None:
            header, mapping = _find_header(rows, ("item_code", "item_name", "unit", "quantity", "comprehensive_price"), detail_aliases)
            if header is not None:
                detail_result = (sheet_name, rows, header, mapping)
    if main_result is None or detail_result is None or main_result[0] == detail_result[0]:
        return None
    main_sheet, main_rows, main_header, main_mapping = main_result
    detail_sheet, detail_rows, detail_header, detail_mapping = detail_result
    for field, aliases in {
        "seq_no": ("序号",),
        "labor_cost": ("人工费",),
        "material_cost": ("材料费",),
        "machinery_cost": ("机械费",),
        "management_cost": ("管理费",),
        "profit": ("利润",),
        "regulatory_fee": ("人工增加费（规费）", "人工增加费"),
        "tax": ("税金",),
        "measure_fee": ("措施费",),
        "note": ("使用说明",),
    }.items():
        found = _find_mapping(main_rows[main_header], {field: aliases})
        if field in found:
            main_mapping[field] = found[field]
    for field, aliases in {
        "seq_no": ("序号",),
        "reference_code": ("参考政府定额编号",),
        "note": ("备注",),
    }.items():
        found = _find_mapping(detail_rows[detail_header], {field: aliases})
        if field in found:
            detail_mapping[field] = found[field]
    return main_sheet, main_rows, main_header, main_mapping, detail_sheet, detail_rows, detail_header, detail_mapping


def _upsert(session, values):
    query = session.query(ProjectListData).filter(
        ProjectListData.source_file == values["source_file"],
        ProjectListData.source_sheet == values["source_sheet"],
    )
    # Sequence numbers distinguish repeated names in a real BOQ.
    if _text(values.get("seq_no")):
        query = query.filter(
            ProjectListData.seq_no == values["seq_no"],
            ProjectListData.item_name == values["item_name"],
            ProjectListData.unit == values["unit"],
            ProjectListData.feature == values.get("feature", ""),
            ProjectListData.quantity == values.get("quantity"),
        )
    else:
        query = query.filter(
            ProjectListData.item_code == values["item_code"],
            ProjectListData.item_name == values["item_name"],
        )
    record = query.first()
    if record is None:
        session.add(ProjectListData(**values))
        return "inserted"
    for field, value in values.items():
        if field not in {"source_file", "source_sheet"}:
            setattr(record, field, value)
    return "updated"


def _enterprise_analysis(components, total, labor, material, machinery, management, profit, extras):
    lines = [
        "企业参考定额表组成明细（按每1个清单计量单位）",
        f"全费用综合单价：{_text(total)}；人工费：{_text(labor)}；材料费：{_text(material)}；机械费：{_text(machinery)}；管理费：{_text(management)}；利润：{_text(profit)}",
    ]
    other = "；".join(f"{key}：{_text(value)}" for key, value in extras.items() if value not in (None, ""))
    if other:
        lines.append("其他全费用项目：" + other)
    lines.append("关联定额子目：")
    for number, component in enumerate(components, 1):
        line = f"{number}. {component['reference_code']} | {component['name']} | {component['unit']} | 单位含量 {component['quantity']} | 子目单价 {_text(component['price'])}"
        if component["note"]:
            line += " | " + component["note"]
        lines.append(line)
    return "\n".join(lines)


def _import_enterprise_reference(sheets, session, file_path, *, source_project, region, period, detected):
    main_sheet, main_rows, main_header, main_mapping, detail_sheet, detail_rows, detail_header, detail_mapping = detected
    inferred_region, inferred_period = _infer_region_period(sheets, file_path.stem)
    region = _text(region) or inferred_region
    period = _text(period) or inferred_period
    source_project = _text(source_project) or _infer_source_project(sheets, file_path.stem)
    component_groups_by_code = {}
    current_code = ""
    current_group = None
    component_count = 0
    for row in detail_rows[detail_header + 1:]:
        code_in_row = _text(_cell(row, detail_mapping, "item_code"))
        code = code_in_row or current_code
        if code_in_row:
            current_code = code
            current_group = []
            component_groups_by_code.setdefault(code, []).append(current_group)
        component = {
            "reference_code": _text(_cell(row, detail_mapping, "reference_code")),
            "name": _text(_cell(row, detail_mapping, "item_name")),
            "unit": _text(_cell(row, detail_mapping, "unit")),
            "quantity": _number(_cell(row, detail_mapping, "quantity")),
            "price": _number(_cell(row, detail_mapping, "comprehensive_price")),
            "note": _text(_cell(row, detail_mapping, "note")),
            "raw": list(row),
        }
        if not code or not component["name"] or not component["unit"] or component["quantity"] is None:
            continue
        if current_group is None:
            continue
        current_group.append(component)
        component_count += 1
    inserted = updated = skipped = 0
    main_occurrence = {}
    for row in main_rows[main_header + 1:]:
        code = _text(_cell(row, main_mapping, "item_code"))
        name = _text(_cell(row, main_mapping, "item_name"))
        unit = _text(_cell(row, main_mapping, "unit"))
        if not code or not name or not unit:
            skipped += 1
            continue
        total = _number(_cell(row, main_mapping, "comprehensive_price"))
        labor = _number(_cell(row, main_mapping, "labor_cost"))
        material = _number(_cell(row, main_mapping, "material_cost"))
        machinery = _number(_cell(row, main_mapping, "machinery_cost"))
        management = _number(_cell(row, main_mapping, "management_cost"))
        profit = _number(_cell(row, main_mapping, "profit"))
        unit_prices, pricing_meta = _normalize_unit_prices(
            quantity=1,
            comprehensive_price=total,
            total_price=None,
            costs={
                "labor_cost": labor,
                "material_cost": material,
                "machinery_cost": machinery,
                "management_cost": management,
                "profit": profit,
            },
        )
        total = unit_prices["comprehensive_price"]
        labor = unit_prices["labor_cost"]
        material = unit_prices["material_cost"]
        machinery = unit_prices["machinery_cost"]
        management = unit_prices["management_cost"]
        profit = unit_prices["profit"]
        extras = {
            "人工增加费（规费）": _cell(row, main_mapping, "regulatory_fee"),
            "税金": _cell(row, main_mapping, "tax"),
            "措施费": _cell(row, main_mapping, "measure_fee"),
        }
        occurrence = main_occurrence.get(code, 0)
        main_occurrence[code] = occurrence + 1
        component_groups = component_groups_by_code.get(code, [])
        components = component_groups[occurrence] if occurrence < len(component_groups) else []
        note = "；".join(value for value in (
            "企业参考定额表，不是正式政府定额",
            "全费用综合单价口径，价格基准期 " + (period or "未识别"),
            "地区 " + (region or "未识别"),
            pricing_meta["note"],
            _text(_cell(row, main_mapping, "note")),
        ) if value)
        raw = {
            "format": "enterprise_reference_quota",
            "source_title": ENTERPRISE_REFERENCE_TITLE,
            "main_sheet": main_sheet,
            "detail_sheet": detail_sheet,
            "main_row": list(row),
            "component_rows": [component["raw"] for component in components],
            "price_basis": {
                "region": region,
                "period": period,
                "scope": "全费用综合单价（含人工、材料、机械、管理费、利润、规费、税金、措施费）",
                **pricing_meta,
            },
        }
        values = {
            "source_file": str(file_path), "source_sheet": main_sheet,
            "source_project": source_project, "region": region, "period": period,
            "seq_no": _text(_cell(row, main_mapping, "seq_no")), "item_code": code,
            "item_name": name, "feature": _combined_text(row, main_mapping, "feature", "work_content"),
            "unit": unit, "quantity": 1.0,
            "analysis": _enterprise_analysis(components, total, labor, material, machinery, management, profit, extras),
            "comprehensive_price": total, "labor_cost": labor, "material_cost": material,
            "machinery_cost": machinery, "management_cost": management, "profit": profit,
            "note": note, "raw_data": json.dumps(raw, ensure_ascii=False, default=str),
        }
        action = _upsert(session, values)
        if action == "inserted":
            inserted += 1
        else:
            updated += 1
    return {"success": True, "format": ENTERPRISE_REFERENCE_TITLE, "imported": inserted, "updated": updated, "skipped": skipped, "component_count": component_count, "region": region, "period": period, "main_sheet": main_sheet, "detail_sheet": detail_sheet}


def import_project_list_data(path: str, session: Session, *, source_project: str = "", region: str = "", period: str = "") -> dict:
    """Import ordinary BOQs or enterprise reference quota workbooks idempotently."""
    file_path = Path(path)
    if not file_path.exists():
        return {"success": False, "error": "文件不存在"}
    try:
        # The specialized Glodon reader is for the modern xlsx export. Older
        # workbooks go through the generic sheet/header detector below.
        specialized = None
        if file_path.suffix.lower() == ".xlsx":
            try:
                specialized = extract_specialized_bill_list(str(file_path))
            except Exception:
                specialized = None
        if specialized is not None:
            inserted = updated = 0
            inferred_region, inferred_period = _infer_region_period([], file_path.stem)
            source_name = _text(source_project) or _infer_source_project([], file_path.stem)
            source_region = _text(region) or inferred_region
            source_period = _text(period) or inferred_period
            for item in specialized.get("items") or []:
                costs = list(item.get("costs") or ()) + [None] * 5
                analysis = _text(item.get("analysis"))
                quantity = _number(item.get("quantity"))
                comprehensive_price = _number(item.get("comprehensive_price"))
                total_price = _number(item.get("total_price"))
                unit_prices, pricing_meta = _normalize_unit_prices(
                    quantity=quantity,
                    comprehensive_price=comprehensive_price,
                    total_price=total_price,
                    costs={
                        "labor_cost": costs[0],
                        "material_cost": costs[1],
                        "machinery_cost": costs[2],
                        "management_cost": costs[3],
                        "profit": costs[4],
                    },
                )
                raw_item = dict(item)
                raw_item["unit_pricing"] = pricing_meta
                values = {
                    "source_file": str(file_path), "source_sheet": item.get("source_tab", ""),
                    "source_project": source_name, "region": source_region, "period": source_period,
                    "seq_no": _text(item.get("seq")), "item_code": _text(item.get("code")),
                    "item_name": _text(item.get("name")), "feature": _text(item.get("feature")),
                    "unit": _text(item.get("unit")), "quantity": quantity,
                    "analysis": analysis,
                    "comprehensive_price": unit_prices["comprehensive_price"],
                    "total_price": total_price,
                    "labor_cost": unit_prices["labor_cost"], "material_cost": unit_prices["material_cost"],
                    "machinery_cost": unit_prices["machinery_cost"], "management_cost": unit_prices["management_cost"],
                    "profit": unit_prices["profit"],
                    "note": "；".join(value for value in (
                        "广联达报价表导入；按每1个清单单位保存单位价格",
                        pricing_meta["note"],
                    ) if value),
                    "raw_data": json.dumps(raw_item, ensure_ascii=False, default=str),
                }
                if not values["item_name"] or not values["unit"] or values["quantity"] is None:
                    continue
                action = _upsert(session, values)
                if action == "inserted":
                    inserted += 1
                else:
                    updated += 1
            session.commit()
            return {
                "success": True,
                "format": specialized.get("format", "广联达报价表"),
                "imported": inserted,
                "updated": updated,
                "skipped": 0,
                "component_count": 0,
                "region": source_region,
                "period": source_period,
            }
        sheets = _read_sheets(file_path)
        detected = _enterprise_format(sheets)
        if detected:
            result = _import_enterprise_reference(
                sheets, session, file_path,
                source_project=source_project, region=region, period=period,
                detected=detected,
            )
            session.commit()
            return result
        inferred_region, inferred_period = _infer_region_period(sheets, file_path.stem)
        source_region = _text(region) or inferred_region
        source_period = _text(period) or inferred_period
        source_name = _text(source_project) or _infer_source_project(sheets, file_path.stem)
        inserted = skipped = 0
        updated = 0
        for sheet_name, rows in sheets:
            if not rows:
                continue
            header_row, mapping, quantity_defaulted = _find_standard_header(rows)
            if header_row is None:
                continue
            # Skip the complete multi-level header block. Some tender
            # workbooks add a formula/legend row immediately below it; rows
            # without an item name/unit/quantity are rejected naturally.
            data_start = header_row + int(mapping.get("_header_depth", 1))
            for row in rows[data_start:]:
                name = _text(_cell(row, mapping, "item_name"))
                unit = _text(_cell(row, mapping, "unit"))
                feature = _combined_text(row, mapping, "feature", "work_content")
                quantity = _number(_cell(row, mapping, "quantity"))
                if quantity is None and quantity_defaulted:
                    quantity = 1.0
                # A reference row must carry a real quantity and a price (or
                # one of the explicit price components). This excludes titles,
                # section totals and empty template rows.
                if (
                    not name
                    or not unit
                    or quantity is None
                    or quantity <= 0
                    or not _has_reference_price(row, mapping)
                    or _is_aggregate_reference_row(name, unit, feature, mapping, row)
                ):
                    skipped += 1
                    continue
                main_material = _number(_cell(row, mapping, "main_material_cost"))
                auxiliary_material = _number(_cell(row, mapping, "auxiliary_material_cost"))
                if main_material is not None or auxiliary_material is not None:
                    material_cost = (main_material or 0.0) + (auxiliary_material or 0.0)
                else:
                    material_cost = _number(_cell(row, mapping, "material_cost"))
                comprehensive_price = _number(_cell(row, mapping, "comprehensive_price"))
                total_price = _number(_cell(row, mapping, "total_price"))
                unit_prices, pricing_meta = _normalize_unit_prices(
                    quantity=quantity,
                    comprehensive_price=comprehensive_price,
                    total_price=total_price,
                    costs={
                        "labor_cost": _cell(row, mapping, "labor_cost"),
                        "material_cost": material_cost,
                        "machinery_cost": _cell(row, mapping, "machinery_cost"),
                        "management_cost": _cell(row, mapping, "management_cost"),
                        "profit": _cell(row, mapping, "profit"),
                    },
                )
                values = dict(
                    source_file=str(file_path), source_sheet=sheet_name,
                    source_project=source_name,
                    region=source_region, period=source_period,
                    seq_no=_text(_cell(row, mapping, "seq_no")),
                    item_code=_text(_cell(row, mapping, "item_code")),
                    item_name=name, feature=feature, unit=unit,
                    quantity=quantity, analysis="；".join(
                        value for value in (
                            _text(_cell(row, mapping, "analysis")),
                            _reference_analysis(row, mapping),
                        ) if value
                    ),
                    comprehensive_price=unit_prices["comprehensive_price"],
                    total_price=total_price,
                    labor_cost=unit_prices["labor_cost"],
                    material_cost=unit_prices["material_cost"],
                    machinery_cost=unit_prices["machinery_cost"],
                    management_cost=unit_prices["management_cost"],
                    profit=unit_prices["profit"],
                    note="；".join(value for value in (
                        _text(_cell(row, mapping, "note")),
                        "按每1个清单单位保存单位价格",
                        pricing_meta["note"],
                    ) if value),
                    raw_data=json.dumps({
                        "row": list(row),
                        "quantity_defaulted": quantity_defaulted,
                        "unit_pricing": pricing_meta,
                    }, ensure_ascii=False, default=str),
                )
                action = _upsert(session, values)
                if action == "inserted":
                    inserted += 1
                else:
                    updated += 1
        session.commit()
        detected_formats = []
        for _sheet_name, _rows in sheets:
            _header, _mapping, _defaulted = _find_standard_header(_rows)
            if _header is not None and _mapping.get("_format"):
                detected_formats.append(_mapping["_format"])
        format_name = (
            "项目清单参考价（含组价明细）"
            if "项目清单参考价（含组价明细）" in detected_formats
            else "普通项目清单"
        )
        return {
            "success": True,
            "format": format_name,
            "imported": inserted,
            "updated": updated,
            "skipped": skipped,
            "valid": inserted + updated,
            "component_count": 0,
            "region": source_region,
            "period": source_period,
        }
    except Exception as error:
        session.rollback()
        return {"success": False, "error": str(error)}


def search_project_list_data(
    session: Session,
    keyword: str = "",
    limit: int | None = None,
    *,
    source_project: str = "",
    region: str = "",
    period: str = "",
):
    keyword = _text(keyword)
    source_project = _text(source_project)
    region = _text(region)
    period = _text(period)
    query = session.query(ProjectListData)
    if source_project:
        query = query.filter(ProjectListData.source_project == source_project)
    if region:
        query = query.filter(ProjectListData.region == region)
    if period:
        query = query.filter(ProjectListData.period == period)
    if keyword:
        pattern = f"%{keyword}%"
        query = query.filter(or_(
            ProjectListData.item_code.ilike(pattern),
            ProjectListData.item_name.ilike(pattern),
            ProjectListData.feature.ilike(pattern),
            ProjectListData.unit.ilike(pattern),
            ProjectListData.analysis.ilike(pattern),
            ProjectListData.source_project.ilike(pattern),
            ProjectListData.source_sheet.ilike(pattern),
            ProjectListData.region.ilike(pattern),
            ProjectListData.period.ilike(pattern),
        ))
    query = query.order_by(ProjectListData.updated_at.desc(), ProjectListData.id.desc())
    if limit is not None:
        query = query.limit(limit)
    return query.all()


def find_project_list_references(
    session: Session,
    name: str,
    feature: str,
    unit: str,
    *,
    region: str = "",
    period: str = "",
    limit: int = 12,
):
    target = f"{_text(name)} {_text(feature)}"
    candidates = []
    for row in session.query(ProjectListData).filter(ProjectListData.item_name != "").limit(10000).all():
        if unit and row.unit and _text(unit) != _text(row.unit):
            continue
        name_score = max(fuzz.ratio(_text(name), row.item_name), fuzz.WRatio(_text(name), row.item_name))
        context_score = fuzz.WRatio(target, f"{row.item_name} {row.feature}")
        score = name_score * 0.65 + context_score * 0.35
        if region and row.region and region in row.region:
            score += 8
        if period and row.period and period == row.period:
            score += 8
        if score >= 55:
            candidates.append((score, row, name_score, context_score))
    candidates.sort(key=lambda value: value[0], reverse=True)
    return candidates[:limit]
