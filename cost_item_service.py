"""固定工程清单库的导入、导出与价格匹配。"""
from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from rapidfuzz import fuzz, process
from sqlalchemy import or_
from sqlalchemy.orm import Session

from .db import write_audit
from .models import CostItem


FIELD_MAP = {
    "数据ID": "source_key",
    "数据来源": "data_source",
    "原始序号": "original_seq",
    "项目名称": "item_name",
    "项目特征": "features",
    "单位": "unit",
    "综合单价": "comprehensive_price",
    "含税价": "tax_inclusive_price",
    "人工费": "labor_cost",
    "材料费": "material_cost",
    "机械费": "machinery_cost",
    "管理费": "management_cost",
    "利润": "profit",
    "完整分类": "full_category",
    "原始类别": "original_category",
    "价格分析": "price_analysis",
    "AI分析": "ai_analysis",
    "备注": "notes",
}
NUMERIC_FIELDS = {
    "original_seq", "comprehensive_price", "tax_inclusive_price", "labor_cost",
    "material_cost", "machinery_cost", "management_cost", "profit",
}


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _number(value, default=0.0):
    if value in (None, ""):
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def import_cost_items_from_excel(
    filepath: str | Path,
    session: Session,
    *,
    sheet_name: str = "全部数据",
    audit: bool = True,
) -> dict:
    """幂等导入固定清单；相同数据 ID 会更新而不会重复。"""
    from openpyxl import load_workbook
    workbook = load_workbook(filepath, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_name] if sheet_name in workbook.sheetnames else workbook.active
        rows = worksheet.iter_rows(values_only=True)
        headers = next(rows, None)
        if not headers:
            return {"success": False, "imported": 0, "updated": 0, "skipped": 0, "error": "Excel 没有数据"}
        header_indexes = {
            FIELD_MAP[_text(header)]: index
            for index, header in enumerate(headers)
            if _text(header) in FIELD_MAP
        }
        required = {"item_name", "comprehensive_price"}
        if not required.issubset(header_indexes):
            missing = "、".join(sorted(required - set(header_indexes)))
            return {"success": False, "imported": 0, "updated": 0, "skipped": 0, "error": f"缺少必要字段: {missing}"}

        existing = {
            item.source_key: item
            for item in session.query(CostItem).all()
        }
        imported = 0
        updated = 0
        skipped = 0
        for row_number, row in enumerate(rows, start=2):
            values = {}
            for field, index in header_indexes.items():
                value = row[index] if index < len(row) else None
                if field in NUMERIC_FIELDS:
                    values[field] = int(_number(value)) if field == "original_seq" else _number(value)
                else:
                    values[field] = _text(value)
            if not values.get("item_name"):
                skipped += 1
                continue
            source_key = values.get("source_key") or f"import-{uuid4().hex}"
            values["source_key"] = source_key
            item = existing.get(source_key)
            if item is None:
                item = CostItem(**values)
                session.add(item)
                existing[source_key] = item
                imported += 1
            else:
                for field, value in values.items():
                    setattr(item, field, value)
                updated += 1
            if (imported + updated) % 500 == 0:
                session.flush()
        session.flush()
        if audit:
            write_audit(
                session,
                "import",
                "cost_item",
                detail=f"固定清单导入: 新增 {imported} 条，更新 {updated} 条，跳过 {skipped} 条",
            )
        return {
            "success": True,
            "imported": imported,
            "updated": updated,
            "skipped": skipped,
            "total": imported + updated,
        }
    finally:
        workbook.close()


def export_cost_items_to_excel(items: list[CostItem], filepath: str | Path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "固定工程清单库"
    headers = list(FIELD_MAP)
    fields = [FIELD_MAP[header] for header in headers]
    worksheet.append(headers)
    fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    for cell in worksheet[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for item in items:
        worksheet.append([getattr(item, field) for field in fields])
    for column in worksheet.columns:
        letter = column[0].column_letter
        worksheet.column_dimensions[letter].width = 18
    worksheet.column_dimensions["E"].width = 42
    workbook.save(filepath)


def find_cost_item_reference(
    session: Session,
    item_name: str,
    unit: str = "",
    category: str = "",
) -> tuple[CostItem | None, float]:
    """优先精确/包含匹配，最后在候选集中做中文模糊匹配。"""
    keyword = _text(item_name)
    if not keyword:
        return None, 0.0
    query = session.query(CostItem)
    if category:
        query = query.filter(CostItem.full_category.contains(category))
    exact = query.filter(CostItem.item_name == keyword).order_by(CostItem.id).first()
    if exact and (not unit or exact.unit == unit):
        return exact, 1.0
    contains = query.filter(
        or_(CostItem.item_name.contains(keyword), CostItem.features.contains(keyword))
    ).limit(100).all()
    if unit:
        same_unit = [item for item in contains if item.unit == unit]
        if same_unit:
            contains = same_unit
    if contains:
        best = max(contains, key=lambda item: fuzz.WRatio(keyword, item.item_name))
        return best, fuzz.WRatio(keyword, best.item_name) / 100
    candidates = query.order_by(CostItem.id).limit(1500).all()
    if unit:
        candidates = [item for item in candidates if item.unit == unit] or candidates
    match = process.extractOne(keyword, {item.id: item.item_name for item in candidates}, scorer=fuzz.WRatio)
    if not match:
        return None, 0.0
    _, score, item_id = match
    if score < 45:
        return None, score / 100
    item = next((candidate for candidate in candidates if candidate.id == item_id), None)
    return item, score / 100
