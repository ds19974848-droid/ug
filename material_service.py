"""Material identity and price metadata helpers shared by imports and crawlers."""

from __future__ import annotations

import re
import unicodedata

from sqlalchemy.orm import Session

from .models import Material, MaterialAlias


def clean_material_text(value: object) -> str:
    """Normalize presentation text without changing meaningful punctuation."""
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = text.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalize_material_identity(value: object) -> str:
    """Return a stable identity key for names such as 'C 30' and 'C30'."""
    return re.sub(r"\s+", "", clean_material_text(value)).casefold()


def build_material_identity_cache(session: Session) -> dict[str, Material]:
    cache: dict[str, Material] = {}
    materials = session.query(Material).all()
    for material in materials:
        cache[normalize_material_identity(material.name)] = material
        for alias in material.aliases:
            cache[normalize_material_identity(alias.alias_name)] = material
    return cache


def find_or_create_material(
    session: Session,
    name: object,
    *,
    unit: object = "",
    spec: object = "",
    category: object = "",
    cache: dict[str, Material] | None = None,
) -> tuple[Material, bool]:
    """Resolve a material by canonical name or alias and record new aliases."""
    canonical_name = clean_material_text(name)
    if not canonical_name:
        raise ValueError("材料名称不能为空")
    identity = normalize_material_identity(canonical_name)
    material = cache.get(identity) if cache is not None else None
    if material is None:
        material = session.query(Material).filter(Material.name == canonical_name).first()
    if material is None:
        material = session.query(MaterialAlias).filter(
            MaterialAlias.alias_name == canonical_name
        ).first()
        material = material.material if material else None
    if material is None and cache is None:
        material = next(
            (
                candidate for candidate in session.query(Material).all()
                if normalize_material_identity(candidate.name) == identity
            ),
            None,
        )
    created = material is None
    if material is None:
        material = Material(
            name=canonical_name,
            category=clean_material_text(category),
            default_unit=clean_material_text(unit),
            spec_template=clean_material_text(spec),
        )
        session.add(material)
        session.flush()
    else:
        if category and not material.category:
            material.category = clean_material_text(category)
        if unit and not material.default_unit:
            material.default_unit = clean_material_text(unit)
        if spec and not material.spec_template:
            material.spec_template = clean_material_text(spec)
        if canonical_name != material.name:
            alias_exists = session.query(MaterialAlias).filter(
                MaterialAlias.material_id == material.id,
                MaterialAlias.alias_name == canonical_name,
            ).first()
            if alias_exists is None:
                session.add(MaterialAlias(
                    material_id=material.id,
                    alias_name=canonical_name,
                    source="自动标准化",
                ))
    if cache is not None:
        cache[identity] = material
        cache[normalize_material_identity(material.name)] = material
    return material, created


def infer_price_basis(*values: object) -> str:
    text = " ".join(clean_material_text(value) for value in values if value)
    if any(token in text for token in ("含税", "税内")):
        return "tax_inclusive"
    if any(token in text for token in ("除税", "不含税", "税前")):
        return "tax_exclusive"
    if any(token in text for token in ("到场", "运至工地", "工地价")):
        return "delivered"
    if "出厂" in text:
        return "ex_factory"
    return "as_published"
