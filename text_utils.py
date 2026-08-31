"""Text utilities shared across matching and normalization logic."""
from __future__ import annotations

import re
import unicodedata
from typing import Dict


def normalize_text(value: object) -> str:
    """Normalize text for comparison/search:
    - NFKC unicode normalization
    - replace non-breaking space
    - collapse whitespace
    - casefold for case-insensitive comparison
    """
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text.casefold()


def score_candidate(needle: Dict[str, str], candidate: Dict[str, str], weights: Dict[str, float] | None = None) -> float:
    """Weighted multi-field score for matching.
    needle and candidate are dicts with keys: name, spec, category, unit, code (optional)
    Returns a score 0.0-1.0
    """
    try:
        from rapidfuzz import fuzz
    except Exception:
        # fallback simple comparison if rapidfuzz not available
        def simple_ratio(a, b):
            a = normalize_text(a)
            b = normalize_text(b)
            if not a or not b:
                return 0.0
            return 1.0 if a == b else 0.0
        class _Faux:
            @staticmethod
            def WRatio(a, b):
                return simple_ratio(a, b) * 100
        fuzz = _Faux()

    if weights is None:
        weights = {"name": 0.6, "spec": 0.2, "category": 0.1, "unit": 0.08, "code": 0.02}

    name_score = (fuzz.WRatio(normalize_text(needle.get("name")), normalize_text(candidate.get("name"))) / 100.0) if needle.get("name") or candidate.get("name") else 0.0
    spec_score = (fuzz.WRatio(normalize_text(needle.get("spec")), normalize_text(candidate.get("spec"))) / 100.0) if needle.get("spec") or candidate.get("spec") else 0.0
    category_score = (fuzz.WRatio(normalize_text(needle.get("category")), normalize_text(candidate.get("category"))) / 100.0) if needle.get("category") or candidate.get("category") else 0.0

    unit_match = False
    if needle.get("unit") and candidate.get("unit"):
        unit_match = normalize_text(needle.get("unit")) == normalize_text(candidate.get("unit"))
    unit_score = 1.0 if unit_match else 0.0

    code_score = 0.0
    if needle.get("code") and candidate.get("code"):
        code_score = 1.0 if normalize_text(needle.get("code")) == normalize_text(candidate.get("code")) else 0.0

    total = (
        weights.get("name", 0.0) * name_score +
        weights.get("spec", 0.0) * spec_score +
        weights.get("category", 0.0) * category_score +
        weights.get("unit", 0.0) * unit_score +
        weights.get("code", 0.0) * code_score
    )

    # penalty for unit mismatch when name is very close
    if name_score > 0.92 and needle.get("unit") and not unit_match:
        total *= 0.92

    return float(total)
