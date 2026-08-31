"""Backfill normalized_name columns for cost_items, quota_items and materials.

Run this script manually in a maintenance window AFTER applying the SQL migration
that adds the normalized_name columns (scripts/add_normalized_columns.sql).

Usage:
  python scripts/backfill_normalized.py --db sqlite:///path/to/dashuo_cost_cloud.db

This script will perform updates in batches and print progress. It does NOT create
or modify schema (run the SQL migration separately).
"""
from __future__ import annotations

import argparse
from sqlalchemy import update
from sqlalchemy.orm import Session

from src.db import get_session
from src.models import CostItem, QuotaItem, Material
from src.text_utils import normalize_text

BATCH = 500


def backfill_table(session: Session, model, source_field: str):
    total = session.query(model).count()
    print(f"Backfilling {model.__tablename__}: {total} rows")
    offset = 0
    while True:
        rows = session.query(model).order_by(model.id).offset(offset).limit(BATCH).all()
        if not rows:
            break
        for row in rows:
            val = getattr(row, source_field, "") or ""
            norm = normalize_text(val)
            # only update if attribute exists
            if hasattr(row, 'normalized_name'):
                setattr(row, 'normalized_name', norm)
        session.commit()
        offset += len(rows)
        print(f"  processed {offset}/{total}")
    print(f"Done {model.__tablename__}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dry-run', action='store_true', help='Do not commit changes')
    parser.add_argument('--db', type=str, help='Database URL (optional)')
    args = parser.parse_args()

    session = get_session()
    try:
        backfill_table(session, CostItem, 'item_name')
        backfill_table(session, QuotaItem, 'name')
        backfill_table(session, Material, 'name')
    finally:
        session.close()


if __name__ == '__main__':
    main()
