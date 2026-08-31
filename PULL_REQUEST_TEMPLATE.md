# PR: Improve matching, normalization, and add normalized_name migration

This pull request implements a series of improvements to make matching of cost
items, quota items and materials more robust and accurate. It also prepares the
schema and tooling for normalized text indexing and later embedding-based
semantic matching.

## Summary of changes

- Add `src/text_utils.py`: `normalize_text` and `score_candidate` helpers.
- Update `material_service.py` to use `normalize_text` for identity/alias handling.
- Improve `cost_item_service.find_cost_item_reference` to:
  - Normalize inputs and candidate names
  - Multi-stage matching: exact -> contains -> fuzzy
  - Prefer same-unit candidates and soft-handle unit mismatches
  - Expose configurable thresholds (min_score, contains_limit, candidate_limit)
- Make `utils.import_prices_from_excel` header lookup tolerant using `normalize_text`.
- Add migration SQL `scripts/add_normalized_columns.sql` to add `normalized_name` columns and indexes.
- Add backfill script `scripts/backfill_normalized.py` to populate `normalized_name` for existing rows.
- Add unit tests `tests/test_matching.py` for core matching behaviors.

## Migration notes

1. Apply `scripts/add_normalized_columns.sql` to add `normalized_name` columns and indexes.
2. Run `scripts/backfill_normalized.py` in a maintenance window to populate values.
3. Optionally create indexes afterwards or concurrently (script already creates indexes).

## How to run tests

- pip install -r requirements.txt
- pip install pytest rapidfuzz
- pytest tests/test_matching.py -q

## Rollback

All code changes are isolated on branch `fix/match-improvements`. To rollback DB schema changes,
restore from backup taken before running the migration SQL.

