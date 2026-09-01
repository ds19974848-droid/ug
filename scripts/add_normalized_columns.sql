-- Migration: add normalized_name columns and indexes
--
-- This script adds a normalized_name column to the main lookup tables and
-- creates indexes to accelerate normalized-text contains/startswith queries.
--
-- WARNING: ALTER TABLE ADD COLUMN in some RDBMS will fail if the column
-- already exists. Verify in your environment or run in a maintenance window
-- with a DB backup available.
--
-- SQLite example (run in maintenance window after DB backup):
--   sqlite3 /path/to/your.db < scripts/add_normalized_columns.sql
--
BEGIN TRANSACTION;

-- cost_items
ALTER TABLE cost_items ADD COLUMN normalized_name VARCHAR(500) DEFAULT '';
CREATE INDEX IF NOT EXISTS idx_cost_item_normalized_name ON cost_items (normalized_name);

-- quota_items
ALTER TABLE quota_items ADD COLUMN normalized_name VARCHAR(500) DEFAULT '';
CREATE INDEX IF NOT EXISTS idx_quota_item_normalized_name ON quota_items (normalized_name);

-- materials
ALTER TABLE materials ADD COLUMN normalized_name VARCHAR(300) DEFAULT '';
CREATE INDEX IF NOT EXISTS idx_materials_normalized_name ON materials (normalized_name);

COMMIT;

-- Notes for other DB engines:
-- PostgreSQL: ALTER TABLE <table> ADD COLUMN IF NOT EXISTS normalized_name VARCHAR(500) DEFAULT '';
-- MySQL: ALTER TABLE <table> ADD COLUMN IF NOT EXISTS normalized_name VARCHAR(500) DEFAULT '';
-- After running the migration, run the backfill script to populate the new columns:
--   python scripts/backfill_normalized.py --db sqlite:///path/to/your.db
--
-- Always test the migration and backfill in a staging environment first and
-- ensure you have a current backup before applying to production.
