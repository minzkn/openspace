-- Migration 003: sheet metadata (merges, row_heights, freeze_panes)
-- 멱등성 보장 위해 ALTER TABLE은 이미 컬럼 존재 시 무시 (SQLite는 에러 발생 가능)
-- init_db.py 또는 start.sh에서 수동 실행 필요

ALTER TABLE template_sheets  ADD COLUMN merges      TEXT;
ALTER TABLE template_sheets  ADD COLUMN row_heights  TEXT;
ALTER TABLE template_sheets  ADD COLUMN freeze_panes TEXT;

ALTER TABLE workspace_sheets ADD COLUMN merges      TEXT;
ALTER TABLE workspace_sheets ADD COLUMN row_heights  TEXT;
ALTER TABLE workspace_sheets ADD COLUMN freeze_panes TEXT;
