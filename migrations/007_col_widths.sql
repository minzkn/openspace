-- Migration 007: add column widths metadata for template/workspace sheets
-- 멱등성: init_db.py가 duplicate column 예외를 무시

ALTER TABLE template_sheets  ADD COLUMN col_widths TEXT;
ALTER TABLE workspace_sheets ADD COLUMN col_widths TEXT;
