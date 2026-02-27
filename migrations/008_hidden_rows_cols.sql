-- Migration 008: Add hidden_rows and hidden_cols to sheets
ALTER TABLE template_sheets ADD COLUMN hidden_rows TEXT;
ALTER TABLE template_sheets ADD COLUMN hidden_cols TEXT;
ALTER TABLE workspace_sheets ADD COLUMN hidden_rows TEXT;
ALTER TABLE workspace_sheets ADD COLUMN hidden_cols TEXT;
