-- Outline groups for template/workspace sheets
ALTER TABLE template_sheets ADD COLUMN outline_rows TEXT;
ALTER TABLE template_sheets ADD COLUMN outline_cols TEXT;
ALTER TABLE workspace_sheets ADD COLUMN outline_rows TEXT;
ALTER TABLE workspace_sheets ADD COLUMN outline_cols TEXT;
