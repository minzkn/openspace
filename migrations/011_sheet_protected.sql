-- Sheet protection toggle for workspace sheets
ALTER TABLE workspace_sheets ADD COLUMN sheet_protected INTEGER DEFAULT 0;
