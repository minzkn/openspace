-- Data validations for template/workspace sheets
ALTER TABLE template_sheets ADD COLUMN data_validations TEXT;
ALTER TABLE workspace_sheets ADD COLUMN data_validations TEXT;
