-- Migration 009: Add hyperlink column to cells
ALTER TABLE template_cells ADD COLUMN hyperlink TEXT;
ALTER TABLE workspace_cells ADD COLUMN hyperlink TEXT;
