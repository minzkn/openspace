-- Migration 005: Add comment column to cell tables
-- Enables cell notes/comments feature

ALTER TABLE workspace_cells ADD COLUMN comment TEXT;
ALTER TABLE template_cells ADD COLUMN comment TEXT;
