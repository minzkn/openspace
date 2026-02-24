-- Migration 006: Add conditional_formats column to sheet tables
-- Enables conditional formatting rules

ALTER TABLE template_sheets ADD COLUMN conditional_formats TEXT;
ALTER TABLE workspace_sheets ADD COLUMN conditional_formats TEXT;
