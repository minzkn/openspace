-- Migration 002: workspaces.template_id, workspace_sheets.template_sheet_id
--   를 nullable + ON DELETE SET NULL 로 변경
-- (서식 삭제 시 연결된 워크스페이스/시트가 고아(orphan)가 되도록 허용)

PRAGMA foreign_keys=OFF;
BEGIN;

-- ── workspaces 테이블 재생성 ──────────────────────────────────────
CREATE TABLE workspaces_new (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    template_id TEXT REFERENCES templates(id) ON DELETE SET NULL,  -- nullable
    status      TEXT NOT NULL DEFAULT 'OPEN'
                CHECK (status IN ('OPEN','CLOSED')),
    created_by  TEXT NOT NULL REFERENCES users(id),
    closed_by   TEXT REFERENCES users(id),
    closed_at   TEXT,
    created_at  TEXT NOT NULL,
    updated_at  TEXT NOT NULL
);

INSERT INTO workspaces_new
    SELECT id, name, template_id, status, created_by, closed_by, closed_at, created_at, updated_at
    FROM workspaces;

DROP TABLE workspaces;
ALTER TABLE workspaces_new RENAME TO workspaces;

-- ── workspace_sheets 테이블 재생성 ───────────────────────────────
CREATE TABLE workspace_sheets_new (
    id                  TEXT PRIMARY KEY,
    workspace_id        TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
    template_sheet_id   TEXT REFERENCES template_sheets(id) ON DELETE SET NULL,  -- nullable
    sheet_index         INTEGER NOT NULL,
    sheet_name          TEXT NOT NULL,
    UNIQUE (workspace_id, sheet_index)
);

INSERT INTO workspace_sheets_new
    SELECT id, workspace_id, template_sheet_id, sheet_index, sheet_name
    FROM workspace_sheets;

DROP TABLE workspace_sheets;
ALTER TABLE workspace_sheets_new RENAME TO workspace_sheets;

-- workspace_cells 은 workspace_sheets(id) 참조, 변화 없음
-- (DROP 후 RENAME 때 기존 인덱스 삭제되므로 재생성)
CREATE INDEX IF NOT EXISTS idx_ws_cells_sheet ON workspace_cells(sheet_id);

COMMIT;
PRAGMA foreign_keys=ON;
