PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS users (
    id          TEXT PRIMARY KEY,
    username    TEXT UNIQUE NOT NULL,
    email       TEXT,
    password_hash TEXT NOT NULL,
    role        TEXT NOT NULL DEFAULT 'USER'
                CHECK (role IN ('SUPER_ADMIN','ADMIN','USER')),
    is_active   INTEGER NOT NULL DEFAULT 1,
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);

CREATE TABLE IF NOT EXISTS user_extra_fields (
    id          TEXT PRIMARY KEY,
    field_name  TEXT UNIQUE NOT NULL,
    label       TEXT NOT NULL,
    field_type  TEXT NOT NULL DEFAULT 'text'
                CHECK (field_type IN ('text','number','date','boolean')),
    is_sensitive INTEGER NOT NULL DEFAULT 0,
    is_required INTEGER NOT NULL DEFAULT 0,
    sort_order  INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);

CREATE TABLE IF NOT EXISTS user_extra_values (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    field_id    TEXT NOT NULL REFERENCES user_extra_fields(id) ON DELETE CASCADE,
    value       TEXT,
    UNIQUE (user_id, field_id)
);

CREATE TABLE IF NOT EXISTS templates (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    description TEXT,
    created_by  TEXT NOT NULL REFERENCES users(id),
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);

CREATE TABLE IF NOT EXISTS template_sheets (
    id           TEXT PRIMARY KEY,
    template_id  TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
    sheet_index  INTEGER NOT NULL,
    sheet_name   TEXT NOT NULL,
    UNIQUE (template_id, sheet_index)
);

CREATE TABLE IF NOT EXISTS template_columns (
    id           TEXT PRIMARY KEY,
    sheet_id     TEXT NOT NULL REFERENCES template_sheets(id) ON DELETE CASCADE,
    col_index    INTEGER NOT NULL,
    col_header   TEXT NOT NULL,
    col_type     TEXT NOT NULL DEFAULT 'text'
                 CHECK (col_type IN ('text','number','date','dropdown','checkbox')),
    col_options  TEXT,
    is_readonly  INTEGER NOT NULL DEFAULT 0,
    width        INTEGER DEFAULT 120,
    UNIQUE (sheet_id, col_index)
);

CREATE TABLE IF NOT EXISTS template_cells (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES template_sheets(id) ON DELETE CASCADE,
    row_index   INTEGER NOT NULL CHECK (row_index >= 0 AND row_index < 10000),
    col_index   INTEGER NOT NULL,
    value       TEXT,
    formula     TEXT,
    style       TEXT,
    UNIQUE (sheet_id, row_index, col_index)
);

CREATE TABLE IF NOT EXISTS workspaces (
    id           TEXT PRIMARY KEY,
    name         TEXT NOT NULL,
    template_id  TEXT NOT NULL REFERENCES templates(id),
    status       TEXT NOT NULL DEFAULT 'OPEN'
                 CHECK (status IN ('OPEN','CLOSED')),
    created_by   TEXT NOT NULL REFERENCES users(id),
    closed_by    TEXT REFERENCES users(id),
    closed_at    TEXT,
    created_at   TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    updated_at   TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);

CREATE TABLE IF NOT EXISTS workspace_sheets (
    id              TEXT PRIMARY KEY,
    workspace_id    TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
    template_sheet_id TEXT NOT NULL REFERENCES template_sheets(id),
    sheet_index     INTEGER NOT NULL,
    sheet_name      TEXT NOT NULL,
    UNIQUE (workspace_id, sheet_index)
);

CREATE TABLE IF NOT EXISTS workspace_cells (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES workspace_sheets(id) ON DELETE CASCADE,
    row_index   INTEGER NOT NULL,
    col_index   INTEGER NOT NULL,
    value       TEXT,
    style       TEXT,
    updated_by  TEXT REFERENCES users(id),
    updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    UNIQUE (sheet_id, row_index, col_index)
);

CREATE INDEX IF NOT EXISTS idx_ws_cells_sheet ON workspace_cells(sheet_id);

CREATE TABLE IF NOT EXISTS change_logs (
    id          TEXT PRIMARY KEY,
    workspace_id TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
    sheet_id    TEXT NOT NULL,
    user_id     TEXT NOT NULL REFERENCES users(id),
    row_index   INTEGER NOT NULL,
    col_index   INTEGER NOT NULL,
    old_value   TEXT,
    new_value   TEXT,
    changed_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);

CREATE INDEX IF NOT EXISTS idx_cl_workspace ON change_logs(workspace_id);

CREATE TABLE IF NOT EXISTS sessions (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    expires_at  TEXT NOT NULL,
    last_seen   TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    ip_address  TEXT,
    user_agent  TEXT
);

CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id);
CREATE INDEX IF NOT EXISTS idx_sessions_expires ON sessions(expires_at);
