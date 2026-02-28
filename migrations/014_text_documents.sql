-- 014_text_documents.sql
-- Text document editor tables

CREATE TABLE IF NOT EXISTS text_documents (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    language TEXT NOT NULL DEFAULT 'plaintext',
    status TEXT NOT NULL DEFAULT 'OPEN',
    version INTEGER NOT NULL DEFAULT 1,
    created_by TEXT NOT NULL REFERENCES users(id),
    closed_by TEXT REFERENCES users(id),
    closed_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS text_document_content (
    id TEXT PRIMARY KEY,
    document_id TEXT NOT NULL REFERENCES text_documents(id) ON DELETE CASCADE,
    content TEXT NOT NULL DEFAULT '',
    version INTEGER NOT NULL DEFAULT 1,
    updated_by TEXT REFERENCES users(id),
    updated_at TEXT NOT NULL,
    UNIQUE(document_id, version)
);

CREATE TABLE IF NOT EXISTS text_document_change_logs (
    id TEXT PRIMARY KEY,
    document_id TEXT NOT NULL REFERENCES text_documents(id) ON DELETE CASCADE,
    user_id TEXT NOT NULL REFERENCES users(id),
    version INTEGER NOT NULL,
    content_size INTEGER NOT NULL DEFAULT 0,
    changed_at TEXT NOT NULL
);
