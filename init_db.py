#!/usr/bin/env python3
"""
DB 초기화 스크립트.
- migrations/001_initial.sql 적용
- migrations/004_must_change_password.sql 적용 (필요 시)
- SUPER_ADMIN 계정 없으면 생성 (admin / admin)
- 멱등성 보장 (이미 존재하면 스킵)
"""
import os
import sys
import uuid
import sqlite3
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from app.config import settings
from app.crypto import crypto

DB_PATH = settings.database_url.replace("sqlite:///", "").replace("sqlite://", "")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def apply_migration(conn):
    sql_path = ROOT / "migrations" / "001_initial.sql"
    script = sql_path.read_text(encoding="utf-8")
    conn.executescript(script)
    print("[init_db] Migration 001 applied.")


def apply_incremental_migrations(conn):
    """002~004 마이그레이션을 순서대로 안전하게 적용 (멱등)."""
    migrations = [
        "002_nullable_template_fk.sql",
        "003_sheet_meta.sql",
        "004_must_change_password.sql",
        "005_cell_comments.sql",
        "006_conditional_formats.sql",
    ]
    for fname in migrations:
        path = ROOT / "migrations" / fname
        if not path.exists():
            continue
        script = path.read_text(encoding="utf-8")
        for stmt in script.split(";"):
            stmt = stmt.strip()
            if not stmt:
                continue
            # 주석 줄 제거 후 실제 SQL이 있는지 확인 (주석으로 시작하는 청크도 실행)
            non_comment = "\n".join(
                line for line in stmt.splitlines()
                if not line.strip().startswith("--")
            ).strip()
            if not non_comment:
                continue
            try:
                conn.execute(stmt)
            except sqlite3.OperationalError as e:
                msg = str(e).lower()
                # 이미 존재하는 컬럼/테이블은 무시
                if "duplicate column" in msg or "already exists" in msg:
                    pass
                else:
                    raise
        conn.commit()
        print(f"[init_db] Migration {fname} applied.")


def create_super_admin(conn):
    cur = conn.execute("SELECT id FROM users WHERE role='SUPER_ADMIN' LIMIT 1")
    if cur.fetchone():
        print("[init_db] SUPER_ADMIN already exists. Skipping.")
        return

    user_id = str(uuid.uuid4())
    password_hash = crypto.hash_password("admin")
    conn.execute(
        """
        INSERT INTO users (id, username, email, password_hash, role, is_active, must_change_password)
        VALUES (?, ?, ?, ?, 'SUPER_ADMIN', 1, 1)
        """,
        (user_id, "admin", None, password_hash),
    )
    conn.commit()
    print(f"[init_db] SUPER_ADMIN created: username=admin, password=admin (id={user_id})")
    print("[init_db] WARNING: Change the admin password immediately after first login!")


def main():
    print(f"[init_db] Using DB: {DB_PATH}")
    conn = get_conn()
    try:
        apply_migration(conn)
        apply_incremental_migrations(conn)
        create_super_admin(conn)
    finally:
        conn.close()
    print("[init_db] Done.")


if __name__ == "__main__":
    main()
