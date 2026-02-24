-- Migration 004: must_change_password flag
ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 0;

-- 기존 SUPER_ADMIN admin 계정에 초기 비밀번호 변경 플래그 설정
UPDATE users SET must_change_password = 1 WHERE username = 'admin' AND role = 'SUPER_ADMIN';
