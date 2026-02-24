# AGENTS.md — OpenSpace Excel Workspace Web Application

> 이 문서는 `REQUIREMENTS.md`를 기반으로 작성된 기술 설계 명세서입니다.
> 구현 에이전트(AI/개발자)는 이 문서를 단일 진실 공급원(Single Source of Truth)으로 삼아 구현합니다.

---

## 목차

1. [프로젝트 개요](#1-프로젝트-개요)
2. [기술 스택 선정](#2-기술-스택-선정)
3. [시스템 아키텍처](#3-시스템-아키텍처)
4. [데이터베이스 설계](#4-데이터베이스-설계)
5. [인증 및 RBAC 설계](#5-인증-및-rbac-설계)
6. [보안 설계](#6-보안-설계)
7. [Excel 기능 설계](#7-excel-기능-설계)
8. [실시간 협업 설계](#8-실시간-협업-설계)
9. [API 명세](#9-api-명세)
10. [프론트엔드 설계](#10-프론트엔드-설계)
11. [프로젝트 디렉터리 구조](#11-프로젝트-디렉터리-구조)
12. [의존성 및 설치](#12-의존성-및-설치)
13. [환경 변수 및 설정](#13-환경-변수-및-설정)
14. [구현 순서 및 태스크 목록](#14-구현-순서-및-태스크-목록)
15. [비기능 요구사항 충족 전략](#15-비기능-요구사항-충족-전략)

---

## 1. 프로젝트 개요

### 1.1 목적

- 여러 사용자가 **동시에 웹 브라우저에서 Excel 스프레드시트를 공동 편집**할 수 있는 협업 플랫폼
- 관리자가 Excel 서식(Template)을 관리하고 Workspace로 게시
- 마감(Closed) 기능으로 데이터 확정 및 접근 통제

### 1.2 주요 시나리오 요약

| # | 행위자 | 시나리오 |
|---|--------|---------|
| 1 | Admin | 계정 생성/수정/삭제, 추가 필드 관리, Excel 일괄 업로드/다운로드 |
| 2 | Admin | Excel 서식(Template) 업로드·편집, 컬럼별 편집 가능 여부 지정 |
| 3 | Admin | Template 선택 → 이름 지정 → Workspace 게시 |
| 4 | User | 게시된 Workspace에 동시 접속, 실시간 공동 편집 |
| 5 | Admin | Workspace 마감 → 사용자 편집 차단 |
| 6 | Admin | 언제든 Workspace/Template/계정 편집, Excel 다운로드/업로드 가능 |

---

## 2. 기술 스택 선정

### 2.1 선정 기준

- **무료/오픈소스** 기술만 사용
- **설치 의존성 최소화** (패키지 관리가 단순해야 함)
- HTML5 표준 준수 프론트엔드
- SQLite3 DB
- 300명 동시 접속 지원

### 2.2 확정 스택

| 계층 | 기술 | 선정 이유 |
|------|------|----------|
| **Runtime** | Python 3.10+ | 표준 라이브러리 풍부, pip 단일 관리, 서버 운영 용이 |
| **Web Framework** | FastAPI 0.111+ | ASGI, WebSocket 내장, 자동 OpenAPI 문서, Pydantic 검증 |
| **ASGI Server** | Uvicorn | FastAPI 공식 서버, 단일 워커로 SQLite + WebSocket 처리 |
| **ORM** | SQLAlchemy 2.0 (Core + ORM) | Python 표준급 ORM, SQLite WAL 모드 지원 |
| **DB** | SQLite 3 (WAL 모드) | 단일 파일, 설치 불필요, WAL로 동시 읽기 성능 향상 |
| **템플릿** | Jinja2 (서버 사이드) | FastAPI 내장 지원, SSR로 초기 로드 속도 우수 |
| **스프레드시트 UI** | Jspreadsheet CE (로컬 번들) | MIT 라이선스, 10000행·64시트 지원, Excel-like UX, 폐쇄망 대응 |
| **Excel I/O** | openpyxl | MIT 라이선스, xlsx 읽기/쓰기, 서식 보존 |
| **비밀번호 해시** | argon2-cffi | OWASP 권장 Argon2id 알고리즘 |
| **암호화** | cryptography (AES-256-GCM) | PyCA 공식 라이브러리, 인증 암호화 |
| **파일 업로드** | python-multipart | FastAPI 파일 업로드 의존성 |
| **설정 관리** | pydantic-settings | .env 파일 자동 파싱, 타입 검증 |

### 2.3 미사용 및 사유

| 제외 기술 | 제외 사유 |
|----------|----------|
| Redis | 설치 의존성 추가, SQLite WAL로 300명 수준 충분 |
| Celery | 비동기 태스크 불필요, 단순 배치 저장으로 대체 |
| React/Vue/Angular | 번들러·Node.js 빌드 체계 필요, Jinja2+VanillaJS로 충분 |
| JWT | HttpOnly 세션 쿠키가 XSS에 더 안전 |
| Nginx (필수) | 개발/소규모엔 Uvicorn 단독 가능, 프로덕션 권장만 |

---

## 3. 시스템 아키텍처

### 3.1 전체 구성도

```
┌─────────────────────────────────────────────────────┐
│                    브라우저 클라이언트                  │
│  HTML5 + Jspreadsheet CE(로컬 번들) + Vanilla JS       │
│  - HTTP REST  ─────────────────────────────┐         │
│  - WebSocket  ──────────────────────────┐  │         │
└────────────────────────────────────────┼──┼─────────┘
                                         │  │
                              WebSocket  │  │ HTTP/HTTPS
                                         ▼  ▼
┌─────────────────────────────────────────────────────┐
│              FastAPI Application (Uvicorn)            │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐             │
│  │  Routers │ │   Auth   │ │   RBAC   │             │
│  │ /api/... │ │Middleware│ │Dependency│             │
│  └──────────┘ └──────────┘ └──────────┘             │
│  ┌──────────────────────────────────────┐            │
│  │           WebSocket Hub              │            │
│  │  (인메모리 브로드캐스트, 단일 프로세스) │            │
│  └──────────────────────────────────────┘            │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐             │
│  │  SQLAlch │ │ CryptoSvc│ │openpyxl  │             │
│  │   ORM    │ │AES-256-  │ │Excel I/O │             │
│  │          │ │  GCM     │ │          │             │
│  └──────────┘ └──────────┘ └──────────┘             │
└──────────────────────┬──────────────────────────────┘
                       │
               ┌───────▼───────┐
               │  SQLite 3     │
               │  (WAL mode)   │
               │  openspace.db │
               └───────────────┘
```

### 3.2 요청 처리 흐름

```
[Browser]
   │
   ├─ GET / → 로그인 체크 → dashboard.html (Jinja2 SSR)
   │
   ├─ POST /api/auth/login → 세션 쿠키 발급 → CSRF 토큰 설정
   │
   ├─ REST API 호출 (X-CSRF-Token 헤더 포함)
   │    └─ Session 검증 → RBAC 검증 → 비즈니스 로직 → JSON 응답
   │
   └─ WS /ws/workspaces/{id}?session_id=xxx
        └─ WebSocket 핸드셰이크 → 세션 검증 → WSHub 등록
             └─ 브라우저 셀 편집 → patch 메시지 → DB 저장 → 전체 브로드캐스트
```

### 3.3 단일 프로세스 원칙

- Uvicorn **workers=1** 고정
- WebSocket Hub를 **인메모리 dict**로 관리 (멀티 프로세스 불가)
- SQLite WAL 모드로 동시 읽기 성능 확보
- 300명 수준에서 단일 프로세스로 충분 (벤치마크: FastAPI + SQLite WAL ≈ 1000+ RPS)

---

## 4. 데이터베이스 설계

### 4.1 SQLite 설정

```sql
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA foreign_keys = ON;
PRAGMA cache_size = -64000;   -- 64MB 캐시
PRAGMA temp_store = MEMORY;
```

### 4.2 테이블 명세

#### `users` — 사용자 계정

```sql
CREATE TABLE users (
    id          TEXT PRIMARY KEY,            -- UUID v4
    username    TEXT UNIQUE NOT NULL,        -- 로그인 ID
    email       TEXT UNIQUE,                 -- 이메일 (암호화 저장)
    password_hash TEXT NOT NULL,             -- Argon2id 해시
    role        TEXT NOT NULL                -- 'SUPER_ADMIN' | 'ADMIN' | 'USER'
                CHECK (role IN ('SUPER_ADMIN','ADMIN','USER')),
    is_active   INTEGER NOT NULL DEFAULT 1,
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);
-- 초기 데이터: SUPER_ADMIN / admin / admin (해시 저장)
```

#### `user_extra_fields` — 관리자가 정의한 추가 계정 필드

```sql
CREATE TABLE user_extra_fields (
    id          TEXT PRIMARY KEY,
    field_name  TEXT UNIQUE NOT NULL,        -- 필드 키 (snake_case)
    label       TEXT NOT NULL,               -- 화면 표시명
    field_type  TEXT NOT NULL DEFAULT 'text' -- 'text'|'number'|'date'|'boolean'
                CHECK (field_type IN ('text','number','date','boolean')),
    is_sensitive INTEGER NOT NULL DEFAULT 0, -- 1이면 AES 암호화 저장
    is_required INTEGER NOT NULL DEFAULT 0,
    sort_order  INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);
```

#### `user_extra_values` — 추가 필드 값

```sql
CREATE TABLE user_extra_values (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    field_id    TEXT NOT NULL REFERENCES user_extra_fields(id) ON DELETE CASCADE,
    value       TEXT,                        -- is_sensitive=1이면 AES-GCM 암호문
    UNIQUE (user_id, field_id)
);
```

#### `templates` — Excel 서식

```sql
CREATE TABLE templates (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    description TEXT,
    created_by  TEXT NOT NULL REFERENCES users(id),
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);
```

#### `template_sheets` — 서식의 시트 목록 (최대 64개)

```sql
CREATE TABLE template_sheets (
    id           TEXT PRIMARY KEY,
    template_id  TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
    sheet_index  INTEGER NOT NULL,           -- 0-based
    sheet_name   TEXT NOT NULL,
    UNIQUE (template_id, sheet_index)
);
```

#### `template_columns` — 시트별 컬럼 정의

```sql
CREATE TABLE template_columns (
    id           TEXT PRIMARY KEY,
    sheet_id     TEXT NOT NULL REFERENCES template_sheets(id) ON DELETE CASCADE,
    col_index    INTEGER NOT NULL,           -- 0-based
    col_header   TEXT NOT NULL,              -- 컬럼 헤더명
    col_type     TEXT NOT NULL DEFAULT 'text'
                 CHECK (col_type IN ('text','number','date','dropdown','checkbox')),
    col_options  TEXT,                       -- JSON: dropdown 선택지 등
    is_readonly  INTEGER NOT NULL DEFAULT 0, -- 1=사용자 편집 불가
    width        INTEGER DEFAULT 120,        -- 픽셀
    UNIQUE (sheet_id, col_index)
);
```

#### `template_cells` — 서식의 초기 데이터 셀 (최대 10000행 × 컬럼 수)

```sql
CREATE TABLE template_cells (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES template_sheets(id) ON DELETE CASCADE,
    row_index   INTEGER NOT NULL,            -- 0-based, 최대 9999
    col_index   INTEGER NOT NULL,            -- 0-based
    value       TEXT,
    formula     TEXT,                        -- Excel 수식 (선택)
    style       TEXT,                        -- JSON: 배경색, 글꼴 등
    UNIQUE (sheet_id, row_index, col_index)
);
```

#### `workspaces` — 게시된 워크스페이스

```sql
CREATE TABLE workspaces (
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
```

#### `workspace_sheets` — 워크스페이스의 실제 시트 (Template 복사본)

```sql
CREATE TABLE workspace_sheets (
    id              TEXT PRIMARY KEY,
    workspace_id    TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
    template_sheet_id TEXT NOT NULL REFERENCES template_sheets(id),
    sheet_index     INTEGER NOT NULL,
    sheet_name      TEXT NOT NULL,
    UNIQUE (workspace_id, sheet_index)
);
```

#### `workspace_cells` — 워크스페이스 실시간 셀 데이터

```sql
CREATE TABLE workspace_cells (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES workspace_sheets(id) ON DELETE CASCADE,
    row_index   INTEGER NOT NULL,
    col_index   INTEGER NOT NULL,
    value       TEXT,
    style       TEXT,                        -- JSON
    updated_by  TEXT REFERENCES users(id),
    updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    UNIQUE (sheet_id, row_index, col_index)
);
CREATE INDEX idx_ws_cells_sheet ON workspace_cells(sheet_id);
```

#### `change_logs` — 변경 이력

```sql
CREATE TABLE change_logs (
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
CREATE INDEX idx_cl_workspace ON change_logs(workspace_id, changed_at DESC);
```

#### `sessions` — 서버 세션

```sql
CREATE TABLE sessions (
    id          TEXT PRIMARY KEY,            -- UUID v4 (세션 토큰)
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    expires_at  TEXT NOT NULL,               -- TTL: 8시간
    last_seen   TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    ip_address  TEXT,
    user_agent  TEXT
);
CREATE INDEX idx_sessions_user ON sessions(user_id);
CREATE INDEX idx_sessions_expires ON sessions(expires_at);
```

### 4.3 데이터 제약

| 항목 | 제한값 | 적용 위치 |
|------|--------|----------|
| 최대 행 수 | 10,000 | template_cells/workspace_cells row_index CHECK |
| 최대 시트 수 | 64 | template_sheets/workspace_sheets INSERT 전 검증 |
| 세션 TTL | 8시간 | sessions.expires_at |
| 비밀번호 최소 길이 | 8자 | API 레이어 Pydantic 검증 |

---

## 5. 인증 및 RBAC 설계

### 5.1 역할 계층

```
SUPER_ADMIN
    │
    ├─ ADMIN (중간 관리자)
    │
    └─ USER (일반 사용자)
```

| 권한 | SUPER_ADMIN | ADMIN | USER |
|------|:-----------:|:-----:|:----:|
| 계정 생성/수정/삭제 | ✅ | ✅* | ❌ |
| SUPER_ADMIN 계정 수정 | ✅ | ❌ | ❌ |
| 추가 필드 관리 | ✅ | ✅ | ❌ |
| 계정 Excel 다운/업로드 | ✅ | ✅ | ❌ |
| Template 관리 | ✅ | ✅ | ❌ |
| Workspace 게시/마감 | ✅ | ✅ | ❌ |
| Workspace 셀 편집 (OPEN) | ✅ | ✅ | ✅ |
| Workspace 셀 편집 (CLOSED) | ✅ | ✅ | ❌ |
| readonly 컬럼 편집 | ✅ | ✅ | ❌ |

*ADMIN은 SUPER_ADMIN 롤을 가진 계정을 대상으로 관리 불가

### 5.2 세션 인증 흐름

```
로그인 요청
   │
   ├─ username/password 수신
   ├─ DB에서 user 조회
   ├─ Argon2id 검증
   ├─ UUID v4 세션 ID 생성 → sessions 테이블 INSERT
   ├─ HttpOnly Secure SameSite=Strict 쿠키 발급: session_id=<UUID>
   └─ CSRF 토큰 설정: csrf_token 쿠키 (httponly=False, JS 읽기 가능)

API 요청 처리
   │
   ├─ session_id 쿠키 → DB 조회 → User 객체 반환
   ├─ X-CSRF-Token 헤더 == csrf_token 쿠키 검증 (POST/PUT/PATCH/DELETE)
   └─ require_role(min_role) dependency → 역할 검사
```

### 5.3 CSRF 방어

- **Double Submit Cookie** 패턴 사용
- 쿠키: `csrf_token` (httponly=False, SameSite=Strict)
- 헤더: `X-CSRF-Token` (모든 상태 변경 요청에 포함)
- FastAPI 미들웨어에서 검증 (GET/HEAD/OPTIONS 제외)

### 5.4 초기 계정

```
username : admin
password : admin
role     : SUPER_ADMIN
```

> **주의**: 첫 로그인 후 반드시 비밀번호를 변경해야 한다는 알림 표시 권장

---

## 6. 보안 설계

### 6.1 비밀번호 해시 — Argon2id

```python
# argon2-cffi 파라미터
time_cost   = 3        # 반복 횟수
memory_cost = 65536    # 64MB
parallelism = 4
hash_len    = 32
```

### 6.2 민감 정보 암호화 — AES-256-GCM

**Envelope Encryption** 구조:

```
KEK (Key Encryption Key)
  └─ .env의 KEK_KEY (32바이트 hex, 최초 생성 후 고정)

DEK (Data Encryption Key)
  └─ DB에 KEK로 암호화하여 저장 (미래 키 로테이션 대비)
  └─ 현재는 단순화: KEK로 직접 필드 암호화

암호화 대상:
  - user_extra_values.value (is_sensitive=1 필드)
  - users.email (선택적)

암호문 포맷: base64(nonce[12] + ciphertext + tag[16])
```

### 6.3 보안 헤더 (FastAPI Middleware)

```
X-Content-Type-Options: nosniff
X-Frame-Options: DENY
X-XSS-Protection: 1; mode=block
Strict-Transport-Security: max-age=31536000
Content-Security-Policy: default-src 'self'; script-src 'self' 'unsafe-inline' 'unsafe-eval'; ...
Referrer-Policy: strict-origin-when-cross-origin
```

### 6.4 입력 검증 원칙

- 모든 API 입력은 Pydantic 모델로 검증
- 파일 업로드: MIME 타입 + 매직 바이트 검증 (xlsx = PK\x03\x04)
- SQL Injection: SQLAlchemy ORM 파라미터 바인딩 사용 (raw SQL 금지)
- Path Traversal: 파일명 sanitize (werkzeug.secure_filename 패턴)
- 행/열 인덱스: 0 ≤ row < 10000, 0 ≤ col < 1024 서버 검증

### 6.5 Rate Limiting

- 로그인 엔드포인트: IP당 10회/분 (인메모리 슬라이딩 윈도우)
- 일반 API: 300명 × 10 req/s = 3000 req/s 수준, Uvicorn으로 충분

### 6.6 로그 및 감사

- 모든 인증 이벤트 로깅 (로그인 성공/실패, 로그아웃)
- change_logs 테이블로 셀 변경 이력 영구 보존
- Python logging → 표준 출력 (운영 환경에서 수집 가능)

---

## 7. Excel 기능 설계

### 7.1 Template 관리

#### 업로드 (.xlsx → DB)

```
1. multipart/form-data로 .xlsx 파일 수신
2. 매직 바이트 검증 (PK\x03\x04)
3. openpyxl.load_workbook() 로드
4. 시트 수 ≤ 64 검증
5. 각 시트:
   a. template_sheets INSERT
   b. 1행을 컬럼 헤더로 파싱 → template_columns INSERT
   c. 2행~10001행 데이터 → template_cells INSERT (배치)
6. 트랜잭션 커밋
```

#### 다운로드 (DB → .xlsx)

```
1. template_sheets 조회 → openpyxl Workbook 생성
2. 각 시트:
   a. 1행: template_columns 헤더 기록
   b. 2행~: template_cells 데이터 기록
   c. 컬럼 너비 적용
   d. readonly 컬럼: 셀 잠금 + 보호 (openpyxl protection)
3. BytesIO 스트리밍 응답
```

### 7.2 Workspace 생성

```
1. Template 선택 + 이름 입력
2. workspaces INSERT
3. template_sheets → workspace_sheets 복사
4. template_cells → workspace_cells 복사 (초기 데이터)
5. status = 'OPEN'
```

### 7.3 Workspace Excel 업로드 (관리자)

- OPEN/CLOSED 상태 불문 관리자는 업로드 가능
- 업로드된 xlsx → workspace_cells 전체 덮어쓰기 (트랜잭션)
- WebSocket으로 모든 접속 클라이언트에 `reload` 이벤트 브로드캐스트

### 7.4 Workspace Excel 다운로드

- 현재 workspace_cells 상태를 xlsx로 다운로드
- 컬럼 헤더는 template_columns에서 가져옴

### 7.5 사용자 계정 Excel 업로드/다운로드

**다운로드 포맷:**

| id | username | role | email | (추가 필드...) |
|----|----------|------|-------|----------------|

- 민감 필드(is_sensitive=1): 다운로드 시 `[ENCRYPTED]` 표시 또는 복호화 후 표시 (설정 가능)

**업로드 처리:**

```
1. 헤더 행으로 필드 매핑
2. id 존재 → UPDATE, id 없음 → INSERT
3. password 컬럼 있으면 Argon2id 해시 후 저장
4. role=SUPER_ADMIN 행: ADMIN이 업로드 시 건너뜀 (RBAC 강제)
```

### 7.6 배치 저장 전략

- 사용자가 셀 편집 시: 즉시 WebSocket 메시지 → 서버 즉시 DB 저장
- 대량 붙여넣기(paste): 클라이언트에서 `batch_patch` 메시지로 묶어 전송
- 서버: `batch_patch` 수신 시 트랜잭션으로 일괄 INSERT OR REPLACE

---

## 8. 실시간 협업 설계

### 8.1 WebSocket 프로토콜

**연결 URL:**
```
ws://host/ws/workspaces/{workspace_id}?session_id={session_uuid}
```

**메시지 포맷 (JSON):**

클라이언트 → 서버:
```json
// 단일 셀 변경
{
  "type": "patch",
  "sheet_id": "<workspace_sheet_id>",
  "row": 5,
  "col": 2,
  "value": "새 값"
}

// 배치 변경 (붙여넣기 등)
{
  "type": "batch_patch",
  "sheet_id": "<workspace_sheet_id>",
  "patches": [
    {"row": 0, "col": 0, "value": "A"},
    {"row": 0, "col": 1, "value": "B"}
  ]
}

// ping (연결 유지)
{"type": "ping"}
```

서버 → 클라이언트:
```json
// 특정 셀 변경 전파
{
  "type": "patch",
  "sheet_id": "<workspace_sheet_id>",
  "row": 5,
  "col": 2,
  "value": "새 값",
  "updated_by": "<username>"
}

// 배치 변경 전파
{
  "type": "batch_patch",
  "sheet_id": "<workspace_sheet_id>",
  "patches": [...],
  "updated_by": "<username>"
}

// 워크스페이스 상태 변경 (마감 등)
{"type": "workspace_status", "status": "CLOSED"}

// 전체 리로드 요청 (xlsx 업로드 후)
{"type": "reload"}

// pong
{"type": "pong"}

// 에러
{"type": "error", "message": "설명"}
```

### 8.2 WSHub (인메모리 브로드캐스터)

```python
class WSHub:
    # workspace_id → Set[WebSocket]
    _rooms: dict[str, set[WebSocket]] = {}

    async def connect(workspace_id, ws): ...
    async def disconnect(workspace_id, ws): ...
    async def broadcast(workspace_id, message, exclude=None): ...
```

### 8.3 충돌 처리 전략

- **Last-Write-Wins**: 동일 셀에 두 사용자가 동시 편집 시 마지막 도착 메시지 우선
- 셀 단위 잠금(Lock)은 미구현 (요구사항 없음)
- 충돌 감지가 필요한 경우: change_logs로 이력 추적 가능

### 8.4 readonly 컬럼 강제

- 클라이언트: Jspreadsheet `columns[i].readOnly = true` 로 UI 잠금
- 서버: patch 수신 시 template_columns.is_readonly 확인, 위반 시 거부
- USER 역할은 readonly 컬럼 편집 불가, ADMIN 이상은 가능

---

## 9. API 명세

### 9.1 인증 API

| Method | Path | Auth | 설명 |
|--------|------|------|------|
| POST | `/api/auth/login` | 없음 | 로그인, 세션 쿠키 발급 |
| POST | `/api/auth/logout` | 필요 | 세션 무효화 |
| GET | `/api/auth/me` | 필요 | 현재 사용자 정보 |
| PATCH | `/api/auth/me/password` | 필요 | 본인 비밀번호 변경 |

### 9.2 계정 관리 API (ADMIN+)

| Method | Path | 설명 |
|--------|------|------|
| GET | `/api/admin/users` | 사용자 목록 (페이지네이션) |
| POST | `/api/admin/users` | 사용자 생성 |
| GET | `/api/admin/users/{id}` | 사용자 조회 |
| PATCH | `/api/admin/users/{id}` | 사용자 수정 |
| DELETE | `/api/admin/users/{id}` | 사용자 삭제 |
| GET | `/api/admin/users/export-xlsx` | 전체 사용자 Excel 다운로드 |
| POST | `/api/admin/users/import-xlsx` | 사용자 Excel 업로드 |
| GET | `/api/admin/users/{id}/field-values` | 추가 필드값 조회 |
| PATCH | `/api/admin/users/{id}/field-values` | 추가 필드값 수정 |

### 9.3 추가 필드 관리 API (ADMIN+)

| Method | Path | 설명 |
|--------|------|------|
| GET | `/api/admin/user-fields` | 필드 목록 |
| POST | `/api/admin/user-fields` | 필드 생성 |
| PATCH | `/api/admin/user-fields/{id}` | 필드 수정 |
| DELETE | `/api/admin/user-fields/{id}` | 필드 삭제 |
| PATCH | `/api/admin/user-fields/reorder` | 정렬 순서 변경 |

### 9.4 Template 관리 API (ADMIN+)

| Method | Path | 설명 |
|--------|------|------|
| GET | `/api/admin/templates` | 서식 목록 |
| POST | `/api/admin/templates` | 서식 생성 (빈 서식) |
| GET | `/api/admin/templates/{id}` | 서식 조회 (시트+컬럼 포함) |
| PATCH | `/api/admin/templates/{id}` | 서식 메타 수정 |
| DELETE | `/api/admin/templates/{id}` | 서식 삭제 |
| POST | `/api/admin/templates/{id}/copy` | 서식 복사 |
| POST | `/api/admin/templates/import-xlsx` | xlsx 업로드 → 서식 생성 |
| GET | `/api/admin/templates/{id}/export-xlsx` | 서식 xlsx 다운로드 |
| PATCH | `/api/admin/templates/{id}/sheets/{sheetId}/columns/{colId}` | 컬럼 속성 수정 (readonly 등) |
| POST | `/api/admin/templates/{id}/sheets/{sheetId}/cells` | 셀 배치 저장 |

### 9.5 Workspace 관리 API

| Method | Path | Auth | 설명 |
|--------|------|------|------|
| GET | `/api/admin/workspaces` | ADMIN+ | Workspace 목록 |
| POST | `/api/admin/workspaces` | ADMIN+ | Workspace 생성 (Template 선택) |
| GET | `/api/admin/workspaces/{id}` | ADMIN+ | Workspace 상세 |
| PATCH | `/api/admin/workspaces/{id}` | ADMIN+ | Workspace 메타 수정 |
| DELETE | `/api/admin/workspaces/{id}` | ADMIN+ | Workspace 삭제 |
| POST | `/api/admin/workspaces/{id}/close` | ADMIN+ | 마감 처리 |
| POST | `/api/admin/workspaces/{id}/reopen` | ADMIN+ | 마감 해제 |
| POST | `/api/admin/workspaces/{id}/import-xlsx` | ADMIN+ | xlsx 업로드 → 덮어쓰기 |
| GET | `/api/admin/workspaces/{id}/export-xlsx` | ADMIN+ | xlsx 다운로드 |
| GET | `/api/workspaces` | USER+ | 접근 가능한 Workspace 목록 |
| GET | `/api/workspaces/{id}` | USER+ | Workspace 상세 (시트 목록) |
| GET | `/api/workspaces/{id}/sheets/{sheetId}/snapshot` | USER+ | 시트 전체 셀 데이터 |
| POST | `/api/workspaces/{id}/sheets/{sheetId}/patches` | USER+ | HTTP fallback 패치 |

### 9.6 WebSocket

| Path | Auth | 설명 |
|------|------|------|
| `WS /ws/workspaces/{id}` | 쿼리 파라미터 session_id | 실시간 협업 |

### 9.7 공통 응답 형식

```json
// 성공
{"data": {...}, "message": "ok"}

// 에러
{"detail": "에러 메시지", "code": "ERROR_CODE"}
```

### 9.8 페이지 라우트 (SSR)

| Path | 설명 |
|------|------|
| `GET /` | 로그인 or 대시보드 리다이렉트 |
| `GET /login` | 로그인 페이지 |
| `GET /dashboard` | 대시보드 |
| `GET /admin/users` | 계정 관리 |
| `GET /admin/user-fields` | 추가 필드 관리 |
| `GET /admin/templates` | 서식 관리 |
| `GET /admin/workspaces` | Workspace 관리 |
| `GET /workspaces` | Workspace 목록 (USER) |
| `GET /workspaces/{id}` | Workspace 편집 화면 |

---

## 10. 프론트엔드 설계

### 10.1 공통 레이아웃

```
┌─────────────────────────────────────────┐
│  OpenSpace        [사용자명] [로그아웃]   │  ← 상단 헤더
├──────────┬──────────────────────────────┤
│          │                              │
│  사이드바 │        메인 콘텐츠           │
│          │                              │
│ - 대시보드│                              │
│ - 계정관리│                              │  (ADMIN+만 표시)
│ - 추가필드│                              │
│ - 서식관리│                              │
│ - WS관리  │                              │
│ ─────── │                              │
│ - WS목록  │                              │  (USER도 표시)
│          │                              │
└──────────┴──────────────────────────────┘
```

### 10.2 페이지별 주요 기능

#### 로그인 (`/login`)
- username/password 폼
- 로그인 실패 횟수 표시
- CSRF 토큰 쿠키 수신

#### 대시보드 (`/dashboard`)
- 통계 카드: 전체 사용자, 서식 수, 워크스페이스 수
- 최근 변경 이력 (ADMIN+)
- 현재 OPEN Workspace 목록

#### 계정 관리 (`/admin/users`)
- 사용자 테이블 (검색, 정렬, 페이지네이션)
- 사용자 추가/수정/삭제 모달
- 추가 필드값 편집 패널
- Excel 업로드/다운로드 버튼

#### 서식 관리 (`/admin/templates`)
- 서식 카드 목록
- xlsx 업로드 → 서식 생성
- 서식 편집 화면:
  - Jspreadsheet로 데이터 편집
  - 컬럼 헤더 클릭 → 컬럼 속성 패널 (타입, readonly, 너비)
- 서식 복사/삭제/다운로드

#### Workspace 관리 (`/admin/workspaces`)
- Workspace 목록 (상태 배지: OPEN/CLOSED)
- 새 Workspace 생성 (서식 선택 드롭다운)
- 마감/재개 토글 버튼
- xlsx 업로드/다운로드

#### Workspace 편집 (`/workspaces/{id}`)

```
┌──────────────────────────────────────────────────┐
│  [워크스페이스명]  상태: OPEN  [다운로드] [업로드] │  ← 헤더
├──────────────────────────────────────────────────┤
│  Sheet1 │ Sheet2 │ Sheet3 │                      │  ← 탭
├──────────────────────────────────────────────────┤
│                                                  │
│         Jspreadsheet 스프레드시트                 │
│         (10000행, 실시간 동기화)                  │
│                                                  │
│  ✓ [수정자: 홍길동 (2행 3열)]                    │  ← 상태 표시줄
└──────────────────────────────────────────────────┘
```

- WebSocket 연결 상태 표시 (녹색/빨강 점)
- 다른 사용자 편집 시 셀 하이라이트 (일시적)
- CLOSED 상태: 편집 불가 오버레이 표시

### 10.3 Jspreadsheet CE 설정

```javascript
jspreadsheet(element, {
    data: snapshotData,          // 서버에서 내려받은 초기 데이터
    columns: columnDefs,         // template_columns 기반
    minDimensions: [cols, 100],  // 최소 100행 표시
    tableOverflow: true,
    tableWidth: '100%',
    tableHeight: 'calc(100vh - 200px)',
    lazyLoading: true,           // 10000행 성능 대응
    loadingSpin: true,
    onchange: handleCellChange,  // WebSocket 전송
    onpaste: handlePaste,        // batch_patch 전송
    license: 'CE',               // Community Edition
});
```

### 10.4 JavaScript 모듈 구조

```
web/static/js/
├── common.js          # CSRF 토큰, fetch 래퍼, 토스트 알림
├── workspace.js       # Jspreadsheet 초기화, WebSocket 연결/처리
├── users.js           # 계정 관리 CRUD
├── templates.js       # 서식 관리 CRUD
└── workspaces.js      # Workspace 관리 CRUD
```

**common.js 핵심 함수:**

```javascript
// CSRF 포함 fetch 래퍼
async function apiFetch(url, options = {}) {
    const csrfToken = getCookie('csrf_token');
    return fetch(url, {
        ...options,
        headers: {
            'Content-Type': 'application/json',
            'X-CSRF-Token': csrfToken,
            ...options.headers,
        },
        credentials: 'same-origin',
    });
}

function showToast(message, type = 'info') { ... }
function showModal(title, content) { ... }
```

---

## 11. 프로젝트 디렉터리 구조

```
openspace/
│
├── app/
│   ├── __init__.py
│   ├── config.py              # pydantic-settings Settings
│   ├── database.py            # SQLAlchemy engine + session + WAL 설정
│   ├── models.py              # ORM 모델 전체
│   ├── crypto.py              # CryptoService: Argon2id + AES-256-GCM
│   ├── auth.py                # 세션 관리, get_current_user dependency
│   ├── rbac.py                # ROLE 상수, require_role, can_manage_user
│   ├── ws_hub.py              # WSHub 싱글턴
│   ├── main.py                # FastAPI 앱 생성, 미들웨어, 페이지 라우트
│   └── routers/
│       ├── __init__.py
│       ├── auth.py            # /api/auth/*
│       ├── users.py           # /api/admin/users/*
│       ├── user_fields.py     # /api/admin/user-fields/*
│       ├── templates.py       # /api/admin/templates/*
│       ├── workspaces.py      # /api/admin/workspaces/* + /api/workspaces/*
│       ├── cells.py           # /api/workspaces/{id}/sheets/{sheetId}/*
│       └── websocket.py       # WS /ws/workspaces/{id}
│
├── web/
│   ├── templates/
│   │   ├── layout.html        # 공통 레이아웃 (헤더, 사이드바)
│   │   ├── login.html
│   │   ├── dashboard.html
│   │   ├── users.html
│   │   ├── user_fields.html
│   │   ├── templates.html
│   │   ├── workspaces.html    # 관리자 Workspace 관리
│   │   ├── workspace_list.html # USER용 목록
│   │   └── workspace.html     # Jspreadsheet 편집 화면
│   └── static/
│       ├── css/
│       │   └── style.css
│       ├── lib/                     # 프론트엔드 로컬 번들 (폐쇄망 대응)
│       │   ├── jsuites.css
│       │   ├── jsuites.js
│       │   ├── jspreadsheet.css
│       │   ├── jspreadsheet.js      # Jspreadsheet CE v4
│       │   └── jspreadsheet-formula.js  # @jspreadsheet/formula v2
│       └── js/
│           ├── common.js
│           ├── workspace.js
│           ├── users.js
│           ├── templates.js
│           └── workspaces.js
│
├── migrations/
│   └── 001_initial.sql        # 초기 스키마 + SUPER_ADMIN INSERT
│
├── init_db.py                 # DB 파일 생성 + 스키마 적용 + 초기 데이터
├── start.sh                   # 전체 시작 스크립트
├── requirements.txt
├── .env.example
├── .gitignore
├── REQUIREMENTS.md
├── AGENTS.md                  # 이 문서
└── README.md
```

---

## 12. 의존성 및 설치

### 12.1 requirements.txt

```
fastapi>=0.111.0
uvicorn[standard]>=0.29.0
sqlalchemy>=2.0.0
pydantic-settings>=2.0.0
argon2-cffi>=23.1.0
cryptography>=42.0.0
openpyxl>=3.1.0
python-multipart>=0.0.9
jinja2>=3.1.0
aiofiles>=23.2.0
```

> **의도적 제외**: Redis, Celery, Node.js, JWT 라이브러리, 외부 인증 서비스

### 12.2 프론트엔드 로컬 번들 (폐쇄망 대응)

모든 프론트엔드 라이브러리는 `web/static/lib/`에 로컬 번들링되어 있어 **외부 인터넷 연결이 불필요**합니다.

```
web/static/lib/
├── jsuites.css              # jSuites UI 컴포넌트 스타일
├── jsuites.js               # jSuites UI 컴포넌트
├── jspreadsheet.css         # Jspreadsheet CE 스타일
├── jspreadsheet.js          # Jspreadsheet CE v4 엔진
└── jspreadsheet-formula.js  # @jspreadsheet/formula v2 수식 엔진
```

HTML 참조 (workspace.html, template_edit.html):
```html
<link rel="stylesheet" href="/static/lib/jsuites.css">
<link rel="stylesheet" href="/static/lib/jspreadsheet.css">
<script src="/static/lib/jsuites.js"></script>
<script src="/static/lib/jspreadsheet.js"></script>
<script src="/static/lib/jspreadsheet-formula.js"></script>
```

> 원본 CDN URL (업데이트 필요 시 참고):
> - `https://cdn.jsdelivr.net/npm/jsuites/dist/jsuites.{css,js}`
> - `https://cdn.jsdelivr.net/npm/jspreadsheet-ce@4/dist/{jspreadsheet.css,index.js}`
> - `https://cdn.jsdelivr.net/npm/@jspreadsheet/formula@2/dist/index.js`

### 12.3 start.sh

```bash
#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Python venv 생성 시도, 실패 시 --break-system-packages 사용
if python3 -m venv .venv 2>/dev/null; then
    source .venv/bin/activate
    pip install --quiet -r requirements.txt
else
    echo "[INFO] venv 사용 불가, 시스템 pip 사용"
    pip3 install --break-system-packages --quiet -r requirements.txt
fi

# .env 파일 없으면 예시에서 복사
if [ ! -f .env ]; then
    cp .env.example .env
    echo "[WARN] .env 파일 생성됨. SECRET_KEY와 KEK_KEY를 반드시 변경하세요!"
fi

# DB 초기화 (이미 존재하면 스킵)
python3 init_db.py

# 서버 시작 (단일 워커 필수)
uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1
```

### 12.4 .env.example

```ini
# 반드시 변경 (64자 hex)
SECRET_KEY=change_me_to_64_hex_chars_random_string_here_000000000000000000

# 필드 암호화 키 (64자 hex = 32바이트)
KEK_KEY=change_me_to_64_hex_chars_for_field_encryption_key_000000000000

# DB 파일 경로
DATABASE_URL=sqlite:///./openspace.db

# 세션 TTL (초)
SESSION_TTL_SECONDS=28800

# 개발 모드 (true: HTTPS 쿠키 비활성화)
DEBUG=false

# 허용 호스트 (쉼표 구분)
ALLOWED_HOSTS=localhost,127.0.0.1
```

---

## 13. 환경 변수 및 설정

### 13.1 config.py (pydantic-settings)

```python
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    secret_key: str
    kek_key: str
    database_url: str = "sqlite:///./openspace.db"
    session_ttl_seconds: int = 28800
    debug: bool = False
    allowed_hosts: list[str] = ["localhost"]

    class Config:
        env_file = ".env"

settings = Settings()
```

### 13.2 데이터베이스 초기화 (init_db.py)

```python
"""
실행: python3 init_db.py
- migrations/001_initial.sql 적용
- SUPER_ADMIN 계정 없으면 생성 (admin/admin)
- 멱등성 보장 (이미 존재하면 스킵)
"""
```

---

## 14. 구현 순서 및 태스크 목록

에이전트는 아래 순서를 따라 구현합니다. 각 태스크는 독립적으로 테스트 가능해야 합니다.

### Phase 1: 기반 인프라

- [ ] **T01** `requirements.txt`, `.env.example`, `.gitignore` 작성
- [ ] **T02** `app/config.py` — Settings 클래스
- [ ] **T03** `app/database.py` — SQLAlchemy engine (WAL), session factory
- [ ] **T04** `app/models.py` — 전체 ORM 모델 정의
- [ ] **T05** `migrations/001_initial.sql` — CREATE TABLE 전체
- [ ] **T06** `app/crypto.py` — CryptoService (Argon2id + AES-256-GCM)
- [ ] **T07** `init_db.py` — DB 초기화 스크립트

### Phase 2: 인증 및 RBAC

- [ ] **T08** `app/auth.py` — 세션 발급/검증/무효화, get_current_user
- [ ] **T09** `app/rbac.py` — ROLE 상수, require_role, can_manage_user
- [ ] **T10** `app/routers/auth.py` — login/logout/me/password 엔드포인트
- [ ] **T11** `app/main.py` — FastAPI 앱, 보안 헤더 미들웨어, CSRF 검증, 정적 파일

### Phase 3: 계정 관리

- [ ] **T12** `app/routers/users.py` — CRUD + Excel 업/다운로드
- [ ] **T13** `app/routers/user_fields.py` — 추가 필드 CRUD
- [ ] **T14** 계정 관련 Jinja2 템플릿 (`users.html`, `user_fields.html`)
- [ ] **T15** `web/static/js/users.js`

### Phase 4: Template 관리

- [ ] **T16** `app/routers/templates.py` — CRUD + xlsx 업/다운로드 + 컬럼 설정
- [ ] **T17** 서식 관련 Jinja2 템플릿 (`templates.html`)
- [ ] **T18** `web/static/js/templates.js`

### Phase 5: Workspace 및 실시간 협업

- [ ] **T19** `app/ws_hub.py` — WSHub 구현
- [ ] **T20** `app/routers/workspaces.py` — CRUD + 마감 + xlsx 업/다운로드
- [ ] **T21** `app/routers/cells.py` — snapshot + HTTP fallback 패치
- [ ] **T22** `app/routers/websocket.py` — WebSocket 핸들러
- [ ] **T23** Workspace 관련 Jinja2 템플릿 (`workspaces.html`, `workspace.html`, `workspace_list.html`)
- [ ] **T24** `web/static/js/workspace.js` — Jspreadsheet + WebSocket

### Phase 6: 공통 UI 및 마무리

- [ ] **T25** `web/templates/layout.html`, `login.html`, `dashboard.html`
- [ ] **T26** `web/static/css/style.css`
- [ ] **T27** `web/static/js/common.js`
- [ ] **T28** `start.sh` 작성 및 전체 통합 테스트
- [ ] **T29** `README.md` 작성

---

## 15. 비기능 요구사항 충족 전략

### 15.1 300명 동시 접속

| 부하 유형 | 전략 |
|----------|------|
| HTTP 요청 | FastAPI + Uvicorn ASGI (비동기 I/O, 단일 워커로 수천 RPS) |
| WebSocket | asyncio 기반 비동기 연결, SQLite WAL 읽기 병렬 처리 |
| DB 쓰기 경쟁 | 배치 저장 + SQLite WAL (쓰기 직렬화, 읽기 비차단) |
| 메모리 | Jspreadsheet lazyLoading, 서버 스냅샷 청크 전송 옵션 |

### 15.2 설치 의존성 최소화 및 폐쇄망 대응

- Python 표준 라이브러리 최대 활용
- pip 패키지 10개 이내
- 추가 서비스(Redis, DB 서버 등) 불필요
- 프론트엔드 빌드 불필요 (로컬 번들 파일 직접 서빙)
- **외부 인터넷 연결 불필요** — 모든 JS/CSS 라이브러리가 `web/static/lib/`에 포함
- CSP 헤더에 외부 도메인 없음 — 폐쇄망 환경에서 완전 동작

### 15.3 보안 체크리스트

- [x] Argon2id 비밀번호 해시
- [x] AES-256-GCM 민감 정보 암호화
- [x] HttpOnly + SameSite=Strict 세션 쿠키
- [x] CSRF Double Submit Cookie
- [x] SQL Injection 방지 (ORM 파라미터 바인딩)
- [x] XSS 방지 (Jinja2 자동 이스케이프, CSP 헤더)
- [x] 파일 업로드 검증 (매직 바이트)
- [x] RBAC 서버 강제 (클라이언트 신뢰 안 함)
- [x] 로그인 실패 제한 (Rate Limiting)
- [x] 보안 응답 헤더
- [x] 입력값 서버 검증 (Pydantic)

### 15.4 데이터 무결성

- SQLite WAL + PRAGMA synchronous=NORMAL (성능/안정성 균형)
- 트랜잭션: xlsx 업로드는 반드시 원자적 실행
- change_logs로 모든 셀 변경 감사 추적
- 소프트 삭제 미사용 (복잡도 최소화), 삭제 전 확인 API 제공

---

## 부록 A. 주요 파일 구현 가이드

### A.1 database.py 핵심

```python
from sqlalchemy import create_engine, event
from sqlalchemy.orm import sessionmaker, declarative_base

DATABASE_URL = settings.database_url

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},
    pool_pre_ping=True,
)

@event.listens_for(engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA synchronous=NORMAL")
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.execute("PRAGMA cache_size=-64000")
    cursor.execute("PRAGMA temp_store=MEMORY")
    cursor.close()

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
```

### A.2 crypto.py 핵심

```python
import os, base64
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError

ph = PasswordHasher(time_cost=3, memory_cost=65536, parallelism=4)

class CryptoService:
    def __init__(self, kek_hex: str):
        self._kek = bytes.fromhex(kek_hex)

    def hash_password(self, password: str) -> str:
        return ph.hash(password)

    def verify_password(self, hash: str, password: str) -> bool:
        try:
            return ph.verify(hash, password)
        except VerifyMismatchError:
            return False

    def encrypt(self, plaintext: str) -> str:
        nonce = os.urandom(12)
        aesgcm = AESGCM(self._kek)
        ct = aesgcm.encrypt(nonce, plaintext.encode(), None)
        return base64.b64encode(nonce + ct).decode()

    def decrypt(self, ciphertext_b64: str) -> str:
        data = base64.b64decode(ciphertext_b64)
        nonce, ct = data[:12], data[12:]
        aesgcm = AESGCM(self._kek)
        return aesgcm.decrypt(nonce, ct, None).decode()
```

### A.3 workspace.js 핵심 흐름

```javascript
// 1. 서버 데이터로 Jspreadsheet 초기화
async function initWorkspace(workspaceId, sheetMeta) {
    const snapshot = await apiFetch(`/api/workspaces/${workspaceId}/sheets/${sheetMeta[0].id}/snapshot`);
    const columns = buildColumnDefs(sheetMeta[0].columns, isClosed, userRole);
    jspreadsheet(document.getElementById('spreadsheet'), {
        data: snapshot.data.cells,
        columns,
        onchange: (el, cell, x, y, value) => sendPatch(x, y, value),
        onpaste: (el, data) => sendBatchPatch(data),
    });
    connectWebSocket(workspaceId);
}

// 2. WebSocket 연결 및 이벤트 처리
function connectWebSocket(workspaceId) {
    const sessionId = document.cookie.match(/session_id=([^;]+)/)?.[1];
    ws = new WebSocket(`ws://${location.host}/ws/workspaces/${workspaceId}?session_id=${sessionId}`);
    ws.onmessage = (event) => {
        const msg = JSON.parse(event.data);
        if (msg.type === 'patch') applyPatch(msg);
        if (msg.type === 'batch_patch') msg.patches.forEach(applyPatch);
        if (msg.type === 'workspace_status') handleStatusChange(msg.status);
        if (msg.type === 'reload') location.reload();
    };
}
```

---

## 부록 B. 제약 사항 및 알려진 한계

| 항목 | 내용 |
|------|------|
| 멀티 프로세스 불가 | WSHub 인메모리 → workers=1 고정 |
| 수식 실행 | openpyxl이 수식 값 미계산, 클라이언트 Jspreadsheet에서 처리 |
| 파일 크기 | xlsx 업로드 최대 50MB 권장 (서버 메모리 고려) |
| 동시 셀 충돌 | Last-Write-Wins, 락 없음 |
| 모바일 지원 | Jspreadsheet는 데스크톱 최적화, 모바일은 제한적 |
| 오프라인 편집 | 미지원 (WebSocket 단절 시 reconnect 로직만) |
| 라이브러리 업데이트 | `web/static/lib/` 파일을 CDN에서 재다운로드하여 교체 필요 |

---

*최종 수정: 2026-02-24*
*기반 문서: REQUIREMENTS.md*
