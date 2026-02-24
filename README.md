# OpenSpace — Excel 협업 워크스페이스

> 여러 사용자가 웹 브라우저에서 실시간으로 Excel 스프레드시트를 공동 편집하는 플랫폼

---

## 목차

1. [주요 기능](#주요-기능)
2. [빠른 시작](#빠른-시작)
3. [환경 변수](#환경-변수)
4. [디렉터리 구조](#디렉터리-구조)
5. [기술 스택](#기술-스택)
6. [아키텍처](#아키텍처)
7. [API 개요](#api-개요)
8. [보안](#보안)
9. [운영 주의사항](#운영-주의사항)
10. [라이선스](#라이선스)

---

## 주요 기능

### Excel 완전 왕복(Round-trip) 지원

| 기능 | 상태 | 비고 |
|------|:----:|------|
| 셀 병합 (Merge) | ✅ | import → 브라우저 표시 → 편집 → export |
| 셀 스타일 (굵기·기울임·밑줄·색상·배경·테두리·정렬) | ✅ | import → 브라우저 표시 → 편집 → export |
| 행 높이 | ✅ | xlsx pt 값 보존 |
| 열 너비 | ✅ | xlsx 실제 값 반영 |
| 숫자 서식 (numFmt) | ✅ | 저장 및 export 지원 |
| 틀 고정 (열) | ✅ | jspreadsheet CE `freezeColumns` |
| 틀 고정 (행) | ❌ | jspreadsheet CE 미지원 |
| 조건부 서식·차트·이미지·Named Range | ❌ | jspreadsheet CE 범위 외 |

### 협업 및 관리

| 기능 | 설명 |
|------|------|
| 실시간 공동 편집 | WebSocket 기반, 셀 수정이 모든 접속자에게 즉시 반영 |
| 포맷 툴바 | 브라우저에서 굵기·색상·정렬·병합·테두리 직접 편집 |
| Template 관리 | xlsx 업로드로 서식(컬럼·스타일·병합) 등록, 컬럼별 편집 가능 여부 설정 |
| Workspace 게시 | Template 기반으로 Workspace 생성 및 사용자 배포 |
| 마감(Close) | 관리자가 마감하면 일반 사용자 편집 차단, 관리자는 xlsx 재업로드 가능 |
| 사용자 추가 필드 | 관리자가 정의한 필드(암호화 저장 가능)를 사용자별로 입력 |
| 권한(RBAC) | SUPER_ADMIN / ADMIN / USER 3단계 |
| 감사 로그 | 셀 값 변경 이력(ChangeLog) 자동 기록 |

---

## 빠른 시작

### 요구 사항

- Python 3.10+
- (선택) Python 가상환경 (`python3-venv`)

### 설치 및 실행

```bash
git clone <repo-url> openspace
cd openspace

cp .env.example .env
# .env 에서 SECRET_KEY, KEK_KEY 반드시 변경 (프로덕션)

./start.sh
```

브라우저에서 `http://localhost:8000` 접속

초기 계정: **admin / admin** (SUPER_ADMIN)

### start.sh 동작

1. Python 가상환경 생성 (`.venv/`) 또는 시스템 pip3 fallback
2. `pip install -r requirements.txt`
3. `init_db.py` 실행 (DB + 기본 계정 생성)
4. `uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1`

---

## 환경 변수

`.env.example`을 복사해 `.env`를 작성합니다.

| 변수 | 설명 | 기본값 |
|------|------|--------|
| `SECRET_KEY` | 세션 서명 키 (64자 hex) | `0x00...` — **반드시 변경** |
| `KEK_KEY` | 필드 암호화 키 (64자 hex) | `0x11...` — **반드시 변경** |
| `DATABASE_URL` | SQLite 경로 | `sqlite:///./openspace.db` |
| `SESSION_TTL_SECONDS` | 세션 유효 시간(초) | `28800` (8시간) |
| `DEBUG` | 개발 모드 (`true` / `false`) | `true` |

> `DEBUG=false` 이면 세션 쿠키에 `Secure` 플래그가 추가됩니다 → HTTPS(Nginx) 필수

---

## 디렉터리 구조

```
openspace/
├── app/
│   ├── config.py          # pydantic-settings: .env 파싱
│   ├── database.py        # SQLAlchemy + SQLite WAL 설정
│   ├── models.py          # ORM 모델 (User, Template, Workspace, Cell, ...)
│   ├── crypto.py          # Argon2id 해시 + AES-256-GCM 필드 암호화
│   ├── auth.py            # 세션 관리 + FastAPI 의존성
│   ├── rbac.py            # 권한 상수 + require_role 의존성
│   ├── ws_hub.py          # 인메모리 WebSocket 브로드캐스터
│   ├── main.py            # FastAPI 앱, 미들웨어, 페이지 라우트
│   └── routers/
│       ├── auth.py        # 로그인/로그아웃/현재 사용자
│       ├── users.py       # 사용자 CRUD + 추가 필드 값
│       ├── user_fields.py # 추가 필드 정의 CRUD
│       ├── templates.py   # Template CRUD + xlsx import/export + 스타일 헬퍼
│       ├── workspaces.py  # Workspace CRUD + xlsx import/export
│       ├── cells.py       # 셀 스냅샷 + HTTP 패치
│       └── websocket.py   # WS 연결 + 실시간 패치 브로드캐스트
├── web/
│   ├── templates/         # Jinja2 HTML 페이지
│   └── static/
│       ├── css/style.css
│       └── js/
│           ├── common.js
│           ├── template_edit.js   # Template 편집 + 포맷 툴바
│           ├── workspace.js       # Workspace 실시간 편집 + 포맷 툴바
│           ├── templates.js
│           ├── workspaces.js
│           ├── users.js
│           └── user_fields.js
├── migrations/
│   ├── 001_initial.sql
│   ├── 002_nullable_template_fk.sql
│   └── 003_sheet_meta.sql         # merges / row_heights / freeze_panes 컬럼
├── init_db.py
├── start.sh
├── requirements.txt
├── .env.example
├── AGENTS.md              # 기술 설계 명세서
└── README.md
```

---

## 기술 스택

| 계층 | 기술 |
|------|------|
| **Runtime** | Python 3.10+ |
| **Web Framework** | FastAPI 0.111+ |
| **ASGI Server** | Uvicorn |
| **ORM** | SQLAlchemy 2.0 |
| **DB** | SQLite 3 (WAL 모드) |
| **템플릿** | Jinja2 (서버 사이드 렌더링) |
| **스프레드시트 UI** | Jspreadsheet CE (CDN, MIT) |
| **Excel I/O** | openpyxl |
| **비밀번호 해시** | argon2-cffi (Argon2id) |
| **암호화** | cryptography (AES-256-GCM) |

의존성 (requirements.txt): `fastapi`, `uvicorn[standard]`, `sqlalchemy`, `pydantic-settings`, `argon2-cffi`, `cryptography`, `openpyxl`, `python-multipart`, `jinja2`, `aiofiles`

---

## 아키텍처

```
브라우저
  ├── HTTP REST  →  FastAPI (app/routers/)
  │                   ├── 인증: Argon2id + HttpOnly 세션 쿠키
  │                   ├── CSRF: Double Submit Cookie
  │                   ├── Template/Workspace CRUD
  │                   └── xlsx import / export (openpyxl)
  │
  └── WebSocket  →  app/routers/websocket.py
                        └── WSHub (인메모리 브로드캐스트)

SQLite (WAL)
  ├── users / sessions / user_extra_fields / user_extra_values
  ├── templates / template_sheets / template_columns / template_cells
  └── workspaces / workspace_sheets / workspace_cells / changelogs
```

### 스타일 JSON 형식

셀 스타일은 `style` TEXT 컬럼에 JSON으로 저장됩니다.

```json
{
  "bold": true,
  "italic": false,
  "underline": false,
  "fontSize": 11,
  "color": "FF0000",
  "bg": "FFFF00",
  "align": "center",
  "valign": "middle",
  "wrap": true,
  "border": {
    "top":    {"style": "thin", "color": "000000"},
    "bottom": {"style": "thin", "color": "000000"},
    "left":   {"style": "thin", "color": "000000"},
    "right":  {"style": "thin", "color": "000000"}
  },
  "numFmt": "0.00%"
}
```

### 병합(Merges) 저장 형식

DB: xlsx 범위 문자열 목록

```json
["A1:B2", "C3:E5"]
```

스냅샷 응답(jspreadsheet 형식):

```json
{"A1": [2, 2], "C3": [3, 3]}
```

---

## API 개요

서버 실행 후 `http://localhost:8000/api/docs` 에서 Swagger UI를 통해 전체 명세를 확인할 수 있습니다.

| 메서드 | 경로 | 설명 |
|--------|------|------|
| `POST` | `/api/auth/login` | 로그인 (세션 쿠키 발급) |
| `POST` | `/api/auth/logout` | 로그아웃 |
| `GET`  | `/api/auth/me` | 현재 사용자 정보 |
| `GET/POST` | `/api/admin/users` | 사용자 목록 조회 / 생성 |
| `GET/PATCH/DELETE` | `/api/admin/users/{id}` | 사용자 조회 / 수정 / 삭제 |
| `GET/PATCH` | `/api/admin/users/{id}/field-values` | 추가 필드 값 조회 / 수정 |
| `GET/POST` | `/api/admin/user-fields` | 추가 필드 정의 목록 / 생성 |
| `POST` | `/api/admin/templates/import-xlsx` | xlsx 업로드로 Template 생성 |
| `GET` | `/api/admin/templates/{id}/export-xlsx` | Template → xlsx 다운로드 |
| `GET` | `/api/admin/templates/{id}/sheets/{sid}/snapshot` | 시트 스냅샷 (셀+스타일+병합) |
| `PATCH` | `/api/admin/templates/{id}/sheets/{sid}/merges` | 병합 정보 저장 |
| `POST` | `/api/admin/workspaces` | Workspace 생성 |
| `POST` | `/api/admin/workspaces/{id}/close` | Workspace 마감 |
| `GET` | `/api/admin/workspaces/{id}/export-xlsx` | Workspace → xlsx 다운로드 |
| `PATCH` | `/api/admin/workspaces/{id}/sheets/{sid}/merges` | 병합 정보 저장 |
| `GET` | `/api/workspaces/{id}/sheets/{sid}/snapshot` | 시트 스냅샷 |
| `POST` | `/api/workspaces/{id}/sheets/{sid}/patches` | 셀 일괄 수정 (HTTP) |
| `WS` | `/ws/workspaces/{id}?session_id=<UUID>` | 실시간 협업 WebSocket |

### WebSocket 메시지 형식

```jsonc
// 클라이언트 → 서버
{"type": "patch", "row": 0, "col": 0, "value": "내용", "style": "{\"bold\":true}"}
{"type": "batch_patch", "patches": [{"row":0,"col":0,"value":"A","style":"..."}]}

// 서버 → 클라이언트 (브로드캐스트)
{"type": "patch", "row": 0, "col": 0, "value": "내용", "style": "{\"bold\":true}", "user_id": 1}
```

---

## 보안

| 항목 | 구현 |
|------|------|
| 비밀번호 | Argon2id (argon2-cffi) |
| 세션 | HttpOnly + SameSite=Lax 쿠키, UUID 세션 ID, DB 저장 |
| CSRF | Double Submit Cookie (X-CSRF-Token 헤더) |
| 필드 암호화 | AES-256-GCM (KEK 기반 per-value 키 파생) |
| XSS | Content-Security-Policy 헤더, Jinja2 자동 이스케이프 |
| 클릭재킹 | X-Frame-Options: DENY |
| RBAC | 서버에서 강제 적용, 클라이언트 값 신뢰 안 함 |
| ADMIN 제한 | ADMIN은 SUPER_ADMIN 계정을 관리할 수 없음 |

---

## 운영 주의사항

1. **workers=1 필수** — SQLite WAL + 인메모리 WSHub는 단일 프로세스에서만 동작합니다.
2. **SECRET_KEY, KEK_KEY 변경** — `.env.example`의 기본값은 개발용입니다. 프로덕션에서 반드시 64자 난수 hex로 교체하세요.
3. **HTTPS 필수 (프로덕션)** — `DEBUG=false` 설정 시 쿠키에 `Secure` 플래그가 추가됩니다. Nginx 등 리버스 프록시로 TLS를 제공해야 합니다.
4. **`.env` git 제외** — `.gitignore`에 `.env`를 추가하세요.
5. **DB 백업** — `openspace.db` 파일을 정기 백업하세요.

### Nginx 리버스 프록시 예시

```nginx
server {
    listen 443 ssl;
    server_name example.com;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

---

## 라이선스

Copyright (c) 2026 JAEHYUK CHO

이 프로젝트는 [MIT License](LICENSE)로 배포됩니다.

서드파티 의존성의 라이선스 정보는 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)를 참고하세요.
