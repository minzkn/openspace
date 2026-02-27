# OpenSpace — Excel 협업 워크스페이스

> 여러 사용자가 웹 브라우저에서 실시간으로 Excel 스프레드시트를 공동 편집하는 플랫폼

---

## 목차

1. [주요 기능](#주요-기능)
2. [빠른 시작](#빠른-시작)
3. [환경 변수](#환경-변수)
4. [프로덕션 배포](#프로덕션-배포)
   - [systemd 서비스 등록](#systemd-서비스-등록)
   - [Nginx 리버스 프록시](#nginx-리버스-프록시)
   - [Apache 리버스 프록시](#apache-리버스-프록시)
   - [SSL/TLS 인증서 설정](#ssltls-인증서-설정)
5. [데이터베이스 관리](#데이터베이스-관리)
6. [디렉터리 구조](#디렉터리-구조)
7. [기술 스택](#기술-스택)
8. [아키텍처](#아키텍처)
9. [API 개요](#api-개요)
10. [보안](#보안)
11. [운영 주의사항](#운영-주의사항)
12. [문제 해결](#문제-해결)
13. [라이선스](#라이선스)

---

## 주요 기능

### Excel 완전 왕복(Round-trip) 지원

| 기능 | 상태 | 비고 |
|------|:----:|------|
| 셀 병합 (Merge) | ✅ | import → 브라우저 표시 → 편집 → export |
| 셀 스타일 (굵기·기울임·밑줄·색상·배경·테두리·정렬) | ✅ | import → 브라우저 표시 → 편집 → export |
| 폰트 이름 (fontName) | ✅ | Arial, Malgun Gothic 등 비-Calibri 폰트 보존 |
| 들여쓰기 (indent) | ✅ | Excel indent level 보존 |
| 텍스트 회전 (textRotation) | ✅ | 각도 값 보존 |
| 밑줄 유형 (underline) | ✅ | single, double, singleAccounting 등 보존 |
| 테마 색상 (theme colors) | ✅ | theme/indexed/rgb 모든 타입 → RGB 변환 보존 |
| 날짜/시간 값 (datetime) | ✅ | ISO 형식 저장, export 시 Excel 날짜 타입 복원 |
| 불리언 값 (boolean) | ✅ | TRUE/FALSE로 저장, export 시 Excel 불리언 복원 |
| 행 높이 | ✅ | xlsx pt 값 보존 |
| 열 너비 | ✅ | xlsx 실제 값 반영 |
| 숫자 서식 (numFmt) | ✅ | 저장 및 export 지원 |
| 틀 고정 (열) | ✅ | jspreadsheet CE `freezeColumns` |
| 틀 고정 (행) | ❌ | jspreadsheet CE 미지원 |
| 규격 위반 xlsx 자동 수정 | ✅ | font family > 14 등 openpyxl 호환성 자동 보정 |
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
- 지원 OS: Linux, macOS, Windows (WSL 권장)

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

> 첫 로그인 후 반드시 비밀번호를 변경하세요.

### start.sh 동작

1. `.env` 파일이 없으면 `.env.example`에서 자동 복사
2. Python 가상환경 생성 (`.venv/`) 또는 시스템 pip3 fallback
3. `pip install -r requirements.txt`
4. `init_db.py` 실행 (DB 스키마 생성 + 마이그레이션 + 기본 계정 생성, 멱등)
5. `uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1`

### 수동 설치 (start.sh 미사용)

```bash
cd openspace

# 가상환경 생성 및 활성화
python3 -m venv .venv
source .venv/bin/activate

# 의존성 설치
pip install -r requirements.txt

# 환경 변수 설정
cp .env.example .env
# .env 파일을 편집하여 SECRET_KEY, KEK_KEY 변경

# DB 초기화
python3 init_db.py

# 서버 시작
uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1
```

### 포트 변경

기본 포트(8000)를 변경하려면 uvicorn 실행 시 `--port` 옵션을 수정합니다.

```bash
uvicorn app.main:app --host 0.0.0.0 --port 9000 --workers 1
```

> `start.sh`를 직접 수정하거나, 수동 실행으로 포트를 지정할 수 있습니다.

---

## 환경 변수

`.env.example`을 복사해 `.env`를 작성합니다.

| 변수 | 설명 | 기본값 |
|------|------|--------|
| `SECRET_KEY` | 세션 서명 키 (64자 hex = 32바이트) | `0x00...` — **반드시 변경** |
| `KEK_KEY` | 필드 암호화 키 (64자 hex = 32바이트) | `0x11...` — **반드시 변경** |
| `DATABASE_URL` | SQLite 경로 | `sqlite:///./openspace.db` |
| `SESSION_TTL_SECONDS` | 세션 유효 시간(초) | `28800` (8시간) |
| `DEBUG` | 개발 모드 (`true` / `false`) | `true` |
| `MAX_UPLOAD_SIZE` | 최대 파일 업로드 크기 (바이트) | `52428800` (50MB) |
| `ALLOWED_HOSTS` | 허용할 호스트 목록 (JSON 배열) | `["*"]` |

### 프로덕션 키 생성

```bash
# SECRET_KEY 생성 (64자 hex)
python3 -c "import secrets; print(secrets.token_hex(32))"

# KEK_KEY 생성 (64자 hex)
python3 -c "import secrets; print(secrets.token_hex(32))"
```

> `DEBUG=false` 이면 세션 쿠키에 `Secure` 플래그가 추가됩니다 → HTTPS 필수
>
> `ALLOWED_HOSTS`를 설정하면 TrustedHostMiddleware가 활성화되어, 지정한 호스트 외의 요청을 차단합니다.

### .env 파일 예시 (프로덕션)

```env
SECRET_KEY=a1b2c3d4e5f6...  # python3 -c "import secrets; print(secrets.token_hex(32))" 로 생성
KEK_KEY=f6e5d4c3b2a1...     # 위와 동일하게 생성
DATABASE_URL=sqlite:///./openspace.db
SESSION_TTL_SECONDS=28800
DEBUG=false
MAX_UPLOAD_SIZE=52428800
ALLOWED_HOSTS=["openspace.example.com", "www.openspace.example.com"]
```

---

## 프로덕션 배포

프로덕션 환경에서는 Uvicorn을 직접 노출하지 않고, 리버스 프록시(Nginx 또는 Apache)를 통해 HTTPS를 제공합니다.

```
[브라우저] ──HTTPS──> [Nginx/Apache :443] ──HTTP──> [Uvicorn :8000]
                                          ──WS───>
```

### systemd 서비스 등록

Uvicorn을 시스템 서비스로 등록하면 서버 재부팅 시 자동 시작됩니다.

#### 1. 서비스 사용자 생성 (선택)

```bash
sudo useradd -r -s /bin/false -d /opt/openspace openspace
sudo mkdir -p /opt/openspace
sudo cp -r . /opt/openspace/
sudo chown -R openspace:openspace /opt/openspace
```

#### 2. 가상환경 및 초기화

```bash
sudo -u openspace bash -c '
  cd /opt/openspace
  python3 -m venv .venv
  source .venv/bin/activate
  pip install -r requirements.txt
  cp .env.example .env
  # .env 편집 후:
  python3 init_db.py
'
```

> `.env` 파일에서 `SECRET_KEY`, `KEK_KEY`를 반드시 변경하고, `DEBUG=false`로 설정하세요.

#### 3. systemd Unit 파일 생성

```bash
sudo tee /etc/systemd/system/openspace.service > /dev/null << 'EOF'
[Unit]
Description=OpenSpace Excel Collaboration Workspace
After=network.target

[Service]
Type=exec
User=openspace
Group=openspace
WorkingDirectory=/opt/openspace
Environment=PATH=/opt/openspace/.venv/bin:/usr/bin:/bin
ExecStart=/opt/openspace/.venv/bin/uvicorn app.main:app \
    --host 127.0.0.1 \
    --port 8000 \
    --workers 1 \
    --log-level info
Restart=on-failure
RestartSec=5

# 보안 강화 옵션
NoNewPrivileges=true
ProtectSystem=strict
ProtectHome=true
ReadWritePaths=/opt/openspace
PrivateTmp=true

[Install]
WantedBy=multi-user.target
EOF
```

> `--host 127.0.0.1`: 리버스 프록시 사용 시 로컬만 바인딩합니다.
> `--workers 1`: SQLite WAL + 인메모리 WSHub 특성상 **반드시 1**이어야 합니다.

#### 4. 서비스 활성화 및 시작

```bash
sudo systemctl daemon-reload
sudo systemctl enable openspace
sudo systemctl start openspace

# 상태 확인
sudo systemctl status openspace

# 로그 확인
sudo journalctl -u openspace -f
```

#### 5. 서비스 관리 명령

```bash
# 서비스 중지
sudo systemctl stop openspace

# 서비스 재시작
sudo systemctl restart openspace

# 로그 보기 (최근 100줄)
sudo journalctl -u openspace -n 100 --no-pager

# 실시간 로그 추적
sudo journalctl -u openspace -f
```

---

### Nginx 리버스 프록시

#### 전제 조건

```bash
# Ubuntu/Debian
sudo apt install nginx

# RHEL/CentOS/Rocky
sudo dnf install nginx
```

#### HTTP 전용 (개발/내부망)

`/etc/nginx/sites-available/openspace` (또는 `/etc/nginx/conf.d/openspace.conf`):

```nginx
upstream openspace_backend {
    server 127.0.0.1:8000;
}

server {
    listen 80;
    server_name openspace.example.com;

    # 최대 업로드 크기 (xlsx 파일)
    client_max_body_size 50M;

    # 요청 타임아웃
    proxy_read_timeout 300s;
    proxy_send_timeout 300s;

    # 정적 파일 직접 제공 (선택 — 성능 최적화)
    location /static/ {
        alias /opt/openspace/web/static/;
        expires 7d;
        add_header Cache-Control "public, immutable";
    }

    # API 및 페이지
    location / {
        proxy_pass http://openspace_backend;
        proxy_http_version 1.1;

        # WebSocket 지원
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";

        # 원본 클라이언트 정보 전달
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

```bash
# 사이트 활성화 (sites-available 방식)
sudo ln -s /etc/nginx/sites-available/openspace /etc/nginx/sites-enabled/

# 설정 검증
sudo nginx -t

# Nginx 재시작
sudo systemctl reload nginx
```

#### HTTPS (프로덕션 권장)

```nginx
upstream openspace_backend {
    server 127.0.0.1:8000;
}

# HTTP → HTTPS 리다이렉트
server {
    listen 80;
    server_name openspace.example.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl http2;
    server_name openspace.example.com;

    # SSL 인증서
    ssl_certificate     /etc/letsencrypt/live/openspace.example.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/openspace.example.com/privkey.pem;

    # SSL 보안 설정
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384;
    ssl_prefer_server_ciphers off;
    ssl_session_cache shared:SSL:10m;
    ssl_session_timeout 1d;

    # OCSP Stapling
    ssl_stapling on;
    ssl_stapling_verify on;
    ssl_trusted_certificate /etc/letsencrypt/live/openspace.example.com/chain.pem;

    # 최대 업로드 크기
    client_max_body_size 50M;

    # 요청 타임아웃
    proxy_read_timeout 300s;
    proxy_send_timeout 300s;

    # 정적 파일 직접 제공 (선택)
    location /static/ {
        alias /opt/openspace/web/static/;
        expires 30d;
        add_header Cache-Control "public, immutable";
    }

    # API 및 페이지
    location / {
        proxy_pass http://openspace_backend;
        proxy_http_version 1.1;

        # WebSocket 지원
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";

        # 원본 클라이언트 정보 전달
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto https;
    }
}
```

---

### Apache 리버스 프록시

#### 전제 조건

```bash
# Ubuntu/Debian
sudo apt install apache2

# 필수 모듈 활성화
sudo a2enmod proxy proxy_http proxy_wstunnel headers rewrite ssl
sudo systemctl restart apache2
```

```bash
# RHEL/CentOS/Rocky
sudo dnf install httpd mod_ssl

# 모듈 확인 (기본 설치에 포함)
httpd -M | grep -E 'proxy|headers|rewrite|ssl'
```

#### HTTP 전용 (개발/내부망)

`/etc/apache2/sites-available/openspace.conf` (Debian 계열) 또는 `/etc/httpd/conf.d/openspace.conf` (RHEL 계열):

```apache
<VirtualHost *:80>
    ServerName openspace.example.com

    # 최대 업로드 크기 (요청 본문)
    LimitRequestBody 52428800

    # 요청 타임아웃
    ProxyTimeout 300
    Timeout 300

    # 정적 파일 직접 제공 (선택)
    Alias /static/ /opt/openspace/web/static/
    <Directory /opt/openspace/web/static/>
        Require all granted
        Options -Indexes
        <IfModule mod_expires.c>
            ExpiresActive On
            ExpiresDefault "access plus 7 days"
        </IfModule>
    </Directory>

    # WebSocket 프록시 (일반 HTTP 프록시보다 먼저 선언)
    RewriteEngine On
    RewriteCond %{HTTP:Upgrade} websocket [NC]
    RewriteCond %{HTTP:Connection} upgrade [NC]
    RewriteRule ^/ws/(.*) ws://127.0.0.1:8000/ws/$1 [P,L]

    # HTTP 리버스 프록시
    ProxyPreserveHost On
    ProxyPass /static/ !
    ProxyPass / http://127.0.0.1:8000/
    ProxyPassReverse / http://127.0.0.1:8000/

    # 원본 클라이언트 정보 전달
    RequestHeader set X-Forwarded-Proto "http"
    RequestHeader set X-Real-IP "%{REMOTE_ADDR}s"
</VirtualHost>
```

```bash
# Debian 계열: 사이트 활성화
sudo a2ensite openspace.conf
sudo apache2ctl configtest
sudo systemctl reload apache2

# RHEL 계열: 설정 검증 및 적용
sudo httpd -t
sudo systemctl reload httpd
```

#### HTTPS (프로덕션 권장)

```apache
# HTTP → HTTPS 리다이렉트
<VirtualHost *:80>
    ServerName openspace.example.com
    RewriteEngine On
    RewriteRule ^(.*)$ https://%{HTTP_HOST}$1 [R=301,L]
</VirtualHost>

<VirtualHost *:443>
    ServerName openspace.example.com

    # SSL 설정
    SSLEngine On
    SSLCertificateFile    /etc/letsencrypt/live/openspace.example.com/fullchain.pem
    SSLCertificateKeyFile /etc/letsencrypt/live/openspace.example.com/privkey.pem

    # SSL 보안 설정
    SSLProtocol all -SSLv3 -TLSv1 -TLSv1.1
    SSLCipherSuite ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384
    SSLHonorCipherOrder off

    # HSTS 헤더 (Uvicorn이 이미 설정하지만 리버스 프록시에서도 추가 가능)
    Header always set Strict-Transport-Security "max-age=31536000"

    # 최대 업로드 크기
    LimitRequestBody 52428800

    # 요청 타임아웃
    ProxyTimeout 300
    Timeout 300

    # 정적 파일 직접 제공 (선택)
    Alias /static/ /opt/openspace/web/static/
    <Directory /opt/openspace/web/static/>
        Require all granted
        Options -Indexes
        <IfModule mod_expires.c>
            ExpiresActive On
            ExpiresDefault "access plus 30 days"
        </IfModule>
    </Directory>

    # WebSocket 프록시 (반드시 일반 ProxyPass보다 먼저)
    RewriteEngine On
    RewriteCond %{HTTP:Upgrade} websocket [NC]
    RewriteCond %{HTTP:Connection} upgrade [NC]
    RewriteRule ^/ws/(.*) ws://127.0.0.1:8000/ws/$1 [P,L]

    # HTTP 리버스 프록시
    ProxyPreserveHost On
    ProxyPass /static/ !
    ProxyPass / http://127.0.0.1:8000/
    ProxyPassReverse / http://127.0.0.1:8000/

    # 원본 클라이언트 정보 전달
    RequestHeader set X-Forwarded-Proto "https"
    RequestHeader set X-Real-IP "%{REMOTE_ADDR}s"
</VirtualHost>
```

> **Apache WebSocket 주의**: `mod_proxy_wstunnel`이 반드시 활성화되어 있어야 합니다.
> `RewriteRule`의 WebSocket 프록시가 일반 `ProxyPass`보다 먼저 선언되어야 정상 동작합니다.

#### Apache SELinux 설정 (RHEL 계열)

SELinux가 활성화된 환경에서는 Apache가 네트워크 연결을 할 수 있도록 허용해야 합니다.

```bash
# Apache가 백엔드로 연결할 수 있도록 허용
sudo setsebool -P httpd_can_network_connect 1

# 정적 파일 디렉터리 접근 허용
sudo semanage fcontext -a -t httpd_sys_content_t "/opt/openspace/web/static(/.*)?"
sudo restorecon -Rv /opt/openspace/web/static/
```

---

### SSL/TLS 인증서 설정

#### Let's Encrypt (무료, 자동 갱신)

```bash
# certbot 설치
sudo apt install certbot                          # Debian/Ubuntu
sudo dnf install certbot                          # RHEL/CentOS

# Nginx 플러그인
sudo apt install python3-certbot-nginx            # Debian/Ubuntu
sudo certbot --nginx -d openspace.example.com

# Apache 플러그인
sudo apt install python3-certbot-apache           # Debian/Ubuntu
sudo certbot --apache -d openspace.example.com

# 수동 발급 (Standalone — 웹 서버 중지 필요)
sudo certbot certonly --standalone -d openspace.example.com

# 갱신 테스트
sudo certbot renew --dry-run
```

인증서 파일 위치:

```
/etc/letsencrypt/live/openspace.example.com/
├── fullchain.pem   ← ssl_certificate (Nginx) / SSLCertificateFile (Apache)
├── privkey.pem     ← ssl_certificate_key (Nginx) / SSLCertificateKeyFile (Apache)
└── chain.pem       ← ssl_trusted_certificate (Nginx, OCSP Stapling)
```

#### 자체 서명 인증서 (내부망/테스트)

```bash
sudo mkdir -p /etc/ssl/openspace

sudo openssl req -x509 -nodes -days 3650 \
    -newkey rsa:2048 \
    -keyout /etc/ssl/openspace/server.key \
    -out /etc/ssl/openspace/server.crt \
    -subj "/CN=openspace.example.com"
```

Nginx/Apache 설정에서 인증서 경로를 위 파일로 변경하면 됩니다.

> 자체 서명 인증서는 브라우저에서 경고가 표시됩니다. 내부망 전용으로만 사용하세요.

---

## 데이터베이스 관리

### DB 파일 위치

기본 설정: 프로젝트 루트의 `openspace.db` (SQLite WAL 모드)

```bash
# 실제 파일 (WAL 모드에서 3개 파일)
openspace.db         # 메인 DB
openspace.db-wal     # Write-Ahead Log
openspace.db-shm     # 공유 메모리
```

### 백업

```bash
# 온라인 백업 (서비스 중단 없이 안전하게 백업)
sqlite3 openspace.db ".backup /backup/openspace_$(date +%Y%m%d_%H%M%S).db"

# 또는 단순 파일 복사 (서비스 중지 후)
sudo systemctl stop openspace
cp openspace.db openspace.db-wal openspace.db-shm /backup/
sudo systemctl start openspace
```

### 자동 백업 (crontab)

```bash
# 매일 03:00에 백업 (최근 30일 보관)
sudo crontab -e
```

```cron
0 3 * * * sqlite3 /opt/openspace/openspace.db ".backup /backup/openspace/openspace_$(date +\%Y\%m\%d).db" && find /backup/openspace/ -name "*.db" -mtime +30 -delete
```

### 복원

```bash
sudo systemctl stop openspace
cp /backup/openspace_20260227.db /opt/openspace/openspace.db
sudo systemctl start openspace
```

### 마이그레이션

`init_db.py`가 실행 시 모든 마이그레이션을 자동으로 적용합니다 (멱등).

```bash
# 수동 마이그레이션 실행
cd /opt/openspace
source .venv/bin/activate
python3 init_db.py
```

현재 마이그레이션 목록:

| 파일 | 내용 |
|------|------|
| `001_initial.sql` | 전체 스키마 (users, templates, workspaces, cells 등) |
| `002_nullable_template_fk.sql` | workspace FK nullable 변경 |
| `003_sheet_meta.sql` | merges, row_heights, freeze_panes 컬럼 추가 |
| `004_must_change_password.sql` | 비밀번호 변경 강제 플래그 |
| `005_cell_comments.sql` | 셀 코멘트 기능 |
| `006_conditional_formats.sql` | 조건부 서식 |
| `007_col_widths.sql` | 열 너비 저장 |

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
│       ├── lib/                         # 프론트엔드 로컬 번들 (폐쇄망 대응)
│       │   ├── jsuites.css
│       │   ├── jsuites.js
│       │   ├── jspreadsheet.css
│       │   ├── jspreadsheet.js          # Jspreadsheet CE v4
│       │   └── jspreadsheet-formula.js  # @jspreadsheet/formula v2
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
│   ├── 003_sheet_meta.sql
│   └── ...
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
| **스프레드시트 UI** | Jspreadsheet CE (로컬 번들, MIT) |
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

셀 스타일은 `style` TEXT 컬럼에 JSON으로 저장됩니다. 변경된 속성만 포함됩니다.

```json
{
  "fontName": "Arial",
  "bold": true,
  "italic": false,
  "underline": "single",
  "strikethrough": false,
  "fontSize": 14,
  "color": "FF0000",
  "bg": "FFFF00",
  "align": "center",
  "valign": "middle",
  "wrap": true,
  "indent": 2,
  "textRotation": 45,
  "border": {
    "top":    {"style": "thin", "color": "000000"},
    "bottom": {"style": "medium", "color": "000000"},
    "left":   {"style": "hair", "color": "000000"},
    "right":  {"style": "dotted", "color": "000000"}
  },
  "numFmt": "0.00%"
}
```

- **fontName**: Calibri 이외의 폰트명만 저장
- **underline**: 타입 문자열 보존 (`single`, `double`, `singleAccounting`, `doubleAccounting`)
- **색상**: 6자리 RGB hex. import 시 theme/indexed/rgb 모든 타입을 RGB로 변환
- **indent**: Excel 들여쓰기 레벨 (정수)
- **textRotation**: Excel 텍스트 회전 각도 (0-180)

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
| XSS | Content-Security-Policy 헤더 (외부 도메인 없음), Jinja2 자동 이스케이프 |
| 클릭재킹 | X-Frame-Options: DENY |
| RBAC | 서버에서 강제 적용, 클라이언트 값 신뢰 안 함 |
| ADMIN 제한 | ADMIN은 SUPER_ADMIN 계정을 관리할 수 없음 |
| HSTS | `DEBUG=false` 시 Strict-Transport-Security 헤더 자동 추가 |
| CSP | `script-src 'self'`, 외부 도메인 차단 |

---

## 운영 주의사항

1. **workers=1 필수** — SQLite WAL + 인메모리 WSHub는 단일 프로세스에서만 동작합니다.
2. **SECRET_KEY, KEK_KEY 변경** — `.env.example`의 기본값은 개발용입니다. 프로덕션에서 반드시 64자 난수 hex로 교체하세요.
3. **HTTPS 필수 (프로덕션)** — `DEBUG=false` 설정 시 쿠키에 `Secure` 플래그가 추가됩니다. Nginx/Apache 리버스 프록시로 TLS를 제공해야 합니다.
4. **`.env` git 제외** — `.gitignore`에 `.env`를 추가하세요.
5. **DB 백업** — `openspace.db` 파일을 정기 백업하세요 ([데이터베이스 관리](#데이터베이스-관리) 참조).
6. **폐쇄망(오프라인) 완전 지원** — 모든 프론트엔드 라이브러리가 `web/static/lib/`에 로컬 번들링되어 있어 외부 인터넷 연결이 필요 없습니다. CDN 의존성이 없으므로 폐쇄망 환경에서도 즉시 사용 가능합니다.
7. **파일 업로드 크기** — `MAX_UPLOAD_SIZE` 환경 변수(기본 50MB)와 리버스 프록시의 `client_max_body_size`(Nginx) 또는 `LimitRequestBody`(Apache) 값을 일치시키세요.

---

## 문제 해결

### 서버가 시작되지 않음

```bash
# 포트 충돌 확인
ss -tlnp | grep 8000

# Python 버전 확인 (3.10+ 필요)
python3 --version

# 의존성 설치 확인
pip list | grep -E 'fastapi|uvicorn|sqlalchemy'
```

### WebSocket 연결 실패

- **Nginx**: `proxy_http_version 1.1`, `Upgrade`, `Connection` 헤더가 올바르게 설정되었는지 확인
- **Apache**: `mod_proxy_wstunnel`이 활성화되었는지 확인 (`a2enmod proxy_wstunnel`)
- **방화벽**: 443(HTTPS) 또는 80(HTTP) 포트가 열려 있는지 확인
- **브라우저 콘솔**: WebSocket 연결 URL이 `ws://` (HTTP) 또는 `wss://` (HTTPS)인지 확인

### CSRF 토큰 오류 (403)

- 브라우저 쿠키가 활성화되어 있는지 확인
- 리버스 프록시에서 쿠키 헤더가 올바르게 전달되는지 확인
- `ProxyPreserveHost On` (Apache) 또는 `proxy_set_header Host $host` (Nginx) 설정 확인

### xlsx 업로드 실패

- 파일 크기가 `MAX_UPLOAD_SIZE` 이하인지 확인
- 리버스 프록시의 업로드 크기 제한 확인:
  - Nginx: `client_max_body_size`
  - Apache: `LimitRequestBody`

### 세션이 유지되지 않음

- `DEBUG=false`인 경우 HTTPS가 필요합니다 (쿠키에 `Secure` 플래그)
- `SESSION_TTL_SECONDS` 값 확인 (기본 8시간)
- 리버스 프록시에서 `Host` 헤더가 올바르게 전달되는지 확인

### DB 잠금 오류 (database is locked)

- `--workers 1` 옵션으로 실행 중인지 확인
- 다른 프로세스가 DB에 접근하고 있지 않은지 확인
- WAL 모드 확인: `sqlite3 openspace.db "PRAGMA journal_mode;"`

### 로그 확인

```bash
# systemd 서비스 로그
sudo journalctl -u openspace -f

# start.sh로 실행 시 — 터미널에 직접 출력

# Nginx 오류 로그
sudo tail -f /var/log/nginx/error.log

# Apache 오류 로그
sudo tail -f /var/log/apache2/error.log     # Debian
sudo tail -f /var/log/httpd/error_log       # RHEL
```

---

## 라이선스

Copyright (c) 2026 JAEHYUK CHO

이 프로젝트는 [MIT License](LICENSE)로 배포됩니다.

서드파티 의존성의 라이선스 정보는 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)를 참고하세요.
