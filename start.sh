#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== OpenSpace 시작 ==="

# .env 파일 없으면 예시에서 복사
if [ ! -f .env ]; then
    cp .env.example .env
    echo "[WARN] .env 파일이 없어서 기본값으로 생성했습니다."
    echo "[WARN] SECRET_KEY 와 KEK_KEY 를 반드시 변경하세요!"
fi

# Python venv 생성 시도
if python3 -m venv .venv 2>/dev/null; then
    echo "[INFO] venv 활성화"
    source .venv/bin/activate
    pip install --quiet --upgrade pip
    pip install --quiet -r requirements.txt
else
    echo "[INFO] venv 사용 불가 - 시스템 pip 사용"
    pip3 install --break-system-packages --quiet -r requirements.txt
fi

# DB 초기화 (멱등)
python3 init_db.py

# 서버 시작 (단일 워커 필수: SQLite WAL + 인메모리 WSHub)
echo ""
echo "=== 서버 시작 ==="
echo "URL: http://0.0.0.0:8000"
echo "초기 계정: admin / admin"
echo "API 문서: http://0.0.0.0:8000/api/docs"
echo ""

exec uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1
