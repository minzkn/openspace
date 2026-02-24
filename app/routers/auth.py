import time
from collections import defaultdict

from fastapi import APIRouter, Depends, HTTPException, Request, Response, status
from pydantic import BaseModel, field_validator
from sqlalchemy.orm import Session as DBSession

from ..database import get_db
from ..models import User
from ..auth import (
    create_session, delete_session, get_current_user, generate_csrf_token
)
from ..crypto import crypto
from ..config import settings

router = APIRouter(prefix="/api/auth", tags=["auth"])

# ------------------------------------------------------------------
# 로그인 Rate Limiting (IP당 10회/분 슬라이딩 윈도우)
# ------------------------------------------------------------------
_login_attempts: dict[str, list[float]] = defaultdict(list)
_RATE_WINDOW = 60   # 초
_RATE_MAX = 10      # 최대 시도 횟수


def _check_rate_limit(ip: str) -> None:
    now = time.time()
    cutoff = now - _RATE_WINDOW
    attempts = [t for t in _login_attempts[ip] if t > cutoff]
    _login_attempts[ip] = attempts
    if len(attempts) >= _RATE_MAX:
        raise HTTPException(
            status_code=429,
            detail="Too many login attempts. Please wait before trying again.",
        )
    _login_attempts[ip].append(now)


def _safe_decrypt_email(email):
    """암호화된 이메일은 복호화, 평문은 그대로 반환."""
    if not email:
        return email
    if crypto.is_encrypted(email):
        try:
            return crypto.decrypt(email)
        except Exception:
            return email
    return email


class LoginRequest(BaseModel):
    username: str
    password: str


class PasswordChangeRequest(BaseModel):
    current_password: str
    new_password: str

    @field_validator("new_password")
    @classmethod
    def validate_new_password(cls, v):
        if len(v) < 8:
            raise ValueError("Password must be at least 8 characters")
        return v


@router.post("/login")
async def login(body: LoginRequest, request: Request, response: Response, db: DBSession = Depends(get_db)):
    ip = request.client.host if request.client else "unknown"
    _check_rate_limit(ip)

    user = db.query(User).filter(User.username == body.username, User.is_active == 1).first()
    if not user or not crypto.verify_password(user.password_hash, body.password):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid credentials")

    ua = request.headers.get("user-agent")
    session_id = create_session(db, user.id, ip, ua)
    csrf_token = generate_csrf_token()

    secure = not settings.debug
    response.set_cookie(
        "session_id", session_id,
        httponly=True, secure=secure, samesite="strict",
        max_age=settings.session_ttl_seconds,
    )
    response.set_cookie(
        "csrf_token", csrf_token,
        httponly=False, secure=secure, samesite="strict",
        max_age=settings.session_ttl_seconds,
    )

    return {
        "data": {
            "id": user.id,
            "username": user.username,
            "role": user.role,
            "must_change_password": bool(user.must_change_password),
        },
        "message": "ok",
    }


@router.post("/logout")
async def logout(request: Request, response: Response, db: DBSession = Depends(get_db)):
    session_id = request.cookies.get("session_id")
    if session_id:
        delete_session(db, session_id)
    response.delete_cookie("session_id")
    response.delete_cookie("csrf_token")
    return {"message": "ok"}


@router.get("/me")
async def me(current_user: User = Depends(get_current_user)):
    return {
        "data": {
            "id": current_user.id,
            "username": current_user.username,
            "email": _safe_decrypt_email(current_user.email),
            "role": current_user.role,
            "is_active": current_user.is_active,
            "must_change_password": bool(current_user.must_change_password),
            "created_at": current_user.created_at,
        }
    }


@router.patch("/me/password")
async def change_password(
    body: PasswordChangeRequest,
    current_user: User = Depends(get_current_user),
    db: DBSession = Depends(get_db),
):
    if not crypto.verify_password(current_user.password_hash, body.current_password):
        raise HTTPException(status_code=400, detail="Current password is incorrect")
    current_user.password_hash = crypto.hash_password(body.new_password)
    current_user.must_change_password = 0  # 비밀번호 변경 완료
    from ..models import _now
    current_user.updated_at = _now()
    db.commit()
    return {"message": "Password changed"}
