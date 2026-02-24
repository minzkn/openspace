import uuid
import secrets
from datetime import datetime, timezone, timedelta
from typing import Optional

from fastapi import Request, HTTPException, status, Depends
from sqlalchemy.orm import Session as DBSession

from .config import settings
from .models import Session, User
from .database import get_db


def _now_str() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def _expires_str() -> str:
    dt = datetime.now(timezone.utc) + timedelta(seconds=settings.session_ttl_seconds)
    return dt.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def create_session(db: DBSession, user_id: str, ip: Optional[str], ua: Optional[str]) -> str:
    session_id = str(uuid.uuid4())
    sess = Session(
        id=session_id,
        user_id=user_id,
        expires_at=_expires_str(),
        last_seen=_now_str(),
        ip_address=ip,
        user_agent=ua,
    )
    db.add(sess)
    db.commit()
    return session_id


def get_session(db: DBSession, session_id: str) -> Optional[Session]:
    now = _now_str()
    sess = db.query(Session).filter(
        Session.id == session_id,
        Session.expires_at > now,
    ).first()
    if not sess:
        return None
    # update last_seen
    sess.last_seen = now
    db.commit()
    return sess


def delete_session(db: DBSession, session_id: str):
    db.query(Session).filter(Session.id == session_id).delete()
    db.commit()


def cleanup_expired_sessions(db: DBSession):
    now = _now_str()
    db.query(Session).filter(Session.expires_at <= now).delete()
    db.commit()


# ------------------------------------------------------------------
# FastAPI dependencies
# ------------------------------------------------------------------

def get_current_user(
    request: Request,
    db: DBSession = Depends(get_db),
) -> User:
    session_id = request.cookies.get("session_id")
    if not session_id:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Not authenticated")
    sess = get_session(db, session_id)
    if not sess:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Session expired")
    user = db.query(User).filter(User.id == sess.user_id, User.is_active == 1).first()
    if not user:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="User not found")
    return user


def get_current_user_optional(
    request: Request,
    db: DBSession = Depends(get_db),
) -> Optional[User]:
    try:
        return get_current_user(request, db)
    except HTTPException:
        return None


def generate_csrf_token() -> str:
    return secrets.token_hex(32)
