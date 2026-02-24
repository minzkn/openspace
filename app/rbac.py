from fastapi import HTTPException, status, Depends
from .models import User
from .auth import get_current_user

# Role constants
SUPER_ADMIN = "SUPER_ADMIN"
ADMIN = "ADMIN"
USER = "USER"

ROLE_RANK = {USER: 0, ADMIN: 1, SUPER_ADMIN: 2}


def role_rank(role: str) -> int:
    return ROLE_RANK.get(role, -1)


def is_admin_or_above(user: User) -> bool:
    return role_rank(user.role) >= role_rank(ADMIN)


def can_manage_user(actor: User, target: User) -> bool:
    """actor가 target을 관리할 수 있는지 확인."""
    if actor.role == SUPER_ADMIN:
        return True
    if actor.role == ADMIN:
        # ADMIN은 SUPER_ADMIN 계정을 관리할 수 없음
        return target.role != SUPER_ADMIN
    return False


def require_role(min_role: str):
    """최소 역할을 요구하는 FastAPI dependency factory."""
    def dependency(current_user: User = Depends(get_current_user)) -> User:
        if role_rank(current_user.role) < role_rank(min_role):
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"Requires {min_role} or above",
            )
        return current_user
    return dependency


require_user = require_role(USER)
require_admin = require_role(ADMIN)
require_super_admin = require_role(SUPER_ADMIN)
