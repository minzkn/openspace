import uuid
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session as DBSession

from ..database import get_db
from ..models import UserExtraField, _now
from ..rbac import require_admin
from ..models import User

router = APIRouter(prefix="/api/admin/user-fields", tags=["user-fields"])


class FieldCreate(BaseModel):
    field_name: str
    label: str
    field_type: str = "text"
    is_sensitive: int = 0
    is_required: int = 0
    sort_order: int = 0


class FieldUpdate(BaseModel):
    label: Optional[str] = None
    field_type: Optional[str] = None
    is_sensitive: Optional[int] = None
    is_required: Optional[int] = None
    sort_order: Optional[int] = None


class ReorderRequest(BaseModel):
    order: list[str]  # list of field IDs in desired order


def _field_to_dict(f: UserExtraField) -> dict:
    return {
        "id": f.id,
        "field_name": f.field_name,
        "label": f.label,
        "field_type": f.field_type,
        "is_sensitive": f.is_sensitive,
        "is_required": f.is_required,
        "sort_order": f.sort_order,
        "created_at": f.created_at,
    }


@router.get("")
async def list_fields(
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    fields = db.query(UserExtraField).order_by(UserExtraField.sort_order).all()
    return {"data": [_field_to_dict(f) for f in fields]}


@router.post("", status_code=201)
async def create_field(
    body: FieldCreate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    if body.field_type not in ("text", "number", "date", "boolean"):
        raise HTTPException(status_code=400, detail="Invalid field_type")
    if db.query(UserExtraField).filter(UserExtraField.field_name == body.field_name).first():
        raise HTTPException(status_code=409, detail="Field name already exists")
    field = UserExtraField(
        id=str(uuid.uuid4()),
        field_name=body.field_name,
        label=body.label,
        field_type=body.field_type,
        is_sensitive=body.is_sensitive,
        is_required=body.is_required,
        sort_order=body.sort_order,
    )
    db.add(field)
    db.commit()
    db.refresh(field)
    return {"data": _field_to_dict(field), "message": "created"}


@router.patch("/{field_id}")
async def update_field(
    field_id: str,
    body: FieldUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    field = db.query(UserExtraField).filter(UserExtraField.id == field_id).first()
    if not field:
        raise HTTPException(status_code=404, detail="Field not found")
    if body.label is not None:
        field.label = body.label
    if body.field_type is not None:
        if body.field_type not in ("text", "number", "date", "boolean"):
            raise HTTPException(status_code=400, detail="Invalid field_type")
        field.field_type = body.field_type
    if body.is_sensitive is not None:
        field.is_sensitive = body.is_sensitive
    if body.is_required is not None:
        field.is_required = body.is_required
    if body.sort_order is not None:
        field.sort_order = body.sort_order
    db.commit()
    return {"data": _field_to_dict(field), "message": "updated"}


@router.delete("/{field_id}", status_code=204)
async def delete_field(
    field_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    field = db.query(UserExtraField).filter(UserExtraField.id == field_id).first()
    if not field:
        raise HTTPException(status_code=404, detail="Field not found")
    db.delete(field)
    db.commit()


@router.patch("/reorder")
async def reorder_fields(
    body: ReorderRequest,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    for i, fid in enumerate(body.order):
        db.query(UserExtraField).filter(UserExtraField.id == fid).update({"sort_order": i})
    db.commit()
    return {"message": "reordered"}
