import uuid
import json
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session as DBSession

from ..database import get_db
from ..models import (
    Workspace, WorkspaceSheet, WorkspaceCell, TemplateSheet, TemplateColumn,
    ChangeLog, _now
)
from ..auth import get_current_user
from ..rbac import require_user, is_admin_or_above
from ..models import User
from ..ws_hub import hub
from .templates import _style_to_css, _range_to_jss, _freeze_to_cols, _pt_to_px
from openpyxl.utils import get_column_letter

router = APIRouter(prefix="/api/workspaces", tags=["cells"])

MAX_ROWS = 10000


class PatchItem(BaseModel):
    row: int
    col: int
    value: Optional[str] = None
    style: Optional[str] = None  # style JSON 문자열 (선택)


class PatchRequest(BaseModel):
    patches: list[PatchItem]


async def _apply_patches(
    workspace_id: str,
    sheet_id: str,
    patches: list[PatchItem],
    current_user: User,
    db: DBSession,
    broadcast: bool = True,
) -> int:
    ws = db.query(Workspace).filter(Workspace.id == workspace_id).first()
    if not ws:
        raise HTTPException(status_code=404, detail="Workspace not found")

    ws_sheet = db.query(WorkspaceSheet).filter(
        WorkspaceSheet.id == sheet_id,
        WorkspaceSheet.workspace_id == workspace_id,
    ).first()
    if not ws_sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")

    # CLOSED 체크
    if ws.status == "CLOSED" and not is_admin_or_above(current_user):
        raise HTTPException(status_code=403, detail="Workspace is closed")

    # 컬럼 readonly 맵
    tmpl_sheet = db.query(TemplateSheet).filter(TemplateSheet.id == ws_sheet.template_sheet_id).first()
    readonly_cols: set[int] = set()
    if tmpl_sheet:
        for c in tmpl_sheet.columns:
            if c.is_readonly:
                readonly_cols.add(c.col_index)

    applied = []
    for p in patches:
        if not (0 <= p.row < MAX_ROWS):
            continue
        if p.col < 0:
            continue
        # readonly 컬럼: 일반 사용자 거부
        if p.col in readonly_cols and not is_admin_or_above(current_user):
            continue

        existing = db.query(WorkspaceCell).filter(
            WorkspaceCell.sheet_id == sheet_id,
            WorkspaceCell.row_index == p.row,
            WorkspaceCell.col_index == p.col,
        ).first()
        old_value = existing.value if existing else None

        if existing:
            if p.value is not None:
                existing.value = p.value
            if p.style is not None:
                existing.style = p.style
            existing.updated_by = current_user.id
            existing.updated_at = _now()
        else:
            db.add(WorkspaceCell(
                id=str(uuid.uuid4()),
                sheet_id=sheet_id,
                row_index=p.row,
                col_index=p.col,
                value=p.value,
                style=p.style,
                updated_by=current_user.id,
                updated_at=_now(),
            ))

        # 변경 이력 (값 변경만)
        if p.value is not None and old_value != p.value:
            db.add(ChangeLog(
                id=str(uuid.uuid4()),
                workspace_id=workspace_id,
                sheet_id=sheet_id,
                user_id=current_user.id,
                row_index=p.row,
                col_index=p.col,
                old_value=old_value,
                new_value=p.value,
            ))
        applied.append({"row": p.row, "col": p.col, "value": p.value, "style": p.style})

    db.commit()
    return len(applied)


@router.get("/{workspace_id}/sheets/{sheet_id}/snapshot")
async def get_snapshot(
    workspace_id: str,
    sheet_id: str,
    current_user: User = Depends(require_user),
    db: DBSession = Depends(get_db),
):
    ws = db.query(Workspace).filter(Workspace.id == workspace_id).first()
    if not ws:
        raise HTTPException(status_code=404, detail="Workspace not found")
    ws_sheet = db.query(WorkspaceSheet).filter(
        WorkspaceSheet.id == sheet_id,
        WorkspaceSheet.workspace_id == workspace_id,
    ).first()
    if not ws_sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")

    cells = db.query(WorkspaceCell).filter(WorkspaceCell.sheet_id == sheet_id).all()

    # 컬럼 메타
    tmpl_sheet = db.query(TemplateSheet).filter(TemplateSheet.id == ws_sheet.template_sheet_id).first()
    num_cols = len(tmpl_sheet.columns) if tmpl_sheet else 0
    max_row = max((c.row_index for c in cells), default=-1) + 1
    num_rows = max(max_row, 100)

    # 2D 배열 + 스타일 맵
    grid = [[""] * num_cols for _ in range(num_rows)]
    styles: dict = {}
    for c in cells:
        if c.row_index < num_rows and c.col_index < num_cols:
            grid[c.row_index][c.col_index] = c.value or ""
            if c.style:
                try:
                    s = json.loads(c.style)
                    if s:
                        cell_name = f"{get_column_letter(c.col_index + 1)}{c.row_index + 1}"
                        styles[cell_name] = _style_to_css(s)
                except Exception:
                    pass

    # 병합 셀: xlsx range strings → jspreadsheet format
    merges: dict = {}
    if ws_sheet.merges:
        try:
            for rng in json.loads(ws_sheet.merges):
                result = _range_to_jss(rng)
                if result:
                    cell_name, dims = result
                    merges[cell_name] = dims
        except Exception:
            pass

    # 행 높이: pt → px
    row_heights_px: dict = {}
    if ws_sheet.row_heights:
        try:
            for k, v in json.loads(ws_sheet.row_heights).items():
                row_heights_px[k] = _pt_to_px(float(v))
        except Exception:
            pass

    # 틀 고정
    freeze_columns = _freeze_to_cols(ws_sheet.freeze_panes)

    return {
        "data": {
            "cells": grid,
            "num_rows": num_rows,
            "num_cols": num_cols,
            "merges": merges,
            "row_heights": row_heights_px,
            "freeze_columns": freeze_columns,
            "styles": styles,
        }
    }


@router.post("/{workspace_id}/sheets/{sheet_id}/patches")
async def http_patches(
    workspace_id: str,
    sheet_id: str,
    body: PatchRequest,
    current_user: User = Depends(require_user),
    db: DBSession = Depends(get_db),
):
    count = await _apply_patches(workspace_id, sheet_id, body.patches, current_user, db)
    if count > 0:
        import asyncio
        msg = {
            "type": "batch_patch",
            "sheet_id": sheet_id,
            "patches": [
                {"row": p.row, "col": p.col, "value": p.value, "style": p.style}
                for p in body.patches
            ],
            "updated_by": current_user.username,
        }
        asyncio.create_task(hub.broadcast(workspace_id, msg))
    return {"message": "ok", "applied": count}
