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
    comment: Optional[str] = None  # 셀 메모 (선택)


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
            if p.comment is not None:
                existing.comment = p.comment if p.comment else None
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
                comment=p.comment if p.comment else None,
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
        applied.append({"row": p.row, "col": p.col, "value": p.value, "style": p.style, "comment": p.comment})

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

    # 2D 배열 + 스타일 맵 + 메모 맵
    grid = [[""] * num_cols for _ in range(num_rows)]
    styles: dict = {}
    comments: dict = {}
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
            if c.comment:
                cell_name = f"{get_column_letter(c.col_index + 1)}{c.row_index + 1}"
                comments[cell_name] = c.comment

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

    # 조건부 서식
    conditional_formats = []
    if ws_sheet.conditional_formats:
        try:
            conditional_formats = json.loads(ws_sheet.conditional_formats)
        except Exception:
            pass

    return {
        "data": {
            "cells": grid,
            "num_rows": num_rows,
            "num_cols": num_cols,
            "merges": merges,
            "row_heights": row_heights_px,
            "freeze_columns": freeze_columns,
            "styles": styles,
            "comments": comments,
            "conditional_formats": conditional_formats,
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
                {"row": p.row, "col": p.col, "value": p.value, "style": p.style, "comment": p.comment}
                for p in body.patches
            ],
            "updated_by": current_user.username,
        }
        asyncio.create_task(hub.broadcast(workspace_id, msg))
    return {"message": "ok", "applied": count}


# ── Row insert / delete ──────────────────────────────────────

class RowInsertRequest(BaseModel):
    row_index: int
    count: int = 1
    direction: str = "above"  # "above" or "below"


class RowDeleteRequest(BaseModel):
    row_indices: list[int]


@router.post("/{workspace_id}/sheets/{sheet_id}/rows/insert")
async def insert_rows(
    workspace_id: str,
    sheet_id: str,
    body: RowInsertRequest,
    current_user: User = Depends(require_user),
    db: DBSession = Depends(get_db),
):
    ws = db.query(Workspace).filter(Workspace.id == workspace_id).first()
    if not ws:
        raise HTTPException(status_code=404, detail="Workspace not found")
    if ws.status == "CLOSED" and not is_admin_or_above(current_user):
        raise HTTPException(status_code=403, detail="Workspace is closed")

    ws_sheet = db.query(WorkspaceSheet).filter(
        WorkspaceSheet.id == sheet_id,
        WorkspaceSheet.workspace_id == workspace_id,
    ).first()
    if not ws_sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")

    insert_at = body.row_index
    count = max(1, min(body.count, 100))  # limit to 100 rows at a time

    # Shift existing cells down
    cells_to_shift = db.query(WorkspaceCell).filter(
        WorkspaceCell.sheet_id == sheet_id,
        WorkspaceCell.row_index >= insert_at,
    ).order_by(WorkspaceCell.row_index.desc()).all()

    for cell in cells_to_shift:
        cell.row_index += count

    # Update merges JSON
    if ws_sheet.merges:
        try:
            merges = json.loads(ws_sheet.merges)
            updated = _shift_merges(merges, insert_at, count)
            ws_sheet.merges = json.dumps(updated)
        except Exception:
            pass

    # Update row_heights JSON
    if ws_sheet.row_heights:
        try:
            rh = json.loads(ws_sheet.row_heights)
            new_rh = {}
            for k, v in rh.items():
                ri = int(k)
                if ri >= insert_at:
                    new_rh[str(ri + count)] = v
                else:
                    new_rh[k] = v
            ws_sheet.row_heights = json.dumps(new_rh)
        except Exception:
            pass

    db.commit()

    # Broadcast to other clients
    import asyncio
    asyncio.create_task(hub.broadcast(workspace_id, {
        "type": "row_insert",
        "sheet_id": sheet_id,
        "row_index": insert_at,
        "count": count,
        "updated_by": current_user.username,
    }, exclude=None))

    return {"message": "rows inserted", "row_index": insert_at, "count": count}


@router.post("/{workspace_id}/sheets/{sheet_id}/rows/delete")
async def delete_rows(
    workspace_id: str,
    sheet_id: str,
    body: RowDeleteRequest,
    current_user: User = Depends(require_user),
    db: DBSession = Depends(get_db),
):
    ws = db.query(Workspace).filter(Workspace.id == workspace_id).first()
    if not ws:
        raise HTTPException(status_code=404, detail="Workspace not found")
    if ws.status == "CLOSED" and not is_admin_or_above(current_user):
        raise HTTPException(status_code=403, detail="Workspace is closed")

    ws_sheet = db.query(WorkspaceSheet).filter(
        WorkspaceSheet.id == sheet_id,
        WorkspaceSheet.workspace_id == workspace_id,
    ).first()
    if not ws_sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")

    indices = sorted(set(body.row_indices))
    if not indices:
        return {"message": "no rows to delete"}

    # Delete cells in target rows
    for ri in indices:
        db.query(WorkspaceCell).filter(
            WorkspaceCell.sheet_id == sheet_id,
            WorkspaceCell.row_index == ri,
        ).delete()

    # Shift cells above deleted rows down
    # Process from bottom to top to avoid conflicts
    for i, ri in enumerate(indices):
        shift = i + 1  # How many rows deleted so far (including this one)
        next_ri = indices[i + 1] if i + 1 < len(indices) else MAX_ROWS
        # Cells between ri+1 and next_ri-1 need to shift up by 'shift'
        cells_to_shift = db.query(WorkspaceCell).filter(
            WorkspaceCell.sheet_id == sheet_id,
            WorkspaceCell.row_index > ri,
            WorkspaceCell.row_index < next_ri,
        ).all()
        for cell in cells_to_shift:
            cell.row_index -= shift

    # Handle remaining cells after the last deleted index
    if indices:
        last_idx = indices[-1]
        total_deleted = len(indices)
        remaining = db.query(WorkspaceCell).filter(
            WorkspaceCell.sheet_id == sheet_id,
            WorkspaceCell.row_index > last_idx,
        ).all()
        for cell in remaining:
            cell.row_index -= total_deleted

    # Update merges JSON
    if ws_sheet.merges:
        try:
            merges = json.loads(ws_sheet.merges)
            for ri in reversed(indices):
                merges = _shift_merges(merges, ri, -1)
            ws_sheet.merges = json.dumps(merges)
        except Exception:
            pass

    # Update row_heights JSON
    if ws_sheet.row_heights:
        try:
            rh = json.loads(ws_sheet.row_heights)
            new_rh = {}
            for k, v in rh.items():
                ri = int(k)
                if ri in indices:
                    continue  # skip deleted rows
                offset = sum(1 for d in indices if d < ri)
                new_rh[str(ri - offset)] = v
            ws_sheet.row_heights = json.dumps(new_rh)
        except Exception:
            pass

    db.commit()

    import asyncio
    asyncio.create_task(hub.broadcast(workspace_id, {
        "type": "row_delete",
        "sheet_id": sheet_id,
        "row_indices": indices,
        "updated_by": current_user.username,
    }, exclude=None))

    return {"message": "rows deleted", "count": len(indices)}


def _shift_merges(merges: list[str], at_row: int, shift: int) -> list[str]:
    """Shift merge ranges when rows are inserted/deleted."""
    from openpyxl.utils import range_boundaries
    updated = []
    for rng in merges:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(rng)
            if shift > 0:
                # Insert: shift ranges at or below at_row
                if min_row > at_row:
                    min_row += shift
                elif min_row == at_row:
                    min_row += shift
                if max_row >= at_row:
                    max_row += shift
            else:
                # Delete: remove ranges that overlap deleted row, shift others
                deleted_row = at_row + 1  # 1-based
                if min_row <= deleted_row <= max_row:
                    if max_row - min_row == 0:
                        continue  # entire merge deleted
                    max_row -= 1
                elif deleted_row < min_row:
                    min_row -= 1
                    max_row -= 1
            start = f"{get_column_letter(min_col)}{min_row}"
            end = f"{get_column_letter(max_col)}{max_row}"
            if start != end:
                updated.append(f"{start}:{end}")
        except Exception:
            updated.append(rng)
    return updated
