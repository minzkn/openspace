import json
import uuid
import logging
from fastapi import APIRouter, WebSocket, WebSocketDisconnect, Depends
from sqlalchemy.orm import Session as DBSession

from ..database import get_db, SessionLocal
from ..models import (
    Workspace, WorkspaceSheet, WorkspaceCell, TemplateSheet,
    ChangeLog, Session as UserSession, User, _now
)
from ..ws_hub import hub
from ..rbac import is_admin_or_above

router = APIRouter()
logger = logging.getLogger(__name__)

MAX_ROWS = 10000


def _get_user_from_session(session_id: str) -> tuple[User | None, DBSession]:
    db = SessionLocal()
    now = _now()
    sess = db.query(UserSession).filter(
        UserSession.id == session_id,
        UserSession.expires_at > now,
    ).first()
    if not sess:
        db.close()
        return None, db
    # WebSocket 연결 동안 세션 만료 방지를 위해 last_seen 갱신
    sess.last_seen = now
    db.commit()
    user = db.query(User).filter(User.id == sess.user_id, User.is_active == 1).first()
    return user, db


@router.websocket("/ws/workspaces/{workspace_id}")
async def ws_endpoint(websocket: WebSocket, workspace_id: str):
    session_id = websocket.query_params.get("session_id")
    if not session_id:
        await websocket.close(code=4001)
        return

    user, db = _get_user_from_session(session_id)
    if not user:
        db.close()
        await websocket.close(code=4001)
        return

    ws = db.query(Workspace).filter(Workspace.id == workspace_id).first()
    if not ws:
        db.close()
        await websocket.close(code=4004)
        return

    await websocket.accept()
    await hub.connect(workspace_id, websocket)

    # 초기 상태 전송
    await websocket.send_json({
        "type": "connected",
        "workspace_status": ws.status,
        "username": user.username,
    })

    try:
        while True:
            raw = await websocket.receive_text()
            try:
                msg = json.loads(raw)
            except json.JSONDecodeError:
                await websocket.send_json({"type": "error", "message": "Invalid JSON"})
                continue

            msg_type = msg.get("type")

            if msg_type == "ping":
                await websocket.send_json({"type": "pong"})
                continue

            if msg_type == "patch":
                await _handle_patch(websocket, workspace_id, msg, user, db, ws)

            elif msg_type == "batch_patch":
                await _handle_batch_patch(websocket, workspace_id, msg, user, db, ws)

            elif msg_type in ("row_insert", "row_delete"):
                # Row operations are handled via REST API
                # WebSocket only receives broadcasts
                pass

            else:
                await websocket.send_json({"type": "error", "message": f"Unknown type: {msg_type}"})

    except WebSocketDisconnect:
        logger.debug(f"WS disconnected: user={user.username} workspace={workspace_id}")
    except Exception as e:
        logger.exception(f"WS error: {e}")
    finally:
        await hub.disconnect(workspace_id, websocket)
        db.close()


async def _handle_patch(websocket, workspace_id, msg, user, db, ws_obj):
    sheet_id = msg.get("sheet_id")
    row = msg.get("row")
    col = msg.get("col")
    value = msg.get("value")
    style = msg.get("style")  # style JSON 문자열 (선택)

    if sheet_id is None or row is None or col is None:
        await websocket.send_json({"type": "error", "message": "Missing fields"})
        return

    # Refresh workspace status
    db.refresh(ws_obj)
    if ws_obj.status == "CLOSED" and not is_admin_or_above(user):
        await websocket.send_json({"type": "error", "message": "Workspace is closed"})
        return

    ws_sheet = db.query(WorkspaceSheet).filter(
        WorkspaceSheet.id == sheet_id,
        WorkspaceSheet.workspace_id == workspace_id,
    ).first()
    if not ws_sheet:
        await websocket.send_json({"type": "error", "message": "Sheet not found"})
        return

    # readonly 체크
    tmpl_sheet = db.query(TemplateSheet).filter(TemplateSheet.id == ws_sheet.template_sheet_id).first()
    if tmpl_sheet:
        for c in tmpl_sheet.columns:
            if c.col_index == col and c.is_readonly and not is_admin_or_above(user):
                await websocket.send_json({"type": "error", "message": "Column is readonly"})
                return

    if not (0 <= row < MAX_ROWS):
        await websocket.send_json({"type": "error", "message": "Row out of range"})
        return

    existing = db.query(WorkspaceCell).filter(
        WorkspaceCell.sheet_id == sheet_id,
        WorkspaceCell.row_index == row,
        WorkspaceCell.col_index == col,
    ).first()
    old_value = existing.value if existing else None

    if existing:
        if value is not None:
            existing.value = value
        if style is not None:
            existing.style = style
        existing.updated_by = user.id
        existing.updated_at = _now()
    else:
        db.add(WorkspaceCell(
            id=str(uuid.uuid4()),
            sheet_id=sheet_id,
            row_index=row,
            col_index=col,
            value=value,
            style=style,
            updated_by=user.id,
            updated_at=_now(),
        ))

    if value is not None and old_value != value:
        db.add(ChangeLog(
            id=str(uuid.uuid4()),
            workspace_id=workspace_id,
            sheet_id=sheet_id,
            user_id=user.id,
            row_index=row,
            col_index=col,
            old_value=old_value,
            new_value=value,
        ))
    db.commit()

    comment = msg.get("comment")  # 셀 메모 (선택)

    # Save comment if provided
    if comment is not None:
        if existing:
            existing.comment = comment if comment else None
        # (for new cells, comment is included in WorkspaceCell creation above)

    broadcast_msg = {
        "type": "patch",
        "sheet_id": sheet_id,
        "row": row,
        "col": col,
        "value": value,
        "style": style,
        "comment": comment,
        "updated_by": user.username,
    }
    await hub.broadcast(workspace_id, broadcast_msg, exclude=websocket)


async def _handle_batch_patch(websocket, workspace_id, msg, user, db, ws_obj):
    sheet_id = msg.get("sheet_id")
    patches = msg.get("patches", [])

    if not sheet_id or not patches:
        await websocket.send_json({"type": "error", "message": "Missing fields"})
        return

    db.refresh(ws_obj)
    if ws_obj.status == "CLOSED" and not is_admin_or_above(user):
        await websocket.send_json({"type": "error", "message": "Workspace is closed"})
        return

    ws_sheet = db.query(WorkspaceSheet).filter(
        WorkspaceSheet.id == sheet_id,
        WorkspaceSheet.workspace_id == workspace_id,
    ).first()
    if not ws_sheet:
        await websocket.send_json({"type": "error", "message": "Sheet not found"})
        return

    tmpl_sheet = db.query(TemplateSheet).filter(TemplateSheet.id == ws_sheet.template_sheet_id).first()
    readonly_cols: set[int] = set()
    if tmpl_sheet:
        for c in tmpl_sheet.columns:
            if c.is_readonly:
                readonly_cols.add(c.col_index)

    applied = []
    for p in patches:
        row = p.get("row")
        col = p.get("col")
        value = p.get("value")
        style = p.get("style")  # style JSON 문자열 (선택)
        comment = p.get("comment")  # 셀 메모 (선택)
        if row is None or col is None:
            continue
        if not (0 <= row < MAX_ROWS):
            continue
        if col in readonly_cols and not is_admin_or_above(user):
            continue

        existing = db.query(WorkspaceCell).filter(
            WorkspaceCell.sheet_id == sheet_id,
            WorkspaceCell.row_index == row,
            WorkspaceCell.col_index == col,
        ).first()
        old_value = existing.value if existing else None

        if existing:
            if value is not None:
                existing.value = value
            if style is not None:
                existing.style = style
            if comment is not None:
                existing.comment = comment if comment else None
            existing.updated_by = user.id
            existing.updated_at = _now()
        else:
            db.add(WorkspaceCell(
                id=str(uuid.uuid4()),
                sheet_id=sheet_id,
                row_index=row,
                col_index=col,
                value=value,
                style=style,
                comment=comment if comment else None,
                updated_by=user.id,
                updated_at=_now(),
            ))

        if value is not None and old_value != value:
            db.add(ChangeLog(
                id=str(uuid.uuid4()),
                workspace_id=workspace_id,
                sheet_id=sheet_id,
                user_id=user.id,
                row_index=row,
                col_index=col,
                old_value=old_value,
                new_value=value,
            ))
        applied.append({"row": row, "col": col, "value": value, "style": style, "comment": comment})

    db.commit()

    if applied:
        broadcast_msg = {
            "type": "batch_patch",
            "sheet_id": sheet_id,
            "patches": applied,
            "updated_by": user.username,
        }
        await hub.broadcast(workspace_id, broadcast_msg, exclude=websocket)
