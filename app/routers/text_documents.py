# SPDX-License-Identifier: MIT
# Copyright (c) 2026 JAEHYUK CHO
import uuid
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session as DBSession

from ..database import get_db
from ..models import (
    TextDocument, TextDocumentContent, TextDocumentChangeLog,
    User, _now
)
from ..auth import get_current_user
from ..rbac import require_admin, require_user, is_admin_or_above
from ..ws_hub import hub

router = APIRouter(tags=["text-documents"])

VALID_LANGUAGES = {
    "plaintext", "javascript", "python", "html", "css",
    "markdown", "xml", "sql", "json",
}


# ---- Pydantic schemas ----

class DocCreate(BaseModel):
    title: str
    language: str = "plaintext"
    content: str = ""


class DocUpdate(BaseModel):
    title: Optional[str] = None
    language: Optional[str] = None


class DocSave(BaseModel):
    content: str
    base_version: int


class BatchDelete(BaseModel):
    ids: list[str]


# ---- Helpers ----

def _doc_summary(doc: TextDocument) -> dict:
    return {
        "id": doc.id,
        "title": doc.title,
        "language": doc.language,
        "status": doc.status,
        "version": doc.version,
        "created_by": doc.created_by,
        "creator_name": doc.creator.username if doc.creator else None,
        "closed_by": doc.closed_by,
        "closed_at": doc.closed_at,
        "created_at": doc.created_at,
        "updated_at": doc.updated_at,
    }


def _doc_detail(doc: TextDocument, db: DBSession) -> dict:
    result = _doc_summary(doc)
    # 최신 버전 내용 로드
    latest = (
        db.query(TextDocumentContent)
        .filter(TextDocumentContent.document_id == doc.id)
        .order_by(TextDocumentContent.version.desc())
        .first()
    )
    result["content"] = latest.content if latest else ""
    return result


# ================================================================
# Admin endpoints
# ================================================================

@router.get("/api/admin/text-documents")
def admin_list_docs(user: User = Depends(require_admin), db: DBSession = Depends(get_db)):
    docs = db.query(TextDocument).order_by(TextDocument.created_at.desc()).all()
    return {"data": [_doc_summary(d) for d in docs]}


@router.post("/api/admin/text-documents")
def admin_create_doc(body: DocCreate, user: User = Depends(require_admin), db: DBSession = Depends(get_db)):
    if body.language not in VALID_LANGUAGES:
        raise HTTPException(400, f"Invalid language: {body.language}")
    if not body.title.strip():
        raise HTTPException(400, "Title is required")

    now = _now()
    doc_id = str(uuid.uuid4())
    doc = TextDocument(
        id=doc_id,
        title=body.title.strip(),
        language=body.language,
        created_by=user.id,
        created_at=now,
        updated_at=now,
    )
    db.add(doc)

    # 초기 내용 저장
    content_id = str(uuid.uuid4())
    db.add(TextDocumentContent(
        id=content_id,
        document_id=doc_id,
        content=body.content,
        version=1,
        updated_by=user.id,
        updated_at=now,
    ))
    db.commit()
    db.refresh(doc)
    return _doc_summary(doc)


@router.patch("/api/admin/text-documents/{doc_id}")
def admin_update_doc(doc_id: str, body: DocUpdate, user: User = Depends(require_admin), db: DBSession = Depends(get_db)):
    doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
    if not doc:
        raise HTTPException(404, "Document not found")

    if body.title is not None:
        if not body.title.strip():
            raise HTTPException(400, "Title is required")
        doc.title = body.title.strip()
    if body.language is not None:
        if body.language not in VALID_LANGUAGES:
            raise HTTPException(400, f"Invalid language: {body.language}")
        doc.language = body.language

    doc.updated_at = _now()
    db.commit()
    db.refresh(doc)
    return _doc_summary(doc)


@router.delete("/api/admin/text-documents/{doc_id}")
def admin_delete_doc(doc_id: str, user: User = Depends(require_admin), db: DBSession = Depends(get_db)):
    doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
    if not doc:
        raise HTTPException(404, "Document not found")
    db.delete(doc)
    db.commit()
    return {"ok": True}


@router.post("/api/admin/text-documents/batch-delete")
def admin_batch_delete(body: BatchDelete, user: User = Depends(require_admin), db: DBSession = Depends(get_db)):
    deleted = 0
    for doc_id in body.ids:
        doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
        if doc:
            db.delete(doc)
            deleted += 1
    db.commit()
    return {"deleted": deleted}


@router.post("/api/admin/text-documents/{doc_id}/close")
def admin_close_doc(doc_id: str, user: User = Depends(require_admin), db: DBSession = Depends(get_db)):
    doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
    if not doc:
        raise HTTPException(404, "Document not found")
    if doc.status == "CLOSED":
        raise HTTPException(400, "Already closed")

    now = _now()
    doc.status = "CLOSED"
    doc.closed_by = user.id
    doc.closed_at = now
    doc.updated_at = now
    db.commit()
    db.refresh(doc)
    return _doc_summary(doc)


@router.post("/api/admin/text-documents/{doc_id}/reopen")
def admin_reopen_doc(doc_id: str, user: User = Depends(require_admin), db: DBSession = Depends(get_db)):
    doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
    if not doc:
        raise HTTPException(404, "Document not found")
    if doc.status == "OPEN":
        raise HTTPException(400, "Already open")

    doc.status = "OPEN"
    doc.closed_by = None
    doc.closed_at = None
    doc.updated_at = _now()
    db.commit()
    db.refresh(doc)
    return _doc_summary(doc)


# ================================================================
# User endpoints
# ================================================================

@router.get("/api/text-documents")
def list_docs(user: User = Depends(require_user), db: DBSession = Depends(get_db)):
    docs = db.query(TextDocument).order_by(TextDocument.created_at.desc()).all()
    return {"data": [_doc_summary(d) for d in docs]}


@router.get("/api/text-documents/{doc_id}")
def get_doc(doc_id: str, user: User = Depends(require_user), db: DBSession = Depends(get_db)):
    doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
    if not doc:
        raise HTTPException(404, "Document not found")
    return _doc_detail(doc, db)


@router.post("/api/text-documents/{doc_id}/save")
async def save_doc(doc_id: str, body: DocSave, user: User = Depends(require_user), db: DBSession = Depends(get_db)):
    doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
    if not doc:
        raise HTTPException(404, "Document not found")

    # CLOSED 상태에서 일반 사용자 수정 차단
    if doc.status == "CLOSED" and not is_admin_or_above(user):
        raise HTTPException(403, "Document is closed")

    # 버전 충돌 체크
    if body.base_version != doc.version:
        latest = (
            db.query(TextDocumentContent)
            .filter(TextDocumentContent.document_id == doc.id)
            .order_by(TextDocumentContent.version.desc())
            .first()
        )
        raise HTTPException(409, detail={
            "message": "Version conflict",
            "server_version": doc.version,
            "content": latest.content if latest else "",
        })

    now = _now()
    new_version = doc.version + 1

    # 새 내용 저장
    db.add(TextDocumentContent(
        id=str(uuid.uuid4()),
        document_id=doc.id,
        content=body.content,
        version=new_version,
        updated_by=user.id,
        updated_at=now,
    ))

    # 변경 로그
    db.add(TextDocumentChangeLog(
        id=str(uuid.uuid4()),
        document_id=doc.id,
        user_id=user.id,
        version=new_version,
        content_size=len(body.content),
        changed_at=now,
    ))

    # 문서 버전 업데이트
    doc.version = new_version
    doc.updated_at = now
    db.commit()

    # 오래된 버전 정리 (최근 50개만 유지)
    old_contents = (
        db.query(TextDocumentContent)
        .filter(TextDocumentContent.document_id == doc.id)
        .order_by(TextDocumentContent.version.desc())
        .offset(50)
        .all()
    )
    for old in old_contents:
        db.delete(old)
    if old_contents:
        db.commit()

    # WebSocket 브로드캐스트
    room = f"textdoc:{doc.id}"
    await hub.broadcast(room, {
        "type": "doc_updated",
        "version": new_version,
        "content": body.content,
        "updated_by": user.username,
    })

    return {
        "ok": True,
        "version": new_version,
        "updated_at": now,
    }


@router.get("/api/text-documents/{doc_id}/history")
def get_history(doc_id: str, user: User = Depends(require_user), db: DBSession = Depends(get_db)):
    doc = db.query(TextDocument).filter(TextDocument.id == doc_id).first()
    if not doc:
        raise HTTPException(404, "Document not found")

    logs = (
        db.query(TextDocumentChangeLog, User)
        .join(User, TextDocumentChangeLog.user_id == User.id)
        .filter(TextDocumentChangeLog.document_id == doc_id)
        .order_by(TextDocumentChangeLog.changed_at.desc())
        .limit(100)
        .all()
    )
    return {
        "data": [
            {
                "version": log.version,
                "username": u.username,
                "content_size": log.content_size,
                "changed_at": log.changed_at,
            }
            for log, u in logs
        ]
    }
