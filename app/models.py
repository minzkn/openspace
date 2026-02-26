# SPDX-License-Identifier: MIT
# Copyright (c) 2026 JAEHYUK CHO
import uuid
from datetime import datetime, timezone
from sqlalchemy import (
    Column, String, Integer, Text, ForeignKey,
    CheckConstraint, UniqueConstraint, Index, Boolean
)
from sqlalchemy.orm import relationship
from .database import Base


def _uuid():
    return str(uuid.uuid4())


def _now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


class User(Base):
    __tablename__ = "users"

    id = Column(String, primary_key=True, default=_uuid)
    username = Column(String, unique=True, nullable=False)
    email = Column(Text)  # 암호화 저장
    password_hash = Column(Text, nullable=False)
    role = Column(
        String,
        nullable=False,
        default="USER",
    )
    is_active = Column(Integer, nullable=False, default=1)
    must_change_password = Column(Integer, nullable=False, default=0)  # 최초 로그인 시 비밀번호 변경 유도
    created_at = Column(String, nullable=False, default=_now)
    updated_at = Column(String, nullable=False, default=_now)

    __table_args__ = (
        CheckConstraint("role IN ('SUPER_ADMIN','ADMIN','USER')", name="ck_users_role"),
    )

    sessions = relationship("Session", back_populates="user", cascade="all, delete-orphan")
    extra_values = relationship("UserExtraValue", back_populates="user", cascade="all, delete-orphan")


class UserExtraField(Base):
    __tablename__ = "user_extra_fields"

    id = Column(String, primary_key=True, default=_uuid)
    field_name = Column(String, unique=True, nullable=False)
    label = Column(String, nullable=False)
    field_type = Column(String, nullable=False, default="text")
    is_sensitive = Column(Integer, nullable=False, default=0)
    is_required = Column(Integer, nullable=False, default=0)
    sort_order = Column(Integer, nullable=False, default=0)
    created_at = Column(String, nullable=False, default=_now)

    __table_args__ = (
        CheckConstraint(
            "field_type IN ('text','number','date','boolean')",
            name="ck_uef_type",
        ),
    )

    values = relationship("UserExtraValue", back_populates="field", cascade="all, delete-orphan")


class UserExtraValue(Base):
    __tablename__ = "user_extra_values"

    id = Column(String, primary_key=True, default=_uuid)
    user_id = Column(String, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    field_id = Column(String, ForeignKey("user_extra_fields.id", ondelete="CASCADE"), nullable=False)
    value = Column(Text)

    __table_args__ = (UniqueConstraint("user_id", "field_id"),)

    user = relationship("User", back_populates="extra_values")
    field = relationship("UserExtraField", back_populates="values")


class Template(Base):
    __tablename__ = "templates"

    id = Column(String, primary_key=True, default=_uuid)
    name = Column(String, nullable=False)
    description = Column(Text)
    created_by = Column(String, ForeignKey("users.id"), nullable=False)
    created_at = Column(String, nullable=False, default=_now)
    updated_at = Column(String, nullable=False, default=_now)

    sheets = relationship("TemplateSheet", back_populates="template", cascade="all, delete-orphan",
                          order_by="TemplateSheet.sheet_index")
    creator = relationship("User")


class TemplateSheet(Base):
    __tablename__ = "template_sheets"

    id = Column(String, primary_key=True, default=_uuid)
    template_id = Column(String, ForeignKey("templates.id", ondelete="CASCADE"), nullable=False)
    sheet_index = Column(Integer, nullable=False)
    sheet_name = Column(String, nullable=False)
    merges = Column(Text)       # JSON list of xlsx range strings e.g. ["A1:B2", "C3:D4"]
    row_heights = Column(Text)  # JSON dict {row_index_str: height_pt} e.g. {"0": 30.0}
    col_widths = Column(Text)   # JSON dict {col_index_str: width_px} e.g. {"0": 150}
    freeze_panes = Column(Text) # xlsx freeze panes string e.g. "B2" (freeze col A)
    conditional_formats = Column(Text)  # JSON array of conditional format rules

    __table_args__ = (UniqueConstraint("template_id", "sheet_index"),)

    template = relationship("Template", back_populates="sheets")
    columns = relationship("TemplateColumn", back_populates="sheet", cascade="all, delete-orphan",
                           order_by="TemplateColumn.col_index")
    cells = relationship("TemplateCell", back_populates="sheet", cascade="all, delete-orphan")


class TemplateColumn(Base):
    __tablename__ = "template_columns"

    id = Column(String, primary_key=True, default=_uuid)
    sheet_id = Column(String, ForeignKey("template_sheets.id", ondelete="CASCADE"), nullable=False)
    col_index = Column(Integer, nullable=False)
    col_header = Column(String, nullable=False)
    col_type = Column(String, nullable=False, default="text")
    col_options = Column(Text)  # JSON
    is_readonly = Column(Integer, nullable=False, default=0)
    width = Column(Integer, default=120)

    __table_args__ = (
        UniqueConstraint("sheet_id", "col_index"),
        CheckConstraint(
            "col_type IN ('text','number','date','dropdown','checkbox')",
            name="ck_tc_type",
        ),
    )

    sheet = relationship("TemplateSheet", back_populates="columns")


class TemplateCell(Base):
    __tablename__ = "template_cells"

    id = Column(String, primary_key=True, default=_uuid)
    sheet_id = Column(String, ForeignKey("template_sheets.id", ondelete="CASCADE"), nullable=False)
    row_index = Column(Integer, nullable=False)
    col_index = Column(Integer, nullable=False)
    value = Column(Text)
    formula = Column(Text)
    style = Column(Text)  # JSON
    comment = Column(Text)  # cell note/comment

    __table_args__ = (
        UniqueConstraint("sheet_id", "row_index", "col_index"),
        CheckConstraint("row_index >= 0 AND row_index < 10000", name="ck_tc_row"),
    )

    sheet = relationship("TemplateSheet", back_populates="cells")


class Workspace(Base):
    __tablename__ = "workspaces"

    id = Column(String, primary_key=True, default=_uuid)
    name = Column(String, nullable=False)
    template_id = Column(String, ForeignKey("templates.id", ondelete="SET NULL"), nullable=True)
    status = Column(String, nullable=False, default="OPEN")
    created_by = Column(String, ForeignKey("users.id"), nullable=False)
    closed_by = Column(String, ForeignKey("users.id"))
    closed_at = Column(String)
    created_at = Column(String, nullable=False, default=_now)
    updated_at = Column(String, nullable=False, default=_now)

    __table_args__ = (
        CheckConstraint("status IN ('OPEN','CLOSED')", name="ck_ws_status"),
    )

    template = relationship("Template")
    creator = relationship("User", foreign_keys=[created_by])
    closer = relationship("User", foreign_keys=[closed_by])
    sheets = relationship("WorkspaceSheet", back_populates="workspace", cascade="all, delete-orphan",
                          order_by="WorkspaceSheet.sheet_index")


class WorkspaceSheet(Base):
    __tablename__ = "workspace_sheets"

    id = Column(String, primary_key=True, default=_uuid)
    workspace_id = Column(String, ForeignKey("workspaces.id", ondelete="CASCADE"), nullable=False)
    template_sheet_id = Column(String, ForeignKey("template_sheets.id", ondelete="SET NULL"), nullable=True)
    sheet_index = Column(Integer, nullable=False)
    sheet_name = Column(String, nullable=False)
    merges = Column(Text)       # JSON list of xlsx range strings e.g. ["A1:B2", "C3:D4"]
    row_heights = Column(Text)  # JSON dict {row_index_str: height_pt} e.g. {"0": 30.0}
    col_widths = Column(Text)   # JSON dict {col_index_str: width_px} e.g. {"0": 150}
    freeze_panes = Column(Text) # xlsx freeze panes string e.g. "B2" (freeze col A)
    conditional_formats = Column(Text)  # JSON array of conditional format rules

    __table_args__ = (UniqueConstraint("workspace_id", "sheet_index"),)

    workspace = relationship("Workspace", back_populates="sheets")
    template_sheet = relationship("TemplateSheet")
    cells = relationship("WorkspaceCell", back_populates="sheet", cascade="all, delete-orphan")


class WorkspaceCell(Base):
    __tablename__ = "workspace_cells"

    id = Column(String, primary_key=True, default=_uuid)
    sheet_id = Column(String, ForeignKey("workspace_sheets.id", ondelete="CASCADE"), nullable=False)
    row_index = Column(Integer, nullable=False)
    col_index = Column(Integer, nullable=False)
    value = Column(Text)
    style = Column(Text)  # JSON
    comment = Column(Text)  # cell note/comment
    updated_by = Column(String, ForeignKey("users.id"))
    updated_at = Column(String, nullable=False, default=_now)

    __table_args__ = (
        UniqueConstraint("sheet_id", "row_index", "col_index"),
        Index("idx_ws_cells_sheet", "sheet_id"),
    )

    sheet = relationship("WorkspaceSheet", back_populates="cells")
    updater = relationship("User")


class ChangeLog(Base):
    __tablename__ = "change_logs"

    id = Column(String, primary_key=True, default=_uuid)
    workspace_id = Column(String, ForeignKey("workspaces.id", ondelete="CASCADE"), nullable=False)
    sheet_id = Column(String, nullable=False)
    user_id = Column(String, ForeignKey("users.id"), nullable=False)
    row_index = Column(Integer, nullable=False)
    col_index = Column(Integer, nullable=False)
    old_value = Column(Text)
    new_value = Column(Text)
    changed_at = Column(String, nullable=False, default=_now)

    __table_args__ = (
        Index("idx_cl_workspace", "workspace_id"),
    )


class Session(Base):
    __tablename__ = "sessions"

    id = Column(String, primary_key=True, default=_uuid)
    user_id = Column(String, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    created_at = Column(String, nullable=False, default=_now)
    expires_at = Column(String, nullable=False)
    last_seen = Column(String, nullable=False, default=_now)
    ip_address = Column(String)
    user_agent = Column(String)

    __table_args__ = (
        Index("idx_sessions_user", "user_id"),
        Index("idx_sessions_expires", "expires_at"),
    )

    user = relationship("User", back_populates="sessions")
