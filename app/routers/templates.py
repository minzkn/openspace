# SPDX-License-Identifier: MIT
# Copyright (c) 2026 JAEHYUK CHO
import io
import json
import re
import uuid
from typing import Optional
from urllib.parse import quote as url_quote
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from sqlalchemy.orm import Session as DBSession
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection

from ..database import get_db
from ..models import (
    Template, TemplateSheet, TemplateColumn, TemplateCell, _now
)
from ..rbac import require_admin
from ..models import User

router = APIRouter(prefix="/api/admin/templates", tags=["templates"])

MAX_SHEETS = 64
MAX_ROWS = 10000


# ── Pydantic models ──────────────────────────────────────────
class TemplateCreate(BaseModel):
    name: str
    description: Optional[str] = None


class TemplateUpdate(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None


class SheetCreate(BaseModel):
    sheet_name: str = "새 시트"


class SheetUpdate(BaseModel):
    sheet_name: Optional[str] = None
    sheet_index: Optional[int] = None


class ColumnCreate(BaseModel):
    col_header: str = "새 컬럼"
    col_type: str = "text"
    is_readonly: int = 0
    width: int = 120


class ColumnUpdate(BaseModel):
    col_header: Optional[str] = None
    col_type: Optional[str] = None
    col_options: Optional[str] = None
    is_readonly: Optional[int] = None
    width: Optional[int] = None


class CellBatchItem(BaseModel):
    row_index: int
    col_index: int
    value: Optional[str] = None
    style: Optional[str] = None
    comment: Optional[str] = None


class MergesUpdate(BaseModel):
    merges: list[str]   # list of xlsx range strings e.g. ["A1:B2", "C3:D4"]


# ── Style helpers ─────────────────────────────────────────────
def _extract_cell_style(cell) -> Optional[str]:
    """openpyxl cell → style JSON 문자열 (변경된 속성만)"""
    style: dict = {}

    # Font
    font = cell.font
    if font:
        if font.bold:
            style['bold'] = True
        if font.italic:
            style['italic'] = True
        if font.underline and font.underline != 'none':
            style['underline'] = True
        if font.strikethrough:
            style['strikethrough'] = True
        if font.size and font.size != 11:
            style['fontSize'] = font.size
        try:
            if font.color and font.color.type == 'rgb':
                rgb = font.color.rgb or ''
                if len(rgb) >= 6 and rgb.upper() not in ('FF000000', '000000'):
                    style['color'] = rgb[-6:].upper()
        except Exception:
            pass

    # Fill
    fill = cell.fill
    if fill and getattr(fill, 'fill_type', None) and fill.fill_type != 'none':
        try:
            fg = fill.fgColor
            if fg and fg.type == 'rgb':
                rgb = fg.rgb or ''
                if len(rgb) >= 6 and rgb.upper() not in ('00000000', 'FFFFFFFF', 'FF000000'):
                    style['bg'] = rgb[-6:].upper()
        except Exception:
            pass

    # Alignment
    align = cell.alignment
    if align:
        if align.horizontal and align.horizontal not in ('general', None):
            style['align'] = align.horizontal
        if align.vertical and align.vertical not in ('bottom', None):
            style['valign'] = align.vertical
        if align.wrap_text:
            style['wrap'] = True

    # Border
    border = cell.border
    if border:
        borders: dict = {}
        for side_name in ('top', 'bottom', 'left', 'right'):
            side = getattr(border, side_name, None)
            if side and side.border_style and side.border_style != 'none':
                entry: dict = {'style': side.border_style}
                try:
                    if side.color and side.color.type == 'rgb':
                        entry['color'] = side.color.rgb[-6:].upper()
                except Exception:
                    entry['color'] = '000000'
                borders[side_name] = entry
        if borders:
            style['border'] = borders

    # Number format
    num_fmt = cell.number_format
    if num_fmt and num_fmt != 'General':
        style['numFmt'] = num_fmt

    return json.dumps(style, ensure_ascii=False) if style else None


def _apply_cell_style(ws_cell, style_json: Optional[str]) -> None:
    """style JSON 문자열 → openpyxl cell 스타일 적용"""
    if not style_json:
        return
    try:
        style = json.loads(style_json)
    except Exception:
        return
    if not style:
        return

    def _to_argb(hex6: str) -> str:
        """6자리 RGB hex → 8자리 ARGB hex (불투명)"""
        return ('FF' + hex6) if len(hex6) == 6 else hex6

    # Font
    font_kwargs: dict = {}
    if style.get('bold'):
        font_kwargs['bold'] = True
    if style.get('italic'):
        font_kwargs['italic'] = True
    if style.get('underline'):
        font_kwargs['underline'] = 'single'
    if style.get('strikethrough'):
        font_kwargs['strike'] = True
    if style.get('fontSize'):
        font_kwargs['size'] = style['fontSize']
    if style.get('color'):
        font_kwargs['color'] = _to_argb(style['color'])
    if font_kwargs:
        ws_cell.font = Font(**font_kwargs)

    # Fill
    if style.get('bg'):
        ws_cell.fill = PatternFill(fill_type='solid', fgColor=_to_argb(style['bg']))

    # Alignment
    align_kwargs: dict = {}
    if style.get('align'):
        align_kwargs['horizontal'] = style['align']
    if style.get('valign'):
        # CSS 'middle' → Excel 'center' 변환
        va = style['valign']
        if va == 'middle':
            va = 'center'
        align_kwargs['vertical'] = va
    if style.get('wrap'):
        align_kwargs['wrap_text'] = True
    if align_kwargs:
        ws_cell.alignment = Alignment(**align_kwargs)

    # Border
    if style.get('border'):
        sides: dict = {}
        for side_name, bd in style['border'].items():
            color = _to_argb(bd.get('color', '000000'))
            sides[side_name] = Side(border_style=bd.get('style', 'thin'), color=color)
        ws_cell.border = Border(**sides)

    # Number format
    if style.get('numFmt'):
        ws_cell.number_format = style['numFmt']


def _style_to_css(style: dict) -> str:
    """style dict → CSS 문자열"""
    parts = []
    if style.get('bold'):
        parts.append('font-weight:bold')
    if style.get('italic'):
        parts.append('font-style:italic')
    decorations = []
    if style.get('underline'):
        decorations.append('underline')
    if style.get('strikethrough'):
        decorations.append('line-through')
    if decorations:
        parts.append('text-decoration:' + ' '.join(decorations))
    if style.get('fontSize'):
        parts.append(f"font-size:{style['fontSize']}pt")
    if style.get('color'):
        parts.append(f"color:#{style['color']}")
    if style.get('bg'):
        parts.append(f"background-color:#{style['bg']}")
    if style.get('align'):
        parts.append(f"text-align:{style['align']}")
    if style.get('valign'):
        # Excel 'center' → CSS 'middle' 변환
        va = style['valign']
        if va == 'center':
            va = 'middle'
        parts.append(f"vertical-align:{va}")
    if style.get('wrap'):
        parts.append('white-space:pre-wrap')
    if style.get('border'):
        width_map = {'thin': '1px', 'medium': '2px', 'thick': '3px',
                     'dashed': '1px', 'dotted': '1px', 'double': '3px'}
        style_map = {'thin': 'solid', 'medium': 'solid', 'thick': 'solid',
                     'dashed': 'dashed', 'dotted': 'dotted', 'double': 'double'}
        for side, bd in style['border'].items():
            bs = bd.get('style', 'thin')
            color = bd.get('color', '000000')
            w = width_map.get(bs, '1px')
            s = style_map.get(bs, 'solid')
            parts.append(f"border-{side}:{w} {s} #{color}")
    if style.get('numFmt'):
        parts.append(f"--num-fmt:{url_quote(style['numFmt'], safe='')}")
    return ';'.join(parts)


def _range_to_jss(range_str: str) -> Optional[tuple]:
    """xlsx 병합 범위 "A1:B2" → (cell_name, [colspan, rowspan])"""
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_str)
        cell_name = f"{get_column_letter(min_col)}{min_row}"
        colspan = max_col - min_col + 1
        rowspan = max_row - min_row + 1
        return cell_name, [colspan, rowspan]
    except Exception:
        return None


def _freeze_to_cols(freeze_str: Optional[str]) -> int:
    """xlsx freeze_panes 문자열 → 고정 열 수"""
    if not freeze_str:
        return 0
    m = re.match(r'^([A-Z]+)\d+$', freeze_str.upper())
    if m:
        return column_index_from_string(m.group(1)) - 1
    return 0


def _pt_to_px(pt: float) -> int:
    """포인트 → 픽셀 (96dpi 기준)"""
    return round(pt * 96 / 72)


# ── Helpers ──────────────────────────────────────────────────
def _template_summary(t: Template) -> dict:
    return {
        "id": t.id,
        "name": t.name,
        "description": t.description,
        "created_by": t.created_by,
        "created_at": t.created_at,
        "updated_at": t.updated_at,
        "sheet_count": len(t.sheets),
    }


def _col_dict(c: TemplateColumn) -> dict:
    return {
        "id": c.id,
        "col_index": c.col_index,
        "col_header": c.col_header,
        "col_type": c.col_type,
        "col_options": json.loads(c.col_options) if c.col_options else None,
        "is_readonly": c.is_readonly,
        "width": c.width,
    }


def _sheet_detail(sheet: TemplateSheet) -> dict:
    return {
        "id": sheet.id,
        "sheet_index": sheet.sheet_index,
        "sheet_name": sheet.sheet_name,
        "columns": [_col_dict(c) for c in sheet.columns],
    }


def _default_columns(sheet_id: str):
    return [
        TemplateColumn(
            id=str(uuid.uuid4()), sheet_id=sheet_id,
            col_index=i, col_header=h,
        )
        for i, h in enumerate(["A", "B", "C", "D", "E"])
    ]


# ── Template CRUD ─────────────────────────────────────────────
@router.get("")
async def list_templates(
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    templates = db.query(Template).order_by(Template.created_at.desc()).all()
    return {"data": [_template_summary(t) for t in templates]}


@router.post("", status_code=201)
async def create_template(
    body: TemplateCreate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    t = Template(id=str(uuid.uuid4()), name=body.name, description=body.description,
                 created_by=current_user.id)
    db.add(t)
    sheet = TemplateSheet(id=str(uuid.uuid4()), template_id=t.id, sheet_index=0, sheet_name="Sheet1")
    db.add(sheet)
    for col in _default_columns(sheet.id):
        db.add(col)
    db.commit()
    db.refresh(t)
    return {"data": _template_summary(t), "message": "created"}


@router.get("/{template_id}")
async def get_template(
    template_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    t = db.query(Template).filter(Template.id == template_id).first()
    if not t:
        raise HTTPException(status_code=404, detail="Template not found")
    return {"data": {**_template_summary(t), "sheets": [_sheet_detail(s) for s in t.sheets]}}


@router.patch("/{template_id}")
async def update_template(
    template_id: str,
    body: TemplateUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    t = db.query(Template).filter(Template.id == template_id).first()
    if not t:
        raise HTTPException(status_code=404, detail="Template not found")
    if body.name is not None:
        t.name = body.name
    if body.description is not None:
        t.description = body.description
    t.updated_at = _now()
    db.commit()
    return {"data": _template_summary(t), "message": "updated"}


@router.delete("/{template_id}", status_code=204)
async def delete_template(
    template_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    t = db.query(Template).filter(Template.id == template_id).first()
    if not t:
        raise HTTPException(status_code=404, detail="Template not found")
    db.delete(t)
    db.commit()


@router.post("/{template_id}/copy", status_code=201)
async def copy_template(
    template_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    src = db.query(Template).filter(Template.id == template_id).first()
    if not src:
        raise HTTPException(status_code=404, detail="Template not found")
    new_t = Template(id=str(uuid.uuid4()), name=f"{src.name} (copy)",
                     description=src.description, created_by=current_user.id)
    db.add(new_t)
    for sheet in src.sheets:
        new_sheet = TemplateSheet(
            id=str(uuid.uuid4()), template_id=new_t.id,
            sheet_index=sheet.sheet_index, sheet_name=sheet.sheet_name,
            merges=sheet.merges, row_heights=sheet.row_heights,
            freeze_panes=sheet.freeze_panes,
            conditional_formats=sheet.conditional_formats,
        )
        db.add(new_sheet)
        for col in sheet.columns:
            db.add(TemplateColumn(id=str(uuid.uuid4()), sheet_id=new_sheet.id,
                                  col_index=col.col_index, col_header=col.col_header,
                                  col_type=col.col_type, col_options=col.col_options,
                                  is_readonly=col.is_readonly, width=col.width))
        for cell in sheet.cells:
            db.add(TemplateCell(id=str(uuid.uuid4()), sheet_id=new_sheet.id,
                                row_index=cell.row_index, col_index=cell.col_index,
                                value=cell.value, formula=cell.formula, style=cell.style,
                                comment=cell.comment))
    db.commit()
    db.refresh(new_t)
    return {"data": _template_summary(new_t), "message": "copied"}


# ── Sheet CRUD ────────────────────────────────────────────────
@router.get("/{template_id}/sheets/{sheet_id}/snapshot")
async def get_template_sheet_snapshot(
    template_id: str,
    sheet_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    """서식 편집기용 셀 스냅샷 반환. 병합·스타일·행높이·틀고정 포함."""
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")

    cols = sorted(sheet.columns, key=lambda c: c.col_index)
    num_cols = len(cols)
    if num_cols == 0:
        return {"data": {"cells": [], "num_rows": 50, "num_cols": 0,
                         "merges": {}, "row_heights": {}, "freeze_columns": 0, "styles": {}}}

    cells = sheet.cells
    max_row = max((c.row_index for c in cells), default=-1) + 1
    num_rows = max(max_row, 50)

    grid = [[""] * num_cols for _ in range(num_rows)]
    styles: dict = {}
    num_formats: dict = {}
    for c in cells:
        if c.row_index < num_rows and c.col_index < num_cols:
            val = c.formula if c.formula else (c.value or "")
            grid[c.row_index][c.col_index] = val
            if c.style:
                try:
                    s = json.loads(c.style)
                    if s:
                        cell_name = f"{get_column_letter(c.col_index + 1)}{c.row_index + 1}"
                        styles[cell_name] = _style_to_css(s)
                        if s.get('numFmt'):
                            num_formats[cell_name] = s['numFmt']
                except Exception:
                    pass

    # 병합 셀: xlsx range strings → jspreadsheet format
    merges: dict = {}
    if sheet.merges:
        try:
            for rng in json.loads(sheet.merges):
                result = _range_to_jss(rng)
                if result:
                    cell_name, dims = result
                    merges[cell_name] = dims
        except Exception:
            pass

    # 행 높이: pt → px
    row_heights_px: dict = {}
    if sheet.row_heights:
        try:
            for k, v in json.loads(sheet.row_heights).items():
                row_heights_px[k] = _pt_to_px(float(v))
        except Exception:
            pass

    # 틀 고정 열 수
    freeze_columns = _freeze_to_cols(sheet.freeze_panes)

    # 셀 메모
    comments: dict = {}
    for c in cells:
        if c.comment and c.row_index < num_rows and c.col_index < num_cols:
            cell_name = f"{get_column_letter(c.col_index + 1)}{c.row_index + 1}"
            comments[cell_name] = c.comment

    # 조건부 서식
    conditional_formats = []
    if sheet.conditional_formats:
        try:
            conditional_formats = json.loads(sheet.conditional_formats)
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
            "num_formats": num_formats,
            "conditional_formats": conditional_formats,
        }
    }


@router.post("/{template_id}/sheets", status_code=201)
async def add_template_sheet(
    template_id: str,
    body: SheetCreate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    t = db.query(Template).filter(Template.id == template_id).first()
    if not t:
        raise HTTPException(status_code=404, detail="Template not found")
    if len(t.sheets) >= MAX_SHEETS:
        raise HTTPException(status_code=400, detail=f"Max {MAX_SHEETS} sheets allowed")

    next_index = max((s.sheet_index for s in t.sheets), default=-1) + 1
    sheet = TemplateSheet(id=str(uuid.uuid4()), template_id=template_id,
                          sheet_index=next_index, sheet_name=body.sheet_name)
    db.add(sheet)
    for col in _default_columns(sheet.id):
        db.add(col)
    t.updated_at = _now()
    db.commit()
    db.refresh(sheet)
    return {"data": _sheet_detail(sheet), "message": "created"}


@router.patch("/{template_id}/sheets/{sheet_id}")
async def update_template_sheet(
    template_id: str,
    sheet_id: str,
    body: SheetUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    if body.sheet_name is not None:
        sheet.sheet_name = body.sheet_name
    if body.sheet_index is not None:
        sheet.sheet_index = body.sheet_index
    db.commit()
    return {"data": _sheet_detail(sheet), "message": "updated"}


@router.patch("/{template_id}/sheets/{sheet_id}/merges")
async def update_sheet_merges(
    template_id: str,
    sheet_id: str,
    body: MergesUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    """병합 셀 목록 저장 (jspreadsheet onmerge 콜백에서 호출)"""
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    sheet.merges = json.dumps(body.merges, ensure_ascii=False)
    db.commit()
    return {"message": "merges saved", "count": len(body.merges)}


class RowHeightsUpdate(BaseModel):
    row_heights: Optional[dict] = None  # e.g. {"0": 30, "5": 50} (row_index_str → pt)


@router.patch("/{template_id}/sheets/{sheet_id}/row-heights")
async def update_sheet_row_heights(
    template_id: str,
    sheet_id: str,
    body: RowHeightsUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    """서식 시트 행 높이 저장"""
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    sheet.row_heights = json.dumps(body.row_heights) if body.row_heights else None
    db.commit()
    return {"message": "row_heights saved"}


class FreezeUpdate(BaseModel):
    freeze_panes: Optional[str] = None


@router.patch("/{template_id}/sheets/{sheet_id}/freeze")
async def update_sheet_freeze(
    template_id: str,
    sheet_id: str,
    body: FreezeUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    """서식 시트 틀 고정 설정"""
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    sheet.freeze_panes = body.freeze_panes
    db.commit()
    return {"message": "freeze saved", "freeze_panes": body.freeze_panes}


class ConditionalFormatsUpdate(BaseModel):
    conditional_formats: Optional[str] = None  # JSON string of rules array


@router.patch("/{template_id}/sheets/{sheet_id}/conditional-formats")
async def update_sheet_conditional_formats(
    template_id: str,
    sheet_id: str,
    body: ConditionalFormatsUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    sheet.conditional_formats = body.conditional_formats
    db.commit()
    return {"message": "conditional formats saved"}


@router.delete("/{template_id}/sheets/{sheet_id}", status_code=204)
async def delete_template_sheet(
    template_id: str,
    sheet_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    t = db.query(Template).filter(Template.id == template_id).first()
    if not t:
        raise HTTPException(status_code=404, detail="Template not found")
    if len(t.sheets) <= 1:
        raise HTTPException(status_code=400, detail="Cannot delete the last sheet")
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    db.delete(sheet)
    t.updated_at = _now()
    db.commit()


# ── Column CRUD ───────────────────────────────────────────────
@router.post("/{template_id}/sheets/{sheet_id}/columns", status_code=201)
async def add_template_column(
    template_id: str,
    sheet_id: str,
    body: ColumnCreate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    next_idx = max((c.col_index for c in sheet.columns), default=-1) + 1
    if body.col_type not in ("text", "number", "date", "dropdown", "checkbox"):
        raise HTTPException(status_code=400, detail="Invalid col_type")
    col = TemplateColumn(
        id=str(uuid.uuid4()), sheet_id=sheet_id,
        col_index=next_idx, col_header=body.col_header,
        col_type=body.col_type, is_readonly=body.is_readonly, width=body.width,
    )
    db.add(col)
    db.commit()
    db.refresh(col)
    return {"data": _col_dict(col), "message": "created"}


@router.patch("/{template_id}/sheets/{sheet_id}/columns/{col_id}")
async def update_column(
    template_id: str,
    sheet_id: str,
    col_id: str,
    body: ColumnUpdate,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    col = db.query(TemplateColumn).filter(
        TemplateColumn.id == col_id,
        TemplateColumn.sheet_id == sheet_id,
    ).first()
    if not col:
        raise HTTPException(status_code=404, detail="Column not found")
    if body.col_header is not None:
        col.col_header = body.col_header
    if body.col_type is not None:
        col.col_type = body.col_type
    if body.col_options is not None:
        col.col_options = body.col_options
    if body.is_readonly is not None:
        col.is_readonly = body.is_readonly
    if body.width is not None:
        col.width = body.width
    db.commit()
    return {"data": _col_dict(col), "message": "updated"}


@router.delete("/{template_id}/sheets/{sheet_id}/columns/{col_id}", status_code=204)
async def delete_template_column(
    template_id: str,
    sheet_id: str,
    col_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    col = db.query(TemplateColumn).filter(
        TemplateColumn.id == col_id,
        TemplateColumn.sheet_id == sheet_id,
    ).first()
    if not col:
        raise HTTPException(status_code=404, detail="Column not found")
    deleted_idx = col.col_index
    db.delete(col)
    # 해당 컬럼 셀 삭제
    db.query(TemplateCell).filter(
        TemplateCell.sheet_id == sheet_id,
        TemplateCell.col_index == deleted_idx,
    ).delete()
    # 뒤 컬럼 인덱스 당기기
    remaining = db.query(TemplateColumn).filter(
        TemplateColumn.sheet_id == sheet_id,
        TemplateColumn.col_index > deleted_idx,
    ).all()
    for c in remaining:
        c.col_index -= 1
    # 뒤 셀 col_index 당기기
    later_cells = db.query(TemplateCell).filter(
        TemplateCell.sheet_id == sheet_id,
        TemplateCell.col_index > deleted_idx,
    ).all()
    for c in later_cells:
        c.col_index -= 1
    # 병합 범위 업데이트
    ts = db.query(TemplateSheet).filter(TemplateSheet.id == sheet_id).first()
    if ts and ts.merges:
        try:
            merges = json.loads(ts.merges)
            from .cells import _shift_merges_col
            updated = _shift_merges_col(merges, deleted_idx, -1)
            ts.merges = json.dumps(updated)
        except Exception:
            pass
    db.commit()


# ── Cell batch save ───────────────────────────────────────────
@router.post("/{template_id}/sheets/{sheet_id}/cells")
async def batch_save_cells(
    template_id: str,
    sheet_id: str,
    body: list[CellBatchItem],
    replace: bool = False,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    sheet = db.query(TemplateSheet).filter(
        TemplateSheet.id == sheet_id,
        TemplateSheet.template_id == template_id,
    ).first()
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")

    # replace=true: 기존 셀 전부 삭제 후 새로 저장 (행/열 삽입·삭제 후 전체 동기화용)
    if replace:
        db.query(TemplateCell).filter(TemplateCell.sheet_id == sheet_id).delete()

    for item in body:
        if not (0 <= item.row_index < MAX_ROWS):
            continue
        # 계산식 여부 분리
        formula = None
        value = item.value
        if item.value and item.value.startswith("="):
            formula = item.value
            value = None

        existing = db.query(TemplateCell).filter(
            TemplateCell.sheet_id == sheet_id,
            TemplateCell.row_index == item.row_index,
            TemplateCell.col_index == item.col_index,
        ).first()
        if existing:
            existing.value = value
            existing.formula = formula
            if item.style is not None:
                existing.style = item.style
            if item.comment is not None:
                existing.comment = item.comment if item.comment else None
        else:
            db.add(TemplateCell(
                id=str(uuid.uuid4()),
                sheet_id=sheet_id,
                row_index=item.row_index,
                col_index=item.col_index,
                value=value,
                formula=formula,
                style=item.style,
                comment=item.comment if item.comment else None,
            ))
    t = db.query(Template).filter(Template.id == template_id).first()
    if t:
        t.updated_at = _now()
    db.commit()
    return {"message": "saved", "count": len(body)}


# ── xlsx import / export ──────────────────────────────────────
@router.post("/import-xlsx", status_code=201)
async def import_template_xlsx(
    file: UploadFile = File(...),
    name: Optional[str] = None,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    content = await file.read()
    if content[:4] != b"PK\x03\x04":
        raise HTTPException(status_code=400, detail="Invalid xlsx file")

    # data_only=False → 계산식 원본 보존
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Cannot open xlsx: {e}")

    if len(wb.sheetnames) > MAX_SHEETS:
        raise HTTPException(status_code=400, detail=f"Too many sheets (max {MAX_SHEETS})")

    template_name = name or (file.filename or "Imported").replace(".xlsx", "")
    t = Template(id=str(uuid.uuid4()), name=template_name, created_by=current_user.id)
    db.add(t)

    for sheet_idx, sheet_name in enumerate(wb.sheetnames):
        ws = wb[sheet_name]
        new_sheet = TemplateSheet(id=str(uuid.uuid4()), template_id=t.id,
                                  sheet_index=sheet_idx, sheet_name=sheet_name)
        db.add(new_sheet)

        all_rows = list(ws.iter_rows(values_only=False))
        if not all_rows:
            db.add(TemplateColumn(id=str(uuid.uuid4()), sheet_id=new_sheet.id, col_index=0, col_header="A"))
            continue

        first_row = all_rows[0]
        # xlsx 전체 시트의 최대 컬럼 수 사용 (첫 행만 보면 뒤쪽 행의 데이터 누락 가능)
        num_cols = ws.max_column or 0
        if num_cols == 0:
            num_cols = max((i for i, cell in enumerate(first_row) if cell.value is not None), default=-1) + 1
        if num_cols == 0:
            num_cols = len(first_row) if first_row else 5

        # 컬럼 헤더는 Excel 열 문자(A, B, C...)
        for ci in range(num_cols):
            col_letter = get_column_letter(ci + 1)
            # 실제 xlsx 열 너비 사용 (px 변환: 1 unit ≈ 7px)
            col_dim = ws.column_dimensions.get(col_letter)
            col_width = int((col_dim.width or 8) * 7) if col_dim else 56
            col_width = max(40, min(col_width, 600))
            db.add(TemplateColumn(
                id=str(uuid.uuid4()),
                sheet_id=new_sheet.id,
                col_index=ci,
                col_header=col_letter,
                width=col_width,
            ))

        # 병합 셀 저장
        merges_list = [str(mr) for mr in ws.merged_cells.ranges]
        if merges_list:
            new_sheet.merges = json.dumps(merges_list)

        # 행 높이 저장 (pt 단위)
        row_heights: dict = {}
        for row_num, rd in ws.row_dimensions.items():
            if rd.height and rd.height > 0:
                ri = row_num - 1  # row_index = Excel row - 1
                row_heights[str(ri)] = rd.height
        if row_heights:
            new_sheet.row_heights = json.dumps(row_heights)

        # 틀 고정
        if ws.freeze_panes:
            new_sheet.freeze_panes = str(ws.freeze_panes)

        # 모든 행을 데이터로 저장
        for ri, row in enumerate(all_rows[:MAX_ROWS]):
            for ci in range(num_cols):
                cell = row[ci] if ci < len(row) else None
                if cell is None or cell.value is None:
                    # 값 없어도 스타일이나 메모 있으면 저장
                    style_json = _extract_cell_style(cell) if cell else None
                    comment_text = cell.comment.text.strip() if cell and cell.comment else None
                    if style_json or comment_text:
                        db.add(TemplateCell(
                            id=str(uuid.uuid4()), sheet_id=new_sheet.id,
                            row_index=ri, col_index=ci,
                            value=None, style=style_json, comment=comment_text,
                        ))
                    continue
                raw = cell.value
                style_json = _extract_cell_style(cell)
                comment_text = cell.comment.text.strip() if cell.comment else None
                if cell.data_type == 'f' or (isinstance(raw, str) and raw.startswith("=")):
                    raw_str = str(raw)
                    formula_str = raw_str if raw_str.startswith("=") else "=" + raw_str
                    db.add(TemplateCell(
                        id=str(uuid.uuid4()), sheet_id=new_sheet.id,
                        row_index=ri, col_index=ci,
                        value=None, formula=formula_str, style=style_json,
                        comment=comment_text,
                    ))
                else:
                    db.add(TemplateCell(
                        id=str(uuid.uuid4()), sheet_id=new_sheet.id,
                        row_index=ri, col_index=ci,
                        value=str(raw), style=style_json,
                        comment=comment_text,
                    ))

    db.commit()
    db.refresh(t)
    return {"data": _template_summary(t), "message": "imported"}


@router.get("/{template_id}/export-xlsx")
async def export_template_xlsx(
    template_id: str,
    current_user: User = Depends(require_admin),
    db: DBSession = Depends(get_db),
):
    t = db.query(Template).filter(Template.id == template_id).first()
    if not t:
        raise HTTPException(status_code=404, detail="Template not found")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet in t.sheets:
        ws = wb.create_sheet(sheet.sheet_name)
        cols = sorted(sheet.columns, key=lambda c: c.col_index)

        # 컬럼 너비
        for col in cols:
            ws.column_dimensions[get_column_letter(col.col_index + 1)].width = (col.width or 120) / 7

        # 행 높이 복원 (pt 단위)
        if sheet.row_heights:
            try:
                for ri_str, height_pt in json.loads(sheet.row_heights).items():
                    ws.row_dimensions[int(ri_str) + 1].height = float(height_pt)
            except Exception:
                pass

        # 틀 고정
        if sheet.freeze_panes:
            ws.freeze_panes = sheet.freeze_panes

        # 셀 데이터 + 스타일
        cells_map: dict[tuple, TemplateCell] = {(c.row_index, c.col_index): c for c in sheet.cells}
        max_row = max((c.row_index for c in sheet.cells), default=-1)
        for ri in range(max_row + 1):
            for ci in range(len(cols)):
                c = cells_map.get((ri, ci))
                if c:
                    val = c.formula if c.formula else c.value
                    if val is not None and not val.startswith("="):
                        # 선행 0이 있는 문자열("00123" 등)은 텍스트로 유지
                        if len(val) > 1 and val.startswith("0") and val != "0" and not val.startswith("0."):
                            pass  # keep as string to preserve leading zeros
                        else:
                            try:
                                val = int(val) if "." not in val else float(val)
                            except (ValueError, TypeError):
                                pass
                    ws_cell = ws.cell(row=ri + 1, column=ci + 1, value=val)
                    _apply_cell_style(ws_cell, c.style)
                    if c.comment:
                        from openpyxl.comments import Comment as XlComment
                        ws_cell.comment = XlComment(c.comment, "")

        # 병합 셀 복원
        if sheet.merges:
            try:
                for rng in json.loads(sheet.merges):
                    ws.merge_cells(rng)
            except Exception:
                pass

        # readonly 컬럼 셀 잠금 (시트 보호 활성화)
        readonly_col_indices = {col.col_index for col in cols if col.is_readonly}
        if readonly_col_indices:
            total_rows = ws.max_row or 1
            for r in range(1, total_rows + 1):
                for ci, col in enumerate(cols):
                    cell = ws.cell(row=r, column=ci + 1)
                    cell.protection = Protection(locked=(col.col_index in readonly_col_indices))
            ws.protection.sheet = True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_name = t.name.replace(" ", "_")[:50]
    from urllib.parse import quote
    encoded_name = quote(safe_name + ".xlsx")
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
    )
