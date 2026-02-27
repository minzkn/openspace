// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   spreadsheet_core.js — 공유 스프레드시트 모듈
   workspace.js 와 template_edit.js 에서 공통으로 사용하는
   유틸리티, 서식 툴바, Undo/Redo, 찾기/바꾸기, 수식 입력줄,
   상태 표시줄, 숫자 서식, 컨텍스트 메뉴, 키보드 단축키 등
   ============================================================ */
'use strict';

var SpreadsheetCore = (function() {

// ── 컬럼 인덱스 ↔ 문자 변환 ─────────────────────────────────
function colIndexToLetter(idx) {
  var letter = '', n = idx;
  while (n >= 0) { letter = String.fromCharCode(65 + (n % 26)) + letter; n = Math.floor(n / 26) - 1; }
  return letter;
}

function letterToColIndex(letter) {
  let result = 0;
  for (let i = 0; i < letter.length; i++) result = result * 26 + (letter.charCodeAt(i) - 64);
  return result - 1;
}

// ── 커스텀 수식 등록 ────────────────────────────────────────
function registerCustomFormulas() {
  if (typeof jspreadsheet === 'undefined') return;
  var fm = (jspreadsheet && jspreadsheet.formula) || (typeof formula !== 'undefined' ? formula : null);
  if (!fm || typeof fm.setFormula !== 'function') return;
  var EPOCH = new Date(1899, 11, 30);
  fm.setFormula({
    TIME: function(h, m, s) { return (Number(h) * 3600 + Number(m) * 60 + Number(s)) / 86400; },
    DATE: function(y, m, d) { var dt = new Date(Number(y), Number(m)-1, Number(d)); return Math.floor((dt-EPOCH)/86400000); },
    TODAY: function() { var t = new Date(); t.setHours(0,0,0,0); return Math.floor((t-EPOCH)/86400000); },
    NOW: function() { return (new Date()-EPOCH)/86400000; },
    HOUR: function(t) { return Math.floor(((Number(t)%1+1)%1)*24); },
    MINUTE: function(t) { return Math.floor(((Number(t)%1+1)%1)*24%1*60); },
    SECOND: function(t) { return Math.round(((Number(t)%1+1)%1)*24%1*3600%60); },
  });
}

// ── 색상 팔레트 (40색) ─────────────────────────────────────
const COLOR_PALETTE = [
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
  'FF8000','8000FF','FF0080','00FF80','800000','008000','000080','808000',
  '800080','008080','C0C0C0','808080','FF9999','99FF99','9999FF','FFFF99',
  'FF99FF','99FFFF','FFCC99','CC99FF','FF99CC','99FFCC','FFCCCC','CCFFCC',
  'CCCCFF','FFFFCC','FFCCFF','CCFFFF','E6E6E6','333333','666666','999999',
];

function initColorSwatches(ctx) {
  ['color-swatches', 'bg-swatches'].forEach((id, idx) => {
    const el = document.getElementById(id);
    if (!el) return;
    COLOR_PALETTE.forEach(hex => {
      const sw = document.createElement('div');
      sw.className = 'color-swatch';
      sw.style.background = '#' + hex;
      sw.title = '#' + hex;
      sw.onclick = (e) => {
        e.stopPropagation();
        if (idx === 0) fmtColor(ctx, hex);
        else fmtBg(ctx, hex);
      };
      el.appendChild(sw);
    });
  });
}

// ── 드롭다운 ────────────────────────────────────────────────
function toggleDropdown(id, e) {
  if (e) e.stopPropagation();
  else if (typeof event !== 'undefined') event.stopPropagation();
  const target = document.getElementById(id);
  if (!target) return;
  const wasOpen = target.classList.contains('open');
  closeAllDropdowns();
  if (!wasOpen) target.classList.add('open');
}

function closeAllDropdowns() {
  document.querySelectorAll('.fmt-dropdown.open').forEach(d => d.classList.remove('open'));
}

// ── CSS 색상 파싱 유틸리티 ───────────────────────────────────
// 브라우저는 CSS 색상을 rgb() 형식으로 정규화하므로 #hex와 rgb() 모두 처리
function _cssColorToHex(v) {
  if (!v) return null;
  v = v.trim();
  if (v.startsWith('#')) {
    var hex = v.substring(1).toUpperCase();
    if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
    return hex;
  }
  var m = v.match(/^rgb[a]?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (m) {
    return ('0'+parseInt(m[1]).toString(16)).slice(-2).toUpperCase() +
           ('0'+parseInt(m[2]).toString(16)).slice(-2).toUpperCase() +
           ('0'+parseInt(m[3]).toString(16)).slice(-2).toUpperCase();
  }
  // 색상 이름 → transparent, inherit 등은 null
  if (v === 'transparent' || v === 'inherit' || v === 'initial') return null;
  return null;
}

// ── CSS ↔ Style Object 변환 ─────────────────────────────────
// 주의: 브라우저가 setAttribute('style', ...) 후 getAttribute('style')로
// 돌려받을 때 CSS를 정규화함 (bold→700, #hex→rgb(), pt→px 등)
function cssToStyleObj(css) {
  const s = {};
  if (!css) return s;
  css.split(';').forEach(part => {
    const colonIdx = part.indexOf(':');
    if (colonIdx < 0) return;
    const k = part.substring(0, colonIdx).trim().toLowerCase();
    const v = part.substring(colonIdx + 1).trim();
    if (!k || !v) return;
    if (k === 'font-weight') {
      if (v === 'bold' || v === '700' || v === '800' || v === '900') s.bold = true;
    }
    if (k === 'font-style' && (v === 'italic' || v === 'oblique')) s.italic = true;
    if (k === 'text-decoration' || k === 'text-decoration-line') {
      if (v.includes('underline')) s.underline = true;
      if (v.includes('line-through')) s.strikethrough = true;
    }
    if (k === 'font-size') {
      var mPt = v.match(/([\d.]+)\s*pt/);
      if (mPt) { s.fontSize = parseFloat(mPt[1]); }
      else {
        var mPx = v.match(/([\d.]+)\s*px/);
        if (mPx) { s.fontSize = Math.round(parseFloat(mPx[1]) * 0.75 * 10) / 10; }
      }
    }
    if (k === 'color') {
      var c = _cssColorToHex(v);
      if (c) s.color = c;
    }
    if (k === 'background-color' || k === 'background') {
      var bg = _cssColorToHex(v);
      if (bg) s.bg = bg;
    }
    if (k === 'text-align') s.align = v;
    if (k === 'vertical-align') {
      s.valign = v; // CSS 값 그대로 유지 (top/middle/bottom)
    }
    if (k === 'white-space' && (v === 'pre-wrap' || v === 'pre-line' || v === 'break-spaces')) s.wrap = true;
    if (k === 'overflow-wrap' && v === 'break-word') s.wrap = true;
    // border-* parsing
    var bm = k.match(/^border-(top|bottom|left|right)$/);
    if (bm) {
      if (!s.border) s.border = {};
      // 브라우저 정규화된 형태: "1px solid rgb(0, 0, 0)" 또는 "1px solid #000000"
      var bParts = v.match(/^([\d.]+\w+)\s+(\w+)\s+(.+)$/);
      if (bParts) {
        var bStyle = _cssBorderStyleToExcel(bParts[2], bParts[1]);
        var bColor = _cssColorToHex(bParts[3]) || '000000';
        s.border[bm[1]] = { style: bStyle, color: bColor };
      }
    }
    // 커스텀 CSS 프로퍼티로 숫자 서식 보존 (URL 인코딩된 값)
    if (k === '--num-fmt') { try { s.numFmt = decodeURIComponent(v); } catch(e2) { s.numFmt = v; } }
  });
  return s;
}

function _cssBorderStyleToExcel(cssStyle, width) {
  var w = parseInt(width) || 1;
  if (cssStyle === 'dashed') return 'dashed';
  if (cssStyle === 'dotted') return 'dotted';
  if (cssStyle === 'double') return 'double';
  if (w >= 3) return 'thick';
  if (w >= 2) return 'medium';
  return 'thin';
}

function styleObjToCss(s) {
  const parts = [];
  if (s.bold) parts.push('font-weight:bold');
  if (s.italic) parts.push('font-style:italic');
  // text-decoration: combine underline + line-through
  const decorations = [];
  if (s.underline) decorations.push('underline');
  if (s.strikethrough) decorations.push('line-through');
  if (decorations.length) parts.push('text-decoration:' + decorations.join(' '));
  if (s.fontSize) parts.push('font-size:' + s.fontSize + 'pt');
  if (s.color) parts.push('color:#' + s.color);
  if (s.bg) parts.push('background-color:#' + s.bg);
  if (s.align) parts.push('text-align:' + s.align);
  if (s.valign) parts.push('vertical-align:' + (s.valign === 'center' ? 'middle' : s.valign));
  if (s.wrap) parts.push('white-space:pre-wrap');
  if (s.border) {
    const wm = {thin:'1px', medium:'2px', thick:'3px', dashed:'1px', dotted:'1px', double:'3px'};
    const sm = {thin:'solid', medium:'solid', thick:'solid', dashed:'dashed', dotted:'dotted', double:'double'};
    for (const [side, bd] of Object.entries(s.border)) {
      const bs = bd.style || 'thin';
      parts.push('border-' + side + ':' + (wm[bs]||'1px') + ' ' + (sm[bs]||'solid') + ' #' + (bd.color||'000000'));
    }
  }
  // 숫자 서식은 CSS가 아니지만 커스텀 프로퍼티로 보존 (CSS round-trip 유지, URL 인코딩)
  if (s.numFmt) parts.push('--num-fmt:' + encodeURIComponent(s.numFmt));
  return parts.join(';');
}

// ── 셀 스타일 조회 및 툴바 상태 ────────────────────────────
function getSelectedCellStyle(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return {};
  try {
    const sel = ctx.getSelection();
    const cellName = colIndexToLetter(sel.x1) + (sel.y1 + 1);
    const cssStr = ss.getStyle(cellName) || '';
    return cssToStyleObj(cssStr);
  } catch(e) { return {}; }
}

function updateToolbarState(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  const s = getSelectedCellStyle(ctx);
  setActive('fmt-bold', !!s.bold);
  setActive('fmt-italic', !!s.italic);
  setActive('fmt-underline', !!s.underline);
  setActive('fmt-strikethrough', !!s.strikethrough);
  setActive('fmt-wrap', !!s.wrap);
  setActive('fmt-align-left', s.align === 'left');
  setActive('fmt-align-center', s.align === 'center');
  setActive('fmt-align-right', s.align === 'right');
  setActive('fmt-valign-top', s.valign === 'top');
  setActive('fmt-valign-middle', s.valign === 'middle');
  setActive('fmt-valign-bottom', s.valign === 'bottom');
  const colorBar = document.getElementById('fmt-color-bar');
  if (colorBar) colorBar.style.background = s.color ? '#' + s.color : '#000000';
  const bgBar = document.getElementById('fmt-bg-bar');
  if (bgBar) bgBar.style.background = s.bg ? '#' + s.bg : 'transparent';
  // font size select
  const fontSizeSel = document.getElementById('fmt-font-size');
  if (fontSizeSel) fontSizeSel.value = s.fontSize ? String(s.fontSize) : '11';
  // number format select
  const numFmtSel = document.getElementById('fmt-num-format');
  if (numFmtSel) numFmtSel.value = s.numFmt || '';
  // formula bar
  updateFormulaBar(ctx);
  // status bar
  updateStatusBar(ctx);
}

function setActive(id, active) {
  const el = document.getElementById(id);
  if (el) el.classList.toggle('active', active);
}

// ── 서식 복사 (Format Painter) ──────────────────────────────
var _painterStyle = null;   // 복사된 스타일 객체 (null = 비활성)
var _painterPersist = false; // 더블클릭 → 지속 모드

function fmtPainterClick(ctx) {
  if (_painterStyle) { _cancelPainter(); return; }
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  _painterStyle = getSelectedCellStyle(ctx);
  _painterPersist = false;
  _activatePainterUI();
}

function fmtPainterDblClick(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  _painterStyle = getSelectedCellStyle(ctx);
  _painterPersist = true;
  _activatePainterUI();
}

function _activatePainterUI() {
  setActive('fmt-painter', true);
  document.body.classList.add('painter-cursor');
}

function _cancelPainter() {
  _painterStyle = null;
  _painterPersist = false;
  setActive('fmt-painter', false);
  document.body.classList.remove('painter-cursor');
}

function applyPainterToSelection(ctx) {
  if (!_painterStyle) return false;
  const ss = ctx.getSpreadsheet();
  if (!ss || !ctx.isEditable()) { _cancelPainter(); return false; }
  const sel = ctx.getSelection();
  const oldStyles = {};
  const styleMap = {};
  for (let r = sel.y1; r <= sel.y2; r++) {
    for (let c = sel.x1; c <= sel.x2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = ss.getStyle(cellName) || '';
      oldStyles[cellName] = cssStr;
      styleMap[cellName] = styleObjToCss(_painterStyle);
    }
  }
  try { ss.setStyle(styleMap); } catch(e) {}
  if (ctx.undoManager) {
    ctx.undoManager.push({
      type: 'style',
      cells: Object.keys(styleMap).map(cn => ({ cellName: cn, oldCss: oldStyles[cn], newCss: styleMap[cn] })),
    });
  }
  if (ctx.onStyleChange) ctx.onStyleChange(styleMap);
  if (!_painterPersist) _cancelPainter();
  updateToolbarState(ctx);
  return true;
}

function isPainterActive() { return !!_painterStyle; }

// ── 스타일 적용 ─────────────────────────────────────────────
function applyStyleToSelection(ctx, styleProp, value) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  if (!ctx.isEditable()) return;
  const sel = ctx.getSelection();

  // Undo: capture old styles
  const oldStyles = {};
  const styleMap = {};
  for (let r = sel.y1; r <= sel.y2; r++) {
    for (let c = sel.x1; c <= sel.x2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = ss.getStyle(cellName) || '';
      oldStyles[cellName] = cssStr;
      const s = cssToStyleObj(cssStr);
      if (value === null) delete s[styleProp];
      else s[styleProp] = value;
      const css = styleObjToCss(s);
      styleMap[cellName] = css;
    }
  }

  try { ss.setStyle(styleMap); } catch(e) {}

  // Push to undo
  if (ctx.undoManager) {
    ctx.undoManager.push({
      type: 'style',
      cells: Object.keys(styleMap).map(cn => ({ cellName: cn, oldCss: oldStyles[cn], newCss: styleMap[cn] })),
    });
  }

  if (ctx.onStyleChange) ctx.onStyleChange(styleMap);
  updateToolbarState(ctx);
}

// ── 서식 버튼 핸들러 ────────────────────────────────────────
function fmtBold(ctx) { const s = getSelectedCellStyle(ctx); applyStyleToSelection(ctx, 'bold', s.bold ? null : true); }
function fmtItalic(ctx) { const s = getSelectedCellStyle(ctx); applyStyleToSelection(ctx, 'italic', s.italic ? null : true); }
function fmtUnderline(ctx) { const s = getSelectedCellStyle(ctx); applyStyleToSelection(ctx, 'underline', s.underline ? null : true); }
function fmtStrikethrough(ctx) { const s = getSelectedCellStyle(ctx); applyStyleToSelection(ctx, 'strikethrough', s.strikethrough ? null : true); }
function fmtColor(ctx, hex) { closeAllDropdowns(); applyStyleToSelection(ctx, 'color', hex); const b = document.getElementById('fmt-color-bar'); if (b) b.style.background = hex ? '#'+hex : '#000000'; }
function fmtBg(ctx, hex) { closeAllDropdowns(); applyStyleToSelection(ctx, 'bg', hex); const b = document.getElementById('fmt-bg-bar'); if (b) b.style.background = hex ? '#'+hex : 'transparent'; }
function fmtAlign(ctx, dir) { const s = getSelectedCellStyle(ctx); applyStyleToSelection(ctx, 'align', s.align === dir ? null : dir); }
function fmtValign(ctx, dir) { const s = getSelectedCellStyle(ctx); applyStyleToSelection(ctx, 'valign', s.valign === dir ? null : dir); }
function fmtWrap(ctx) { const s = getSelectedCellStyle(ctx); applyStyleToSelection(ctx, 'wrap', s.wrap ? null : true); }
function fmtFontSize(ctx, size) { applyStyleToSelection(ctx, 'fontSize', size ? parseFloat(size) : null); }
function fmtNumFormat(ctx, fmt) { applyStyleToSelection(ctx, 'numFmt', fmt || null); }

// ── 병합 ───────────────────────────────────────────────────
function fmtMerge(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) { if (typeof showToast === 'function') showToast('스프레드시트가 로드되지 않았습니다', 'error'); return; }
  if (!ctx.canMerge()) return;
  const sel = ctx.getSelection();
  if (sel.x1 === sel.x2 && sel.y1 === sel.y2) {
    if (typeof showToast === 'function') showToast('병합할 셀 범위를 선택하세요 (2개 이상)', 'error');
    return;
  }
  try {
    var cellName = colIndexToLetter(sel.x1) + (sel.y1 + 1);
    var colspan = sel.x2 - sel.x1 + 1;
    var rowspan = sel.y2 - sel.y1 + 1;
    ss.setMerge(cellName, colspan, rowspan);
    if (ctx.onMergeChange) ctx.onMergeChange();
  } catch(e) { if (typeof showToast === 'function') showToast('병합 실패: ' + e.message, 'error'); }
}

function fmtUnmerge(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  if (!ctx.canMerge()) return;
  const sel = ctx.getSelection();
  try {
    var merges = ss.getMerge();
    if (!merges || Object.keys(merges).length === 0) {
      if (typeof showToast === 'function') showToast('선택한 셀에 병합이 없습니다', 'error');
      return;
    }
    // 선택 범위 내 또는 선택 셀을 포함하는 병합 찾기
    var found = null;
    for (var key in merges) {
      var m = key.match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      var mc = letterToColIndex(m[1]);
      var mr = parseInt(m[2]) - 1;
      var span = merges[key];
      var mcEnd = mc + (span[0] || 1) - 1;
      var mrEnd = mr + (span[1] || 1) - 1;
      // 선택된 셀이 이 병합 범위 안에 있는지 확인
      if (sel.x1 >= mc && sel.x1 <= mcEnd && sel.y1 >= mr && sel.y1 <= mrEnd) {
        found = key;
        break;
      }
    }
    if (!found) {
      if (typeof showToast === 'function') showToast('선택한 셀에 병합이 없습니다', 'error');
      return;
    }
    ss.removeMerge(found);
    if (ctx.onMergeChange) ctx.onMergeChange();
  } catch(e) { if (typeof showToast === 'function') showToast('병합 해제 실패: ' + e.message, 'error'); }
}

// ── 테두리 ─────────────────────────────────────────────────
function fmtBorder(ctx, preset, borderStyle, borderColor) {
  closeAllDropdowns();
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  if (!ctx.isEditable()) return;
  const sel = ctx.getSelection();
  const bStyle = borderStyle || 'thin';
  const bColor = borderColor || '000000';
  const bd = { style: bStyle, color: bColor };
  const oldStyles = {};
  const styleMap = {};
  for (let r = sel.y1; r <= sel.y2; r++) {
    for (let c = sel.x1; c <= sel.x2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = ss.getStyle(cellName) || '';
      oldStyles[cellName] = cssStr;
      const s = cssToStyleObj(cssStr);
      if (preset === 'none') { delete s.border; }
      else if (preset === 'all') { s.border = { top: bd, bottom: bd, left: bd, right: bd }; }
      else if (preset === 'outer') {
        if (!s.border) s.border = {};
        if (r === sel.y1) s.border.top = bd;
        if (r === sel.y2) s.border.bottom = bd;
        if (c === sel.x1) s.border.left = bd;
        if (c === sel.x2) s.border.right = bd;
      }
      else if (preset === 'bottom') { if (!s.border) s.border = {}; s.border.bottom = bd; }
      else if (preset === 'top') { if (!s.border) s.border = {}; s.border.top = bd; }
      else if (preset === 'left') { if (!s.border) s.border = {}; s.border.left = bd; }
      else if (preset === 'right') { if (!s.border) s.border = {}; s.border.right = bd; }
      styleMap[cellName] = styleObjToCss(s);
    }
  }
  try { ss.setStyle(styleMap); } catch(e) {}
  if (ctx.undoManager) {
    ctx.undoManager.push({
      type: 'style',
      cells: Object.keys(styleMap).map(cn => ({ cellName: cn, oldCss: oldStyles[cn], newCss: styleMap[cn] })),
    });
  }
  if (ctx.onStyleChange) ctx.onStyleChange(styleMap);
}

// ============================================================
// ── Undo/Redo Manager ───────────────────────────────────────
// ============================================================
class UndoManager {
  constructor(maxSize) {
    this.undoStack = [];
    this.redoStack = [];
    this.maxSize = maxSize || 100;
    this._recording = true;
  }

  push(action) {
    if (!this._recording) return;
    this.undoStack.push(action);
    if (this.undoStack.length > this.maxSize) this.undoStack.shift();
    this.redoStack = []; // clear redo on new action
  }

  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }

  undo(ctx) {
    if (!this.canUndo()) return;
    const action = this.undoStack.pop();
    this._recording = false;
    try {
      this._apply(ctx, action, true);
    } finally {
      this._recording = true;
    }
    this.redoStack.push(action);
  }

  redo(ctx) {
    if (!this.canRedo()) return;
    const action = this.redoStack.pop();
    this._recording = false;
    try {
      this._apply(ctx, action, false);
    } finally {
      this._recording = true;
    }
    this.undoStack.push(action);
  }

  _apply(ctx, action, isUndo) {
    const ss = ctx.getSpreadsheet();
    if (!ss) return;

    // setValueFromCoords가 onchange 트리거 → handleCellChange 중복 전송 방지
    if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
    try {
      if (action.type === 'value') {
        const changes = action.changes;
        for (const c of changes) {
          const val = isUndo ? c.oldVal : c.newVal;
          try { ss.setValueFromCoords(c.col, c.row, val || '', true); } catch(e) {}
        }
        ctx.onUndoRedoValue(changes, isUndo);
      }
      else if (action.type === 'style') {
        const styleMap = {};
        for (const c of action.cells) {
          styleMap[c.cellName] = isUndo ? c.oldCss : c.newCss;
        }
        try { ss.setStyle(styleMap); } catch(e) {}
        ctx.onStyleChange(styleMap);
      }
      else if (action.type === 'batch') {
        if (action.values) {
          for (const c of action.values) {
            const val = isUndo ? c.oldVal : c.newVal;
            try { ss.setValueFromCoords(c.col, c.row, val || '', true); } catch(e) {}
          }
          ctx.onUndoRedoValue(action.values, isUndo);
        }
        if (action.styles) {
          const styleMap = {};
          for (const c of action.styles) {
            styleMap[c.cellName] = isUndo ? c.oldCss : c.newCss;
          }
          try { ss.setStyle(styleMap); } catch(e) {}
          ctx.onStyleChange(styleMap);
        }
      }
    } finally {
      if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false;
    }
  }

  clear() {
    this.undoStack = [];
    this.redoStack = [];
  }
}


// ============================================================
// ── 키보드 단축키 ───────────────────────────────────────────
// ============================================================
function registerShortcuts(ctx) {
  document.addEventListener('keydown', function(e) {
    // Don't capture when typing in inputs/textareas
    const tag = e.target.tagName;
    const isInput = (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT');

    // Formula bar Enter/Escape handled separately
    if (e.target.id === 'formula-input') return;

    // Escape in inputs (modal, find panel etc.): let the element handle it
    if (e.key === 'Escape' && isInput) return;

    // Escape: cancel format painter or close find panel
    if (e.key === 'Escape') {
      if (_painterStyle) { _cancelPainter(); return; }
      closeFindPanel();
      return;
    }

    // Allow find panel shortcuts even in inputs
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      openFindPanel();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'h') {
      e.preventDefault();
      openFindPanel(true);
      return;
    }

    if (isInput) return;

    const ctrl = e.ctrlKey || e.metaKey;

    // Ctrl+Z: Undo
    if (ctrl && !e.shiftKey && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      if (ctx.undoManager) ctx.undoManager.undo(ctx);
      return;
    }
    // Ctrl+Y or Ctrl+Shift+Z: Redo
    if ((ctrl && e.key.toLowerCase() === 'y') || (ctrl && e.shiftKey && e.key.toLowerCase() === 'z')) {
      e.preventDefault();
      if (ctx.undoManager) ctx.undoManager.redo(ctx);
      return;
    }
    // Ctrl+B: Bold
    if (ctrl && e.key.toLowerCase() === 'b') {
      e.preventDefault();
      fmtBold(ctx);
      return;
    }
    // Ctrl+I: Italic
    if (ctrl && e.key.toLowerCase() === 'i') {
      e.preventDefault();
      fmtItalic(ctx);
      return;
    }
    // Ctrl+U: Underline
    if (ctrl && e.key.toLowerCase() === 'u') {
      e.preventDefault();
      fmtUnderline(ctx);
      return;
    }
    // Delete: clear selected cells
    if (e.key === 'Delete') {
      deleteSelectedCells(ctx);
      return;
    }
    // Ctrl+Shift+V: Paste Special
    if (ctrl && e.shiftKey && e.key.toLowerCase() === 'v') {
      e.preventDefault();
      showPasteSpecialDialog(ctx);
      return;
    }
    // Ctrl+C: capture internal clipboard
    if (ctrl && !e.shiftKey && e.key.toLowerCase() === 'c') {
      _captureInternalClipboard(ctx);
      // don't prevent default - let browser copy text too
    }
    // Ctrl+X: capture internal clipboard (cut)
    if (ctrl && !e.shiftKey && e.key.toLowerCase() === 'x') {
      _captureInternalClipboard(ctx);
    }
  });
}

function deleteSelectedCells(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss || !ctx.isEditable()) return;
  const sel = ctx.getSelection();
  const changes = [];
  if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
  try {
    for (let r = sel.y1; r <= sel.y2; r++) {
      for (let c = sel.x1; c <= sel.x2; c++) {
        let oldVal = '';
        try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
        if (oldVal !== '') {
          changes.push({ row: r, col: c, oldVal: oldVal, newVal: '' });
          try { ss.setValueFromCoords(c, r, '', true); } catch(e) {}
        }
      }
    }
  } finally {
    if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false;
  }
  if (changes.length > 0) {
    if (ctx.undoManager) ctx.undoManager.push({ type: 'value', changes });
    ctx.onDeleteCells(changes);
  }
}

// ============================================================
// ── 선택하여 붙여넣기 (Paste Special) ──────────────────────
// ============================================================
var _internalClipboard = null; // {data: [[{value, style}]], x1, y1, x2, y2}

function _captureInternalClipboard(ctx) {
  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  var sel = ctx.getSelection();
  var clipData = [];
  for (var r = sel.y1; r <= sel.y2; r++) {
    var row = [];
    for (var c = sel.x1; c <= sel.x2; c++) {
      var val = '';
      try { val = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
      var cellName = colIndexToLetter(c) + (r + 1);
      var style = null;
      try { style = ss.getStyle(cellName) || null; } catch(e) {}
      row.push({ value: val, style: style });
    }
    clipData.push(row);
  }
  _internalClipboard = { data: clipData, x1: sel.x1, y1: sel.y1, x2: sel.x2, y2: sel.y2 };
}

function showPasteSpecialDialog(ctx) {
  if (!_internalClipboard || !ctx.isEditable()) return;
  var old = document.getElementById('paste-special-dialog');
  if (old) old.remove();

  var dialog = document.createElement('div');
  dialog.id = 'paste-special-dialog';
  dialog.className = 'paste-special-dialog';
  dialog.innerHTML =
    '<div class="ps-title">선택하여 붙여넣기</div>' +
    '<label class="ps-opt"><input type="radio" name="ps-mode" value="all" checked> 모두</label>' +
    '<label class="ps-opt"><input type="radio" name="ps-mode" value="values"> 값만</label>' +
    '<label class="ps-opt"><input type="radio" name="ps-mode" value="formats"> 서식만</label>' +
    '<label class="ps-opt"><input type="checkbox" id="ps-transpose"> 전치(행/열 바꿈)</label>' +
    '<div class="ps-btn-bar">' +
    '<button class="ps-ok">확인</button>' +
    '<button class="ps-cancel">취소</button>' +
    '</div>';

  document.body.appendChild(dialog);
  // 중앙 위치
  dialog.style.position = 'fixed';
  dialog.style.left = '50%';
  dialog.style.top = '50%';
  dialog.style.transform = 'translate(-50%, -50%)';

  dialog.querySelector('.ps-ok').addEventListener('click', function() {
    var mode = dialog.querySelector('input[name="ps-mode"]:checked').value;
    var transpose = dialog.querySelector('#ps-transpose').checked;
    executePasteSpecial(ctx, mode, transpose);
    dialog.remove();
  });
  dialog.querySelector('.ps-cancel').addEventListener('click', function() {
    dialog.remove();
  });
}

function executePasteSpecial(ctx, mode, transpose) {
  var ss = ctx.getSpreadsheet();
  if (!ss || !_internalClipboard) return;
  var sel = ctx.getSelection();
  var startRow = sel.y1, startCol = sel.x1;
  var clipData = _internalClipboard.data;

  if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
  var changes = [];
  try {
    for (var r = 0; r < clipData.length; r++) {
      for (var c = 0; c < clipData[r].length; c++) {
        var destR = transpose ? startRow + c : startRow + r;
        var destC = transpose ? startCol + r : startCol + c;
        var item = clipData[r][c];

        if (mode === 'all' || mode === 'values') {
          var oldVal = '';
          try { oldVal = ss.getValueFromCoords(destC, destR) || ''; } catch(e) {}
          try { ss.setValueFromCoords(destC, destR, item.value, true); } catch(e) {}
          changes.push({ row: destR, col: destC, oldVal: oldVal, newVal: item.value });
        }
        if (mode === 'all' || mode === 'formats') {
          if (item.style) {
            var cellName = colIndexToLetter(destC) + (destR + 1);
            try { ss.setStyle(cellName, item.style); } catch(e) {}
          }
        }
      }
    }
  } finally {
    if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false;
  }
  if (changes.length > 0 && ctx.onPasteSpecial) {
    ctx.onPasteSpecial(changes);
  }
}

// ============================================================
// ── 찾기 / 바꾸기 ───────────────────────────────────────────
// ============================================================
let _findState = { results: [], current: -1, lastQuery: '', lastCase: false };

function openFindPanel(showReplace) {
  const panel = document.getElementById('find-panel');
  if (!panel) return;
  panel.classList.add('open');
  const replaceRow = document.getElementById('find-replace-row');
  if (replaceRow) replaceRow.style.display = showReplace ? 'flex' : 'none';
  const input = document.getElementById('find-input');
  if (input) { input.focus(); input.select(); }
}

function closeFindPanel() {
  const panel = document.getElementById('find-panel');
  if (!panel) return;
  panel.classList.remove('open');
  _findState = { results: [], current: -1, lastQuery: '', lastCase: false };
  updateFindCount();
}

function findNext(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  const query = (document.getElementById('find-input') || {}).value || '';
  if (!query) return;
  const caseSensitive = (document.getElementById('find-case') || {}).checked || false;

  // rebuild results if query or case-sensitivity changed
  if (query !== _findState.lastQuery || caseSensitive !== _findState.lastCase) {
    _buildFindResults(ctx, query, caseSensitive);
  }
  if (_findState.results.length === 0) { updateFindCount(); return; }

  _findState.current = (_findState.current + 1) % _findState.results.length;
  const r = _findState.results[_findState.current];
  try { ss.updateSelectionFromCoords(r.col, r.row, r.col, r.row); } catch(e) {}
  updateFindCount();
}

function findPrev(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  const query = (document.getElementById('find-input') || {}).value || '';
  if (!query) return;
  const caseSensitive = (document.getElementById('find-case') || {}).checked || false;

  if (query !== _findState.lastQuery || caseSensitive !== _findState.lastCase) {
    _buildFindResults(ctx, query, caseSensitive);
  }
  if (_findState.results.length === 0) { updateFindCount(); return; }

  _findState.current = (_findState.current - 1 + _findState.results.length) % _findState.results.length;
  const r = _findState.results[_findState.current];
  try { ss.updateSelectionFromCoords(r.col, r.row, r.col, r.row); } catch(e) {}
  updateFindCount();
}

function replaceCurrent(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss || !ctx.isEditable()) return;
  if (_findState.results.length === 0 || _findState.current < 0) return;
  const replaceVal = (document.getElementById('replace-input') || {}).value || '';
  const r = _findState.results[_findState.current];
  let oldVal = '';
  try {
    const data = ss.getData();
    if (data && data[r.row] && data[r.row][r.col] !== undefined && data[r.row][r.col] !== null) {
      oldVal = String(data[r.row][r.col]);
    }
  } catch(e) {}
  const query = (document.getElementById('find-input') || {}).value || '';
  const caseSensitive = (document.getElementById('find-case') || {}).checked || false;
  const regex = new RegExp(_escapeRegex(query), caseSensitive ? 'g' : 'gi');
  const newVal = oldVal.replace(regex, replaceVal);
  if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
  try { ss.setValueFromCoords(r.col, r.row, newVal, true); } catch(e) {}
  finally { if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false; }
  if (ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes: [{ row: r.row, col: r.col, oldVal, newVal }] });
  }
  if (ctx.onReplaceChange) ctx.onReplaceChange([{ row: r.row, col: r.col, oldVal, newVal }]);
  _findState.lastQuery = '';
  findNext(ctx);
}

function replaceAll(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss || !ctx.isEditable()) return;
  const query = (document.getElementById('find-input') || {}).value || '';
  const replaceVal = (document.getElementById('replace-input') || {}).value || '';
  if (!query) return;
  const caseSensitive = (document.getElementById('find-case') || {}).checked || false;
  _buildFindResults(ctx, query, caseSensitive);
  if (_findState.results.length === 0) return;
  const regex = new RegExp(_escapeRegex(query), caseSensitive ? 'g' : 'gi');
  const changes = [];
  const data = ss.getData();
  if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
  try {
    for (const r of _findState.results) {
      let oldVal = '';
      try {
        if (data && data[r.row] && data[r.row][r.col] !== undefined && data[r.row][r.col] !== null) {
          oldVal = String(data[r.row][r.col]);
        }
      } catch(e) {}
      const newVal = oldVal.replace(regex, replaceVal);
      if (newVal !== oldVal) {
        try { ss.setValueFromCoords(r.col, r.row, newVal, true); } catch(e) {}
        changes.push({ row: r.row, col: r.col, oldVal, newVal });
      }
    }
  } finally {
    if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false;
  }
  if (changes.length > 0) {
    if (ctx.undoManager) ctx.undoManager.push({ type: 'value', changes });
    if (ctx.onReplaceChange) ctx.onReplaceChange(changes);
  }
  _findState = { results: [], current: -1, lastQuery: '' };
  updateFindCount();
  if (typeof showToast === 'function') showToast(changes.length + '개 항목이 바뀌었습니다', 'success');
}

function _buildFindResults(ctx, query, caseSensitive) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  _findState.results = [];
  _findState.current = -1;
  _findState.lastQuery = query;
  _findState.lastCase = caseSensitive;
  const data = ss.getData();
  const q = caseSensitive ? query : query.toLowerCase();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < (data[r] || []).length; c++) {
      const val = String(data[r][c] || '');
      const cmp = caseSensitive ? val : val.toLowerCase();
      if (cmp.includes(q)) {
        _findState.results.push({ row: r, col: c });
      }
    }
  }
}

function _escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function updateFindCount() {
  const el = document.getElementById('find-count');
  if (!el) return;
  if (_findState.results.length === 0) {
    el.textContent = '결과 없음';
  } else {
    el.textContent = (_findState.current + 1) + ' / ' + _findState.results.length;
  }
}

// ============================================================
// ── 수식 입력줄 (Formula Bar) ──────────────────────────────
// ============================================================
function updateFormulaBar(ctx) {
  const cellRef = document.getElementById('formula-cell-ref');
  const input = document.getElementById('formula-input');
  if (!cellRef || !input) return;
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  const sel = ctx.getSelection();
  // Cell reference display (name box)
  const startName = colIndexToLetter(sel.x1) + (sel.y1 + 1);
  var refText;
  if (sel.x1 === sel.x2 && sel.y1 === sel.y2) {
    refText = startName;
  } else {
    refText = startName + ':' + colIndexToLetter(sel.x2) + (sel.y2 + 1);
  }
  // input 또는 div 지원
  if (cellRef.tagName === 'INPUT') {
    cellRef.value = refText;
  } else {
    cellRef.textContent = refText;
  }
  // Raw value (formula or value)
  try {
    let rawVal = ss.getValueFromCoords(sel.x1, sel.y1);
    // For formulas, try to get the raw formula from data
    const data = ss.options.data;
    if (data && data[sel.y1] && data[sel.y1][sel.x1] !== undefined) {
      const dataVal = data[sel.y1][sel.x1];
      if (typeof dataVal === 'string' && dataVal.startsWith('=')) {
        rawVal = dataVal;
      }
    }
    input.value = rawVal || '';
  } catch(e) {
    input.value = '';
  }
}

function handleFormulaBarEnter(ctx, inputEl) {
  const ss = ctx.getSpreadsheet();
  if (!ss || !ctx.isEditable()) return;
  const sel = ctx.getSelection();
  const newVal = inputEl.value;
  let oldVal = '';
  try { oldVal = ss.getValueFromCoords(sel.x1, sel.y1) || ''; } catch(e) {}
  if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
  try { ss.setValueFromCoords(sel.x1, sel.y1, newVal, true); } catch(e) {}
  finally { if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false; }
  if (ctx.undoManager && oldVal !== newVal) {
    ctx.undoManager.push({ type: 'value', changes: [{ row: sel.y1, col: sel.x1, oldVal, newVal }] });
  }
  if (oldVal !== newVal && ctx.onFormulaBarChange) {
    ctx.onFormulaBarChange(sel.y1, sel.x1, newVal);
  }
  inputEl.blur();
}

// ── 이름 상자 셀 이동 (Name Box Navigation) ──────────────────
function initNameBox(ctx) {
  const cellRef = document.getElementById('formula-cell-ref');
  if (!cellRef || cellRef.tagName !== 'INPUT') return;
  cellRef.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      navigateToCell(ctx, cellRef.value.trim().toUpperCase());
      cellRef.blur();
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      updateFormulaBar(ctx);
      cellRef.blur();
    }
  });
  cellRef.addEventListener('focus', function() { cellRef.select(); });
}

function navigateToCell(ctx, ref) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  // 범위 지원: "A1:D10" 또는 단일 셀 "A1"
  var rangeMatch = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (rangeMatch) {
    var c1 = letterToColIndex(rangeMatch[1]);
    var r1 = parseInt(rangeMatch[2]) - 1;
    var c2 = letterToColIndex(rangeMatch[3]);
    var r2 = parseInt(rangeMatch[4]) - 1;
    if (r1 < 0 || c1 < 0 || r2 < 0 || c2 < 0) return;
    try { ss.updateSelectionFromCoords(c1, r1, c2, r2); } catch(e) {}
    return;
  }
  var cellMatch = ref.match(/^([A-Z]+)(\d+)$/);
  if (cellMatch) {
    var col = letterToColIndex(cellMatch[1]);
    var row = parseInt(cellMatch[2]) - 1;
    if (row < 0 || col < 0) return;
    try { ss.updateSelectionFromCoords(col, row, col, row); } catch(e) {}
  }
}

// ============================================================
// ── 상태 표시줄 (Status Bar) ────────────────────────────────
// ============================================================
function updateStatusBar(ctx) {
  const bar = document.getElementById('status-bar');
  if (!bar) return;
  const ss = ctx.getSpreadsheet();
  if (!ss) { bar.innerHTML = ''; return; }
  const sel = ctx.getSelection();
  // 단일 셀 선택 시 빈 표시
  if (sel.x1 === sel.x2 && sel.y1 === sel.y2) {
    bar.innerHTML = '';
    return;
  }
  let count = 0, numCount = 0, sum = 0, min = Infinity, max = -Infinity;
  for (let r = sel.y1; r <= sel.y2; r++) {
    for (let c = sel.x1; c <= sel.x2; c++) {
      let val = '';
      try { val = ss.getValueFromCoords(c, r); } catch(e) {}
      if (val !== '' && val !== null && val !== undefined) {
        count++;
        const num = Number(val);
        if (!isNaN(num) && val !== '') {
          numCount++;
          sum += num;
          if (num < min) min = num;
          if (num > max) max = num;
        }
      }
    }
  }
  const spans = [];
  spans.push('<span class="stat-item">개수: <b>' + count + '</b></span>');
  if (numCount > 0) {
    spans.push('<span class="stat-item">합계: <b>' + _formatNum(sum) + '</b></span>');
    spans.push('<span class="stat-item">평균: <b>' + _formatNum(sum / numCount) + '</b></span>');
    spans.push('<span class="stat-item">최소: <b>' + _formatNum(min) + '</b></span>');
    spans.push('<span class="stat-item">최대: <b>' + _formatNum(max) + '</b></span>');
  }
  bar.innerHTML = spans.join('<span class="stat-sep">|</span>');
}

function _formatNum(n) {
  if (Number.isInteger(n)) return n.toLocaleString();
  return n.toLocaleString(undefined, { maximumFractionDigits: 4 });
}

// ============================================================
// ── 숫자 서식 포맷팅 유틸리티 ───────────────────────────────
// ============================================================
const NUM_FORMATS = {
  '': null, // General
  '#,##0': { type: 'number', decimals: 0 },
  '#,##0.00': { type: 'number', decimals: 2 },
  '0.00%': { type: 'percent', decimals: 2 },
  '\\u20A9#,##0': { type: 'currency', decimals: 0, symbol: '\u20A9' },
  'yyyy-mm-dd': { type: 'date' },
  'yyyy-mm-dd hh:mm': { type: 'datetime' },
};

function formatNumber(value, numFmt) {
  if (!numFmt || value === null || value === undefined || value === '') return null;
  var num = Number(value);
  if (isNaN(num)) return null;

  // 정확히 일치하는 간단한 서식 먼저 확인
  if (numFmt === '#,##0') return Math.round(num).toLocaleString();
  if (numFmt === '#,##0.00') return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (numFmt === '0.00%') return (num * 100).toFixed(2) + '%';
  if (numFmt === '\u20A9#,##0') return '\u20A9' + Math.round(num).toLocaleString();
  if (numFmt === 'yyyy-mm-dd') {
    var d = _serialToDate(num);
    return d ? _fmtDate(d) : null;
  }
  if (numFmt === 'yyyy-mm-dd hh:mm') {
    var d2 = _serialToDate(num);
    return d2 ? _fmtDateTime(d2) : null;
  }

  // 복합 Excel 서식 코드 파싱 (예: "_-* #,##0_-;\-* #,##0_-;_-* \"-\"_-;_-@_-")
  // 세미콜론으로 분리: [양수; 음수; 0; 텍스트]
  var sections = numFmt.split(';');
  var section = num > 0 ? sections[0] : (num < 0 ? (sections[1] || sections[0]) : (sections[2] || sections[0]));
  if (!section) return null;

  // 핵심 숫자 패턴 추출 (_, *, 리터럴 문자열 제거)
  var pattern = section
    .replace(/_./g, '')           // _x (공백 문자)
    .replace(/\*./g, '')           // *x (반복 문자)
    .replace(/"[^"]*"/g, '')       // "리터럴" 문자열
    .replace(/\\/g, '')            // 이스케이프
    .trim();

  // 날짜 패턴 감지
  if (/[ymd]/i.test(pattern) && /[ymd].*[ymd]/i.test(pattern)) {
    var d3 = _serialToDate(num);
    return d3 ? _fmtDate(d3) : null;
  }

  // % 패턴 감지
  if (pattern.indexOf('%') >= 0) {
    var decMatch = pattern.match(/0\.(0+)/);
    var decimals = decMatch ? decMatch[1].length : 0;
    return (num * 100).toFixed(decimals) + '%';
  }

  // #,##0 패턴 (소수점 포함 여부)
  var decimalMatch = pattern.match(/#,##0\.(0+)/);
  if (decimalMatch) {
    var dec = decimalMatch[1].length;
    var result = (num < 0 ? -num : num).toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec });
    return num < 0 ? '-' + result : result;
  }
  if (pattern.indexOf('#,##0') >= 0) {
    var result2 = Math.round(num < 0 ? -num : num).toLocaleString();
    return num < 0 ? '-' + result2 : result2;
  }

  // 0.00 패턴 (콤마 없이)
  var zeroDecMatch = pattern.match(/0\.(0+)/);
  if (zeroDecMatch) {
    return num.toFixed(zeroDecMatch[1].length);
  }

  return null;
}

function _serialToDate(serial) {
  if (serial < 1) return null;
  const EPOCH = new Date(1899, 11, 30);
  const d = new Date(EPOCH.getTime() + serial * 86400000);
  return d;
}

function _fmtDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + dd;
}

function _fmtDateTime(d) {
  return _fmtDate(d) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
}

// ============================================================
// ── 컨텍스트 메뉴 ──────────────────────────────────────────
// ============================================================
function buildContextMenu(ctx) {
  return function(obj, x, y, e) {
    var items = [];
    // x, y 안전한 정수 변환 (Jspreadsheet CE v4는 null/undefined 전달 가능)
    var cx = (x !== null && x !== undefined) ? parseInt(x) : null;
    var cy = (y !== null && y !== undefined) ? parseInt(y) : null;

    // 복사 (항상 허용) / 잘라내기·붙여넣기 (편집 가능 시에만)
    if (ctx.isEditable()) {
      items.push({ title: '잘라내기', onclick: function() { obj.cut(); } });
    }
    items.push({ title: '복사', onclick: function() { obj.copy(); } });
    if (ctx.isEditable()) {
      items.push({ title: '붙여넣기', onclick: function() {
        navigator.clipboard.readText().then(function(text) {
          if (text) obj.paste(obj.selectedCell[0], obj.selectedCell[1], text);
        }).catch(function() {});
      }});
      items.push({ title: '선택하여 붙여넣기...', onclick: function() {
        showPasteSpecialDialog(ctx);
      }});
    }

    if (ctx.isEditable() && cy !== null && !isNaN(cy)) {
      items.push({ type: 'line' }); // separator
      items.push({ title: '위에 행 삽입', onclick: function() {
        if (ctx.onRowInsert) ctx.onRowInsert(cy, 'above');
        else obj.insertRow(1, cy, true);
      }});
      items.push({ title: '아래에 행 삽입', onclick: function() {
        if (ctx.onRowInsert) ctx.onRowInsert(cy, 'below');
        else obj.insertRow(1, cy, false);
      }});
      items.push({ title: '행 삭제', onclick: function() {
        var sel = ctx.getSelection();
        var rows = [];
        if (sel) for (var r = sel.y1; r <= sel.y2; r++) rows.push(r);
        else rows.push(cy);
        if (ctx.onRowsDelete) ctx.onRowsDelete(rows);
        else if (ctx.onRowDelete) { for (var i = rows.length - 1; i >= 0; i--) ctx.onRowDelete(rows[i]); }
        else { for (var i = rows.length - 1; i >= 0; i--) obj.deleteRow(rows[i]); }
      }});
    }

    if (ctx.isEditable() && cx !== null && !isNaN(cx)) {
      items.push({ type: 'line' });
      items.push({ title: '왼쪽에 열 삽입', onclick: function() {
        if (ctx.onColumnInsert) ctx.onColumnInsert(cx, 'before');
        else obj.insertColumn(1, cx, true);
      }});
      items.push({ title: '오른쪽에 열 삽입', onclick: function() {
        if (ctx.onColumnInsert) ctx.onColumnInsert(cx, 'after');
        else obj.insertColumn(1, cx, false);
      }});
      items.push({ title: '열 삭제', onclick: function() {
        var sel = ctx.getSelection();
        var cols = [];
        if (sel) for (var c = sel.x1; c <= sel.x2; c++) cols.push(c);
        else cols.push(cx);
        if (ctx.onColumnsDelete) ctx.onColumnsDelete(cols);
        else if (ctx.onColumnDelete) { for (var i = cols.length - 1; i >= 0; i--) ctx.onColumnDelete(cols[i]); }
        else { for (var i = cols.length - 1; i >= 0; i--) obj.deleteColumn(cols[i]); }
      }});
    }

    if (ctx.canMerge()) {
      items.push({ type: 'line' });
      items.push({ title: '셀 병합', onclick: function() { fmtMerge(ctx); } });
      items.push({ title: '병합 해제', onclick: function() { fmtUnmerge(ctx); } });
    }

    if (ctx.isEditable() && cx !== null && !isNaN(cx)) {
      items.push({ type: 'line' });
      items.push({ title: '오름차순 정렬', onclick: function() { sortColumn(ctx, cx, true); } });
      items.push({ title: '내림차순 정렬', onclick: function() { sortColumn(ctx, cx, false); } });
    }

    // 컬럼 속성 (onColumnProps 콜백이 있고 열이 식별 가능할 때)
    if (ctx.onColumnProps && cx !== null && !isNaN(cx)) {
      items.push({ type: 'line' });
      items.push({ title: '컬럼 속성', onclick: function() { ctx.onColumnProps(cx); } });
    }

    // 메모 추가/편집
    if (ctx.isEditable() && cx !== null && cy !== null && !isNaN(cx) && !isNaN(cy)) {
      items.push({ type: 'line' });
      var cellName = colIndexToLetter(cx) + (cy + 1);
      var hasComment = _commentsMap[cellName];
      items.push({ title: hasComment ? '메모 편집' : '메모 추가', onclick: function() { editCellComment(ctx, cx, cy); } });
      if (hasComment) {
        items.push({ title: '메모 삭제', onclick: function() {
          delete _commentsMap[cellName];
          addCommentIndicators(ctx, _commentsMap);
          if (ctx.onCommentChange) ctx.onCommentChange(cy, cx, '');
        }});
      }
    }

    // 하이퍼링크
    if (ctx.isEditable() && cx !== null && cy !== null && !isNaN(cx) && !isNaN(cy)) {
      items.push({ type: 'line' });
      var hlCellName = colIndexToLetter(cx) + (cy + 1);
      var hasLink = _hyperlinksMap && _hyperlinksMap[hlCellName];
      items.push({ title: hasLink ? '하이퍼링크 편집' : '하이퍼링크 삽입', onclick: function() {
        editHyperlink(ctx, cx, cy);
      }});
      if (hasLink) {
        items.push({ title: '하이퍼링크 삭제', onclick: function() {
          delete _hyperlinksMap[hlCellName];
          applyHyperlinkStyles(ctx);
          if (ctx.onHyperlinkChange) ctx.onHyperlinkChange(cy, cx, '');
        }});
      }
    }

    // 행/열 숨기기 (관리자 전용)
    if (ctx.canMerge && ctx.canMerge()) {
      if (cy !== null && !isNaN(cy)) {
        items.push({ type: 'line' });
        items.push({ title: '행 숨기기', onclick: function() {
          var sel = ctx.getSelection();
          var rows = [];
          if (sel) for (var r = sel.y1; r <= sel.y2; r++) rows.push(r);
          else rows.push(cy);
          if (ctx.onHideRows) ctx.onHideRows(rows);
        }});
        items.push({ title: '행 숨기기 해제', onclick: function() {
          if (ctx.onUnhideRows) ctx.onUnhideRows();
        }});
      }
      if (cx !== null && !isNaN(cx)) {
        items.push({ title: '열 숨기기', onclick: function() {
          var sel = ctx.getSelection();
          var cols = [];
          if (sel) for (var c = sel.x1; c <= sel.x2; c++) cols.push(c);
          else cols.push(cx);
          if (ctx.onHideCols) ctx.onHideCols(cols);
        }});
        items.push({ title: '열 숨기기 해제', onclick: function() {
          if (ctx.onUnhideCols) ctx.onUnhideCols();
        }});
      }

      // 틀 고정 설정 (관리자)
      items.push({ type: 'line' });
      items.push({ title: '틀 고정 설정', onclick: function() {
        if (ctx.onFreezeSetup) ctx.onFreezeSetup();
      }});
      // 시트 보호 토글 (관리자)
      if (ctx.onSheetProtection) {
        items.push({ title: '시트 보호 설정', onclick: function() {
          ctx.onSheetProtection();
        }});
      }
      // 인쇄 설정 (관리자)
      if (ctx.onPrintSetup) {
        items.push({ title: '인쇄 설정', onclick: function() {
          ctx.onPrintSetup();
        }});
      }
    }

    // 자동 필터 (모든 사용자)
    items.push({ type: 'line' });
    if (_filterActive) {
      items.push({ title: '자동 필터 해제', onclick: function() {
        clearAutoFilter(ctx);
      }});
    } else {
      items.push({ title: '자동 필터', onclick: function() {
        initAutoFilter(ctx, 0);
      }});
    }

    return items;
  };
}

// ============================================================
// ── 열 정렬 (클라이언트 사이드) ────────────────────────────
// ============================================================
function sortColumn(ctx, colIdx, ascending) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  // 병합 셀이 있으면 정렬 시 데이터 손상 가능 → 경고
  try {
    var merges = ss.getMerge();
    if (merges && Object.keys(merges).length > 0) {
      if (!confirm('병합된 셀이 있어 정렬 시 데이터가 손상될 수 있습니다.\n계속하시겠습니까?')) return;
    }
  } catch(e) {}
  const data = ss.getData();
  if (!data || data.length === 0) return;

  // Find last non-empty row
  let lastRow = 0;
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < (data[r] || []).length; c++) {
      if (data[r][c] !== '' && data[r][c] !== null && data[r][c] !== undefined) {
        lastRow = r;
      }
    }
  }

  // Extract rows with data
  // 정렬 전 행별 스타일 캡처
  const rows = [];
  for (let r = 0; r <= lastRow; r++) {
    const rowStyles = {};
    for (let c = 0; c < (data[r] || []).length; c++) {
      var cn = colIndexToLetter(c) + (r + 1);
      try {
        var st = ss.getStyle(cn);
        if (st) rowStyles[c] = st;
      } catch(e) {}
    }
    rows.push({ idx: r, data: (data[r] || []).slice(), styles: rowStyles });
  }

  // Sort by target column
  rows.sort((a, b) => {
    let va = a.data[colIdx];
    let vb = b.data[colIdx];
    if (va === '' || va === null || va === undefined) va = null;
    if (vb === '' || vb === null || vb === undefined) vb = null;
    if (va === null && vb === null) return 0;
    if (va === null) return 1;
    if (vb === null) return -1;
    const na = Number(va), nb = Number(vb);
    if (!isNaN(na) && !isNaN(nb)) {
      return ascending ? na - nb : nb - na;
    }
    const sa = String(va), sb = String(vb);
    return ascending ? sa.localeCompare(sb) : sb.localeCompare(sa);
  });

  // Apply sorted data + styles
  const changes = [];
  if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
  try {
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].data.length; c++) {
        let oldVal = '';
        try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
        const newVal = rows[r].data[c] || '';
        if (oldVal !== newVal) {
          try { ss.setValueFromCoords(c, r, newVal, true); } catch(e) {}
          changes.push({ row: r, col: c, oldVal, newVal: String(newVal) });
        }
        // 스타일도 원래 행에서 이동
        var cn = colIndexToLetter(c) + (r + 1);
        var srcStyle = rows[r].styles[c] || '';
        try { ss.setStyle(cn, srcStyle); } catch(e) {}
      }
    }
  } finally {
    if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false;
  }
  if (changes.length > 0) {
    if (ctx.undoManager) ctx.undoManager.push({ type: 'value', changes });
    if (ctx.onSort) ctx.onSort(changes);
  }
}

// ============================================================
// ── 하이퍼링크 관리 ─────────────────────────────────────────
// ============================================================
var _hyperlinksMap = {};

function setHyperlinksMap(map) { _hyperlinksMap = map || {}; }
function getHyperlinksMap() { return _hyperlinksMap; }

function editHyperlink(ctx, col, row) {
  var cellName = colIndexToLetter(col) + (row + 1);
  var currentUrl = _hyperlinksMap[cellName] || '';
  var url = prompt('하이퍼링크 URL을 입력하세요 (http://, https://, mailto:):', currentUrl);
  if (url === null) return; // 취소
  url = url.trim();
  if (url && !url.match(/^(https?:\/\/|mailto:|#)/i)) {
    if (typeof showToast === 'function') showToast('허용되지 않는 URL 프로토콜입니다', 'error');
    return;
  }
  if (url) {
    _hyperlinksMap[cellName] = url;
  } else {
    delete _hyperlinksMap[cellName];
  }
  applyHyperlinkStyles(ctx);
  if (ctx.onHyperlinkChange) ctx.onHyperlinkChange(row, col, url);
}

function applyHyperlinkStyles(ctx) {
  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  var container = document.getElementById('spreadsheet');
  if (!container) return;
  // 기존 하이퍼링크 클래스 제거
  container.querySelectorAll('.cell-hyperlink').forEach(function(el) {
    el.classList.remove('cell-hyperlink');
  });
  // 하이퍼링크 셀에 클래스 추가
  for (var cellName in _hyperlinksMap) {
    try {
      var m = cellName.match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      var ci = letterToColIndex(m[1]);
      var ri = parseInt(m[2]) - 1;
      var td = ss.getCell(ci, ri);
      if (td) td.classList.add('cell-hyperlink');
    } catch(e) {}
  }
}

function handleHyperlinkClick(ctx, e) {
  if (!e.ctrlKey && !e.metaKey) return;
  var td = e.target.closest('td.cell-hyperlink');
  if (!td) return;
  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  // 셀 좌표 찾기
  try {
    var x = td.getAttribute('data-x');
    var y = td.getAttribute('data-y');
    if (x === null || y === null) return;
    var cellName = colIndexToLetter(parseInt(x)) + (parseInt(y) + 1);
    var url = _hyperlinksMap[cellName];
    if (url) {
      window.open(url, '_blank', 'noopener,noreferrer');
      e.preventDefault();
    }
  } catch(e2) {}
}

// ============================================================
// ── 행/열 숨기기 ────────────────────────────────────────────
// ============================================================
function applyHiddenRows(ctx, hiddenRows) {
  var ss = ctx.getSpreadsheet();
  if (!ss || !ss.rows) return;
  for (var i = 0; i < ss.rows.length; i++) {
    if (ss.rows[i]) {
      ss.rows[i].style.display = hiddenRows.includes(i) ? 'none' : '';
    }
  }
}

function applyHiddenCols(ctx, hiddenCols) {
  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  var container = document.getElementById('spreadsheet');
  if (!container) return;
  // 열 숨기기: colgroup + thead td + tbody td
  var colgroup = ss.colgroup;
  if (colgroup) {
    for (var i = 0; i < colgroup.length; i++) {
      if (colgroup[i]) {
        colgroup[i].style.display = hiddenCols.includes(i) ? 'none' : '';
      }
    }
  }
  // thead
  var thead = container.querySelector('thead');
  if (thead) {
    var ths = thead.querySelectorAll('td');
    ths.forEach(function(th, idx) {
      // 첫 번째 td는 행 번호 열이므로 idx-1로 컬럼 매핑
      if (idx > 0) {
        th.style.display = hiddenCols.includes(idx - 1) ? 'none' : '';
      }
    });
  }
  // tbody
  var trs = container.querySelectorAll('tbody tr');
  trs.forEach(function(tr) {
    var tds = tr.querySelectorAll('td');
    tds.forEach(function(td, idx) {
      if (idx > 0) {
        td.style.display = hiddenCols.includes(idx - 1) ? 'none' : '';
      }
    });
  });
}

// ============================================================
// ── 행 고정 (Freeze Rows) ──────────────────────────────────
// ============================================================
var _freezeRowCount = 0;

function applyFreezeRows(ctx, numRows) {
  _freezeRowCount = numRows || 0;
  if (_freezeRowCount <= 0) return;
  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  var container = document.getElementById('spreadsheet');
  if (!container) return;

  // jspreadsheet CE 사용 시 테이블 구조: content div > table > thead + tbody
  var contentDiv = container.querySelector('.jexcel_content');
  if (!contentDiv) return;

  // 고정 행의 누적 높이를 계산하여 sticky top 설정
  _applyFreezeStickyRows(contentDiv, _freezeRowCount);

  // lazy loading으로 행이 재생성될 수 있으므로 MutationObserver 등록
  if (!container._freezeObserver) {
    container._freezeObserver = new MutationObserver(function() {
      if (_freezeRowCount > 0) {
        var cd = container.querySelector('.jexcel_content');
        if (cd) _applyFreezeStickyRows(cd, _freezeRowCount);
      }
    });
    var tbody = contentDiv.querySelector('tbody');
    if (tbody) {
      container._freezeObserver.observe(tbody, { childList: true });
    }
  }
}

function _applyFreezeStickyRows(contentDiv, numRows) {
  var tbody = contentDiv.querySelector('tbody');
  if (!tbody) return;
  var rows = tbody.querySelectorAll('tr');
  var cumHeight = 0;
  // thead 높이 (컬럼 헤더)
  var thead = contentDiv.querySelector('thead');
  var theadH = thead ? thead.offsetHeight : 0;

  for (var i = 0; i < rows.length; i++) {
    var tr = rows[i];
    if (i < numRows) {
      var topVal = theadH + cumHeight;
      tr.style.position = 'sticky';
      tr.style.top = topVal + 'px';
      tr.style.zIndex = '3';
      tr.classList.add('freeze-row');
      // 마지막 고정 행에 구분선 표시
      if (i === numRows - 1) {
        tr.classList.add('freeze-row-last');
      } else {
        tr.classList.remove('freeze-row-last');
      }
      cumHeight += tr.offsetHeight || 25;
    } else {
      tr.style.position = '';
      tr.style.top = '';
      tr.style.zIndex = '';
      tr.classList.remove('freeze-row', 'freeze-row-last');
    }
  }
}

function clearFreezeRows() {
  _freezeRowCount = 0;
  var container = document.getElementById('spreadsheet');
  if (!container) return;
  if (container._freezeObserver) {
    container._freezeObserver.disconnect();
    container._freezeObserver = null;
  }
  var rows = container.querySelectorAll('tbody tr');
  rows.forEach(function(tr) {
    tr.style.position = '';
    tr.style.top = '';
    tr.style.zIndex = '';
    tr.classList.remove('freeze-row', 'freeze-row-last');
  });
}

// ============================================================
// ── 행/열 그룹화 (Outline/Grouping) ────────────────────────
// ============================================================
var _outlineRows = {}; // {row_index_str: level}
var _outlineCols = {}; // {col_index_str: level}
var _collapsedGroups = new Set(); // collapsed group identifiers

function applyOutlines(ctx, rowOutlines, colOutlines) {
  _outlineRows = rowOutlines || {};
  _outlineCols = colOutlines || {};
  _collapsedGroups.clear();
  _renderOutlineButtons(ctx);
}

function clearOutlines() {
  _outlineRows = {};
  _outlineCols = {};
  _collapsedGroups.clear();
  var container = document.getElementById('spreadsheet');
  if (container) {
    container.querySelectorAll('.outline-btn').forEach(function(b) { b.remove(); });
  }
}

function _renderOutlineButtons(ctx) {
  var container = document.getElementById('spreadsheet');
  if (!container) return;
  // 기존 버튼 제거
  container.querySelectorAll('.outline-btn').forEach(function(b) { b.remove(); });

  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  var contentDiv = container.querySelector('.jexcel_content');
  if (!contentDiv) return;

  // 행 그룹 버튼 (행 번호 옆에 +/- 표시)
  var groups = _findGroups(_outlineRows);
  groups.forEach(function(group) {
    var lastRow = group.end;
    var tr = ss.rows && ss.rows[lastRow];
    if (!tr) return;
    var btn = document.createElement('span');
    btn.className = 'outline-btn outline-row-btn';
    var isCollapsed = _collapsedGroups.has('r' + group.start + '-' + group.end);
    btn.textContent = isCollapsed ? '+' : '-';
    btn.title = isCollapsed ? '행 그룹 펼치기' : '행 그룹 접기';
    btn.dataset.start = group.start;
    btn.dataset.end = group.end;
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      var key = 'r' + group.start + '-' + group.end;
      if (_collapsedGroups.has(key)) {
        _collapsedGroups.delete(key);
        _expandRowGroup(ss, group.start, group.end);
        btn.textContent = '-';
        btn.title = '행 그룹 접기';
      } else {
        _collapsedGroups.add(key);
        _collapseRowGroup(ss, group.start, group.end);
        btn.textContent = '+';
        btn.title = '행 그룹 펼치기';
      }
    });
    // 첫 번째 td (행 번호) 옆에 배치
    var firstTd = tr.querySelector('td:first-child');
    if (firstTd) {
      firstTd.style.position = 'relative';
      firstTd.appendChild(btn);
    }
  });
}

function _findGroups(outlines) {
  // 연속된 같은 레벨 행들을 그룹으로 묶기
  var groups = [];
  var entries = Object.entries(outlines)
    .map(function(e) { return { idx: parseInt(e[0]), level: e[1] }; })
    .sort(function(a, b) { return a.idx - b.idx; });
  if (entries.length === 0) return groups;
  var start = entries[0].idx, end = entries[0].idx, level = entries[0].level;
  for (var i = 1; i < entries.length; i++) {
    if (entries[i].idx === end + 1 && entries[i].level === level) {
      end = entries[i].idx;
    } else {
      groups.push({ start: start, end: end, level: level });
      start = entries[i].idx;
      end = entries[i].idx;
      level = entries[i].level;
    }
  }
  groups.push({ start: start, end: end, level: level });
  return groups;
}

function _collapseRowGroup(ss, start, end) {
  if (!ss || !ss.rows) return;
  for (var i = start; i <= end; i++) {
    if (ss.rows[i]) {
      ss.rows[i].style.display = 'none';
      ss.rows[i].classList.add('outline-collapsed');
    }
  }
}

function _expandRowGroup(ss, start, end) {
  if (!ss || !ss.rows) return;
  for (var i = start; i <= end; i++) {
    if (ss.rows[i] && ss.rows[i].classList.contains('outline-collapsed')) {
      ss.rows[i].style.display = '';
      ss.rows[i].classList.remove('outline-collapsed');
    }
  }
}

// ============================================================
// ── 자동 필터 (AutoFilter) ─────────────────────────────────
// ============================================================
var _filterState = {}; // { colIndex: { type: 'values', selected: Set } }
var _filterActive = false;
var _filterHeaderRow = -1; // 필터 헤더가 적용된 행 (0-based, -1 = 미적용)

function initAutoFilter(ctx, headerRow) {
  _filterState = {};
  _filterActive = true;
  _filterHeaderRow = (headerRow !== undefined) ? headerRow : 0;
  _renderFilterButtons(ctx);
}

function clearAutoFilter(ctx) {
  _filterState = {};
  _filterActive = false;
  _filterHeaderRow = -1;
  // 필터 버튼 제거
  var container = document.getElementById('spreadsheet');
  if (container) {
    container.querySelectorAll('.af-btn').forEach(function(b) { b.remove(); });
  }
  // 모든 행 표시
  var ss = ctx.getSpreadsheet();
  if (ss && ss.rows) {
    for (var i = 0; i < ss.rows.length; i++) {
      if (ss.rows[i] && ss.rows[i].style.display === 'none') {
        // 숨겨진 행(hidden_rows)은 유지
        if (!ss.rows[i].classList.contains('filter-hidden')) {
          continue;
        }
        ss.rows[i].style.display = '';
        ss.rows[i].classList.remove('filter-hidden');
      }
    }
  }
}

function isAutoFilterActive() { return _filterActive; }

function _renderFilterButtons(ctx) {
  var container = document.getElementById('spreadsheet');
  if (!container) return;
  // 기존 버튼 제거
  container.querySelectorAll('.af-btn').forEach(function(b) { b.remove(); });

  var ss = ctx.getSpreadsheet();
  if (!ss) return;

  var contentDiv = container.querySelector('.jexcel_content');
  if (!contentDiv) return;

  var thead = contentDiv.querySelector('thead');
  if (!thead) return;
  var headerCells = thead.querySelectorAll('td');

  // 각 열 헤더에 필터 드롭다운 버튼 추가
  headerCells.forEach(function(th, idx) {
    if (idx === 0) return; // 행 번호 열 스킵
    var colIdx = idx - 1;
    var btn = document.createElement('span');
    btn.className = 'af-btn';
    btn.innerHTML = '&#x25BC;';
    btn.title = '필터';
    if (_filterState[colIdx]) {
      btn.classList.add('af-btn-active');
    }
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      e.preventDefault();
      _showFilterDropdown(ctx, colIdx, btn);
    });
    th.style.position = 'relative';
    th.appendChild(btn);
  });
}

function _showFilterDropdown(ctx, colIdx, anchorEl) {
  // 기존 드롭다운 닫기
  var old = document.getElementById('af-dropdown');
  if (old) old.remove();

  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  var data = ss.getData();

  // 해당 열의 고유 값 수집
  var uniqueVals = new Set();
  for (var r = 0; r < data.length; r++) {
    if (r === _filterHeaderRow) continue;
    var val = (data[r] && data[r][colIdx] !== undefined && data[r][colIdx] !== null)
      ? String(data[r][colIdx]) : '';
    uniqueVals.add(val);
  }

  var sorted = Array.from(uniqueVals).sort(function(a, b) {
    var na = parseFloat(a), nb = parseFloat(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });

  // 현재 선택 상태
  var curFilter = _filterState[colIdx];
  var selected = curFilter ? curFilter.selected : new Set(sorted);

  // 드롭다운 DOM 생성
  var dropdown = document.createElement('div');
  dropdown.id = 'af-dropdown';
  dropdown.className = 'af-dropdown';

  // 텍스트 검색
  var search = document.createElement('input');
  search.type = 'text';
  search.placeholder = '검색...';
  search.className = 'af-search';
  dropdown.appendChild(search);

  // 전체 선택/해제
  var allLabel = document.createElement('label');
  allLabel.className = 'af-item af-item-all';
  var allCb = document.createElement('input');
  allCb.type = 'checkbox';
  allCb.checked = selected.size === sorted.length;
  allLabel.appendChild(allCb);
  allLabel.appendChild(document.createTextNode(' (전체 선택)'));
  dropdown.appendChild(allLabel);

  // 값 목록
  var listDiv = document.createElement('div');
  listDiv.className = 'af-list';
  var checkboxes = [];
  sorted.forEach(function(val) {
    var label = document.createElement('label');
    label.className = 'af-item';
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = selected.has(val);
    cb.dataset.val = val;
    label.appendChild(cb);
    label.appendChild(document.createTextNode(' ' + (val || '(빈 셀)')));
    listDiv.appendChild(label);
    checkboxes.push({ cb: cb, label: label, val: val });
  });
  dropdown.appendChild(listDiv);

  // 검색 필터링
  search.addEventListener('input', function() {
    var q = search.value.toLowerCase();
    checkboxes.forEach(function(item) {
      item.label.style.display = item.val.toLowerCase().includes(q) ? '' : 'none';
    });
  });

  // 전체 선택 체크박스
  allCb.addEventListener('change', function() {
    checkboxes.forEach(function(item) {
      if (item.label.style.display !== 'none') {
        item.cb.checked = allCb.checked;
      }
    });
  });

  // 버튼 영역
  var btnBar = document.createElement('div');
  btnBar.className = 'af-btn-bar';
  var okBtn = document.createElement('button');
  okBtn.textContent = '적용';
  okBtn.className = 'af-ok';
  okBtn.addEventListener('click', function() {
    var newSelected = new Set();
    checkboxes.forEach(function(item) {
      if (item.cb.checked) newSelected.add(item.val);
    });
    if (newSelected.size === sorted.length) {
      delete _filterState[colIdx]; // 전체 선택 = 필터 해제
    } else {
      _filterState[colIdx] = { type: 'values', selected: newSelected };
    }
    _applyFilters(ctx);
    _renderFilterButtons(ctx);
    dropdown.remove();
  });
  var cancelBtn = document.createElement('button');
  cancelBtn.textContent = '취소';
  cancelBtn.className = 'af-cancel';
  cancelBtn.addEventListener('click', function() { dropdown.remove(); });
  var clearBtn = document.createElement('button');
  clearBtn.textContent = '필터 해제';
  clearBtn.className = 'af-clear';
  clearBtn.addEventListener('click', function() {
    delete _filterState[colIdx];
    _applyFilters(ctx);
    _renderFilterButtons(ctx);
    dropdown.remove();
  });
  btnBar.appendChild(okBtn);
  btnBar.appendChild(cancelBtn);
  btnBar.appendChild(clearBtn);
  dropdown.appendChild(btnBar);

  // 위치 설정
  var rect = anchorEl.getBoundingClientRect();
  dropdown.style.position = 'fixed';
  dropdown.style.left = rect.left + 'px';
  dropdown.style.top = rect.bottom + 'px';

  document.body.appendChild(dropdown);

  // 외부 클릭으로 닫기
  setTimeout(function() {
    document.addEventListener('mousedown', function handler(e) {
      if (!dropdown.contains(e.target)) {
        dropdown.remove();
        document.removeEventListener('mousedown', handler);
      }
    });
  }, 0);
}

function _applyFilters(ctx) {
  var ss = ctx.getSpreadsheet();
  if (!ss || !ss.rows) return;
  var data = ss.getData();
  var hasActiveFilter = Object.keys(_filterState).length > 0;

  for (var r = 0; r < data.length; r++) {
    if (r === _filterHeaderRow) continue;
    if (!ss.rows[r]) continue;
    var visible = true;
    if (hasActiveFilter) {
      for (var colIdx in _filterState) {
        var filter = _filterState[colIdx];
        var val = (data[r] && data[r][colIdx] !== undefined && data[r][colIdx] !== null)
          ? String(data[r][colIdx]) : '';
        if (!filter.selected.has(val)) {
          visible = false;
          break;
        }
      }
    }
    if (visible) {
      if (ss.rows[r].classList.contains('filter-hidden')) {
        ss.rows[r].style.display = '';
        ss.rows[r].classList.remove('filter-hidden');
      }
    } else {
      ss.rows[r].style.display = 'none';
      ss.rows[r].classList.add('filter-hidden');
    }
  }
}

// ============================================================
// ── 데이터 유효성 검사 (Data Validation) ────────────────────
// ============================================================
var _dataValidations = []; // [{ranges, type, operator, formula1, formula2, allowBlank, ...}]

function setDataValidations(rules) {
  _dataValidations = rules || [];
}

function getDataValidations() { return _dataValidations; }

/**
 * 셀 편집 시 유효성 검증
 * @returns {string|null} 오류 메시지 (null이면 유효)
 */
function validateCellValue(cellName, value) {
  if (!_dataValidations || _dataValidations.length === 0) return null;
  for (var i = 0; i < _dataValidations.length; i++) {
    var rule = _dataValidations[i];
    if (!_cellInRanges(cellName, rule.ranges)) continue;
    var err = _checkRule(rule, value);
    if (err) return err;
  }
  return null;
}

function _cellInRanges(cellName, rangesStr) {
  if (!rangesStr) return false;
  var ranges = rangesStr.split(/\s+/);
  for (var i = 0; i < ranges.length; i++) {
    if (_cellInRange(cellName, ranges[i])) return true;
  }
  return false;
}

function _cellInRange(cellName, rangeStr) {
  // "A1:D10" or "A1"
  var parts = rangeStr.split(':');
  var start = _parseCellRef(parts[0]);
  var end = parts.length > 1 ? _parseCellRef(parts[1]) : start;
  var cell = _parseCellRef(cellName);
  if (!start || !end || !cell) return false;
  return cell.col >= start.col && cell.col <= end.col &&
         cell.row >= start.row && cell.row <= end.row;
}

function _parseCellRef(ref) {
  var m = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  return { col: letterToColIndex(m[1].toUpperCase()), row: parseInt(m[2]) };
}

function _checkRule(rule, value) {
  var type = rule.type || 'none';
  if (type === 'none') return null;
  if (rule.allowBlank && (value === '' || value === null || value === undefined)) return null;

  var errMsg = rule.error || '입력값이 유효하지 않습니다.';

  if (type === 'list') {
    var items = (rule.formula1 || '').split(',').map(function(s) { return s.trim(); });
    if (items.indexOf(String(value)) === -1) {
      return errMsg;
    }
    return null;
  }

  if (type === 'whole' || type === 'decimal') {
    var num = parseFloat(value);
    if (isNaN(num)) return errMsg;
    if (type === 'whole' && num !== Math.floor(num)) return errMsg;
    var v1 = parseFloat(rule.formula1);
    var v2 = parseFloat(rule.formula2);
    var op = rule.operator || 'between';
    if (!_compareOp(op, num, v1, v2)) return errMsg;
    return null;
  }

  if (type === 'textLength') {
    var len = String(value || '').length;
    var v1 = parseInt(rule.formula1) || 0;
    var v2 = parseInt(rule.formula2) || 0;
    var op = rule.operator || 'between';
    if (!_compareOp(op, len, v1, v2)) return errMsg;
    return null;
  }

  if (type === 'date') {
    var d = new Date(value);
    if (isNaN(d.getTime())) return errMsg;
    return null;
  }

  return null; // 미지원 타입은 통과
}

function _compareOp(op, val, v1, v2) {
  switch (op) {
    case 'between': return val >= v1 && val <= v2;
    case 'notBetween': return val < v1 || val > v2;
    case 'equal': return val === v1;
    case 'notEqual': return val !== v1;
    case 'greaterThan': return val > v1;
    case 'lessThan': return val < v1;
    case 'greaterThanOrEqual': return val >= v1;
    case 'lessThanOrEqual': return val <= v1;
    default: return true;
  }
}

/**
 * list 타입 유효성이 있는 셀에 드롭다운 표시
 */
function showValidationDropdown(ctx, col, row) {
  var cellName = colIndexToLetter(col) + (row + 1);
  for (var i = 0; i < _dataValidations.length; i++) {
    var rule = _dataValidations[i];
    if (rule.type === 'list' && _cellInRanges(cellName, rule.ranges)) {
      _showListDropdown(ctx, col, row, rule);
      return true;
    }
  }
  return false;
}

function _showListDropdown(ctx, col, row, rule) {
  var old = document.getElementById('dv-dropdown');
  if (old) old.remove();

  var items = (rule.formula1 || '').split(',').map(function(s) { return s.trim(); });
  var ss = ctx.getSpreadsheet();
  if (!ss) return;

  var td = ss.records[row] && ss.records[row][col];
  if (!td) return;

  var rect = td.getBoundingClientRect();
  var dropdown = document.createElement('div');
  dropdown.id = 'dv-dropdown';
  dropdown.className = 'dv-dropdown';
  dropdown.style.position = 'fixed';
  dropdown.style.left = rect.left + 'px';
  dropdown.style.top = rect.bottom + 'px';
  dropdown.style.minWidth = rect.width + 'px';

  items.forEach(function(item) {
    var opt = document.createElement('div');
    opt.className = 'dv-option';
    opt.textContent = item;
    opt.addEventListener('click', function() {
      ss.setValueFromCoords(col, row, item, true);
      dropdown.remove();
    });
    dropdown.appendChild(opt);
  });

  document.body.appendChild(dropdown);
  setTimeout(function() {
    document.addEventListener('mousedown', function handler(e) {
      if (!dropdown.contains(e.target)) {
        dropdown.remove();
        document.removeEventListener('mousedown', handler);
      }
    });
  }, 0);
}

function hideValidationDropdown() {
  var old = document.getElementById('dv-dropdown');
  if (old) old.remove();
}

// ============================================================
// ── 자동 채우기 (Autofill) ─────────────────────────────────
// ============================================================
function initAutofill(ctx) {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  let handle = document.getElementById('autofill-handle');
  if (!handle) {
    handle = document.createElement('div');
    handle.id = 'autofill-handle';
    handle.className = 'autofill-handle';
    container.appendChild(handle);
  }

  function positionHandle() {
    const ss = ctx.getSpreadsheet();
    if (!ss || !ctx.isEditable()) { handle.style.display = 'none'; return; }
    const sel = ctx.getSelection();
    try {
      const td = container.querySelector('td[data-x="' + sel.x2 + '"][data-y="' + sel.y2 + '"]');
      if (td) {
        const rect = td.getBoundingClientRect();
        const cRect = container.getBoundingClientRect();
        handle.style.display = 'block';
        handle.style.left = (rect.right - cRect.left - 4) + 'px';
        handle.style.top = (rect.bottom - cRect.top - 4) + 'px';
      } else {
        handle.style.display = 'none';
      }
    } catch(e) { handle.style.display = 'none'; }
  }

  ctx._positionAutofillHandle = positionHandle;

  handle.addEventListener('mousedown', function(e) {
    e.preventDefault();
    e.stopPropagation();
    const ss = ctx.getSpreadsheet();
    if (!ss) return;
    const sel = ctx.getSelection();
    // 드래그 방향 결정용 변수
    var fillDir = null;  // 'down', 'up', 'right', 'left'
    var fillEnd = { row: sel.y2, col: sel.x2 };
    handle.classList.add('dragging');

    function onMouseMove(ev) {
      if (!ctx.getSpreadsheet()) return;
      // 마우스가 올라간 셀 찾기
      var target = document.elementFromPoint(ev.clientX, ev.clientY);
      if (!target) return;
      var td = target.closest ? target.closest('td[data-x][data-y]') : null;
      if (!td) return;
      var cx = parseInt(td.getAttribute('data-x'));
      var cy = parseInt(td.getAttribute('data-y'));
      if (isNaN(cx) || isNaN(cy)) return;

      // 선택 영역 밖인 방향 결정 (가장 크게 벗어난 방향 우선)
      var dDown = cy > sel.y2 ? cy - sel.y2 : 0;
      var dUp = cy < sel.y1 ? sel.y1 - cy : 0;
      var dRight = cx > sel.x2 ? cx - sel.x2 : 0;
      var dLeft = cx < sel.x1 ? sel.x1 - cx : 0;
      var maxD = Math.max(dDown, dUp, dRight, dLeft);
      if (maxD === 0) {
        // 선택 영역 내부 — 원래 선택 복원
        fillDir = null;
        fillEnd = { row: sel.y2, col: sel.x2 };
        try { ss.updateSelectionFromCoords(sel.x1, sel.y1, sel.x2, sel.y2); } catch(e) {}
        return;
      }

      if (maxD === dDown) {
        fillDir = 'down'; fillEnd = { row: cy, col: sel.x2 };
        try { ss.updateSelectionFromCoords(sel.x1, sel.y1, sel.x2, cy); } catch(e) {}
      } else if (maxD === dUp) {
        fillDir = 'up'; fillEnd = { row: cy, col: sel.x2 };
        try { ss.updateSelectionFromCoords(sel.x1, cy, sel.x2, sel.y2); } catch(e) {}
      } else if (maxD === dRight) {
        fillDir = 'right'; fillEnd = { row: sel.y2, col: cx };
        try { ss.updateSelectionFromCoords(sel.x1, sel.y1, cx, sel.y2); } catch(e) {}
      } else {
        fillDir = 'left'; fillEnd = { row: sel.y2, col: cx };
        try { ss.updateSelectionFromCoords(cx, sel.y1, sel.x2, sel.y2); } catch(e) {}
      }
    }

    function onMouseUp() {
      handle.classList.remove('dragging');
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      if (fillDir) {
        _performAutofill(ctx, sel.x1, sel.x2, sel.y1, sel.y2, fillDir, fillEnd);
      }
      positionHandle();
    }

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  });
}

// dir: 'down' | 'up' | 'right' | 'left'
function _performAutofill(ctx, x1, x2, y1, y2, dir, end) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  const changes = [];

  if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = true;
  try {
    if (dir === 'down' || dir === 'up') {
      // 세로 채우기: 각 열에 대해 소스 열 값으로 패턴 감지
      for (let c = x1; c <= x2; c++) {
        const srcVals = [];
        for (let r = y1; r <= y2; r++) {
          try { srcVals.push(ss.getValueFromCoords(c, r) || ''); } catch(e) { srcVals.push(''); }
        }
        const pattern = _detectPattern(srcVals);
        if (dir === 'down') {
          for (let r = y2 + 1; r <= end.row; r++) {
            const offset = r - y1;
            const newVal = _patternValue(pattern, srcVals, offset);
            let oldVal = '';
            try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
            try { ss.setValueFromCoords(c, r, newVal); } catch(e) {}
            changes.push({ row: r, col: c, oldVal, newVal });
          }
        } else {
          // up: y1-1 부터 end.row 까지 역순
          for (let r = y1 - 1; r >= end.row; r--) {
            const offset = y1 - r;  // 1, 2, 3, ...
            const newVal = _patternValueReverse(pattern, srcVals, offset);
            let oldVal = '';
            try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
            try { ss.setValueFromCoords(c, r, newVal); } catch(e) {}
            changes.push({ row: r, col: c, oldVal, newVal });
          }
        }
      }
    } else {
      // 가로 채우기: 각 행에 대해 소스 행 값으로 패턴 감지
      for (let r = y1; r <= y2; r++) {
        const srcVals = [];
        for (let c = x1; c <= x2; c++) {
          try { srcVals.push(ss.getValueFromCoords(c, r) || ''); } catch(e) { srcVals.push(''); }
        }
        const pattern = _detectPattern(srcVals);
        if (dir === 'right') {
          for (let c = x2 + 1; c <= end.col; c++) {
            const offset = c - x1;
            const newVal = _patternValue(pattern, srcVals, offset);
            let oldVal = '';
            try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
            try { ss.setValueFromCoords(c, r, newVal); } catch(e) {}
            changes.push({ row: r, col: c, oldVal, newVal });
          }
        } else {
          // left: x1-1 부터 end.col 까지 역순
          for (let c = x1 - 1; c >= end.col; c--) {
            const offset = x1 - c;  // 1, 2, 3, ...
            const newVal = _patternValueReverse(pattern, srcVals, offset);
            let oldVal = '';
            try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
            try { ss.setValueFromCoords(c, r, newVal); } catch(e) {}
            changes.push({ row: r, col: c, oldVal, newVal });
          }
        }
      }
    }
  } finally {
    if (typeof _suppressOnChange !== 'undefined') _suppressOnChange = false;
  }

  if (changes.length > 0) {
    if (ctx.undoManager) ctx.undoManager.push({ type: 'value', changes });
    if (ctx.onAutofill) ctx.onAutofill(changes);
  }
}

// 정방향 패턴 값 (down/right)
function _patternValue(pattern, srcVals, offset) {
  if (pattern.type === 'number_seq') {
    return String(pattern.start + pattern.step * offset);
  }
  return srcVals[offset % srcVals.length];
}

// 역방향 패턴 값 (up/left)
function _patternValueReverse(pattern, srcVals, offset) {
  if (pattern.type === 'number_seq') {
    return String(pattern.start - pattern.step * offset);
  }
  // 반복 패턴: 역순으로 순환
  var len = srcVals.length;
  var idx = ((len - (offset % len)) % len);
  return srcVals[idx];
}

function _detectPattern(values) {
  if (values.length === 0) return { type: 'repeat' };
  if (values.some(v => v === '' || v === null || v === undefined)) return { type: 'repeat' };
  const nums = values.map(Number);
  if (nums.every(n => !isNaN(n))) {
    if (nums.length >= 2) {
      const step = nums[1] - nums[0];
      const isSequence = nums.every((n, i) => i === 0 || Math.abs(n - nums[i-1] - step) < 1e-10);
      if (isSequence) return { type: 'number_seq', start: nums[0], step: step };
    }
    if (nums.length === 1) return { type: 'number_seq', start: nums[0], step: 1 };
  }
  return { type: 'repeat' };
}

// ============================================================
// ── 셀 메모/노트 ──────────────────────────────────────────
// ============================================================
// 셀 메모 데이터 저장 (cellName -> comment text)
var _commentsMap = {};

function showCommentTooltip(ctx, cellName, comment) {
  let tooltip = document.getElementById('cell-comment-tooltip');
  if (!tooltip) {
    tooltip = document.createElement('div');
    tooltip.id = 'cell-comment-tooltip';
    tooltip.className = 'cell-comment-tooltip';
    document.body.appendChild(tooltip);
  }
  tooltip.textContent = comment;
  tooltip.style.display = 'block';
}

function hideCommentTooltip() {
  const tooltip = document.getElementById('cell-comment-tooltip');
  if (tooltip) tooltip.style.display = 'none';
}

var _commentDelegationInstalled = false;

function addCommentIndicators(ctx, comments) {
  _commentsMap = comments || {};
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  // Remove existing indicators
  container.querySelectorAll('.cell-comment-indicator').forEach(el => el.remove());

  // 이벤트 위임: 컨테이너에 한 번만 등록 (리스너 누적 방지)
  if (!_commentDelegationInstalled) {
    _commentDelegationInstalled = true;
    container.addEventListener('mouseenter', function(e) {
      const td = e.target.closest('td[data-x][data-y]');
      if (!td) return;
      const col = parseInt(td.getAttribute('data-x'));
      const row = parseInt(td.getAttribute('data-y'));
      const cellName = colIndexToLetter(col) + (row + 1);
      const c = _commentsMap[cellName];
      if (c) {
        const tooltip = _getOrCreateTooltip();
        tooltip.textContent = c;
        const rect = td.getBoundingClientRect();
        tooltip.style.left = (rect.right + 4) + 'px';
        tooltip.style.top = rect.top + 'px';
        tooltip.style.display = 'block';
      }
    }, true);
    container.addEventListener('mouseleave', function(e) {
      const td = e.target.closest('td[data-x][data-y]');
      if (td) hideCommentTooltip();
    }, true);
  }

  // Add indicators for cells with comments
  for (const [cellName, comment] of Object.entries(_commentsMap)) {
    const m = cellName.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;
    const col = letterToColIndex(m[1]);
    const row = parseInt(m[2]) - 1;
    const td = container.querySelector('td[data-x="' + col + '"][data-y="' + row + '"]');
    if (td) {
      td.style.position = 'relative';
      let indicator = td.querySelector('.cell-comment-indicator');
      if (!indicator) {
        indicator = document.createElement('div');
        indicator.className = 'cell-comment-indicator';
        td.appendChild(indicator);
      }
    }
  }
}

function _getOrCreateTooltip() {
  let tooltip = document.getElementById('cell-comment-tooltip');
  if (!tooltip) {
    tooltip = document.createElement('div');
    tooltip.id = 'cell-comment-tooltip';
    tooltip.className = 'cell-comment-tooltip';
    document.body.appendChild(tooltip);
  }
  return tooltip;
}

function editCellComment(ctx, col, row) {
  const cellName = colIndexToLetter(col) + (row + 1);
  const existing = _commentsMap[cellName] || '';
  const newComment = prompt('셀 메모를 입력하세요:', existing);
  if (newComment === null) return; // cancelled
  if (newComment === existing) return; // no change
  // Update local map
  if (newComment) {
    _commentsMap[cellName] = newComment;
  } else {
    delete _commentsMap[cellName];
  }
  // Re-apply indicators
  addCommentIndicators(ctx, _commentsMap);
  // Notify context for server save
  if (ctx.onCommentChange) {
    ctx.onCommentChange(row, col, newComment || '');
  }
}

function getCommentsMap() {
  return _commentsMap;
}

// ============================================================
// ── 조건부 서식 (클라이언트 측 평가) ────────────────────────
// ============================================================
var _conditionalRules = [];

function applyConditionalFormats(ctx, rules) {
  _conditionalRules = rules || [];
  if (!_conditionalRules.length) return;
  const ss = ctx.getSpreadsheet();
  if (!ss) return;

  const styleUpdates = {};
  for (const rule of _conditionalRules) {
    // xlsx-extracted format: {type, range, operator, formula, format, colors, color, iconStyle}
    if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet') {
      _applyAdvancedCF(ctx, ss, rule);
      continue;
    }
    // xlsx-extracted cellIs/expression format
    if (rule.type === 'cellIs' || rule.type === 'expression') {
      _applyXlsxCF(ctx, ss, rule, styleUpdates);
      continue;
    }
    // Legacy format: {range, rule, value, value2, style}
    if (!rule.range || !rule.rule || !rule.style) continue;
    const cells = _parseCellRange(rule.range);
    for (const [col, row] of cells) {
      let val = '';
      try { val = ss.getValueFromCoords(col, row); } catch(e) { continue; }
      if (_evaluateRule(val, rule.rule, rule.value, rule.value2)) {
        const cellName = colIndexToLetter(col) + (row + 1);
        const css = _condStyleToCss(rule.style);
        if (css) {
          const existingCss = styleUpdates[cellName] || (ss.getStyle(cellName) || '');
          styleUpdates[cellName] = existingCss + ';' + css;
        }
      }
    }
  }
  if (Object.keys(styleUpdates).length > 0) {
    try { ss.setStyle(styleUpdates); } catch(e) {}
  }
}

function _applyXlsxCF(ctx, ss, rule, styleUpdates) {
  if (!rule.range) return;
  var cells = _parseCellRange(rule.range);
  for (var i = 0; i < cells.length; i++) {
    var col = cells[i][0], row = cells[i][1];
    var val = '';
    try { val = ss.getValueFromCoords(col, row); } catch(e) { continue; }
    var match = false;
    if (rule.type === 'cellIs') {
      var op = rule.operator || 'equal';
      var f = rule.formula || [];
      match = _evaluateRule(val, _xlsxOpMap(op), f[0], f[1]);
    } else if (rule.type === 'expression' && rule.formula && rule.formula[0]) {
      match = false; // 수식 평가는 클라이언트에서 제한적
    }
    if (match && rule.format) {
      var cellName = colIndexToLetter(col) + (row + 1);
      var css = _condStyleToCss(rule.format);
      if (css) {
        var existing = styleUpdates[cellName] || (ss.getStyle(cellName) || '');
        styleUpdates[cellName] = existing + ';' + css;
      }
    }
  }
}

function _xlsxOpMap(op) {
  var map = {
    'equal': 'equalTo', 'notEqual': 'notEqualTo',
    'greaterThan': 'greaterThan', 'lessThan': 'lessThan',
    'greaterThanOrEqual': 'greaterThanOrEqual', 'lessThanOrEqual': 'lessThanOrEqual',
    'between': 'between', 'notBetween': 'notBetween',
  };
  return map[op] || op;
}

function _applyAdvancedCF(ctx, ss, rule) {
  if (!rule.range) return;
  var cells = _parseCellRange(rule.range);
  if (cells.length === 0) return;

  // 범위 내 숫자 값 수집
  var numVals = [];
  for (var i = 0; i < cells.length; i++) {
    var val = '';
    try { val = ss.getValueFromCoords(cells[i][0], cells[i][1]); } catch(e) {}
    var num = parseFloat(val);
    numVals.push(isNaN(num) ? null : num);
  }
  var validNums = numVals.filter(function(n) { return n !== null; });
  if (validNums.length === 0) return;
  var minVal = Math.min.apply(null, validNums);
  var maxVal = Math.max.apply(null, validNums);
  var range = maxVal - minVal || 1;

  if (rule.type === 'colorScale') {
    var colors = rule.colors || [];
    if (colors.length < 2) return;
    for (var i = 0; i < cells.length; i++) {
      if (numVals[i] === null) continue;
      var pct = (numVals[i] - minVal) / range;
      var bg = _interpolateColor(colors, pct);
      var cellName = colIndexToLetter(cells[i][0]) + (cells[i][1] + 1);
      try { ss.setStyle(cellName, 'background-color: #' + bg); } catch(e) {}
    }
  } else if (rule.type === 'dataBar') {
    var barColor = rule.color || '4472C4';
    for (var i = 0; i < cells.length; i++) {
      if (numVals[i] === null) continue;
      var pct = Math.round(((numVals[i] - minVal) / range) * 100);
      var cellName = colIndexToLetter(cells[i][0]) + (cells[i][1] + 1);
      var gradient = 'background: linear-gradient(to right, #' + barColor + ' ' + pct + '%, transparent ' + pct + '%)';
      try { ss.setStyle(cellName, gradient); } catch(e) {}
    }
  } else if (rule.type === 'iconSet') {
    // 아이콘: 유니코드 심볼 사용
    var icons = ['🔴', '🟡', '🟢']; // 기본 3단계
    for (var i = 0; i < cells.length; i++) {
      if (numVals[i] === null) continue;
      var pct = ((numVals[i] - minVal) / range) * 100;
      var iconIdx = pct < 33 ? 0 : (pct < 67 ? 1 : 2);
      var td = ss.records && ss.records[cells[i][1]] && ss.records[cells[i][1]][cells[i][0]];
      if (td) {
        var existing = td.innerHTML || '';
        if (existing.indexOf('cf-icon') < 0) {
          td.innerHTML = '<span class="cf-icon">' + icons[iconIdx] + '</span> ' + existing;
        }
      }
    }
  }
}

function _interpolateColor(colors, pct) {
  if (colors.length === 2) {
    return _blendHex(colors[0], colors[1], pct);
  }
  // 3+ colors
  if (pct <= 0.5) {
    return _blendHex(colors[0], colors[1], pct * 2);
  } else {
    return _blendHex(colors[1], colors[2] || colors[1], (pct - 0.5) * 2);
  }
}

function _blendHex(hex1, hex2, t) {
  var r1 = parseInt(hex1.substr(0,2), 16), g1 = parseInt(hex1.substr(2,2), 16), b1 = parseInt(hex1.substr(4,2), 16);
  var r2 = parseInt(hex2.substr(0,2), 16), g2 = parseInt(hex2.substr(2,2), 16), b2 = parseInt(hex2.substr(4,2), 16);
  var r = Math.round(r1 + (r2 - r1) * t), g = Math.round(g1 + (g2 - g1) * t), b = Math.round(b1 + (b2 - b1) * t);
  return ((r < 16 ? '0' : '') + r.toString(16)) + ((g < 16 ? '0' : '') + g.toString(16)) + ((b < 16 ? '0' : '') + b.toString(16));
}

function reapplyConditionalFormats(ctx) {
  if (_conditionalRules.length > 0) {
    applyConditionalFormats(ctx, _conditionalRules);
  }
}

function _parseCellRange(rangeStr) {
  // Parse "A1:C10" or "A1" into array of [col, row] pairs
  const cells = [];
  const parts = rangeStr.split(':');
  if (parts.length === 1) {
    const m = rangeStr.match(/^([A-Z]+)(\d+)$/);
    if (m) cells.push([letterToColIndex(m[1]), parseInt(m[2]) - 1]);
  } else if (parts.length === 2) {
    const m1 = parts[0].match(/^([A-Z]+)(\d+)$/);
    const m2 = parts[1].match(/^([A-Z]+)(\d+)$/);
    if (m1 && m2) {
      const c1 = letterToColIndex(m1[1]), r1 = parseInt(m1[2]) - 1;
      const c2 = letterToColIndex(m2[1]), r2 = parseInt(m2[2]) - 1;
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          cells.push([c, r]);
        }
      }
    }
  }
  return cells;
}

function _evaluateRule(cellValue, ruleType, ruleValue, ruleValue2) {
  const num = Number(cellValue);
  const rv = Number(ruleValue);
  const rv2 = Number(ruleValue2);
  switch (ruleType) {
    case 'greaterThan': return !isNaN(num) && !isNaN(rv) && num > rv;
    case 'lessThan': return !isNaN(num) && !isNaN(rv) && num < rv;
    case 'equalTo': return (!isNaN(num) && !isNaN(rv)) ? num === rv : String(cellValue) === String(ruleValue);
    case 'notEqualTo': return (!isNaN(num) && !isNaN(rv)) ? num !== rv : String(cellValue) !== String(ruleValue);
    case 'greaterThanOrEqual': return !isNaN(num) && !isNaN(rv) && num >= rv;
    case 'lessThanOrEqual': return !isNaN(num) && !isNaN(rv) && num <= rv;
    case 'between': return !isNaN(num) && !isNaN(rv) && !isNaN(rv2) && num >= rv && num <= rv2;
    case 'notBetween': return !isNaN(num) && !isNaN(rv) && !isNaN(rv2) && (num < rv || num > rv2);
    case 'contains': return String(cellValue || '').includes(String(ruleValue || ''));
    case 'notContains': return !String(cellValue || '').includes(String(ruleValue || ''));
    case 'isEmpty': return cellValue === '' || cellValue === null || cellValue === undefined;
    case 'isNotEmpty': return cellValue !== '' && cellValue !== null && cellValue !== undefined;
    default: return false;
  }
}

function _condStyleToCss(style) {
  const parts = [];
  if (style.bg) parts.push('background-color: #' + style.bg);
  if (style.color) parts.push('color: #' + style.color);
  if (style.bold) parts.push('font-weight: bold');
  if (style.italic) parts.push('font-style: italic');
  return parts.join('; ') + (parts.length ? ';' : '');
}

// ============================================================
// ── 인쇄 ──────────────────────────────────────────────────
// ============================================================
function printSpreadsheet() {
  window.print();
}

// ── 열 헤더 갱신 ─────────────────────────────────────────────
// 열 삽입/삭제 후 jspreadsheet 헤더를 A, B, C... 순서로 재설정
function refreshColumnHeaders(ctx) {
  var ss = ctx.getSpreadsheet();
  if (!ss) return;
  var headers = ss.headers || [];
  var cols = ss.options.columns || [];
  for (var i = 0; i < headers.length; i++) {
    var letter = colIndexToLetter(i);
    if (headers[i]) {
      headers[i].textContent = letter;
      headers[i].setAttribute('title', letter);
    }
    if (cols[i]) cols[i].title = letter;
  }
}

// ============================================================
// ── 공개 API ────────────────────────────────────────────────
// ============================================================
return {
  // Utilities
  colIndexToLetter,
  letterToColIndex,
  registerCustomFormulas,
  COLOR_PALETTE,
  initColorSwatches,
  toggleDropdown,
  closeAllDropdowns,
  // Style conversion
  cssToStyleObj,
  styleObjToCss,
  // Toolbar
  getSelectedCellStyle,
  updateToolbarState,
  setActive,
  applyStyleToSelection,
  // Format buttons
  fmtBold,
  fmtItalic,
  fmtUnderline,
  fmtStrikethrough,
  fmtColor,
  fmtBg,
  fmtAlign,
  fmtValign,
  fmtWrap,
  fmtFontSize,
  fmtNumFormat,
  fmtMerge,
  fmtUnmerge,
  fmtBorder,
  // Format Painter
  fmtPainterClick,
  fmtPainterDblClick,
  applyPainterToSelection,
  isPainterActive,
  // Undo/Redo
  UndoManager,
  // Shortcuts
  registerShortcuts,
  deleteSelectedCells,
  // Find/Replace
  openFindPanel,
  closeFindPanel,
  findNext,
  findPrev,
  replaceCurrent,
  replaceAll,
  // Formula bar
  updateFormulaBar,
  handleFormulaBarEnter,
  // Name box
  initNameBox,
  navigateToCell,
  // Status bar
  updateStatusBar,
  // Number format
  formatNumber,
  // Context menu
  buildContextMenu,
  // Sort
  sortColumn,
  // Autofill
  initAutofill,
  // Comments
  showCommentTooltip,
  hideCommentTooltip,
  addCommentIndicators,
  editCellComment,
  getCommentsMap,
  // Conditional Formats
  applyConditionalFormats,
  reapplyConditionalFormats,
  // Print
  printSpreadsheet,
  // Column headers
  refreshColumnHeaders,
  // Hyperlinks
  setHyperlinksMap,
  getHyperlinksMap,
  editHyperlink,
  applyHyperlinkStyles,
  handleHyperlinkClick,
  // Hidden rows/cols
  applyHiddenRows,
  applyHiddenCols,
  // Freeze rows
  applyFreezeRows,
  clearFreezeRows,
  // AutoFilter
  initAutoFilter,
  clearAutoFilter,
  isAutoFilterActive,
  // Data Validation
  setDataValidations,
  getDataValidations,
  validateCellValue,
  showValidationDropdown,
  hideValidationDropdown,
  // Paste Special
  showPasteSpecialDialog,
  executePasteSpecial,
  // Outline/Grouping
  applyOutlines,
  clearOutlines,
};

})();
