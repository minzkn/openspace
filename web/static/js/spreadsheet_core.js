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
  if (!ss) return;
  if (!ctx.canMerge()) return;
  const sel = ctx.getSelection();
  if (sel.x1 === sel.x2 && sel.y1 === sel.y2) return; // 단일 셀 병합 불가
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
    var cellName = colIndexToLetter(sel.x1) + (sel.y1 + 1);
    ss.removeMerge(cellName);
    if (ctx.onMergeChange) ctx.onMergeChange();
  } catch(e) { if (typeof showToast === 'function') showToast('병합 해제 실패', 'error'); }
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
  const styleMap = {};
  for (let r = sel.y1; r <= sel.y2; r++) {
    for (let c = sel.x1; c <= sel.x2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = ss.getStyle(cellName) || '';
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

    if (action.type === 'value') {
      // { type: 'value', changes: [{row, col, oldVal, newVal}] }
      const changes = action.changes;
      for (const c of changes) {
        const val = isUndo ? c.oldVal : c.newVal;
        try { ss.setValueFromCoords(c.col, c.row, val || '', true); } catch(e) {}
      }
      ctx.onUndoRedoValue(changes, isUndo);
    }
    else if (action.type === 'style') {
      // { type: 'style', cells: [{cellName, oldCss, newCss}] }
      const styleMap = {};
      for (const c of action.cells) {
        styleMap[c.cellName] = isUndo ? c.oldCss : c.newCss;
      }
      try { ss.setStyle(styleMap); } catch(e) {}
      ctx.onStyleChange(styleMap);
    }
    else if (action.type === 'batch') {
      // combined value+style action
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

    // Escape: close find panel
    if (e.key === 'Escape') {
      closeFindPanel();
      return;
    }

    // Formula bar Enter/Escape handled separately
    if (e.target.id === 'formula-input') return;

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
  });
}

function deleteSelectedCells(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss || !ctx.isEditable()) return;
  const sel = ctx.getSelection();
  const changes = [];
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
  if (changes.length > 0) {
    if (ctx.undoManager) ctx.undoManager.push({ type: 'value', changes });
    ctx.onDeleteCells(changes);
  }
}

// ============================================================
// ── 찾기 / 바꾸기 ───────────────────────────────────────────
// ============================================================
let _findState = { results: [], current: -1, lastQuery: '' };

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
  _findState = { results: [], current: -1, lastQuery: '' };
  updateFindCount();
}

function findNext(ctx) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  const query = (document.getElementById('find-input') || {}).value || '';
  if (!query) return;
  const caseSensitive = (document.getElementById('find-case') || {}).checked || false;

  // rebuild results if query changed
  if (query !== _findState.lastQuery) {
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

  if (query !== _findState.lastQuery) {
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
  try { oldVal = ss.getValueFromCoords(r.col, r.row) || ''; } catch(e) {}
  const query = (document.getElementById('find-input') || {}).value || '';
  const caseSensitive = (document.getElementById('find-case') || {}).checked || false;
  const regex = new RegExp(_escapeRegex(query), caseSensitive ? 'g' : 'gi');
  const newVal = oldVal.replace(regex, replaceVal);
  try { ss.setValueFromCoords(r.col, r.row, newVal); } catch(e) {}
  if (ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes: [{ row: r.row, col: r.col, oldVal, newVal }] });
  }
  // rebuild results
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
  for (const r of _findState.results) {
    let oldVal = '';
    try { oldVal = ss.getValueFromCoords(r.col, r.row) || ''; } catch(e) {}
    const newVal = oldVal.replace(regex, replaceVal);
    if (newVal !== oldVal) {
      try { ss.setValueFromCoords(r.col, r.row, newVal); } catch(e) {}
      changes.push({ row: r.row, col: r.col, oldVal, newVal });
    }
  }
  if (changes.length > 0 && ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes });
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
  // Cell reference display
  const startName = colIndexToLetter(sel.x1) + (sel.y1 + 1);
  if (sel.x1 === sel.x2 && sel.y1 === sel.y2) {
    cellRef.textContent = startName;
  } else {
    cellRef.textContent = startName + ':' + colIndexToLetter(sel.x2) + (sel.y2 + 1);
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
  try { ss.setValueFromCoords(sel.x1, sel.y1, newVal); } catch(e) {}
  if (ctx.undoManager && oldVal !== newVal) {
    ctx.undoManager.push({ type: 'value', changes: [{ row: sel.y1, col: sel.x1, oldVal, newVal }] });
  }
  // Blur to return focus to spreadsheet
  inputEl.blur();
}

// ============================================================
// ── 상태 표시줄 (Status Bar) ────────────────────────────────
// ============================================================
function updateStatusBar(ctx) {
  const bar = document.getElementById('status-bar');
  if (!bar) return;
  const ss = ctx.getSpreadsheet();
  if (!ss) { bar.textContent = ''; return; }
  const sel = ctx.getSelection();
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
  const parts = [];
  parts.push('개수: ' + count);
  if (numCount > 0) {
    parts.push('합계: ' + _formatNum(sum));
    parts.push('평균: ' + _formatNum(sum / numCount));
    parts.push('최소: ' + _formatNum(min));
    parts.push('최대: ' + _formatNum(max));
  }
  bar.textContent = parts.join('   |   ');
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

    // 잘라내기 / 복사 / 붙여넣기
    items.push({ title: '잘라내기', onclick: function() { obj.cut(); } });
    items.push({ title: '복사', onclick: function() { obj.copy(); } });
    items.push({ title: '붙여넣기', onclick: function() {
      navigator.clipboard.readText().then(function(text) {
        if (text) obj.paste(obj.selectedCell[0], obj.selectedCell[1], text);
      }).catch(function() {});
    }});

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
        if (ctx.onRowDelete) ctx.onRowDelete(cy);
        else obj.deleteRow(cy);
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
        if (ctx.onColumnDelete) ctx.onColumnDelete(cx);
        else obj.deleteColumn(cx);
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

    return items;
  };
}

// ============================================================
// ── 열 정렬 (클라이언트 사이드) ────────────────────────────
// ============================================================
function sortColumn(ctx, colIdx, ascending) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
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
  const rows = [];
  for (let r = 0; r <= lastRow; r++) {
    rows.push({ idx: r, data: (data[r] || []).slice() });
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

  // Apply sorted data
  const changes = [];
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].data.length; c++) {
      let oldVal = '';
      try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
      const newVal = rows[r].data[c] || '';
      if (oldVal !== newVal) {
        try { ss.setValueFromCoords(c, r, newVal, true); } catch(e) {}
        changes.push({ row: r, col: c, oldVal, newVal: String(newVal) });
      }
    }
  }
  if (changes.length > 0 && ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes });
  }
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

  let isDragging = false;
  let startRow = 0, startCol = 0, endRow = 0;

  function positionHandle() {
    const ss = ctx.getSpreadsheet();
    if (!ss || !ctx.isEditable()) { handle.style.display = 'none'; return; }
    const sel = ctx.getSelection();
    try {
      const cellName = colIndexToLetter(sel.x2) + (sel.y2 + 1);
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

  // Reposition after selection changes
  ctx._positionAutofillHandle = positionHandle;

  handle.addEventListener('mousedown', function(e) {
    e.preventDefault();
    e.stopPropagation();
    isDragging = true;
    const sel = ctx.getSelection();
    startRow = sel.y1;
    startCol = sel.x1;
    endRow = sel.y2;
    handle.classList.add('dragging');

    function onMouseMove(ev) {
      const ss = ctx.getSpreadsheet();
      if (!ss) return;
      // Find which row the mouse is over
      const tds = container.querySelectorAll('td[data-x="' + sel.x2 + '"]');
      for (const td of tds) {
        const rect = td.getBoundingClientRect();
        if (ev.clientY >= rect.top && ev.clientY <= rect.bottom) {
          const newRow = parseInt(td.getAttribute('data-y'));
          if (!isNaN(newRow) && newRow > sel.y2) {
            endRow = newRow;
            // Visual highlight
            try { ss.updateSelectionFromCoords(sel.x1, sel.y1, sel.x2, newRow); } catch(e) {}
          }
          break;
        }
      }
    }

    function onMouseUp() {
      isDragging = false;
      handle.classList.remove('dragging');
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      if (endRow > sel.y2) {
        _performAutofill(ctx, sel.x1, sel.x2, sel.y1, sel.y2, endRow);
      }
      positionHandle();
    }

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  });
}

function _performAutofill(ctx, x1, x2, y1, y2, targetRow) {
  const ss = ctx.getSpreadsheet();
  if (!ss) return;
  const srcHeight = y2 - y1 + 1;
  const changes = [];

  for (let c = x1; c <= x2; c++) {
    // Collect source values
    const srcVals = [];
    for (let r = y1; r <= y2; r++) {
      try { srcVals.push(ss.getValueFromCoords(c, r) || ''); } catch(e) { srcVals.push(''); }
    }

    // Detect pattern
    const pattern = _detectPattern(srcVals);

    for (let r = y2 + 1; r <= targetRow; r++) {
      const offset = r - y1;
      let newVal;
      if (pattern.type === 'number_seq') {
        newVal = String(pattern.start + pattern.step * offset);
      } else {
        // repeat pattern
        newVal = srcVals[offset % srcHeight];
      }
      let oldVal = '';
      try { oldVal = ss.getValueFromCoords(c, r) || ''; } catch(e) {}
      try { ss.setValueFromCoords(c, r, newVal); } catch(e) {}
      changes.push({ row: r, col: c, oldVal, newVal });
    }
  }

  if (changes.length > 0 && ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes });
  }
}

function _detectPattern(values) {
  if (values.length === 0) return { type: 'repeat' };
  // Check if all values are numbers
  const nums = values.map(Number);
  if (nums.every(n => !isNaN(n) && values[0] !== '')) {
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

function addCommentIndicators(ctx, comments) {
  _commentsMap = comments || {};
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  // Remove existing indicators
  container.querySelectorAll('.cell-comment-indicator').forEach(el => el.remove());
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
      // Hover tooltip
      td.addEventListener('mouseenter', function() {
        const c = _commentsMap[cellName];
        if (c) {
          const tooltip = _getOrCreateTooltip();
          tooltip.textContent = c;
          const rect = td.getBoundingClientRect();
          tooltip.style.left = (rect.right + 4) + 'px';
          tooltip.style.top = rect.top + 'px';
          tooltip.style.display = 'block';
        }
      });
      td.addEventListener('mouseleave', hideCommentTooltip);
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
    if (!rule.range || !rule.rule || !rule.style) continue;
    const cells = _parseCellRange(rule.range);
    for (const [col, row] of cells) {
      let val = '';
      try { val = ss.getValueFromCoords(col, row); } catch(e) { continue; }
      if (_evaluateRule(val, rule.rule, rule.value, rule.value2)) {
        const cellName = colIndexToLetter(col) + (row + 1);
        const css = _condStyleToCss(rule.style);
        if (css) {
          styleUpdates[cellName] = (styleUpdates[cellName] || '') + css;
        }
      }
    }
  }
  if (Object.keys(styleUpdates).length > 0) {
    try { ss.setStyle(styleUpdates); } catch(e) {}
  }
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
    case 'equalTo': return String(cellValue) === String(ruleValue);
    case 'notEqualTo': return String(cellValue) !== String(ruleValue);
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
};

})();
