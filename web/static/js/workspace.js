// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   workspace.js — Jspreadsheet + WebSocket 실시간 협업
   SpreadsheetCore 공유 모듈 사용

   셀 번호 체계:
     col_index 0 → A열, 1 → B열, 2 → C열, ...
     row_index 0 → 행1, 1 → 행2, ...
   ============================================================ */
'use strict';

let workspaceData = null;
let currentSheetIndex = 0;
let sheets = [];
let spreadsheet = null;
let ws = null;
let wsReconnectTimer = null;
let isClosed = false;
let pendingPatches = [];
let batchTimer = null;
let renamingWsSheetIndex = -1;

// 현재 선택 범위
let selX1 = 0, selY1 = 0, selX2 = 0, selY2 = 0;

// 숫자 서식 맵 (cellName → numFmt 문자열)
let numFmtMap = {};

const BATCH_DELAY = 100;

// 수식 엔진
if (typeof formula !== 'undefined') {
  jspreadsheet.setExtensions({ formula });
}
SpreadsheetCore.registerCustomFormulas();

// ── SpreadsheetCore 컨텍스트 ────────────────────────────────
const ctx = {
  getSpreadsheet: () => spreadsheet,
  getSelection: () => ({ x1: selX1, y1: selY1, x2: selX2, y2: selY2 }),
  isEditable: () => !isClosed || IS_ADMIN,
  canMerge: () => IS_ADMIN,
  getCurrentSheet: () => sheets[currentSheetIndex],
  onStyleChange: (styleMap) => {
    // CSS → style JSON → WebSocket patch
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    const patches = [];
    for (const [cellName, css] of Object.entries(styleMap)) {
      const m = cellName.match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      const col = SpreadsheetCore.letterToColIndex(m[1]);
      const row = parseInt(m[2]) - 1;
      const s = SpreadsheetCore.cssToStyleObj(css);
      patches.push({ row, col, value: null, style: JSON.stringify(s) });
      // numFmtMap 갱신 (숫자 서식 실시간 표시용)
      if (s.numFmt) {
        numFmtMap[cellName] = s.numFmt;
      } else {
        delete numFmtMap[cellName];
      }
    }
    if (patches.length > 0) sendBatchPatch(sheet.id, patches);
  },
  onMergeChange: () => { saveMerges(); },
  onUndoRedoValue: (changes, isUndo) => {
    // Undo/Redo 값 변경 시 WebSocket으로 전송
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    const patches = changes.map(c => ({
      row: c.row, col: c.col,
      value: isUndo ? c.oldVal : c.newVal,
      style: null,
    }));
    if (patches.length > 0) sendBatchPatch(sheet.id, patches);
  },
  onDeleteCells: (changes) => {
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    const patches = changes.map(c => ({ row: c.row, col: c.col, value: '', style: null }));
    if (patches.length > 0) sendBatchPatch(sheet.id, patches);
  },
  onRowInsert: (rowIndex, direction) => { insertRowApi(rowIndex, direction); },
  onRowDelete: (rowIndex) => { deleteRowApi(rowIndex); },
  onColumnInsert: (colIndex, direction) => { insertColApi(colIndex, direction); },
  onColumnDelete: (colIndex) => { deleteColApi(colIndex); },
  onCommentChange: (row, col, comment) => { saveComment(row, col, comment); },
  undoManager: new SpreadsheetCore.UndoManager(),
};

// ── 전역 함수 바인딩 (HTML onclick 호환) ────────────────────
function colIndexToLetter(idx) { return SpreadsheetCore.colIndexToLetter(idx); }
function letterToColIndex(l) { return SpreadsheetCore.letterToColIndex(l); }
function toggleDropdown(id) { SpreadsheetCore.toggleDropdown(id); }
function closeAllDropdowns() { SpreadsheetCore.closeAllDropdowns(); }
function fmtBold() { SpreadsheetCore.fmtBold(ctx); }
function fmtItalic() { SpreadsheetCore.fmtItalic(ctx); }
function fmtUnderline() { SpreadsheetCore.fmtUnderline(ctx); }
function fmtStrikethrough() { SpreadsheetCore.fmtStrikethrough(ctx); }
function fmtColor(hex) { SpreadsheetCore.fmtColor(ctx, hex); }
function fmtBg(hex) { SpreadsheetCore.fmtBg(ctx, hex); }
function fmtAlign(dir) { SpreadsheetCore.fmtAlign(ctx, dir); }
function fmtValign(dir) { SpreadsheetCore.fmtValign(ctx, dir); }
function fmtWrap() { SpreadsheetCore.fmtWrap(ctx); }
function fmtFontSize(size) { SpreadsheetCore.fmtFontSize(ctx, size); }
function fmtNumFormat(fmt) { SpreadsheetCore.fmtNumFormat(ctx, fmt); }
function fmtMerge() { SpreadsheetCore.fmtMerge(ctx); }
function fmtUnmerge() { SpreadsheetCore.fmtUnmerge(ctx); }
function fmtBorder(preset) { SpreadsheetCore.fmtBorder(ctx, preset); }
function fmtBorderStyled(preset) {
  const style = (document.getElementById('border-style-select') || {}).value || 'thin';
  const color = (document.getElementById('border-color-val') || {}).value || '000000';
  SpreadsheetCore.fmtBorder(ctx, preset, style, color);
}
function findNext() { SpreadsheetCore.findNext(ctx); }
function findPrev() { SpreadsheetCore.findPrev(ctx); }
function replaceCurrent() { SpreadsheetCore.replaceCurrent(ctx); }
function replaceAll() { SpreadsheetCore.replaceAll(ctx); }
function printSheet() { SpreadsheetCore.printSpreadsheet(); }

document.addEventListener('DOMContentLoaded', init);

function init() {
  const dataEl = document.getElementById('workspace-data');
  if (!dataEl) return;
  workspaceData = JSON.parse(dataEl.textContent);
  sheets = workspaceData.sheets || [];
  isClosed = workspaceData.status === 'CLOSED';
  SpreadsheetCore.initColorSwatches(ctx);
  SpreadsheetCore.registerShortcuts(ctx);
  initFormulaBar();
  renderTabs();
  if (sheets.length > 0) loadSheet(0);
  connectWebSocket();
  document.addEventListener('click', SpreadsheetCore.closeAllDropdowns);
}

function initFormulaBar() {
  const input = document.getElementById('formula-input');
  if (!input) return;
  input.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      SpreadsheetCore.handleFormulaBarEnter(ctx, input);
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      SpreadsheetCore.updateFormulaBar(ctx);
      input.blur();
    }
  });
}

// ── 시트 탭 렌더링 ────────────────────────────────────────────
function renderTabs() {
  const tabsEl = document.getElementById('sheet-tabs');
  tabsEl.innerHTML = sheets.map((s, i) => {
    const isActive = i === currentSheetIndex;
    const delBtn = (IS_ADMIN && sheets.length > 1)
      ? `<span class="tab-del" onclick="deleteWsSheet(${i})" title="삭제">\u00d7</span>`
      : '';
    return `<div class="sheet-tab ${isActive ? 'active' : ''}"
        onclick="handleTabClick(event, ${i})">
      <span ${IS_ADMIN ? `ondblclick="renameWsSheet(${i})" title="더블클릭으로 이름 변경"` : ''}>${esc(s.sheet_name)}</span>${delBtn}
    </div>`;
  }).join('');
}

function handleTabClick(e, index) {
  if (e.target.classList.contains('tab-del')) return;
  switchSheet(index);
}

function switchSheet(index) {
  if (index === currentSheetIndex && spreadsheet) return;
  currentSheetIndex = index;
  ctx.undoManager.clear();
  renderTabs();
  loadSheet(index);
}

// ── 시트 로드 ─────────────────────────────────────────────────
async function loadSheet(index) {
  const sheet = sheets[index];
  if (!sheet) return;

  const container = document.getElementById('spreadsheet');

  // ★ destroy를 innerHTML 변경 전에 수행 (DOM 정리 순서 중요)
  if (spreadsheet) {
    try { jspreadsheet.destroy(container); } catch(e) {}
    spreadsheet = null;
  }
  container.innerHTML = '<div style="padding:20px;color:#64748b">로딩 중...</div>';

  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/snapshot`
  );
  if (!res.ok) {
    container.innerHTML = '<div style="color:red;padding:20px">로딩 실패</div>';
    return;
  }
  const { data } = await res.json();

  const isEditable = !isClosed || IS_ADMIN;
  let columns = buildColumnDefs(sheet.columns, isEditable);
  // snapshot에서 num_cols가 더 크면 컬럼 보충
  if (columns.length < data.num_cols) {
    for (let ci = columns.length; ci < data.num_cols; ci++) {
      columns.push({ title: SpreadsheetCore.colIndexToLetter(ci), width: 120, type: 'text', readOnly: false });
    }
  }
  // 최소 5열 보장
  if (columns.length === 0) {
    for (let ci = 0; ci < Math.max(data.num_cols, 5); ci++) {
      columns.push({ title: SpreadsheetCore.colIndexToLetter(ci), width: 120, type: 'text', readOnly: false });
    }
  }
  const numRows = Math.max(data.num_rows, 100);

  const gridData = [];
  for (let r = 0; r < numRows; r++) {
    const row = (data.cells[r] || []).slice();
    while (row.length < columns.length) row.push('');
    gridData.push(row);
  }

  const mergeCells = data.merges && Object.keys(data.merges).length > 0 ? data.merges : undefined;
  const freezeColumns = data.freeze_columns > 0 ? data.freeze_columns : undefined;

  // 숫자 서식 맵 초기화
  numFmtMap = data.num_formats || {};

  // ★ 컨테이너 완전 초기화 + 스크롤 위치 리셋
  container.innerHTML = '';
  container.scrollLeft = 0;
  container.scrollTop = 0;

  // ★ 부모 요소에서 안정적인 너비 계산
  var parentEl = container.parentElement || container;
  var tw = parentEl.offsetWidth || container.offsetWidth || (window.innerWidth - 260);
  if (tw < 100) tw = window.innerWidth - 260;
  var th = window.innerHeight - 280;
  if (th < 200) th = 400;

  spreadsheet = jspreadsheet(container, {
    data: gridData,
    columns,
    minDimensions: [columns.length, numRows],
    tableOverflow: true,
    tableWidth: tw + 'px',
    tableHeight: th + 'px',
    lazyLoading: true,
    loadingSpin: true,
    editable: isEditable,
    allowInsertColumn: true,
    allowDeleteColumn: true,
    mergeCells,
    freezeColumns,
    onchange: handleCellChange,
    onpaste: handlePaste,
    onbeforechange: handleBeforeChange,
    onselection: handleSelection,
    onmerge: IS_ADMIN ? handleMerge : undefined,
    onresizerow: IS_ADMIN ? handleResizeRow : undefined,
    contextMenu: SpreadsheetCore.buildContextMenu(ctx),
    updateTable: function(instance, cell, col, row, val, label, cellName) {
      // cellName이 없는 경우 직접 계산 (jspreadsheet 버전 호환)
      var cn = cellName || (SpreadsheetCore.colIndexToLetter(col) + (row + 1));
      var fmt = numFmtMap[cn];
      if (fmt) {
        var formatted = SpreadsheetCore.formatNumber(val, fmt);
        if (formatted !== null) cell.innerHTML = formatted;
      }
    },
  });

  // ★ 시트 전환 후 스크롤 위치 초기화
  var wrapper = container.querySelector('.jexcel_content');
  if (wrapper) { wrapper.scrollLeft = 0; wrapper.scrollTop = 0; }

  // 스타일 적용
  if (data.styles && Object.keys(data.styles).length > 0) {
    try { spreadsheet.setStyle(data.styles); } catch(e) {}
  }

  // 행 높이 적용
  if (data.row_heights) {
    Object.entries(data.row_heights).forEach(([riStr, px]) => {
      try { spreadsheet.setHeight(parseInt(riStr), px); } catch(e) {}
    });
  }

  // 포맷 툴바 표시 여부
  const toolbar = document.getElementById('format-toolbar');
  if (toolbar) toolbar.style.display = isEditable ? 'flex' : 'none';
  const formulaBar = document.getElementById('formula-bar');
  if (formulaBar) formulaBar.style.display = isEditable ? 'flex' : 'none';

  // 셀 메모 표시
  if (data.comments && Object.keys(data.comments).length > 0) {
    SpreadsheetCore.addCommentIndicators(ctx, data.comments);
  }

  // 조건부 서식 적용
  if (data.conditional_formats && data.conditional_formats.length > 0) {
    setTimeout(() => SpreadsheetCore.applyConditionalFormats(ctx, data.conditional_formats), 100);
  }

  // 자동 채우기 핸들 초기화
  SpreadsheetCore.initAutofill(ctx);
}

function buildColumnDefs(columns, isEditable) {
  return columns.map(c => ({
    title: c.col_header,
    width: c.width || 120,
    type: mapColType(c.col_type),
    readOnly: c.is_readonly && !IS_ADMIN,
  }));
}

function mapColType(t) {
  return { text: 'text', number: 'numeric', date: 'calendar',
           checkbox: 'checkbox', dropdown: 'dropdown' }[t] || 'text';
}

// ── 셀 변경 핸들러 ────────────────────────────────────────────
function handleBeforeChange(instance, cell, x, y, value) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return value;
  const col = sheet.columns[x];
  if (col && col.is_readonly && !IS_ADMIN) return false;
  // Capture old value for undo
  try {
    const oldVal = instance.getValueFromCoords(parseInt(x), parseInt(y)) || '';
    cell._undoOldVal = oldVal;
  } catch(e) {}
  return value;
}

function handleCellChange(instance, cell, x, y, value) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  let rawValue = value;
  try {
    const gridData = instance.options.data;
    const iy = parseInt(y), ix = parseInt(x);
    if (gridData && iy < gridData.length && gridData[iy] && ix < gridData[iy].length) {
      const raw = gridData[iy][ix];
      if (raw !== undefined && raw !== null) rawValue = String(raw);
    }
  } catch(e) {}

  // Undo support
  const oldVal = cell._undoOldVal || '';
  if (ctx.undoManager && oldVal !== rawValue) {
    ctx.undoManager.push({ type: 'value', changes: [{ row: parseInt(y), col: parseInt(x), oldVal, newVal: rawValue }] });
  }

  queuePatch(sheet.id, parseInt(y), parseInt(x), rawValue, null);
}

function handlePaste(instance, data) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const patches = [];
  const undoChanges = [];
  data.forEach(item => {
    let rawValue = item[3];
    let oldVal = '';
    try {
      const gridData = instance.options.data;
      const iy = item[0], ix = item[1];
      oldVal = instance.getValueFromCoords(ix, iy) || '';
      if (gridData && iy < gridData.length && gridData[iy] && ix < gridData[iy].length) {
        const raw = gridData[iy][ix];
        if (raw !== undefined && raw !== null) rawValue = String(raw);
      }
    } catch(e) {}
    patches.push({ row: item[0], col: item[1], value: rawValue });
    undoChanges.push({ row: item[0], col: item[1], oldVal, newVal: rawValue });
  });
  if (undoChanges.length > 0 && ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes: undoChanges });
  }
  if (patches.length > 0) sendBatchPatch(sheet.id, patches);
}

function handleSelection(el, x1, y1, x2, y2) {
  selX1 = x1; selY1 = y1; selX2 = x2; selY2 = y2;
  SpreadsheetCore.updateToolbarState(ctx);
  if (ctx._positionAutofillHandle) ctx._positionAutofillHandle();
}

// ── 패치 큐 ───────────────────────────────────────────────────
function queuePatch(sheetId, row, col, value, style) {
  pendingPatches.push({ sheetId, row, col, value, style });
  clearTimeout(batchTimer);
  batchTimer = setTimeout(flushPatches, BATCH_DELAY);
}

function flushPatches() {
  if (!pendingPatches.length) return;
  const bySheet = {};
  pendingPatches.forEach(p => {
    if (!bySheet[p.sheetId]) bySheet[p.sheetId] = [];
    bySheet[p.sheetId].push({ row: p.row, col: p.col, value: p.value, style: p.style });
  });
  pendingPatches = [];
  Object.entries(bySheet).forEach(([sheetId, patches]) => {
    if (patches.length === 1) {
      const p = patches[0];
      sendPatch(sheetId, p.row, p.col, p.value, p.style);
    } else {
      sendBatchPatch(sheetId, patches);
    }
  });
}

function sendPatch(sheetId, row, col, value, style) {
  if (ws && ws.readyState === WebSocket.OPEN)
    ws.send(JSON.stringify({ type: 'patch', sheet_id: sheetId, row, col, value, style }));
}

function sendBatchPatch(sheetId, patches) {
  if (ws && ws.readyState === WebSocket.OPEN)
    ws.send(JSON.stringify({ type: 'batch_patch', sheet_id: sheetId, patches }));
}

// ── WebSocket ─────────────────────────────────────────────────
function connectWebSocket() {
  if (ws) {
    ws.onclose = null; ws.onerror = null;
    if (ws.readyState !== WebSocket.CLOSED) ws.close();
    ws = null;
  }
  clearTimeout(wsReconnectTimer);
  setConnStatus('connecting');
  const sessionId = getCookie('session_id');
  const proto = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const url = `${proto}//${location.host}/ws/workspaces/${workspaceData.id}?session_id=${encodeURIComponent(sessionId || '')}`;

  ws = new WebSocket(url);
  ws.onopen = () => { setConnStatus('connected'); clearTimeout(wsReconnectTimer); };
  ws.onclose = () => { setConnStatus('disconnected'); wsReconnectTimer = setTimeout(connectWebSocket, 3000); };
  ws.onerror = () => setConnStatus('disconnected');
  ws.onmessage = (event) => {
    let msg;
    try { msg = JSON.parse(event.data); } catch { return; }
    handleWsMessage(msg);
  };
}

function handleWsMessage(msg) {
  if (msg.type === 'pong') return;

  if (msg.type === 'patch') {
    applyRemotePatch(msg);
    showActivity(`${msg.updated_by} 편집 중`);
    return;
  }
  if (msg.type === 'batch_patch') {
    msg.patches.forEach(p => applyRemotePatch({ ...p, sheet_id: msg.sheet_id, updated_by: msg.updated_by }));
    showActivity(`${msg.updated_by} 편집 중`);
    return;
  }
  if (msg.type === 'row_insert' || msg.type === 'row_delete') {
    handleRemoteRowOp(msg);
    return;
  }
  if (msg.type === 'col_insert' || msg.type === 'col_delete') {
    handleRemoteColOp(msg);
    return;
  }
  if (msg.type === 'workspace_status') {
    isClosed = msg.status === 'CLOSED';
    updateStatusBadge(msg.status);
    if (isClosed && !IS_ADMIN) {
      showToast('관리자가 워크스페이스를 마감했습니다.', 'warning');
      loadSheet(currentSheetIndex);
    } else if (!isClosed) {
      showToast('워크스페이스가 재개되었습니다.', 'success');
      loadSheet(currentSheetIndex);
    }
    return;
  }
  if (msg.type === 'reload') {
    showToast('데이터가 업로드되었습니다. 새로고침합니다.', 'info');
    setTimeout(() => loadSheet(currentSheetIndex), 800);
    return;
  }
  if (msg.type === 'sheet_added') {
    sheets.push(msg.sheet);
    renderTabs();
    showToast(`시트 "${msg.sheet.sheet_name}" 추가됨`, 'info');
    return;
  }
  if (msg.type === 'sheet_deleted') {
    const idx = sheets.findIndex(s => s.id === msg.sheet_id);
    if (idx >= 0) {
      const wasActive = idx === currentSheetIndex;
      sheets.splice(idx, 1);
      if (currentSheetIndex >= sheets.length) currentSheetIndex = sheets.length - 1;
      else if (idx < currentSheetIndex) currentSheetIndex--;
      renderTabs();
      if (wasActive) loadSheet(currentSheetIndex);
      showToast('시트가 삭제됨', 'info');
    }
    return;
  }
  if (msg.type === 'sheet_renamed') {
    const s = sheets.find(s => s.id === msg.sheet_id);
    if (s) { s.sheet_name = msg.sheet_name; renderTabs(); }
    return;
  }
  if (msg.type === 'error') showToast(msg.message || '오류 발생', 'error');
}

function applyRemotePatch(msg) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet || msg.sheet_id !== sheet.id) return;
  if (!spreadsheet) return;
  if (msg.value !== undefined && msg.value !== null) {
    try { spreadsheet.setValueFromCoords(msg.col, msg.row, msg.value || '', true); } catch(e) {}
  }
  if (msg.style) {
    try {
      const cellName = SpreadsheetCore.colIndexToLetter(msg.col) + (msg.row + 1);
      const s = JSON.parse(msg.style);
      const css = SpreadsheetCore.styleObjToCss(s);
      spreadsheet.setStyle({ [cellName]: css });
      // numFmtMap 갱신 (원격 사용자의 숫자 서식 변경)
      if (s.numFmt) {
        numFmtMap[cellName] = s.numFmt;
      } else {
        delete numFmtMap[cellName];
      }
    } catch(e) {}
  }
  if (msg.comment !== undefined && msg.comment !== null) {
    const cellName = SpreadsheetCore.colIndexToLetter(msg.col) + (msg.row + 1);
    const commentsMap = SpreadsheetCore.getCommentsMap();
    if (msg.comment) {
      commentsMap[cellName] = msg.comment;
    } else {
      delete commentsMap[cellName];
    }
    SpreadsheetCore.addCommentIndicators(ctx, commentsMap);
  }
}

function handleRemoteRowOp(msg) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet || msg.sheet_id !== sheet.id) return;
  if (!spreadsheet) return;
  // 자신이 발행한 메시지는 이미 로컬 적용됨 → 스킵
  if (msg.updated_by === CURRENT_USER) return;
  if (msg.type === 'row_insert') {
    try { spreadsheet.insertRow(msg.count || 1, msg.row_index, true); } catch(e) {}
  } else if (msg.type === 'row_delete') {
    const indices = msg.row_indices || [msg.row_index];
    for (let i = indices.length - 1; i >= 0; i--) {
      try { spreadsheet.deleteRow(indices[i]); } catch(e) {}
    }
  }
}

// ── 행 삽입/삭제 API ──────────────────────────────────────────
async function insertRowApi(rowIndex, direction) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const actualRow = direction === 'below' ? rowIndex + 1 : rowIndex;
  // 로컬 즉시 적용
  try { spreadsheet.insertRow(1, actualRow, true); } catch(e) {}
  // 서버 API
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/rows/insert`,
    { method: 'POST', body: JSON.stringify({ row_index: actualRow, count: 1, direction }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '행 삽입 실패', 'error');
  }
}

async function deleteRowApi(rowIndex) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  try { spreadsheet.deleteRow(rowIndex); } catch(e) {}
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/rows/delete`,
    { method: 'POST', body: JSON.stringify({ row_indices: [rowIndex] }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '행 삭제 실패', 'error');
  }
}

// ── 열 삽입/삭제 API ──────────────────────────────────────────
async function insertColApi(colIndex, direction) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  // jspreadsheet CE v4: insertColumn(num, colIndex, insertBefore)
  // insertBefore=1 → 왼쪽, insertBefore=0 → 오른쪽
  try { spreadsheet.insertColumn(1, colIndex, direction === 'before' ? 1 : 0); } catch(e) {}
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/cols/insert`,
    { method: 'POST', body: JSON.stringify({ col_index: colIndex, count: 1, direction }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '열 삽입 실패', 'error');
  }
}

async function deleteColApi(colIndex) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  try { spreadsheet.deleteColumn(colIndex); } catch(e) {}
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/cols/delete`,
    { method: 'POST', body: JSON.stringify({ col_indices: [colIndex] }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '열 삭제 실패', 'error');
  }
}

function handleRemoteColOp(msg) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet || msg.sheet_id !== sheet.id) return;
  if (!spreadsheet) return;
  // 자신이 발행한 메시지는 이미 로컬 적용됨 → 스킵
  if (msg.updated_by === CURRENT_USER) return;
  if (msg.type === 'col_insert') {
    try { spreadsheet.insertColumn(msg.count || 1, msg.col_index, true); } catch(e) {}
  } else if (msg.type === 'col_delete') {
    const indices = msg.col_indices || [msg.col_index];
    for (let i = indices.length - 1; i >= 0; i--) {
      try { spreadsheet.deleteColumn(indices[i]); } catch(e) {}
    }
  }
}

// ── 셀 메모 저장 ─────────────────────────────────────────────
async function saveComment(row, col, comment) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/patches`,
    { method: 'POST', body: JSON.stringify({ patches: [{ row, col, comment: comment || '' }] }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '메모 저장 실패', 'error');
  }
}

// ── UI helpers ────────────────────────────────────────────────
function setConnStatus(status) {
  const dot = document.getElementById('ws-conn-indicator');
  if (!dot) return;
  dot.className = `conn-dot ${status}`;
  dot.title = { connected: '연결됨', disconnected: '연결 끊김', connecting: '연결 중...' }[status];
}

let activityTimer;
function showActivity(text) {
  const el = document.getElementById('ws-activity');
  if (!el) return;
  el.textContent = text;
  clearTimeout(activityTimer);
  activityTimer = setTimeout(() => { el.textContent = ''; }, 2000);
}

function updateStatusBadge(status) {
  const badge = document.getElementById('ws-status-badge');
  if (badge) { badge.textContent = status; badge.className = `badge ${status.toLowerCase()}`; }
  const closeBtn = document.getElementById('close-btn');
  if (closeBtn) closeBtn.textContent = status === 'OPEN' ? '마감' : '재개';
}

// ── 관리자: 워크스페이스 시트 관리 ───────────────────────────
async function addWsSheet() {
  const name = prompt('새 시트 이름을 입력하세요:', `Sheet${sheets.length + 1}`);
  if (!name) return;
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets`,
    { method: 'POST', body: JSON.stringify({ sheet_name: name }) }
  );
  if (res.ok) {
    const { data: newSheet } = await res.json();
    sheets.push(newSheet);
    renderTabs();
    switchSheet(sheets.length - 1);
    showToast(`"${name}" 시트가 추가되었습니다`, 'success');
  } else {
    const e = await res.json();
    showToast(e.detail || '추가 실패', 'error');
  }
}

function renameWsSheet(index) {
  renamingWsSheetIndex = index;
  showModalFromTemplate('시트 이름 변경', 'rename-ws-sheet-tpl');
  setTimeout(() => {
    const inp = document.getElementById('f-ws-sheet-name');
    if (inp) { inp.value = sheets[index].sheet_name; inp.focus(); inp.select(); }
  }, 50);
}

async function submitRenameWsSheet(e) {
  e.preventDefault();
  const newName = document.getElementById('f-ws-sheet-name').value.trim();
  if (!newName) return;
  const sheet = sheets[renamingWsSheetIndex];
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}`,
    { method: 'PATCH', body: JSON.stringify({ sheet_name: newName }) }
  );
  if (res.ok) {
    sheet.sheet_name = newName;
    closeModal();
    renderTabs();
    showToast('이름이 변경되었습니다', 'success');
  } else {
    const e2 = await res.json();
    showToast(e2.detail || '변경 실패', 'error');
  }
}

async function deleteWsSheet(index) {
  if (sheets.length <= 1) { showToast('마지막 시트는 삭제할 수 없습니다', 'warning'); return; }
  if (!confirm(`"${sheets[index].sheet_name}" 시트를 삭제하시겠습니까?`)) return;
  const sheet = sheets[index];
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}`,
    { method: 'DELETE' }
  );
  if (res.ok || res.status === 204) {
    sheets.splice(index, 1);
    if (currentSheetIndex >= sheets.length) currentSheetIndex = sheets.length - 1;
    renderTabs();
    loadSheet(currentSheetIndex);
    showToast('시트가 삭제되었습니다', 'success');
  } else {
    const e = await res.json();
    showToast(e.detail || '삭제 실패', 'error');
  }
}

// ── 관리자: 마감/재개/다운로드/업로드 ────────────────────────
async function toggleClose() {
  const action = isClosed ? 'reopen' : 'close';
  const label = isClosed ? '재개' : '마감';
  if (!confirm(`워크스페이스를 ${label}하시겠습니까?`)) return;
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/${action}`,
    { method: 'POST' }
  );
  if (!res.ok) { const e = await res.json(); showToast(e.detail || '오류', 'error'); }
}

function exportWorkspace() {
  const a = document.createElement('a');
  a.href = `/api/admin/workspaces/${workspaceData.id}/export-xlsx`;
  a.download = `${workspaceData.name}.xlsx`;
  a.click();
}

async function importWorkspace(input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const res = await apiFetch(`/api/admin/workspaces/${workspaceData.id}/import-xlsx`, {
    method: 'POST', body: fd,
    headers: { 'X-CSRF-Token': getCookie('csrf_token') },
  });
  input.value = '';
  if (res.ok) showToast('업로드 완료, 모든 사용자 화면이 새로고침됩니다', 'success');
  else { const e = await res.json(); showToast(e.detail || '업로드 실패', 'error'); }
}

// ── 병합 ──────────────────────────────────────────────────────
function handleMerge(el, x, y, colspan, rowspan) {
  setTimeout(saveMerges, 100);
}

async function saveMerges() {
  if (!spreadsheet || !IS_ADMIN) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  let mergeMap = {};
  try { mergeMap = spreadsheet.getMerge() || {}; } catch(e) {}
  const mergesList = [];
  for (const [cellName, dims] of Object.entries(mergeMap)) {
    const m = cellName.match(/^([A-Z]+)(\d+)$/);
    if (!m || !dims || dims.length < 2) continue;
    const startCol = SpreadsheetCore.letterToColIndex(m[1]);
    const startRow = parseInt(m[2]) - 1;
    const endCol = startCol + dims[0] - 1;
    const endRow = startRow + dims[1] - 1;
    const endName = SpreadsheetCore.colIndexToLetter(endCol) + (endRow + 1);
    if (cellName !== endName) mergesList.push(`${cellName}:${endName}`);
  }
  await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/merges`,
    { method: 'PATCH', body: JSON.stringify({ merges: mergesList }) }
  );
}

// ── 행 높이 변경 핸들러 ─────────────────────────────────────
let rowHeightSaveTimer = null;
function handleResizeRow(el, row, height) {
  // 디바운스: 연속 리사이즈 시 마지막 변경만 저장
  clearTimeout(rowHeightSaveTimer);
  rowHeightSaveTimer = setTimeout(saveRowHeights, 500);
}

async function saveRowHeights() {
  if (!spreadsheet || !IS_ADMIN) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const rowHeights = {};
  // jspreadsheet CE v4: getHeight(row) 또는 rows 배열에서 높이 추출
  try {
    const rows = spreadsheet.rows;
    if (rows) {
      for (let i = 0; i < rows.length; i++) {
        if (rows[i] && rows[i].style && rows[i].style.height) {
          const px = parseFloat(rows[i].style.height);
          if (px && px !== 0) {
            // px → pt 변환 (Excel 행 높이는 pt 단위)
            rowHeights[String(i)] = Math.round(px / 1.333 * 10) / 10;
          }
        }
      }
    }
  } catch(e) {}
  await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/row-heights`,
    { method: 'PATCH', body: JSON.stringify({ row_heights: rowHeights }) }
  );
}

// Keep-alive
setInterval(() => {
  if (ws && ws.readyState === WebSocket.OPEN)
    ws.send(JSON.stringify({ type: 'ping' }));
}, 25000);
