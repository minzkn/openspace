// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   template_edit.js — 서식(Template) 전체 편집기
   SpreadsheetCore 공유 모듈 사용

   셀 번호 체계:
     col_index 0 → A열, 1 → B열, ...
     row_index 0 → 1행, 1 → 2행, ...
   ============================================================ */
'use strict';

let templateData = null;
let sheets = [];
let currentSheetIndex = 0;
let spreadsheet = null;
let selectedCol = -1;
let tabClickTimer = null;

let savePending = [];
let saveTimer = null;
const SAVE_DELAY = 800;

// 현재 선택 범위
let selX1 = 0, selY1 = 0, selX2 = 0, selY2 = 0;
// 붙여넣기 시작 위치 (onbeforepaste에서 캡처, onpaste에서 사용)
let _pasteStartRow = 0, _pasteStartCol = 0;

// 숫자 서식 맵 (cellName → numFmt 문자열)
let numFmtMap = {};

// 붙여넣기 중 onchange 억제 플래그 (중복 undo 항목 + 이중 저장 방지)
var _suppressOnChange = false;

// loadSheet 경합 방지 카운터
let _loadSheetSeq = 0;

// 수식 엔진
if (typeof formula !== 'undefined') {
  jspreadsheet.setExtensions({ formula });
}
SpreadsheetCore.registerCustomFormulas();

// ── SpreadsheetCore 컨텍스트 ────────────────────────────────
const ctx = {
  getSpreadsheet: () => spreadsheet,
  getSelection: () => ({ x1: selX1, y1: selY1, x2: selX2, y2: selY2 }),
  isEditable: () => true,
  canMerge: () => true,
  getCurrentSheet: () => sheets[currentSheetIndex],
  onStyleChange: (styleMap) => { saveStyleBatch(styleMap); },
  onMergeChange: () => { saveMerges(); },
  onUndoRedoValue: (changes, isUndo) => {
    // Undo/Redo 시 서버에 셀 값 저장
    const batch = changes.map(c => ({
      row_index: c.row,
      col_index: c.col,
      value: isUndo ? (c.oldVal ?? null) : (c.newVal ?? null),
    }));
    flushBatch(batch);
  },
  onDeleteCells: (changes) => {
    const batch = changes.map(c => ({
      row_index: c.row,
      col_index: c.col,
      value: null,
    }));
    flushBatch(batch);
  },
  onAutofill: (changes) => {
    const batch = changes.map(c => ({
      row_index: c.row,
      col_index: c.col,
      value: c.newVal ?? null,
    }));
    if (batch.length > 0) flushBatch(batch);
  },
  onSort: (changes) => {
    const batch = changes.map(c => ({
      row_index: c.row,
      col_index: c.col,
      value: c.newVal ?? null,
    }));
    if (batch.length > 0) flushBatch(batch);
  },
  onFormulaBarChange: (row, col, value) => {
    enqueueSave(row, col, value);
  },
  onReplaceChange: (changes) => {
    const batch = changes.map(c => ({
      row_index: c.row,
      col_index: c.col,
      value: c.newVal ?? null,
    }));
    if (batch.length > 0) flushBatch(batch);
  },
  onCommentChange: (row, col, comment) => { saveCommentTemplate(row, col, comment); },
  onRowInsert: (rowIndex, direction) => {
    const ss = ctx.getSpreadsheet();
    if (!ss) return;
    ss.insertRow(1, rowIndex, direction === 'above');
    saveAllCells();
  },
  onRowDelete: (rowIndex) => {
    const ss = ctx.getSpreadsheet();
    if (!ss) return;
    ss.deleteRow(rowIndex);
    saveAllCells();
  },
  onRowsDelete: (rowIndices) => {
    const ss = ctx.getSpreadsheet();
    if (!ss) return;
    for (let i = rowIndices.length - 1; i >= 0; i--) {
      ss.deleteRow(rowIndices[i]);
    }
    saveAllCells();
  },
  onColumnInsert: (colIndex, direction) => {
    const ss = ctx.getSpreadsheet();
    if (!ss) return;
    ss.insertColumn(1, colIndex, direction === 'before');
    SpreadsheetCore.refreshColumnHeaders(ctx);
    saveAllCells();
  },
  onColumnDelete: (colIndex) => {
    const ss = ctx.getSpreadsheet();
    if (!ss) return;
    ss.deleteColumn(colIndex);
    SpreadsheetCore.refreshColumnHeaders(ctx);
    saveAllCells();
  },
  onColumnsDelete: (colIndices) => {
    const ss = ctx.getSpreadsheet();
    if (!ss) return;
    const sorted = colIndices.slice().sort((a, b) => b - a);
    for (const ci of sorted) { ss.deleteColumn(ci); }
    SpreadsheetCore.refreshColumnHeaders(ctx);
    saveAllCells();
  },
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

// ── 초기화 ────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const dataEl = document.getElementById('template-data');
  if (!dataEl) return;
  templateData = JSON.parse(dataEl.textContent);
  sheets = templateData.sheets || [];
  SpreadsheetCore.initColorSwatches(ctx);
  SpreadsheetCore.registerShortcuts(ctx);
  initFormulaBar();
  renderTabs();
  if (sheets.length > 0) loadSheet(0);
  document.addEventListener('click', SpreadsheetCore.closeAllDropdowns);
});

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
  const wrap = document.getElementById('sheet-tabs');
  wrap.innerHTML = '';
  sheets.forEach((s, i) => {
    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (i === currentSheetIndex ? ' active' : '');
    tab.innerHTML =
      `<span>${esc(s.sheet_name)}</span>` +
      (sheets.length > 1
        ? `<span class="tab-del" onclick="deleteSheet(${i})" title="시트 삭제">\u00d7</span>`
        : '');
    tab.addEventListener('click', (e) => {
      if (e.target.classList.contains('tab-del')) return;
      if (tabClickTimer) clearTimeout(tabClickTimer);
      tabClickTimer = setTimeout(() => { tabClickTimer = null; switchSheet(i); }, 250);
    });
    tab.addEventListener('dblclick', (e) => {
      if (tabClickTimer) { clearTimeout(tabClickTimer); tabClickTimer = null; }
      if (i !== currentSheetIndex) switchSheet(i);
      renameSheet(i);
    });
    wrap.appendChild(tab);
  });
  const addBtn = document.createElement('button');
  addBtn.className = 'sheet-add-btn';
  addBtn.textContent = '+';
  addBtn.title = '시트 추가';
  addBtn.onclick = addSheet;
  wrap.appendChild(addBtn);
}

function switchSheet(index) {
  if (index === currentSheetIndex && spreadsheet) return;
  flushSave();
  currentSheetIndex = index;
  ctx.undoManager.clear();
  renderTabs();
  loadSheet(index);
}

// ── 시트 로드 ─────────────────────────────────────────────────
async function loadSheet(index) {
  const sheet = sheets[index];
  if (!sheet) return;
  selectedCol = -1;
  hideColPanel();

  const container = document.getElementById('spreadsheet');
  const mySeq = ++_loadSheetSeq;  // 경합 방지

  // ★ destroy를 innerHTML 변경 전에 수행 (DOM 정리 순서 중요)
  if (spreadsheet) {
    try { jspreadsheet.destroy(container); } catch(e) {}
    spreadsheet = null;
  }
  container.innerHTML = '<div style="padding:20px;color:#64748b">로딩 중...</div>';

  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/snapshot`
  );
  if (mySeq !== _loadSheetSeq) return;  // 경합 방지: 이미 다른 시트 로딩 시작됨
  if (!res.ok) {
    container.innerHTML = '<div style="color:red;padding:20px">로딩 실패</div>';
    return;
  }
  const { data } = await res.json();

  const columns = buildColumns(sheet.columns);
  const numRows = Math.max(data.num_rows, 50);

  const gridData = [];
  for (let r = 0; r < numRows; r++) {
    const row = (data.cells[r] || []).slice();
    while (row.length < columns.length) row.push('');
    gridData.push(row);
  }

  const mergeCells = data.merges && Object.keys(data.merges).length > 0 ? data.merges : {};
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
  var th = window.innerHeight - 360;
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
    allowInsertColumn: true,
    allowDeleteColumn: true,
    allowInsertRow: true,
    allowDeleteRow: true,
    mergeCells,
    freezeColumns,
    onchange: handleCellChange,
    onbeforepaste: function() {
      _suppressOnChange = true;
      _pasteStartRow = selY1;
      _pasteStartCol = selX1;
      clearTimeout(window._pasteResetTimer);
      window._pasteResetTimer = setTimeout(function() { _suppressOnChange = false; }, 1000);
    },
    onpaste: handlePaste,
    onbeforechange: handleBeforeChange,
    onselection: handleSelection,
    onmerge: handleMerge,
    onresizerow: handleResizeRow,
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

  attachHeaderClickListeners();

  // 셀 메모 표시 (시트 전환 시 이전 시트 메모 잔존 방지 위해 항상 호출)
  SpreadsheetCore.addCommentIndicators(ctx, data.comments || {});

  // 조건부 서식 적용
  if (data.conditional_formats && data.conditional_formats.length > 0) {
    setTimeout(() => SpreadsheetCore.applyConditionalFormats(ctx, data.conditional_formats), 100);
  }

  // 자동 채우기 핸들 초기화
  SpreadsheetCore.initAutofill(ctx);
}

function buildColumns(cols) {
  return cols.map(c => ({
    title: c.col_header,
    width: c.width || 120,
    type: mapType(c.col_type),
    readOnly: false,
  }));
}

function mapType(t) {
  return { text: 'text', number: 'numeric', date: 'calendar', checkbox: 'checkbox', dropdown: 'dropdown' }[t] || 'text';
}

// ── 헤더 클릭 ─────────────────────────────────────────────────
function attachHeaderClickListeners() {
  const container = document.getElementById('spreadsheet');
  container.querySelectorAll('thead td').forEach((th, i) => {
    if (i === 0) return;
    const colIdx = i - 1;
    th.style.cursor = 'pointer';
    th.addEventListener('click', () => selectColumn(colIdx));
  });
}

function selectColumn(colIdx) {
  selectedCol = colIdx;
  const sheet = sheets[currentSheetIndex];
  const col = sheet.columns[colIdx];
  if (!col) return;
  document.getElementById('cp-header').value = col.col_header;
  document.getElementById('cp-type').value = col.col_type;
  document.getElementById('cp-width').value = col.width || 120;
  document.getElementById('cp-readonly').checked = !!col.is_readonly;
  showColPanel();
}

function handleSelection(el, x1, y1, x2, y2) {
  selX1 = x1; selY1 = y1; selX2 = x2; selY2 = y2;
  if (x1 === x2 && y1 === 0) selectColumn(x1);
  SpreadsheetCore.updateToolbarState(ctx);
  if (ctx._positionAutofillHandle) ctx._positionAutofillHandle();
}

// ── 컬럼 패널 ─────────────────────────────────────────────────
function showColPanel() { document.getElementById('col-panel').classList.add('visible'); }
function hideColPanel() { document.getElementById('col-panel').classList.remove('visible'); }

async function applyColProps() {
  if (selectedCol < 0) { showToast('헤더를 클릭하여 컬럼을 선택하세요', 'warning'); return; }
  const sheet = sheets[currentSheetIndex];
  const col = sheet.columns[selectedCol];
  if (!col) return;

  const payload = {
    col_header: document.getElementById('cp-header').value,
    col_type: document.getElementById('cp-type').value,
    width: parseInt(document.getElementById('cp-width').value) || 120,
    is_readonly: document.getElementById('cp-readonly').checked ? 1 : 0,
  };

  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/columns/${col.id}`,
    { method: 'PATCH', body: JSON.stringify(payload) }
  );
  if (res.ok) {
    Object.assign(col, payload);
    showToast('컬럼 속성이 저장되었습니다', 'success');
    const container = document.getElementById('spreadsheet');
    const ths = container.querySelectorAll('thead td');
    if (ths[selectedCol + 1]) ths[selectedCol + 1].textContent = payload.col_header;
  } else {
    const e = await res.json();
    showToast(e.detail || '저장 실패', 'error');
  }
}

async function addColumn() {
  const sheet = sheets[currentSheetIndex];
  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/columns`,
    { method: 'POST', body: JSON.stringify({ col_header: `Col${sheet.columns.length + 1}` }) }
  );
  if (res.ok) {
    const { data: newCol } = await res.json();
    sheet.columns.push(newCol);
    showToast('컬럼이 추가되었습니다', 'success');
    await loadSheet(currentSheetIndex);
  } else {
    const e = await res.json();
    showToast(e.detail || '추가 실패', 'error');
  }
}

async function deleteColumn() {
  if (selectedCol < 0) { showToast('헤더를 클릭하여 컬럼을 선택하세요', 'warning'); return; }
  const sheet = sheets[currentSheetIndex];
  const col = sheet.columns[selectedCol];
  if (!col) return;
  if (!confirm(`"${col.col_header}" 컬럼과 해당 셀 데이터를 삭제하시겠습니까?`)) return;

  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/columns/${col.id}`,
    { method: 'DELETE' }
  );
  if (res.ok || res.status === 204) {
    sheet.columns.splice(selectedCol, 1);
    selectedCol = -1;
    hideColPanel();
    showToast('컬럼이 삭제되었습니다', 'success');
    await loadSheet(currentSheetIndex);
  } else {
    const e = await res.json();
    showToast(e.detail || '삭제 실패', 'error');
  }
}

// ── 시트 CRUD ─────────────────────────────────────────────────
async function addSheet() {
  const name = prompt('새 시트 이름을 입력하세요:', `Sheet${sheets.length + 1}`);
  if (!name) return;
  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets`,
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

let renamingSheetIndex = -1;

function renameSheet(index) {
  renamingSheetIndex = index;
  showModalFromTemplate('시트 이름 변경', 'rename-sheet-tpl');
  setTimeout(() => {
    const inp = document.getElementById('f-sheet-name');
    if (inp) { inp.value = sheets[index].sheet_name; inp.focus(); inp.select(); }
  }, 50);
}

async function submitRenameSheet(e) {
  e.preventDefault();
  const newName = document.getElementById('f-sheet-name').value.trim();
  if (!newName) return;
  const sheet = sheets[renamingSheetIndex];
  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}`,
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

async function deleteSheet(index) {
  if (sheets.length <= 1) { showToast('마지막 시트는 삭제할 수 없습니다', 'warning'); return; }
  if (!confirm(`"${sheets[index].sheet_name}" 시트와 모든 셀 데이터를 삭제하시겠습니까?`)) return;
  const sheet = sheets[index];
  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}`,
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

// ── 셀 변경 → 자동 저장 ──────────────────────────────────────
function handleBeforeChange(instance, cell, x, y, value) {
  if (_suppressOnChange) return value;
  try {
    const oldVal = instance.getValueFromCoords(parseInt(x), parseInt(y)) || '';
    cell._undoOldVal = oldVal;
  } catch(e) {}
  return value;
}

function handleCellChange(instance, cell, x, y, value) {
  if (_suppressOnChange) return;
  let rawValue = value;
  try {
    const raw = instance.options.data[parseInt(y)][parseInt(x)];
    if (raw !== undefined) rawValue = raw;
  } catch(e) {}
  // Undo support
  const oldVal = cell._undoOldVal || '';
  if (ctx.undoManager && oldVal !== rawValue) {
    ctx.undoManager.push({ type: 'value', changes: [{ row: parseInt(y), col: parseInt(x), oldVal, newVal: rawValue }] });
  }
  enqueueSave(parseInt(y), parseInt(x), rawValue);
}

function handlePaste(instance, data) {
  clearTimeout(window._pasteResetTimer);
  _suppressOnChange = false;
  // data는 2D 텍스트 배열: [[val1, val2], [val3, val4], ...]
  const undoChanges = [];
  const gridData = instance.options.data;
  data.forEach((rowData, ri) => {
    if (!Array.isArray(rowData)) return;
    const row = _pasteStartRow + ri;
    rowData.forEach((cellVal, ci) => {
      const col = _pasteStartCol + ci;
      let rawValue = cellVal != null ? String(cellVal) : '';
      try {
        if (gridData && row < gridData.length && gridData[row] && col < gridData[row].length) {
          const raw = gridData[row][col];
          if (raw !== undefined && raw !== null) rawValue = String(raw);
        }
      } catch(e) {}
      undoChanges.push({ row, col, oldVal: '', newVal: rawValue });
      enqueueSave(row, col, rawValue);
    });
  });
  if (undoChanges.length > 0 && ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes: undoChanges });
  }
}

function enqueueSave(row, col, value) {
  savePending = savePending.filter(p => !(p.row_index === row && p.col_index === col));
  savePending.push({ row_index: row, col_index: col, value: value ?? null });
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flushSave, SAVE_DELAY);
  setSaveStatus('입력 중...');
}

async function flushSave() {
  if (!savePending.length) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const batch = savePending.splice(0);
  setSaveStatus('저장 중...');
  showSaveIndicator();

  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/cells`,
    { method: 'POST', body: JSON.stringify(batch) }
  );
  hideSaveIndicator();
  if (res.ok) setSaveStatus('저장됨');
  else setSaveStatus('저장 실패');
  setTimeout(() => setSaveStatus(''), 2000);
}

async function flushBatch(batch) {
  if (!batch.length) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  try {
    const res = await apiFetch(
      `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/cells`,
      { method: 'POST', body: JSON.stringify(batch) }
    );
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      showToast(e.detail || '셀 저장 실패', 'error');
    }
  } catch(e) {
    showToast('서버 연결 실패 — 변경사항이 저장되지 않을 수 있습니다', 'error');
  }
}

// 행/열 삽입·삭제 후 전체 셀 재동기화 (replace=true로 기존 셀 삭제 후 재저장)
async function saveAllCells() {
  const ss = ctx.getSpreadsheet();
  const sheet = sheets[currentSheetIndex];
  if (!ss || !sheet) return;

  // 대기 중인 저장 버리기 (전체 동기화로 대체)
  savePending = [];
  clearTimeout(saveTimer);

  const data = ss.getData();
  const batch = [];
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < (data[r] || []).length; c++) {
      const val = data[r][c];
      if (val !== '' && val !== null && val !== undefined) {
        batch.push({ row_index: r, col_index: c, value: String(val) });
      }
    }
  }

  setSaveStatus('동기화 중...');
  showSaveIndicator();
  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/cells?replace=true`,
    { method: 'POST', body: JSON.stringify(batch) }
  );
  hideSaveIndicator();
  if (res.ok) setSaveStatus('저장됨');
  else setSaveStatus('저장 실패');
  setTimeout(() => setSaveStatus(''), 2000);
}

function setSaveStatus(msg) {
  const el = document.getElementById('save-status');
  if (el) el.textContent = msg;
}
function showSaveIndicator() { const el = document.getElementById('save-indicator'); if (el) el.classList.add('show'); }
function hideSaveIndicator() { const el = document.getElementById('save-indicator'); if (el) el.classList.remove('show'); }

function downloadTemplate() {
  const a = document.createElement('a');
  a.href = `/api/admin/templates/${templateData.id}/export-xlsx`;
  a.download = `${templateData.name}.xlsx`;
  a.click();
}

window.addEventListener('beforeunload', (e) => {
  if (savePending.length > 0) {
    e.preventDefault();
    e.returnValue = '저장되지 않은 변경사항이 있습니다.';
  }
});

// ── 스타일 서버 저장 ────────────────────────────────────────
async function saveStyleBatch(styleMap) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const batch = [];
  for (const [cellName, css] of Object.entries(styleMap)) {
    const m = cellName.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;
    const col_index = SpreadsheetCore.letterToColIndex(m[1]);
    const row_index = parseInt(m[2]) - 1;
    const s = SpreadsheetCore.cssToStyleObj(css);
    batch.push({ row_index, col_index, style: JSON.stringify(s) });
    // numFmtMap 갱신 (숫자 서식 실시간 표시용)
    if (s.numFmt) {
      numFmtMap[cellName] = s.numFmt;
    } else {
      delete numFmtMap[cellName];
    }
  }
  if (!batch.length) return;
  await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/cells`,
    { method: 'POST', body: JSON.stringify(batch) }
  );
}

// ── 셀 메모 저장 ─────────────────────────────────────────────
async function saveCommentTemplate(row, col, comment) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/cells`,
    { method: 'POST', body: JSON.stringify([{ row_index: row, col_index: col, comment: comment || '' }]) }
  );
}

// ── 병합 핸들러 ───────────────────────────────────────────────
function handleMerge(el, x, y, colspan, rowspan) {
  setTimeout(saveMerges, 100);
}

async function saveMerges() {
  if (!spreadsheet) return;
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
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/merges`,
    { method: 'PATCH', body: JSON.stringify({ merges: mergesList }) }
  );
}

// ── 행 높이 변경 핸들러 ─────────────────────────────────────
let rowHeightSaveTimer = null;
function handleResizeRow(el, row, height) {
  clearTimeout(rowHeightSaveTimer);
  rowHeightSaveTimer = setTimeout(saveRowHeights, 500);
}

async function saveRowHeights() {
  if (!spreadsheet) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const rowHeights = {};
  try {
    const rows = spreadsheet.rows;
    if (rows) {
      for (let i = 0; i < rows.length; i++) {
        if (rows[i] && rows[i].style && rows[i].style.height) {
          const px = parseFloat(rows[i].style.height);
          if (px && px !== 0) {
            rowHeights[String(i)] = Math.round(px / 1.333 * 10) / 10;
          }
        }
      }
    }
  } catch(e) {}
  await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/row-heights`,
    { method: 'PATCH', body: JSON.stringify({ row_heights: rowHeights }) }
  );
}
