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
let tabClickTimer = null;

// 다중 탭 구분을 위한 고유 탭 ID (동일 사용자의 여러 탭에서 브로드캐스트 충돌 방지)
const TAB_ID = crypto.randomUUID ? crypto.randomUUID() : (Date.now().toString(36) + Math.random().toString(36).slice(2));

// 현재 선택 범위
let selX1 = 0, selY1 = 0, selX2 = 0, selY2 = 0;
// 붙여넣기 시작 위치 (onbeforepaste에서 캡처, onpaste에서 사용)
let _pasteStartRow = 0, _pasteStartCol = 0;

// 숫자 서식 맵 (cellName → numFmt 문자열)
let numFmtMap = {};

// 프로그래밍적 셀 변경 중 onchange 억제 플래그 (피드백 루프 및 중복 전송 방지)
// var 사용: spreadsheet_core.js IIFE에서도 접근 가능하도록 window 프로퍼티로 등록
var _suppressOnChange = false;

// loadSheet 경합 방지 카운터 (빠른 시트 전환 시 이전 API 응답이 현재 시트를 덮어쓰는 것 방지)
let _loadSheetSeq = 0;

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
  onPasteSpecial: (changes) => {
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    const patches = changes.map(c => ({ row: c.row, col: c.col, value: c.newVal, style: null }));
    if (patches.length > 0) sendBatchPatch(sheet.id, patches);
  },
  onAutofill: (changes) => {
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    const patches = changes.map(c => ({ row: c.row, col: c.col, value: c.newVal, style: null }));
    if (patches.length > 0) sendBatchPatch(sheet.id, patches);
  },
  onSort: (changes) => {
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    const patches = changes.map(c => ({ row: c.row, col: c.col, value: c.newVal, style: null }));
    if (patches.length > 0) sendBatchPatch(sheet.id, patches);
  },
  onFormulaBarChange: (row, col, value) => {
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    queuePatch(sheet.id, row, col, value, null);
  },
  onReplaceChange: (changes) => {
    const sheet = sheets[currentSheetIndex];
    if (!sheet) return;
    const patches = changes.map(c => ({ row: c.row, col: c.col, value: c.newVal, style: null }));
    if (patches.length > 0) sendBatchPatch(sheet.id, patches);
  },
  onRowInsert: (rowIndex, direction) => { insertRowApi(rowIndex, direction); },
  onRowDelete: (rowIndex) => { deleteRowApi(rowIndex); },
  onRowsDelete: (rowIndices) => { deleteRowsApi(rowIndices); },
  onColumnInsert: (colIndex, direction) => { insertColApi(colIndex, direction); },
  onColumnDelete: (colIndex) => { deleteColApi(colIndex); },
  onColumnsDelete: (colIndices) => { deleteColsApi(colIndices); },
  onCommentChange: (row, col, comment) => { saveComment(row, col, comment); },
  onHyperlinkChange: (row, col, url) => { saveHyperlink(row, col, url); },
  onHideRows: IS_ADMIN ? (rows) => { hideRows(rows); } : undefined,
  onUnhideRows: IS_ADMIN ? () => { unhideRows(); } : undefined,
  onHideCols: IS_ADMIN ? (cols) => { hideCols(cols); } : undefined,
  onUnhideCols: IS_ADMIN ? () => { unhideCols(); } : undefined,
  onColumnProps: IS_ADMIN ? (colIndex) => { showColumnPropsModal(colIndex); } : undefined,
  onFreezeSetup: IS_ADMIN ? () => { showFreezeDialog(); } : undefined,
  onSheetProtection: IS_ADMIN ? () => { toggleSheetProtection(); } : undefined,
  onPrintSetup: IS_ADMIN ? () => { showPrintSettingsDialog(); } : undefined,
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
function fmtPainter() { SpreadsheetCore.fmtPainterClick(ctx); }
function fmtPainterDbl() { SpreadsheetCore.fmtPainterDblClick(ctx); }

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
  SpreadsheetCore.initNameBox(ctx);
  renderTabs();
  if (sheets.length > 0) loadSheet(0);
  connectWebSocket();
  document.addEventListener('click', SpreadsheetCore.closeAllDropdowns);
  // 페이지 닫기 시 미전송 패치 플러시 (데이터 손실 방지)
  window.addEventListener('beforeunload', function() {
    flushPatches();
  });
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
    const drag = IS_ADMIN ? `draggable="true" data-tab-idx="${i}"` : '';
    return `<div class="sheet-tab ${isActive ? 'active' : ''}" ${drag}
        onclick="handleTabClick(event, ${i})" ${IS_ADMIN ? `ondblclick="handleTabDblClick(event, ${i})"` : ''}
        ${IS_ADMIN ? `oncontextmenu="showTabContextMenu(event, ${i})"` : ''}>
      <span>${esc(s.sheet_name)}</span>${delBtn}
    </div>`;
  }).join('');
  if (IS_ADMIN) {
    const addBtn = document.createElement('button');
    addBtn.className = 'sheet-add-btn';
    addBtn.textContent = '+';
    addBtn.title = '시트 추가';
    addBtn.onclick = addWsSheet;
    tabsEl.appendChild(addBtn);
    // Drag-and-drop reorder
    initTabDragDrop(tabsEl);
  }
}

function initTabDragDrop(tabsEl) {
  let dragIdx = -1;
  tabsEl.querySelectorAll('.sheet-tab[draggable]').forEach(tab => {
    tab.addEventListener('dragstart', function(e) {
      dragIdx = parseInt(this.dataset.tabIdx);
      this.classList.add('tab-dragging');
      e.dataTransfer.effectAllowed = 'move';
    });
    tab.addEventListener('dragend', function() {
      this.classList.remove('tab-dragging');
      tabsEl.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('tab-drag-over'));
    });
    tab.addEventListener('dragover', function(e) {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      this.classList.add('tab-drag-over');
    });
    tab.addEventListener('dragleave', function() {
      this.classList.remove('tab-drag-over');
    });
    tab.addEventListener('drop', function(e) {
      e.preventDefault();
      this.classList.remove('tab-drag-over');
      const dropIdx = parseInt(this.dataset.tabIdx);
      if (dragIdx < 0 || dragIdx === dropIdx) return;
      // Reorder sheets array
      const [moved] = sheets.splice(dragIdx, 1);
      sheets.splice(dropIdx, 0, moved);
      // Adjust current sheet index
      if (currentSheetIndex === dragIdx) {
        currentSheetIndex = dropIdx;
      } else if (dragIdx < currentSheetIndex && dropIdx >= currentSheetIndex) {
        currentSheetIndex--;
      } else if (dragIdx > currentSheetIndex && dropIdx <= currentSheetIndex) {
        currentSheetIndex++;
      }
      renderTabs();
      saveSheetOrder();
    });
  });
}

async function saveSheetOrder() {
  const order = sheets.map(s => s.id);
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets-order`,
    { method: 'PATCH', body: JSON.stringify({ order }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '순서 저장 실패', 'error');
  }
}

function handleTabClick(e, index) {
  if (e.target.classList.contains('tab-del')) return;
  if (tabClickTimer) clearTimeout(tabClickTimer);
  tabClickTimer = setTimeout(() => { tabClickTimer = null; switchSheet(index); }, 250);
}

function handleTabDblClick(e, index) {
  if (tabClickTimer) { clearTimeout(tabClickTimer); tabClickTimer = null; }
  if (index !== currentSheetIndex) switchSheet(index);
  renameWsSheet(index);
}

function showTabContextMenu(e, index) {
  e.preventDefault();
  // 기존 컨텍스트 메뉴 제거
  var old = document.getElementById('tab-context-menu');
  if (old) old.remove();

  var menu = document.createElement('div');
  menu.id = 'tab-context-menu';
  menu.className = 'tab-context-menu';
  menu.innerHTML =
    '<div class="tcm-item" data-action="copy">시트 복사</div>' +
    '<div class="tcm-item" data-action="rename">이름 변경</div>' +
    (sheets.length > 1 ? '<div class="tcm-item tcm-danger" data-action="delete">시트 삭제</div>' : '');

  menu.style.position = 'fixed';
  menu.style.left = e.clientX + 'px';
  menu.style.top = e.clientY + 'px';
  document.body.appendChild(menu);

  menu.addEventListener('click', function(ev) {
    var action = ev.target.dataset.action;
    if (action === 'copy') copySheet(index);
    else if (action === 'rename') renameWsSheet(index);
    else if (action === 'delete') deleteWsSheet(index);
    menu.remove();
  });

  setTimeout(function() {
    document.addEventListener('mousedown', function handler(ev) {
      if (!menu.contains(ev.target)) {
        menu.remove();
        document.removeEventListener('mousedown', handler);
      }
    });
  }, 0);
}

async function copySheet(index) {
  var sheet = sheets[index];
  if (!sheet) return;
  var res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/copy`,
    { method: 'POST' }
  );
  if (res.ok) {
    var data = await res.json();
    sheets.push({
      id: data.id,
      sheet_name: data.name,
      sheet_index: data.sheet_index,
      columns: sheet.columns,
    });
    renderTabs();
    switchSheet(sheets.length - 1);
    showToast('시트가 복사되었습니다.');
  } else {
    var e = await res.json().catch(() => ({}));
    showToast(e.detail || '시트 복사 실패', 'error');
  }
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
  const mySeq = ++_loadSheetSeq;  // 경합 방지: 이 호출의 고유 시퀀스 번호

  // ★ destroy를 innerHTML 변경 전에 수행 (DOM 정리 순서 중요)
  if (spreadsheet) {
    try { jspreadsheet.destroy(container); } catch(e) {}
    spreadsheet = null;
  }
  container.innerHTML = '<div style="padding:20px;color:#64748b">로딩 중...</div>';

  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/snapshot`
  );
  // 경합 방지: API 응답 도착 시 다른 loadSheet가 이미 시작되었으면 폐기
  if (mySeq !== _loadSheetSeq) return;
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
    columnResize: true,
    rowResize: true,
    allowInsertColumn: true,
    allowDeleteColumn: true,
    mergeCells,
    freezeColumns,
    onchange: handleCellChange,
    onbeforepaste: function() {
      _suppressOnChange = true;
      // 붙여넣기 시작 위치 캡처 (onpaste에서 사용)
      _pasteStartRow = selY1;
      _pasteStartCol = selX1;
      // Safety: onpaste가 호출되지 않는 예외 상황 대비 (편집 불가 방지)
      clearTimeout(window._pasteResetTimer);
      window._pasteResetTimer = setTimeout(function() { _suppressOnChange = false; }, 1000);
    },
    onpaste: handlePaste,
    onbeforechange: handleBeforeChange,
    onselection: handleSelection,
    onmerge: IS_ADMIN ? handleMerge : undefined,
    onresizerow: IS_ADMIN ? handleResizeRow : undefined,
    onresizecolumn: IS_ADMIN ? handleResizeColumn : undefined,
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

  // 열 너비 적용
  if (data.col_widths) {
    Object.entries(data.col_widths).forEach(([ciStr, px]) => {
      try { spreadsheet.setWidth(parseInt(ciStr), px); } catch(e) {}
    });
  }

  // 포맷 툴바 표시 여부
  const toolbar = document.getElementById('format-toolbar');
  if (toolbar) toolbar.style.display = isEditable ? 'flex' : 'none';
  const formulaBar = document.getElementById('formula-bar');
  if (formulaBar) formulaBar.style.display = isEditable ? 'flex' : 'none';

  // 셀 메모 표시 (시트 전환 시 이전 시트 메모 잔존 방지 위해 항상 호출)
  SpreadsheetCore.addCommentIndicators(ctx, data.comments || {});

  // 조건부 서식 적용
  if (data.conditional_formats && data.conditional_formats.length > 0) {
    setTimeout(() => SpreadsheetCore.applyConditionalFormats(ctx, data.conditional_formats), 100);
  }

  // 하이퍼링크 표시
  SpreadsheetCore.setHyperlinksMap(data.hyperlinks || {});
  setTimeout(() => SpreadsheetCore.applyHyperlinkStyles(ctx), 100);

  // 데이터 유효성 검사 로드
  SpreadsheetCore.setDataValidations(data.data_validations || []);

  // 시트 보호 상태 저장
  var curSheet = sheets[currentSheetIndex];
  if (curSheet) curSheet.sheet_protected = data.sheet_protected || false;

  // 인쇄 설정 로드
  _printSettings = data.print_settings || null;
  applyPrintCSS(_printSettings);

  // 행/열 그룹 적용
  SpreadsheetCore.clearOutlines();
  if ((data.outline_rows && Object.keys(data.outline_rows).length > 0) ||
      (data.outline_cols && Object.keys(data.outline_cols).length > 0)) {
    setTimeout(() => SpreadsheetCore.applyOutlines(ctx, data.outline_rows || {}, data.outline_cols || {}), 200);
  }

  // 숨겨진 행/열 적용
  _hiddenRows = data.hidden_rows || [];
  _hiddenCols = data.hidden_cols || [];
  if (_hiddenRows.length > 0) {
    setTimeout(() => SpreadsheetCore.applyHiddenRows(ctx, _hiddenRows), 100);
  }
  if (_hiddenCols.length > 0) {
    setTimeout(() => SpreadsheetCore.applyHiddenCols(ctx, _hiddenCols), 100);
  }

  // 행 고정 적용
  SpreadsheetCore.clearFreezeRows();
  if (data.freeze_rows > 0) {
    setTimeout(() => SpreadsheetCore.applyFreezeRows(ctx, data.freeze_rows), 150);
  }

  // 자동 채우기 핸들 초기화
  SpreadsheetCore.initAutofill(ctx);

  // 하이퍼링크 클릭 핸들러
  container.addEventListener('click', function(e) {
    SpreadsheetCore.handleHyperlinkClick(ctx, e);
  });
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
  if (_suppressOnChange) return value;
  if (isClosed && !IS_ADMIN) return false;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return value;
  const col = sheet.columns[x];
  if (col && col.is_readonly && !IS_ADMIN) return false;
  // 데이터 유효성 검증
  var cellName = SpreadsheetCore.colIndexToLetter(parseInt(x)) + (parseInt(y) + 1);
  var dvErr = SpreadsheetCore.validateCellValue(cellName, value);
  if (dvErr) {
    showToast(dvErr, 'error');
    return false;
  }
  // Capture old value for undo
  try {
    const oldVal = instance.getValueFromCoords(parseInt(x), parseInt(y)) || '';
    cell._undoOldVal = oldVal;
  } catch(e) {}
  return value;
}

function handleCellChange(instance, cell, x, y, value) {
  if (_suppressOnChange) return;
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
  // onbeforepaste에서 _suppressOnChange=true 설정 → handleCellChange 중복 방지
  clearTimeout(window._pasteResetTimer);
  _suppressOnChange = false;
  if (isClosed && !IS_ADMIN) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  // data는 2D 텍스트 배열: [[val1, val2], [val3, val4], ...]
  // _pasteStartRow/_pasteStartCol은 onbeforepaste에서 캡처한 시작 위치
  const patches = [];
  const undoChanges = [];
  const gridData = instance.options.data;
  data.forEach((rowData, ri) => {
    if (!Array.isArray(rowData)) return;
    const row = _pasteStartRow + ri;
    rowData.forEach((cellVal, ci) => {
      const col = _pasteStartCol + ci;
      // 읽기 전용 컬럼은 붙여넣기 대상에서 제외 (ADMIN 제외)
      const colDef = sheet.columns[col];
      if (colDef && colDef.is_readonly && !IS_ADMIN) return;
      // gridData에서 실제 적용된 값 읽기 (jspreadsheet가 이미 적용)
      let rawValue = cellVal != null ? String(cellVal) : '';
      try {
        if (gridData && row < gridData.length && gridData[row] && col < gridData[row].length) {
          const raw = gridData[row][col];
          if (raw !== undefined && raw !== null) rawValue = String(raw);
        }
      } catch(e) {}
      patches.push({ row, col, value: rawValue, style: null });
      undoChanges.push({ row, col, oldVal: '', newVal: rawValue });
    });
  });
  if (undoChanges.length > 0 && ctx.undoManager) {
    ctx.undoManager.push({ type: 'value', changes: undoChanges });
  }
  if (patches.length > 0) sendBatchPatch(sheet.id, patches);
}

function handleSelection(el, x1, y1, x2, y2) {
  selX1 = x1; selY1 = y1; selX2 = x2; selY2 = y2;
  // Format Painter: 선택 시 자동 적용
  if (SpreadsheetCore.isPainterActive()) {
    SpreadsheetCore.applyPainterToSelection(ctx);
  }
  SpreadsheetCore.updateToolbarState(ctx);
  if (ctx._positionAutofillHandle) ctx._positionAutofillHandle();
  // 데이터 유효성 list 드롭다운 (단일 셀 선택 시)
  SpreadsheetCore.hideValidationDropdown();
  if (x1 === x2 && y1 === y2) {
    SpreadsheetCore.showValidationDropdown(ctx, parseInt(x1), parseInt(y1));
  }
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
    if (!bySheet[p.sheetId]) bySheet[p.sheetId] = {};
    // 같은 셀에 대한 패치는 마지막 값만 유지 (중복 제거)
    bySheet[p.sheetId][`${p.row},${p.col}`] = { row: p.row, col: p.col, value: p.value, style: p.style };
  });
  pendingPatches = [];
  Object.entries(bySheet).forEach(([sheetId, patchMap]) => {
    const patches = Object.values(patchMap);
    if (patches.length === 1) {
      const p = patches[0];
      sendPatch(sheetId, p.row, p.col, p.value, p.style);
    } else {
      sendBatchPatch(sheetId, patches);
    }
  });
}

function sendPatch(sheetId, row, col, value, style) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify({ type: 'patch', sheet_id: sheetId, row, col, value, style }));
  } else {
    // WebSocket 미연결 시 REST API fallback (데이터 손실 방지)
    _sendPatchesViaRest(sheetId, [{ row, col, value, style }]);
  }
}

function sendBatchPatch(sheetId, patches) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify({ type: 'batch_patch', sheet_id: sheetId, patches }));
  } else {
    _sendPatchesViaRest(sheetId, patches);
  }
}

async function _sendPatchesViaRest(sheetId, patches) {
  try {
    const res = await apiFetch(
      `/api/workspaces/${workspaceData.id}/sheets/${sheetId}/patches`,
      { method: 'POST', body: JSON.stringify({ patches }) }
    );
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      showToast(e.detail || '저장 실패 (REST)', 'error');
    }
  } catch(e) {
    showToast('서버 연결 실패 — 변경사항이 저장되지 않을 수 있습니다', 'error');
  }
}

// ── WebSocket ─────────────────────────────────────────────────
let _wsConnectedOnce = false;  // 재접속 감지용

function connectWebSocket() {
  if (ws) {
    ws.onclose = null; ws.onerror = null;
    if (ws.readyState !== WebSocket.CLOSED) ws.close();
    ws = null;
  }
  clearTimeout(wsReconnectTimer);
  setConnStatus('connecting');
  const proto = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const url = `${proto}//${location.host}/ws/workspaces/${workspaceData.id}`;

  ws = new WebSocket(url);
  ws.onopen = () => {
    setConnStatus('connected');
    clearTimeout(wsReconnectTimer);
    if (_wsConnectedOnce) {
      // 재접속 시 누락된 패치 복구를 위해 현재 시트 새로고침
      loadSheet(currentSheetIndex);
    }
    _wsConnectedOnce = true;
  };
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
    flushPatches();  // 대기 중인 변경사항 즉시 전송
    showToast('데이터가 업로드되었습니다. 새로고침합니다.', 'info');
    setTimeout(() => loadSheet(currentSheetIndex), 800);
    return;
  }
  if (msg.type === 'sheet_added') {
    if (!sheets.find(s => s.id === msg.sheet.id)) {
      sheets.push(msg.sheet);
      renderTabs();
    }
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
  if (msg.type === 'sheet_config_updated') {
    // 다른 탭에서 병합/행높이/틀고정 변경 → 해당 시트 리로드
    if (msg.tab_id === TAB_ID) return;  // 자신의 탭에서 발행한 변경은 스킵
    const sheet = sheets[currentSheetIndex];
    if (sheet && sheet.id === msg.sheet_id) {
      loadSheet(currentSheetIndex);
    }
    return;
  }
  if (msg.type === 'sheets_reordered') {
    const orderMap = {};
    msg.order.forEach((id, idx) => { orderMap[id] = idx; });
    const curSheetId = sheets[currentSheetIndex] ? sheets[currentSheetIndex].id : null;
    sheets.sort((a, b) => (orderMap[a.id] ?? 999) - (orderMap[b.id] ?? 999));
    if (curSheetId) {
      currentSheetIndex = sheets.findIndex(s => s.id === curSheetId);
      if (currentSheetIndex < 0) currentSheetIndex = 0;
    }
    renderTabs();
    return;
  }
  if (msg.type === 'sheet_renamed') {
    const s = sheets.find(s => s.id === msg.sheet_id);
    if (s) { s.sheet_name = msg.sheet_name; renderTabs(); }
    return;
  }
  if (msg.type === 'sheet_added') {
    // 다른 사용자가 시트를 추가한 경우 페이지 새로고침
    if (!sheets.find(s => s.id === msg.sheet_id)) {
      location.reload();
    }
    return;
  }
  if (msg.type === 'error') showToast(msg.message || '오류 발생', 'error');
}

function applyRemotePatch(msg) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet || msg.sheet_id !== sheet.id) return;
  if (!spreadsheet) return;
  _suppressOnChange = true;
  try {
    if (msg.value !== undefined && msg.value !== null) {
      try { spreadsheet.setValueFromCoords(msg.col, msg.row, msg.value || '', true); } catch(e) {}
    }
    if (msg.style) {
      try {
        const cellName = SpreadsheetCore.colIndexToLetter(msg.col) + (msg.row + 1);
        const s = JSON.parse(msg.style);
        const css = SpreadsheetCore.styleObjToCss(s);
        spreadsheet.setStyle({ [cellName]: css });
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
  } finally {
    _suppressOnChange = false;
  }
}

function handleRemoteRowOp(msg) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet || msg.sheet_id !== sheet.id) return;
  if (!spreadsheet) return;
  // 자신의 탭에서 발행한 메시지는 이미 로컬 적용됨 → 스킵 (tab_id로 정확한 탭 식별)
  if (msg.tab_id && msg.tab_id === TAB_ID) return;
  _suppressOnChange = true;
  try {
    if (msg.type === 'row_insert') {
      try { spreadsheet.insertRow(msg.count || 1, msg.row_index, true); } catch(e) {}
    } else if (msg.type === 'row_delete') {
      const indices = msg.row_indices || [msg.row_index];
      for (let i = indices.length - 1; i >= 0; i--) {
        try { spreadsheet.deleteRow(indices[i]); } catch(e) {}
      }
    }
  } finally {
    _suppressOnChange = false;
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

async function deleteRowsApi(rowIndices) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  // 로컬: 뒤에서부터 삭제 (인덱스 밀림 방지)
  for (let i = rowIndices.length - 1; i >= 0; i--) {
    try { spreadsheet.deleteRow(rowIndices[i]); } catch(e) {}
  }
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/rows/delete`,
    { method: 'POST', body: JSON.stringify({ row_indices: rowIndices }) }
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
  try {
    spreadsheet.insertColumn(1, colIndex, direction === 'before' ? 1 : 0);
    SpreadsheetCore.refreshColumnHeaders(ctx);
  } catch(e) {
    console.error('insertColumn error:', e);
    showToast('열 삽입 실패 (로컬): ' + e.message, 'error');
  }
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
  deleteColsApi([colIndex]);
}

async function deleteColsApi(colIndices) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  // 로컬: 뒤에서부터 삭제 (인덱스 밀림 방지)
  const sorted = colIndices.slice().sort((a, b) => b - a);
  for (const ci of sorted) {
    try { spreadsheet.deleteColumn(ci); } catch(e) {}
  }
  SpreadsheetCore.refreshColumnHeaders(ctx);
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/cols/delete`,
    { method: 'POST', body: JSON.stringify({ col_indices: colIndices }) }
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
  // 자신의 탭에서 발행한 메시지는 이미 로컬 적용됨 → 스킵 (tab_id로 정확한 탭 식별)
  if (msg.tab_id && msg.tab_id === TAB_ID) return;
  _suppressOnChange = true;
  try {
    if (msg.type === 'col_insert') {
      try { spreadsheet.insertColumn(msg.count || 1, msg.col_index, true); } catch(e) {}
    } else if (msg.type === 'col_delete') {
      const indices = msg.col_indices || [msg.col_index];
      for (let i = indices.length - 1; i >= 0; i--) {
        try { spreadsheet.deleteColumn(indices[i]); } catch(e) {}
      }
    }
    SpreadsheetCore.refreshColumnHeaders(ctx);
  } finally {
    _suppressOnChange = false;
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
    if (!sheets.find(s => s.id === newSheet.id)) {
      sheets.push(newSheet);
    }
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
    if (inp) {
      inp.value = sheets[index].sheet_name;
      inp.focus();
      inp.select();
      inp.addEventListener('keydown', function(ev) {
        if (ev.key === 'Enter') { ev.preventDefault(); submitRenameWsSheet(); }
        if (ev.key === 'Escape') { ev.preventDefault(); closeModal(); }
      });
    }
  }, 50);
}

async function submitRenameWsSheet() {
  try {
    const inp = document.getElementById('f-ws-sheet-name');
    if (!inp) return;
    const newName = inp.value.trim();
    if (!newName) { showToast('시트 이름을 입력하세요', 'error'); return; }
    const sheet = sheets[renamingWsSheetIndex];
    if (!sheet) { showToast('시트를 찾을 수 없습니다', 'error'); return; }
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
      const e2 = await res.json().catch(() => ({}));
      showToast(e2.detail || `변경 실패 (HTTP ${res.status})`, 'error');
    }
  } catch (err) {
    showToast('시트 이름 변경 실패: ' + err.message, 'error');
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
    const delIdx = sheets.findIndex(s => s.id === sheet.id);
    if (delIdx >= 0) sheets.splice(delIdx, 1);
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

// ── 열 너비 변경 핸들러 ─────────────────────────────────────
let colWidthSaveTimer = null;
function handleResizeColumn(el, col, width) {
  clearTimeout(colWidthSaveTimer);
  colWidthSaveTimer = setTimeout(saveColWidths, 500);
  if (ctx._positionAutofillHandle) ctx._positionAutofillHandle();
}

async function saveColWidths() {
  if (!spreadsheet || !IS_ADMIN) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const colWidths = {};
  try {
    const colgroup = spreadsheet.colgroup;
    if (colgroup) {
      for (let i = 0; i < colgroup.length; i++) {
        const w = colgroup[i] && colgroup[i].getAttribute('width');
        if (w) colWidths[String(i)] = parseInt(w);
      }
    }
  } catch(e) {}
  await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/col-widths`,
    { method: 'PATCH', body: JSON.stringify({ col_widths: colWidths }) }
  );
}

// ── 행 높이 변경 핸들러 ─────────────────────────────────────
let rowHeightSaveTimer = null;
function handleResizeRow(el, row, height) {
  // 디바운스: 연속 리사이즈 시 마지막 변경만 저장
  clearTimeout(rowHeightSaveTimer);
  rowHeightSaveTimer = setTimeout(saveRowHeights, 500);
  if (ctx._positionAutofillHandle) ctx._positionAutofillHandle();
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

// ── 컬럼 속성 모달 (관리자 전용) ────────────────────────────
let selectedCol = -1;

function showColumnPropsModal(colIdx) {
  selectedCol = colIdx;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const col = sheet.columns[colIdx];
  if (!col) return;

  showModalFromTemplate('컬럼 속성: ' + col.col_header, 'col-props-tpl');
  setTimeout(() => {
    const hdr = document.getElementById('cp-header');
    const tp = document.getElementById('cp-type');
    const wd = document.getElementById('cp-width');
    const ro = document.getElementById('cp-readonly');
    if (hdr) hdr.value = col.col_header;
    if (tp) tp.value = col.col_type;
    if (wd) wd.value = col.width || 120;
    if (ro) ro.checked = !!col.is_readonly;
  }, 50);
}

async function applyColProps() {
  if (selectedCol < 0) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const col = sheet.columns[selectedCol];
  if (!col) return;
  const payload = {
    col_header: document.getElementById('cp-header').value,
    col_type: document.getElementById('cp-type').value,
    width: parseInt(document.getElementById('cp-width').value) || 120,
    is_readonly: document.getElementById('cp-readonly').checked ? 1 : 0,
  };
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/columns/${col.col_index}`,
    { method: 'PATCH', body: JSON.stringify(payload) }
  );
  if (res.ok) {
    const result = await res.json();
    // 서버 응답으로 컬럼 데이터 갱신
    Object.assign(col, payload);
    if (result.data && result.data.id) col.id = result.data.id;
    if (result.template_sheet_id) sheet.template_sheet_id = result.template_sheet_id;
    closeModal();
    showToast('컬럼 속성이 저장되었습니다', 'success');
    const container = document.getElementById('spreadsheet');
    const ths = container.querySelectorAll('thead td');
    if (ths[selectedCol + 1]) ths[selectedCol + 1].textContent = payload.col_header;
  } else {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '저장 실패', 'error');
  }
}

async function deleteColumnFromModal() {
  if (selectedCol < 0) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const col = sheet.columns[selectedCol];
  if (!col) return;
  if (!confirm(`"${col.col_header}" 컬럼과 해당 셀 데이터를 삭제하시겠습니까?`)) return;
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/columns/${col.col_index}`,
    { method: 'DELETE' }
  );
  if (res.ok || res.status === 204) {
    sheet.columns.splice(selectedCol, 1);
    selectedCol = -1;
    closeModal();
    showToast('컬럼이 삭제되었습니다', 'success');
    loadSheet(currentSheetIndex);
  } else {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '삭제 실패', 'error');
  }
}

// ── 하이퍼링크 저장 ─────────────────────────────────────────
async function saveHyperlink(row, col, url) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/patches`,
    { method: 'POST', body: JSON.stringify({ patches: [{ row, col, hyperlink: url || '' }] }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '하이퍼링크 저장 실패', 'error');
  }
}

// ── 행/열 숨기기 관리 ───────────────────────────────────────
let _hiddenRows = [];
let _hiddenCols = [];

async function hideRows(rows) {
  rows.forEach(r => { if (!_hiddenRows.includes(r)) _hiddenRows.push(r); });
  _hiddenRows.sort((a, b) => a - b);
  SpreadsheetCore.applyHiddenRows(ctx, _hiddenRows);
  await saveHidden();
}

async function unhideRows() {
  _hiddenRows = [];
  SpreadsheetCore.applyHiddenRows(ctx, []);
  await saveHidden();
}

async function hideCols(cols) {
  cols.forEach(c => { if (!_hiddenCols.includes(c)) _hiddenCols.push(c); });
  _hiddenCols.sort((a, b) => a - b);
  SpreadsheetCore.applyHiddenCols(ctx, _hiddenCols);
  await saveHidden();
}

async function unhideCols() {
  _hiddenCols = [];
  SpreadsheetCore.applyHiddenCols(ctx, []);
  await saveHidden();
}

async function saveHidden() {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/hidden`,
    { method: 'PATCH', body: JSON.stringify({ hidden_rows: _hiddenRows, hidden_cols: _hiddenCols }) }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    showToast(e.detail || '숨기기 저장 실패', 'error');
  }
}

// ── 틀 고정 설정 다이얼로그 ────────────────────────────────
function showFreezeDialog() {
  var sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  // 현재 freeze_panes에서 행/열 추출
  var fp = sheet.freeze_panes || '';
  var curCols = 0, curRows = 0;
  if (fp) {
    var m = fp.match(/^([A-Z]+)(\d+)$/i);
    if (m) {
      curCols = SpreadsheetCore.letterToColIndex(m[1]);
      curRows = Math.max(0, parseInt(m[2]) - 1);
    }
  }
  var input = prompt(
    '틀 고정 설정\n고정할 행 수, 열 수를 쉼표로 입력하세요.\n예: 2,1 → 2행 고정 + 1열 고정\n해제하려면 0,0 또는 빈 값을 입력하세요.',
    curRows + ',' + curCols
  );
  if (input === null) return;
  input = input.trim();
  var freezeRows = 0, freezeCols = 0;
  if (input) {
    var parts = input.split(',');
    freezeRows = parseInt(parts[0]) || 0;
    freezeCols = parts.length > 1 ? (parseInt(parts[1]) || 0) : 0;
  }
  var freezePanes = null;
  if (freezeRows > 0 || freezeCols > 0) {
    freezePanes = SpreadsheetCore.colIndexToLetter(freezeCols) + (freezeRows + 1);
  }
  saveFreeze(freezePanes);
}

async function saveFreeze(freezePanes) {
  var sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  var res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/freeze`,
    { method: 'PATCH', body: JSON.stringify({ freeze_panes: freezePanes }) }
  );
  if (res.ok) {
    sheet.freeze_panes = freezePanes;
    showToast('틀 고정 설정이 저장되었습니다.');
    // 시트 다시 로드
    loadSheet(currentSheetIndex);
  } else {
    var e = await res.json().catch(() => ({}));
    showToast(e.detail || '틀 고정 저장 실패', 'error');
  }
}

// ── 인쇄 설정 ────────────────────────────────────────────────
var _printSettings = null;

function showPrintSettingsDialog() {
  var sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  var ps = _printSettings || { paperSize: 'A4', orientation: 'portrait', scale: 100, margins: { top: 20, bottom: 20, left: 15, right: 15 } };

  var old = document.getElementById('print-settings-dialog');
  if (old) old.remove();

  var dialog = document.createElement('div');
  dialog.id = 'print-settings-dialog';
  dialog.className = 'paste-special-dialog';
  dialog.innerHTML =
    '<div class="ps-title">인쇄 설정</div>' +
    '<label class="ps-opt">용지 크기: <select id="ps-paper"><option value="A4"' + (ps.paperSize==='A4'?' selected':'') + '>A4</option><option value="A3"' + (ps.paperSize==='A3'?' selected':'') + '>A3</option><option value="Letter"' + (ps.paperSize==='Letter'?' selected':'') + '>Letter</option></select></label>' +
    '<label class="ps-opt">방향: <select id="ps-orient"><option value="portrait"' + (ps.orientation==='portrait'?' selected':'') + '>세로</option><option value="landscape"' + (ps.orientation==='landscape'?' selected':'') + '>가로</option></select></label>' +
    '<label class="ps-opt">배율(%): <input type="number" id="ps-scale" value="' + (ps.scale||100) + '" min="10" max="400" style="width:60px"></label>' +
    '<div class="ps-btn-bar"><button class="ps-ok">저장</button><button class="ps-cancel">취소</button></div>';

  dialog.style.position = 'fixed';
  dialog.style.left = '50%';
  dialog.style.top = '50%';
  dialog.style.transform = 'translate(-50%, -50%)';
  document.body.appendChild(dialog);

  dialog.querySelector('.ps-ok').addEventListener('click', async function() {
    var settings = {
      paperSize: document.getElementById('ps-paper').value,
      orientation: document.getElementById('ps-orient').value,
      scale: parseInt(document.getElementById('ps-scale').value) || 100,
      margins: ps.margins,
    };
    var res = await apiFetch(
      `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/print-settings`,
      { method: 'PATCH', body: JSON.stringify({ print_settings: settings }) }
    );
    if (res.ok) {
      _printSettings = settings;
      applyPrintCSS(settings);
      showToast('인쇄 설정이 저장되었습니다.');
    } else {
      showToast('인쇄 설정 저장 실패', 'error');
    }
    dialog.remove();
  });
  dialog.querySelector('.ps-cancel').addEventListener('click', function() { dialog.remove(); });
}

function applyPrintCSS(settings) {
  var existing = document.getElementById('print-page-style');
  if (existing) existing.remove();
  if (!settings) return;
  var size = settings.paperSize || 'A4';
  var orient = settings.orientation || 'portrait';
  var scale = settings.scale || 100;
  var m = settings.margins || {};
  var css = '@page { size: ' + size + ' ' + orient + '; margin: ' +
    (m.top||20) + 'mm ' + (m.right||15) + 'mm ' + (m.bottom||20) + 'mm ' + (m.left||15) + 'mm; }' +
    ' @media print { body { transform: scale(' + (scale/100) + '); transform-origin: top left; } }';
  var style = document.createElement('style');
  style.id = 'print-page-style';
  style.textContent = css;
  document.head.appendChild(style);
}

// ── 시트 보호 토글 ──────────────────────────────────────────
async function toggleSheetProtection() {
  var sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  var isProtected = sheet.sheet_protected || false;
  var newVal = !isProtected;
  var msg = newVal ? '시트를 보호하시겠습니까?\n잠긴(locked) 셀은 일반 사용자가 편집할 수 없게 됩니다.' : '시트 보호를 해제하시겠습니까?';
  if (!confirm(msg)) return;
  var res = await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/protection`,
    { method: 'PATCH', body: JSON.stringify({ protected: newVal }) }
  );
  if (res.ok) {
    sheet.sheet_protected = newVal;
    showToast(newVal ? '시트 보호가 활성화되었습니다.' : '시트 보호가 해제되었습니다.');
  } else {
    var e = await res.json().catch(() => ({}));
    showToast(e.detail || '시트 보호 변경 실패', 'error');
  }
}

// Keep-alive
setInterval(() => {
  if (ws && ws.readyState === WebSocket.OPEN)
    ws.send(JSON.stringify({ type: 'ping' }));
}, 25000);
