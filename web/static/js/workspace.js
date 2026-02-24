/* ============================================================
   workspace.js — Jspreadsheet + WebSocket 실시간 협업

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

const BATCH_DELAY = 100;

// 수식 엔진
if (typeof formula !== 'undefined') {
  jspreadsheet.setExtensions({ formula });
}

// ── 커스텀 수식 ────────────────────────────────────────────────
(function registerCustomFormulas() {
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
})();

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

document.addEventListener('DOMContentLoaded', init);

function init() {
  const dataEl = document.getElementById('workspace-data');
  if (!dataEl) return;
  workspaceData = JSON.parse(dataEl.textContent);
  sheets = workspaceData.sheets || [];
  isClosed = workspaceData.status === 'CLOSED';
  initColorSwatches();
  renderTabs();
  if (sheets.length > 0) loadSheet(0);
  connectWebSocket();
  // 전역 클릭으로 드롭다운 닫기
  document.addEventListener('click', closeAllDropdowns);
}

// ── 시트 탭 렌더링 ────────────────────────────────────────────
function renderTabs() {
  const tabsEl = document.getElementById('sheet-tabs');
  tabsEl.innerHTML = sheets.map((s, i) => {
    const isActive = i === currentSheetIndex;
    const delBtn = (IS_ADMIN && sheets.length > 1)
      ? `<span class="tab-del" onclick="deleteWsSheet(${i})" title="삭제">×</span>`
      : '';
    return `<div class="sheet-tab ${isActive ? 'active' : ''}"
        onclick="handleTabClick(event, ${i})"
        ondblclick="${IS_ADMIN ? `renameWsSheet(${i})` : ''}">
      <span>${esc(s.sheet_name)}</span>${delBtn}
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
  renderTabs();
  loadSheet(index);
}

// ── 시트 로드 ─────────────────────────────────────────────────
async function loadSheet(index) {
  const sheet = sheets[index];
  if (!sheet) return;

  const container = document.getElementById('spreadsheet');
  container.innerHTML = '<div style="padding:20px;color:#64748b">로딩 중...</div>';
  if (spreadsheet) {
    try { jspreadsheet.destroy(container); } catch(e) {}
    spreadsheet = null;
  }

  const res = await apiFetch(
    `/api/workspaces/${workspaceData.id}/sheets/${sheet.id}/snapshot`
  );
  if (!res.ok) {
    container.innerHTML = '<div style="color:red;padding:20px">로딩 실패</div>';
    return;
  }
  const { data } = await res.json();

  const isEditable = !isClosed || IS_ADMIN;
  const columns = buildColumnDefs(sheet.columns, isEditable);
  const numRows = Math.max(data.num_rows, 100);

  const gridData = [];
  for (let r = 0; r < numRows; r++) {
    const row = (data.cells[r] || []).slice();
    while (row.length < columns.length) row.push('');
    gridData.push(row);
  }

  // 병합
  const mergeCells = data.merges && Object.keys(data.merges).length > 0 ? data.merges : undefined;
  // 틀 고정
  const freezeColumns = data.freeze_columns > 0 ? data.freeze_columns : undefined;

  container.innerHTML = '';
  spreadsheet = jspreadsheet(container, {
    data: gridData,
    columns,
    minDimensions: [columns.length, numRows],
    tableOverflow: true,
    tableWidth: (container.offsetWidth || (window.innerWidth - 260)) + 'px',
    tableHeight: (window.innerHeight - 200) + 'px',
    lazyLoading: true,
    loadingSpin: true,
    editable: isEditable,
    allowInsertColumn: false,
    allowDeleteColumn: false,
    mergeCells,
    freezeColumns,
    onchange: handleCellChange,
    onpaste: handlePaste,
    onbeforechange: handleBeforeChange,
    onselection: handleSelection,
    onmerge: IS_ADMIN ? handleMerge : undefined,
  });

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
  queuePatch(sheet.id, parseInt(y), parseInt(x), rawValue, null);
}

function handlePaste(instance, data) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const patches = [];
  data.forEach(item => {
    let rawValue = item[3];
    try {
      const gridData = instance.options.data;
      const iy = item[0], ix = item[1];
      if (gridData && iy < gridData.length && gridData[iy] && ix < gridData[iy].length) {
        const raw = gridData[iy][ix];
        if (raw !== undefined && raw !== null) rawValue = String(raw);
      }
    } catch(e) {}
    patches.push({ row: item[0], col: item[1], value: rawValue });
  });
  if (patches.length > 0) sendBatchPatch(sheet.id, patches);
}

function handleSelection(el, x1, y1, x2, y2) {
  selX1 = x1; selY1 = y1; selX2 = x2; selY2 = y2;
  updateToolbarState();
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
  // 값 적용
  if (msg.value !== undefined && msg.value !== null) {
    try { spreadsheet.setValueFromCoords(msg.col, msg.row, msg.value || '', true); } catch(e) {}
  }
  // 스타일 적용
  if (msg.style) {
    try {
      const cellName = colIndexToLetter(msg.col) + (msg.row + 1);
      const s = JSON.parse(msg.style);
      const css = styleObjToCss(s);
      spreadsheet.setStyle({ [cellName]: css });
    } catch(e) {}
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

// Keep-alive
setInterval(() => {
  if (ws && ws.readyState === WebSocket.OPEN)
    ws.send(JSON.stringify({ type: 'ping' }));
}, 25000);

// ============================================================
// ── 포맷 툴바 ─────────────────────────────────────────────────
// ============================================================

const COLOR_PALETTE = [
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
  'FF8000','8000FF','FF0080','00FF80','800000','008000','000080','808000',
  '800080','008080','C0C0C0','808080','FF9999','99FF99','9999FF','FFFF99',
  'FF99FF','99FFFF','FFCC99','CC99FF','FF99CC','99FFCC','FFCCCC','CCFFCC',
  'CCCCFF','FFFFCC','FFCCFF','CCFFFF','E6E6E6','333333','666666','999999',
];

function initColorSwatches() {
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
        if (idx === 0) fmtColor(hex);
        else fmtBg(hex);
      };
      el.appendChild(sw);
    });
  });
}

function toggleDropdown(id) {
  event.stopPropagation();
  const target = document.getElementById(id);
  if (!target) return;
  const wasOpen = target.classList.contains('open');
  closeAllDropdowns();
  if (!wasOpen) target.classList.add('open');
}

function closeAllDropdowns() {
  document.querySelectorAll('.fmt-dropdown.open').forEach(d => d.classList.remove('open'));
}

function cssToStyleObj(css) {
  const s = {};
  if (!css) return s;
  css.split(';').forEach(part => {
    const [k, v] = part.split(':').map(x => x.trim());
    if (!k || !v) return;
    if (k === 'font-weight' && v === 'bold') s.bold = true;
    if (k === 'font-style' && v === 'italic') s.italic = true;
    if (k === 'text-decoration' && v === 'underline') s.underline = true;
    if (k === 'color') s.color = v.replace('#','');
    if (k === 'background-color') s.bg = v.replace('#','');
    if (k === 'text-align') s.align = v;
    if (k === 'white-space' && v === 'pre-wrap') s.wrap = true;
  });
  return s;
}

function styleObjToCss(s) {
  const parts = [];
  if (s.bold) parts.push('font-weight:bold');
  if (s.italic) parts.push('font-style:italic');
  if (s.underline) parts.push('text-decoration:underline');
  if (s.fontSize) parts.push(`font-size:${s.fontSize}pt`);
  if (s.color) parts.push(`color:#${s.color}`);
  if (s.bg) parts.push(`background-color:#${s.bg}`);
  if (s.align) parts.push(`text-align:${s.align}`);
  if (s.valign) parts.push(`vertical-align:${s.valign}`);
  if (s.wrap) parts.push('white-space:pre-wrap');
  if (s.border) {
    const wm = {thin:'1px', medium:'2px', thick:'3px', dashed:'1px', dotted:'1px'};
    const sm = {thin:'solid', medium:'solid', thick:'solid', dashed:'dashed', dotted:'dotted'};
    for (const [side, bd] of Object.entries(s.border)) {
      const bs = bd.style || 'thin';
      parts.push(`border-${side}:${wm[bs]||'1px'} ${sm[bs]||'solid'} #${bd.color||'000000'}`);
    }
  }
  return parts.join(';');
}

function getSelectedCellStyle() {
  if (!spreadsheet) return {};
  try {
    const cellName = colIndexToLetter(selX1) + (selY1 + 1);
    const cssStr = spreadsheet.getStyle(cellName) || '';
    return cssToStyleObj(cssStr);
  } catch(e) { return {}; }
}

function updateToolbarState() {
  if (!spreadsheet) return;
  const s = getSelectedCellStyle();
  setActive('fmt-bold', !!s.bold);
  setActive('fmt-italic', !!s.italic);
  setActive('fmt-underline', !!s.underline);
  setActive('fmt-wrap', !!s.wrap);
  setActive('fmt-align-left', s.align === 'left');
  setActive('fmt-align-center', s.align === 'center');
  setActive('fmt-align-right', s.align === 'right');
  const colorBar = document.getElementById('fmt-color-bar');
  if (colorBar) colorBar.style.background = s.color ? '#' + s.color : '#000000';
  const bgBar = document.getElementById('fmt-bg-bar');
  if (bgBar) bgBar.style.background = s.bg ? '#' + s.bg : 'transparent';
}

function setActive(id, active) {
  const el = document.getElementById(id);
  if (el) el.classList.toggle('active', active);
}

function applyStyleToSelection(styleProp, value) {
  if (!spreadsheet) return;
  const isEditable = !isClosed || IS_ADMIN;
  if (!isEditable) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;

  const styleMap = {};
  const patches = [];
  for (let r = selY1; r <= selY2; r++) {
    for (let c = selX1; c <= selX2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = spreadsheet.getStyle(cellName) || '';
      const s = cssToStyleObj(cssStr);
      if (value === null) delete s[styleProp];
      else s[styleProp] = value;
      const css = styleObjToCss(s);
      styleMap[cellName] = css;
      patches.push({ row: r, col: c, value: null, style: JSON.stringify(s) });
    }
  }
  try { spreadsheet.setStyle(styleMap); } catch(e) {}
  // 서버에 스타일 패치 전송 (값은 null — style 필드만)
  if (patches.length > 0 && ws && ws.readyState === WebSocket.OPEN) {
    sendBatchPatch(sheet.id, patches);
  }
  updateToolbarState();
}

function fmtBold() { const s = getSelectedCellStyle(); applyStyleToSelection('bold', s.bold ? null : true); }
function fmtItalic() { const s = getSelectedCellStyle(); applyStyleToSelection('italic', s.italic ? null : true); }
function fmtUnderline() { const s = getSelectedCellStyle(); applyStyleToSelection('underline', s.underline ? null : true); }
function fmtColor(hex) { closeAllDropdowns(); applyStyleToSelection('color', hex); const b = document.getElementById('fmt-color-bar'); if (b) b.style.background = hex ? '#'+hex : '#000000'; }
function fmtBg(hex) { closeAllDropdowns(); applyStyleToSelection('bg', hex); const b = document.getElementById('fmt-bg-bar'); if (b) b.style.background = hex ? '#'+hex : 'transparent'; }
function fmtAlign(dir) { const s = getSelectedCellStyle(); applyStyleToSelection('align', s.align === dir ? null : dir); }
function fmtWrap() { const s = getSelectedCellStyle(); applyStyleToSelection('wrap', s.wrap ? null : true); }

// ── 병합 ──────────────────────────────────────────────────────
function fmtMerge() {
  if (!spreadsheet || !IS_ADMIN) return;
  try {
    spreadsheet.setMerge(selX1, selY1, selX2 - selX1 + 1, selY2 - selY1 + 1);
    saveMerges();
  } catch(e) { showToast('병합 실패: ' + e.message, 'error'); }
}

function fmtUnmerge() {
  if (!spreadsheet || !IS_ADMIN) return;
  try {
    spreadsheet.removeMerge(selX1, selY1);
    saveMerges();
  } catch(e) { showToast('병합 해제 실패', 'error'); }
}

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
    const startCol = letterToColIndex(m[1]);
    const startRow = parseInt(m[2]) - 1;
    const endCol = startCol + dims[0] - 1;
    const endRow = startRow + dims[1] - 1;
    const endName = colIndexToLetter(endCol) + (endRow + 1);
    if (cellName !== endName) mergesList.push(`${cellName}:${endName}`);
  }
  await apiFetch(
    `/api/admin/workspaces/${workspaceData.id}/sheets/${sheet.id}/merges`,
    { method: 'PATCH', body: JSON.stringify({ merges: mergesList }) }
  );
}

// ── 테두리 ────────────────────────────────────────────────────
function fmtBorder(preset) {
  closeAllDropdowns();
  if (!spreadsheet) return;
  const isEditable = !isClosed || IS_ADMIN;
  if (!isEditable) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  const thin = { style: 'thin', color: '000000' };
  const styleMap = {};
  const patches = [];
  for (let r = selY1; r <= selY2; r++) {
    for (let c = selX1; c <= selX2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = spreadsheet.getStyle(cellName) || '';
      const s = cssToStyleObj(cssStr);
      if (preset === 'none') { delete s.border; }
      else if (preset === 'all') { s.border = { top: thin, bottom: thin, left: thin, right: thin }; }
      else if (preset === 'outer') {
        if (!s.border) s.border = {};
        if (r === selY1) s.border.top = thin;
        if (r === selY2) s.border.bottom = thin;
        if (c === selX1) s.border.left = thin;
        if (c === selX2) s.border.right = thin;
      } else if (preset === 'bottom') {
        if (!s.border) s.border = {};
        s.border.bottom = thin;
      }
      const css = styleObjToCss(s);
      styleMap[cellName] = css;
      patches.push({ row: r, col: c, value: null, style: JSON.stringify(s) });
    }
  }
  try { spreadsheet.setStyle(styleMap); } catch(e) {}
  if (patches.length > 0 && ws && ws.readyState === WebSocket.OPEN) {
    sendBatchPatch(sheet.id, patches);
  }
}
