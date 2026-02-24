/* ============================================================
   template_edit.js
   서식(Template) 전체 편집기
   - 다중 시트 탭 (추가/삭제/이름변경)
   - 컬럼 추가/삭제/속성 편집
   - 계산식(=SUM 등) 지원 — Jspreadsheet CE 네이티브
   - 셀 변경 시 자동 저장 (디바운스 800ms)
   - 셀 스타일(굵기·색·정렬·테두리) 편집
   - 셀 병합/해제
   - 행 높이·열 너비·틀 고정 표시

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

let savePending = [];
let saveTimer = null;
const SAVE_DELAY = 800;

// 현재 선택 범위 (onselection에서 업데이트)
let selX1 = 0, selY1 = 0, selX2 = 0, selY2 = 0;

// 수식 엔진
if (typeof formula !== 'undefined') {
  jspreadsheet.setExtensions({ formula });
}

// ── 커스텀 수식 등록 ──────────────────────────────────────────
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

// ── 초기화 ────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const dataEl = document.getElementById('template-data');
  if (!dataEl) return;
  templateData = JSON.parse(dataEl.textContent);
  sheets = templateData.sheets || [];
  initColorSwatches();
  renderTabs();
  if (sheets.length > 0) loadSheet(0);
  // 전역 클릭으로 드롭다운 닫기
  document.addEventListener('click', closeAllDropdowns);
});

// ── 시트 탭 렌더링 ────────────────────────────────────────────
function renderTabs() {
  const wrap = document.getElementById('sheet-tabs');
  wrap.innerHTML = '';
  sheets.forEach((s, i) => {
    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (i === currentSheetIndex ? ' active' : '');
    tab.innerHTML =
      `<span ondblclick="renameSheet(${i})" title="더블클릭으로 이름 변경">${esc(s.sheet_name)}</span>` +
      (sheets.length > 1
        ? `<span class="tab-del" onclick="deleteSheet(${i})" title="시트 삭제">×</span>`
        : '');
    tab.addEventListener('click', (e) => {
      if (e.target.classList.contains('tab-del')) return;
      switchSheet(i);
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
  container.innerHTML = '<div style="padding:20px;color:#64748b">로딩 중...</div>';
  if (spreadsheet) {
    try { jspreadsheet.destroy(container); } catch(e) {}
    spreadsheet = null;
  }

  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/snapshot`
  );
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

  // 병합 셀 옵션
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
    tableHeight: (window.innerHeight - 300) + 'px',
    lazyLoading: true,
    loadingSpin: true,
    allowInsertColumn: false,
    allowDeleteColumn: false,
    allowInsertRow: true,
    allowDeleteRow: true,
    mergeCells,
    freezeColumns,
    onchange: handleCellChange,
    onpaste: handlePaste,
    onselection: handleSelection,
    onmerge: handleMerge,
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

  attachHeaderClickListeners();
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
  updateToolbarState();
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
function handleCellChange(instance, cell, x, y, value) {
  let rawValue = value;
  try {
    const raw = instance.options.data[parseInt(y)][parseInt(x)];
    if (raw !== undefined) rawValue = raw;
  } catch(e) {}
  enqueueSave(parseInt(y), parseInt(x), rawValue);
}

function handlePaste(instance, data) {
  data.forEach(item => {
    let rawValue = item[3];
    try {
      const raw = instance.options.data[item[0]][item[1]];
      if (raw !== undefined) rawValue = raw;
    } catch(e) {}
    enqueueSave(item[0], item[1], rawValue);
  });
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
  const batch = savePending.splice(0);
  setSaveStatus('저장 중...');
  showSaveIndicator();

  const res = await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/cells`,
    { method: 'POST', body: JSON.stringify(batch) }
  );
  hideSaveIndicator();
  if (res.ok) setSaveStatus('저장됨 ✓');
  else setSaveStatus('저장 실패 ✗');
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

// ============================================================
// ── 포맷 툴바 ─────────────────────────────────────────────────
// ============================================================

// 색상 팔레트 (40색)
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

// 현재 선택된 셀의 스타일 JSON 읽기
function getSelectedCellStyle() {
  if (!spreadsheet) return {};
  try {
    const cellName = colIndexToLetter(selX1) + (selY1 + 1);
    const cssStr = spreadsheet.getStyle(cellName) || '';
    return cssToStyleObj(cssStr);
  } catch(e) {
    return {};
  }
}

// CSS 문자열 → style dict (역변환, 단순)
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

// style dict → CSS string
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

// 툴바 버튼 상태 업데이트
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

// 선택 범위 순회해 스타일 적용
function applyStyleToSelection(styleProp, value) {
  if (!spreadsheet) return;
  const styleMap = {};
  for (let r = selY1; r <= selY2; r++) {
    for (let c = selX1; c <= selX2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = spreadsheet.getStyle(cellName) || '';
      const s = cssToStyleObj(cssStr);
      if (value === null) {
        delete s[styleProp];
      } else {
        s[styleProp] = value;
      }
      styleMap[cellName] = styleObjToCss(s);
    }
  }
  try { spreadsheet.setStyle(styleMap); } catch(e) {}
  saveStyleBatch(styleMap);
  updateToolbarState();
}

// 스타일을 서버에 저장
async function saveStyleBatch(styleMap) {
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  // CSS → style JSON으로 변환해 batch_save_cells 형식으로 전송
  const batch = [];
  for (const [cellName, css] of Object.entries(styleMap)) {
    const m = cellName.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;
    const col_index = letterToColIndex(m[1]);
    const row_index = parseInt(m[2]) - 1;
    const s = cssToStyleObj(css);
    batch.push({ row_index, col_index, style: JSON.stringify(s) });
  }
  if (!batch.length) return;
  await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/cells`,
    { method: 'POST', body: JSON.stringify(batch) }
  );
}

function letterToColIndex(letter) {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
}

// ── 포맷 버튼 핸들러 ──────────────────────────────────────────
function fmtBold() {
  const s = getSelectedCellStyle();
  applyStyleToSelection('bold', s.bold ? null : true);
}
function fmtItalic() {
  const s = getSelectedCellStyle();
  applyStyleToSelection('italic', s.italic ? null : true);
}
function fmtUnderline() {
  const s = getSelectedCellStyle();
  applyStyleToSelection('underline', s.underline ? null : true);
}
function fmtColor(hex) {
  closeAllDropdowns();
  applyStyleToSelection('color', hex);
  const bar = document.getElementById('fmt-color-bar');
  if (bar) bar.style.background = hex ? '#' + hex : '#000000';
}
function fmtBg(hex) {
  closeAllDropdowns();
  applyStyleToSelection('bg', hex);
  const bar = document.getElementById('fmt-bg-bar');
  if (bar) bar.style.background = hex ? '#' + hex : 'transparent';
}
function fmtAlign(dir) {
  const s = getSelectedCellStyle();
  applyStyleToSelection('align', s.align === dir ? null : dir);
}
function fmtWrap() {
  const s = getSelectedCellStyle();
  applyStyleToSelection('wrap', s.wrap ? null : true);
}

// ── 병합 핸들러 ───────────────────────────────────────────────
function fmtMerge() {
  if (!spreadsheet) return;
  try {
    spreadsheet.setMerge(selX1, selY1, selX2 - selX1 + 1, selY2 - selY1 + 1);
    saveMerges();
  } catch(e) { showToast('병합 실패: ' + e.message, 'error'); }
}

function fmtUnmerge() {
  if (!spreadsheet) return;
  try {
    spreadsheet.removeMerge(selX1, selY1);
    saveMerges();
  } catch(e) { showToast('병합 해제 실패', 'error'); }
}

function handleMerge(el, x, y, colspan, rowspan) {
  // onmerge 콜백: 병합 변경 후 서버 저장
  setTimeout(saveMerges, 100);
}

async function saveMerges() {
  if (!spreadsheet) return;
  const sheet = sheets[currentSheetIndex];
  if (!sheet) return;
  // jspreadsheet getMerge(null) → 전체 병합 맵 {A1: [colspan,rowspan], ...}
  let mergeMap = {};
  try { mergeMap = spreadsheet.getMerge() || {}; } catch(e) {}
  // jspreadsheet 형식 → xlsx 범위 문자열 변환
  const mergesList = [];
  for (const [cellName, dims] of Object.entries(mergeMap)) {
    const m = cellName.match(/^([A-Z]+)(\d+)$/);
    if (!m || !dims || dims.length < 2) continue;
    const startCol = letterToColIndex(m[1]);
    const startRow = parseInt(m[2]) - 1;
    const endCol = startCol + dims[0] - 1;
    const endRow = startRow + dims[1] - 1;
    const endName = colIndexToLetter(endCol) + (endRow + 1);
    if (cellName !== endName) {
      mergesList.push(`${cellName}:${endName}`);
    }
  }
  await apiFetch(
    `/api/admin/templates/${templateData.id}/sheets/${sheet.id}/merges`,
    { method: 'PATCH', body: JSON.stringify({ merges: mergesList }) }
  );
}

// ── 테두리 프리셋 ─────────────────────────────────────────────
function fmtBorder(preset) {
  closeAllDropdowns();
  if (!spreadsheet) return;
  const thin = { style: 'thin', color: '000000' };
  const styleMap = {};
  for (let r = selY1; r <= selY2; r++) {
    for (let c = selX1; c <= selX2; c++) {
      const cellName = colIndexToLetter(c) + (r + 1);
      const cssStr = spreadsheet.getStyle(cellName) || '';
      const s = cssToStyleObj(cssStr);
      if (preset === 'none') {
        delete s.border;
      } else if (preset === 'all') {
        s.border = { top: thin, bottom: thin, left: thin, right: thin };
      } else if (preset === 'outer') {
        if (!s.border) s.border = {};
        if (r === selY1) s.border.top = thin;
        if (r === selY2) s.border.bottom = thin;
        if (c === selX1) s.border.left = thin;
        if (c === selX2) s.border.right = thin;
      } else if (preset === 'bottom') {
        if (!s.border) s.border = {};
        s.border.bottom = thin;
      }
      styleMap[cellName] = styleObjToCss(s);
    }
  }
  try { spreadsheet.setStyle(styleMap); } catch(e) {}
  saveStyleBatch(styleMap);
}
