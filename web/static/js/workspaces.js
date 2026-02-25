// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   workspaces.js — 워크스페이스 관리 (admin)
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => loadWorkspaces());

async function loadWorkspaces() {
  const res = await apiFetch('/api/admin/workspaces');
  if (!res.ok) return;
  const { data: workspaces } = await res.json();
  renderTable(workspaces);
}

function renderTable(workspaces) {
  const tbody = document.getElementById('ws-tbody');
  if (!workspaces.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="loading">워크스페이스가 없습니다.</td></tr>';
    updateSelection();
    return;
  }
  tbody.innerHTML = workspaces.map(w => `
    <tr>
      <td><input type="checkbox" class="row-check" value="${w.id}" onchange="updateSelection()"></td>
      <td>
        <a href="/workspaces/${w.id}" style="color:var(--primary); font-weight:500">${esc(w.name)}</a>
      </td>
      <td><span class="badge ${w.status.toLowerCase()}">${w.status}</span></td>
      <td>${w.sheet_count}</td>
      <td>${fmtDate(w.created_at)}</td>
      <td>
        <div class="actions">
          <span class="action-link" onclick="${w.status === 'OPEN' ? 'closeWS' : 'reopenWS'}('${w.id}')">
            ${w.status === 'OPEN' ? '마감' : '재개'}
          </span>
          <span class="action-link" onclick="downloadWS('${w.id}','${esc(w.name)}')">다운로드</span>
          <label class="action-link">
            업로드<input type="file" accept=".xlsx" style="display:none" onchange="importWS(this,'${w.id}')">
          </label>
          <span class="action-link danger" onclick="deleteWS('${w.id}','${esc(w.name)}')">삭제</span>
        </div>
      </td>
    </tr>
  `).join('');
  updateSelection();
}

function toggleSelectAll(master) {
  document.querySelectorAll('.row-check').forEach(cb => { cb.checked = master.checked; });
  updateSelection();
}

function updateSelection() {
  const checks = document.querySelectorAll('.row-check');
  const checked = document.querySelectorAll('.row-check:checked');
  const cnt = checked.length;
  const bar = document.getElementById('batch-toolbar');
  const cntEl = document.getElementById('sel-count');
  const all = document.getElementById('select-all');
  if (bar) bar.style.display = cnt > 0 ? '' : 'none';
  if (cntEl) cntEl.textContent = cnt;
  if (all) all.checked = checks.length > 0 && checks.length === cnt;
}

async function batchDeleteWorkspaces() {
  const ids = Array.from(document.querySelectorAll('.row-check:checked')).map(cb => cb.value);
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}개 워크스페이스를 삭제하시겠습니까?`)) return;
  const res = await apiFetch('/api/admin/workspaces/batch-delete', {
    method: 'POST',
    body: JSON.stringify({ ids }),
  });
  if (res.ok) {
    const { deleted } = await res.json();
    showToast(`${deleted}개 워크스페이스가 삭제되었습니다`, 'success');
    loadWorkspaces();
  } else {
    const err = await res.json();
    showToast(err.detail || '일괄 삭제 실패', 'error');
  }
}

async function showCreateWSModal() {
  // 서식 목록 로드
  const res = await apiFetch('/api/admin/templates');
  if (!res.ok) { showToast('서식 목록을 불러오지 못했습니다', 'error'); return; }
  const { data: templates } = await res.json();

  showModalFromTemplate('새 워크스페이스', 'ws-form-tpl');
  const sel = document.getElementById('f-ws-template');
  templates.forEach(t => {
    const opt = document.createElement('option');
    opt.value = t.id;
    opt.textContent = t.name;
    sel.appendChild(opt);
  });
}

async function submitWSForm(e) {
  e.preventDefault();
  const payload = {
    name: document.getElementById('f-ws-name').value,
    template_id: document.getElementById('f-ws-template').value,
  };
  const res = await apiFetch('/api/admin/workspaces', { method: 'POST', body: JSON.stringify(payload) });
  if (res.ok) {
    showToast('워크스페이스 생성 완료', 'success');
    closeModal();
    loadWorkspaces();
  } else {
    const err = await res.json();
    showToast(err.detail || '생성 실패', 'error');
  }
}

async function closeWS(wsId) {
  if (!confirm('이 워크스페이스를 마감하시겠습니까? 사용자 편집이 차단됩니다.')) return;
  const res = await apiFetch(`/api/admin/workspaces/${wsId}/close`, { method: 'POST' });
  if (res.ok) { showToast('마감되었습니다', 'success'); loadWorkspaces(); }
  else { const e = await res.json(); showToast(e.detail || '오류', 'error'); }
}

async function reopenWS(wsId) {
  const res = await apiFetch(`/api/admin/workspaces/${wsId}/reopen`, { method: 'POST' });
  if (res.ok) { showToast('재개되었습니다', 'success'); loadWorkspaces(); }
  else { const e = await res.json(); showToast(e.detail || '오류', 'error'); }
}

function downloadWS(wsId, name) {
  const a = document.createElement('a');
  a.href = `/api/admin/workspaces/${wsId}/export-xlsx`;
  a.download = `${name}.xlsx`;
  a.click();
}

async function importWS(input, wsId) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const res = await apiFetch(`/api/admin/workspaces/${wsId}/import-xlsx`, {
    method: 'POST',
    body: fd,
    headers: { 'X-CSRF-Token': getCookie('csrf_token') },
  });
  input.value = '';
  if (res.ok) { showToast('업로드 완료', 'success'); loadWorkspaces(); }
  else { const e = await res.json(); showToast(e.detail || '업로드 실패', 'error'); }
}

async function deleteWS(wsId, name) {
  if (!confirm(`"${name}" 워크스페이스를 삭제하시겠습니까?`)) return;
  const res = await apiFetch(`/api/admin/workspaces/${wsId}`, { method: 'DELETE' });
  if (res.ok || res.status === 204) { showToast('삭제되었습니다', 'success'); loadWorkspaces(); }
  else { const e = await res.json(); showToast(e.detail || '삭제 실패', 'error'); }
}
