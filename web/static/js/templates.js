// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   templates.js — 서식 관리
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => loadTemplates());

async function loadTemplates() {
  const res = await apiFetch('/api/admin/templates');
  if (!res.ok) return;
  const { data: templates } = await res.json();
  renderTable(templates);
}

function renderTable(templates) {
  const tbody = document.getElementById('templates-tbody');
  if (!templates.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="loading">등록된 서식이 없습니다. xlsx 파일을 업로드하거나 새 서식을 만드세요.</td></tr>';
    return;
  }
  tbody.innerHTML = templates.map(t => `
    <tr>
      <td>
        <a href="/admin/templates/${t.id}/edit" style="color:var(--primary); font-weight:500">${esc(t.name)}</a>
      </td>
      <td>${t.sheet_count}</td>
      <td style="color:var(--text-muted)">${esc(t.description || '')}</td>
      <td>${fmtDate(t.updated_at)}</td>
      <td>
        <div class="actions">
          <a class="action-link" href="/admin/templates/${t.id}/edit" style="text-decoration:none">편집기</a>
          <span class="action-link" onclick="downloadTemplate('${t.id}','${esc(t.name)}')">다운로드</span>
          <span class="action-link" onclick="copyTemplate('${t.id}')">복사</span>
          <span class="action-link" onclick="showEditTemplateModal(${JSON.stringify(t).replace(/"/g,'&quot;')})">이름변경</span>
          <span class="action-link danger" onclick="deleteTemplate('${t.id}','${esc(t.name)}')">삭제</span>
        </div>
      </td>
    </tr>
  `).join('');
}

function showCreateTemplateModal() {
  showModalFromTemplate('새 서식', 'tmpl-form-tpl');
  document.getElementById('tmpl-id').value = '';
}

function showEditTemplateModal(t) {
  showModalFromTemplate('서식 편집', 'tmpl-form-tpl');
  document.getElementById('tmpl-id').value = t.id;
  document.getElementById('f-tmpl-name').value = t.name;
  document.getElementById('f-tmpl-desc').value = t.description || '';
}

async function submitTemplateForm(e) {
  e.preventDefault();
  const tid = document.getElementById('tmpl-id').value;
  const payload = {
    name: document.getElementById('f-tmpl-name').value,
    description: document.getElementById('f-tmpl-desc').value || null,
  };

  let res;
  if (tid) {
    res = await apiFetch(`/api/admin/templates/${tid}`, { method: 'PATCH', body: JSON.stringify(payload) });
  } else {
    res = await apiFetch('/api/admin/templates', { method: 'POST', body: JSON.stringify(payload) });
  }

  if (res.ok) {
    showToast('저장되었습니다', 'success');
    closeModal();
    loadTemplates();
  } else {
    const err = await res.json();
    showToast(err.detail || '저장 실패', 'error');
  }
}

async function copyTemplate(templateId) {
  const res = await apiFetch(`/api/admin/templates/${templateId}/copy`, { method: 'POST' });
  if (res.ok) {
    showToast('복사되었습니다', 'success');
    loadTemplates();
  } else {
    const err = await res.json();
    showToast(err.detail || '복사 실패', 'error');
  }
}

async function deleteTemplate(templateId, name) {
  if (!confirm(`"${name}" 서식을 삭제하시겠습니까?`)) return;
  const res = await apiFetch(`/api/admin/templates/${templateId}`, { method: 'DELETE' });
  if (res.ok || res.status === 204) {
    showToast('삭제되었습니다', 'success');
    loadTemplates();
  } else {
    const err = await res.json();
    showToast(err.detail || '삭제 실패', 'error');
  }
}

function downloadTemplate(templateId, name) {
  const a = document.createElement('a');
  a.href = `/api/admin/templates/${templateId}/export-xlsx`;
  a.download = `${name}.xlsx`;
  a.click();
}

async function importTemplate(input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const res = await apiFetch('/api/admin/templates/import-xlsx', {
    method: 'POST',
    body: fd,
    headers: { 'X-CSRF-Token': getCookie('csrf_token') },
  });
  input.value = '';
  if (res.ok) {
    const data = await res.json();
    showToast(`"${data.data.name}" 서식 가져오기 완료`, 'success');
    loadTemplates();
  } else {
    const err = await res.json();
    showToast(err.detail || '가져오기 실패', 'error');
  }
}
