// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   text_documents.js — 텍스트 문서 관리 (admin)
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => loadDocs());

async function loadDocs() {
  const res = await apiFetch('/api/admin/text-documents');
  if (!res.ok) return;
  const { data: docs } = await res.json();
  renderTable(docs);
}

const LANG_LABELS = {
  plaintext: 'Plain Text', javascript: 'JavaScript', python: 'Python',
  html: 'HTML', css: 'CSS', markdown: 'Markdown', xml: 'XML', sql: 'SQL', json: 'JSON',
};

function renderTable(docs) {
  const tbody = document.getElementById('doc-tbody');
  if (!docs.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="loading">텍스트 문서가 없습니다.</td></tr>';
    updateSelection();
    return;
  }
  tbody.innerHTML = docs.map(d => `
    <tr>
      <td><input type="checkbox" class="row-check" value="${d.id}" onchange="updateSelection()"></td>
      <td>
        <a href="/text-documents/${d.id}" style="color:var(--primary); font-weight:500">${esc(d.title)}</a>
      </td>
      <td>${LANG_LABELS[d.language] || d.language}</td>
      <td><span class="badge ${d.status.toLowerCase()}">${d.status}</span></td>
      <td>v${d.version}</td>
      <td>${fmtDate(d.created_at)}</td>
      <td>
        <div class="actions">
          <span class="action-link" onclick="${d.status === 'OPEN' ? 'closeDoc' : 'reopenDoc'}('${d.id}')">
            ${d.status === 'OPEN' ? '마감' : '재개'}
          </span>
          <span class="action-link danger" onclick="deleteDoc('${d.id}','${esc(d.title)}')">삭제</span>
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

async function batchDeleteDocs() {
  const ids = Array.from(document.querySelectorAll('.row-check:checked')).map(cb => cb.value);
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}개 문서를 삭제하시겠습니까?`)) return;
  const res = await apiFetch('/api/admin/text-documents/batch-delete', {
    method: 'POST',
    body: JSON.stringify({ ids }),
  });
  if (res.ok) {
    const { deleted } = await res.json();
    showToast(`${deleted}개 문서가 삭제되었습니다`, 'success');
    loadDocs();
  } else {
    const err = await res.json();
    showToast(err.detail || '일괄 삭제 실패', 'error');
  }
}

function showCreateDocModal() {
  showModalFromTemplate('새 텍스트 문서', 'doc-form-tpl');
}

async function submitDocForm(e) {
  e.preventDefault();
  const payload = {
    title: document.getElementById('f-doc-title').value,
    language: document.getElementById('f-doc-lang').value,
  };
  const res = await apiFetch('/api/admin/text-documents', {
    method: 'POST',
    body: JSON.stringify(payload),
  });
  if (res.ok) {
    const doc = await res.json();
    showToast('문서 생성 완료', 'success');
    closeModal();
    // 바로 에디터로 이동
    window.location.href = `/text-documents/${doc.id}`;
  } else {
    const err = await res.json();
    showToast(err.detail || '생성 실패', 'error');
  }
}

async function closeDoc(docId) {
  if (!confirm('이 문서를 마감하시겠습니까? 일반 사용자 편집이 차단됩니다.')) return;
  const res = await apiFetch(`/api/admin/text-documents/${docId}/close`, { method: 'POST' });
  if (res.ok) { showToast('마감되었습니다', 'success'); loadDocs(); }
  else { const e = await res.json(); showToast(e.detail || '오류', 'error'); }
}

async function reopenDoc(docId) {
  const res = await apiFetch(`/api/admin/text-documents/${docId}/reopen`, { method: 'POST' });
  if (res.ok) { showToast('재개되었습니다', 'success'); loadDocs(); }
  else { const e = await res.json(); showToast(e.detail || '오류', 'error'); }
}

async function deleteDoc(docId, title) {
  if (!confirm(`"${title}" 문서를 삭제하시겠습니까?`)) return;
  const res = await apiFetch(`/api/admin/text-documents/${docId}`, { method: 'DELETE' });
  if (res.ok || res.status === 204) { showToast('삭제되었습니다', 'success'); loadDocs(); }
  else { const e = await res.json(); showToast(e.detail || '삭제 실패', 'error'); }
}
