/* ============================================================
   users.js — 계정 관리
   ============================================================ */
let currentPage = 1;
const PAGE_SIZE = 50;

document.addEventListener('DOMContentLoaded', () => loadUsers());

async function loadUsers(page) {
  if (page) currentPage = page;
  const q = document.getElementById('search-input')?.value || '';
  const res = await apiFetch(`/api/admin/users?page=${currentPage}&page_size=${PAGE_SIZE}&q=${encodeURIComponent(q)}`);
  if (!res.ok) return;
  const json = await res.json();
  renderTable(json.data);
  renderPagination(json.total, json.page, json.page_size);
}

function renderTable(users) {
  const tbody = document.getElementById('users-tbody');
  if (!users.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="loading">사용자가 없습니다.</td></tr>';
    return;
  }
  tbody.innerHTML = users.map(u => `
    <tr>
      <td><strong>${esc(u.username)}</strong></td>
      <td>${esc(u.email || '')}</td>
      <td><span class="role-badge role-${u.role.toLowerCase()}">${u.role}</span></td>
      <td>${u.is_active ? '<span style="color:var(--success)">활성</span>' : '<span style="color:var(--danger)">비활성</span>'}</td>
      <td>${fmtDate(u.created_at)}</td>
      <td>
        <div class="actions">
          <span class="action-link" onclick="showEditModal('${u.id}')">편집</span>
          <span class="action-link danger" onclick="deleteUser('${u.id}','${esc(u.username)}')">삭제</span>
        </div>
      </td>
    </tr>
  `).join('');
}

function renderPagination(total, page, pageSize) {
  const totalPages = Math.ceil(total / pageSize);
  const el = document.getElementById('pagination');
  if (totalPages <= 1) { el.innerHTML = ''; return; }
  let html = '';
  for (let i = 1; i <= totalPages; i++) {
    html += `<button class="page-btn ${i === page ? 'active' : ''}" onclick="loadUsers(${i})">${i}</button>`;
  }
  el.innerHTML = html;
}

function showCreateUserModal() {
  showModalFromTemplate('사용자 추가', 'user-form-tpl');
  document.getElementById('pw-hint').textContent = '(8자 이상, 필수)';
  document.getElementById('f-password').required = true;
}

async function showEditModal(userId) {
  const res = await apiFetch(`/api/admin/users/${userId}`);
  if (!res.ok) return;
  const { data: u } = await res.json();
  showModalFromTemplate('사용자 편집', 'user-form-tpl');
  document.getElementById('user-id').value = u.id;
  document.getElementById('f-username').value = u.username;
  document.getElementById('f-email').value = u.email || '';
  document.getElementById('f-role').value = u.role;
  document.getElementById('f-is-active').value = u.is_active;
  document.getElementById('f-password').required = false;
  document.getElementById('pw-hint').textContent = '(변경 시에만 입력)';
  document.getElementById('user-submit-btn').textContent = '수정';
}

async function submitUserForm(e) {
  e.preventDefault();
  const uid = document.getElementById('user-id').value;
  const payload = {
    username: document.getElementById('f-username').value,
    email: document.getElementById('f-email').value || null,
    role: document.getElementById('f-role').value,
    is_active: parseInt(document.getElementById('f-is-active').value),
  };
  const pw = document.getElementById('f-password').value;
  if (pw) payload.password = pw;

  let res;
  if (uid) {
    res = await apiFetch(`/api/admin/users/${uid}`, { method: 'PATCH', body: JSON.stringify(payload) });
  } else {
    if (!pw) { showToast('비밀번호를 입력하세요', 'error'); return; }
    payload.password = pw;
    res = await apiFetch('/api/admin/users', { method: 'POST', body: JSON.stringify(payload) });
  }

  if (res.ok) {
    showToast(uid ? '수정되었습니다' : '생성되었습니다', 'success');
    closeModal();
    loadUsers();
  } else {
    const err = await res.json();
    showToast(err.detail || '저장 실패', 'error');
  }
}

async function deleteUser(userId, username) {
  if (!confirm(`"${username}" 사용자를 삭제하시겠습니까?`)) return;
  const res = await apiFetch(`/api/admin/users/${userId}`, { method: 'DELETE' });
  if (res.ok || res.status === 204) {
    showToast('삭제되었습니다', 'success');
    loadUsers();
  } else {
    const err = await res.json();
    showToast(err.detail || '삭제 실패', 'error');
  }
}

async function exportUsers() {
  const a = document.createElement('a');
  a.href = '/api/admin/users/export-xlsx';
  a.download = 'users.xlsx';
  a.click();
}

async function importUsers(input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const res = await apiFetch('/api/admin/users/import-xlsx', {
    method: 'POST',
    body: fd,
    headers: { 'X-CSRF-Token': getCookie('csrf_token') },
  });
  input.value = '';
  if (res.ok) {
    const data = await res.json();
    showToast(`가져오기 완료: 추가 ${data.created}, 수정 ${data.updated}, 건너뜀 ${data.skipped}`, 'success');
    loadUsers();
  } else {
    const err = await res.json();
    showToast(err.detail || '가져오기 실패', 'error');
  }
}
