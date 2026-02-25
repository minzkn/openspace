// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   common.js — CSRF fetch wrapper, toast, modal, utils
   ============================================================ */

function getCookie(name) {
  const m = document.cookie.match(new RegExp('(?:^|;)\\s*' + name + '=([^;]*)'));
  return m ? decodeURIComponent(m[1]) : null;
}

async function apiFetch(url, options = {}) {
  const csrfToken = getCookie('csrf_token');
  const method = (options.method || 'GET').toUpperCase();
  const headers = { ...options.headers };

  if (!headers['Content-Type'] && !(options.body instanceof FormData)) {
    headers['Content-Type'] = 'application/json';
  }
  if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(method) && csrfToken) {
    headers['X-CSRF-Token'] = csrfToken;
  }

  const res = await fetch(url, {
    ...options,
    headers,
    credentials: 'same-origin',
  });

  if (res.status === 401) {
    window.location.href = '/login';
    throw new Error('Unauthorized');
  }
  return res;
}

// ---- Toast ----
function showToast(message, type = 'info', duration = 3000) {
  const container = document.getElementById('toast-container');
  if (!container) return;
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = message;
  container.appendChild(el);
  setTimeout(() => el.remove(), duration);
}

// ---- Modal ----
function showModal(title, contentHtml) {
  document.getElementById('modal-title').textContent = title;
  document.getElementById('modal-body').innerHTML = contentHtml;
  document.getElementById('modal-overlay').classList.remove('hidden');
}

function showModalFromTemplate(title, templateId) {
  const tpl = document.getElementById(templateId);
  if (!tpl) return;
  showModal(title, tpl.innerHTML);
}

function closeModal(event) {
  if (event && event.target !== document.getElementById('modal-overlay')) return;
  document.getElementById('modal-overlay').classList.add('hidden');
  document.getElementById('modal-body').innerHTML = '';
}

// ---- Logout ----
async function logout() {
  await apiFetch('/api/auth/logout', { method: 'POST' });
  window.location.href = '/login';
}

// ---- Debounce ----
function debounce(fn, delay) {
  let timer;
  return function(...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

// ---- Date format ----
function fmtDate(iso) {
  if (!iso) return '';
  return iso.replace('T', ' ').substring(0, 16);
}

// ---- Password Change ----
function openPasswordModal() {
  document.getElementById('pw-modal').classList.remove('hidden');
}
function closePwModal() {
  document.getElementById('pw-modal').classList.add('hidden');
  document.getElementById('pw-current').value = '';
  document.getElementById('pw-new').value = '';
  document.getElementById('pw-confirm').value = '';
}
async function submitPasswordChange() {
  const current = document.getElementById('pw-current').value;
  const nw = document.getElementById('pw-new').value;
  const confirm_ = document.getElementById('pw-confirm').value;
  if (!current || !nw) { showToast('모든 필드를 입력하세요.', 'error'); return; }
  if (nw.length < 8) { showToast('새 비밀번호는 8자 이상이어야 합니다.', 'error'); return; }
  if (nw !== confirm_) { showToast('새 비밀번호가 일치하지 않습니다.', 'error'); return; }
  const res = await apiFetch('/api/auth/me/password', {
    method: 'PATCH',
    body: JSON.stringify({ current_password: current, new_password: nw }),
  });
  if (res.ok) {
    showToast('비밀번호가 변경되었습니다.', 'success');
    closePwModal();
    var alert = document.getElementById('pw-change-alert');
    if (alert) alert.style.display = 'none';
  } else {
    const d = await res.json();
    showToast(d.detail || '변경 실패', 'error');
  }
}

// ---- Escape HTML ----
function esc(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
