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

// ---- Escape HTML ----
function esc(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
