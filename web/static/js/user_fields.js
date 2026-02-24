/* ============================================================
   user_fields.js — 추가 필드 관리
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => loadFields());

async function loadFields() {
  const res = await apiFetch('/api/admin/user-fields');
  if (!res.ok) return;
  const { data: fields } = await res.json();
  renderTable(fields);
}

function renderTable(fields) {
  const tbody = document.getElementById('fields-tbody');
  if (!fields.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="loading">등록된 필드가 없습니다.</td></tr>';
    return;
  }
  tbody.innerHTML = fields.map((f, i) => `
    <tr>
      <td>${i + 1}</td>
      <td><code>${esc(f.field_name)}</code></td>
      <td>${esc(f.label)}</td>
      <td>${f.field_type}</td>
      <td>${f.is_sensitive ? '&#128274; 암호화' : '-'}</td>
      <td>${f.is_required ? '필수' : '-'}</td>
      <td>
        <div class="actions">
          <span class="action-link" onclick="showEditFieldModal(${JSON.stringify(f).replace(/"/g,'&quot;')})">편집</span>
          <span class="action-link danger" onclick="deleteField('${f.id}','${esc(f.label)}')">삭제</span>
        </div>
      </td>
    </tr>
  `).join('');
}

function showCreateFieldModal() {
  showModalFromTemplate('필드 추가', 'field-form-tpl');
  document.getElementById('field-id').value = '';
  document.getElementById('f-field-name').disabled = false;
}

function showEditFieldModal(field) {
  showModalFromTemplate('필드 편집', 'field-form-tpl');
  document.getElementById('field-id').value = field.id;
  document.getElementById('f-field-name').value = field.field_name;
  document.getElementById('f-field-name').disabled = true;
  document.getElementById('f-label').value = field.label;
  document.getElementById('f-field-type').value = field.field_type;
  document.getElementById('f-is-sensitive').checked = !!field.is_sensitive;
  document.getElementById('f-is-required').checked = !!field.is_required;
}

async function submitFieldForm(e) {
  e.preventDefault();
  const fid = document.getElementById('field-id').value;
  const payload = {
    label: document.getElementById('f-label').value,
    field_type: document.getElementById('f-field-type').value,
    is_sensitive: document.getElementById('f-is-sensitive').checked ? 1 : 0,
    is_required: document.getElementById('f-is-required').checked ? 1 : 0,
  };

  let res;
  if (fid) {
    res = await apiFetch(`/api/admin/user-fields/${fid}`, { method: 'PATCH', body: JSON.stringify(payload) });
  } else {
    payload.field_name = document.getElementById('f-field-name').value;
    payload.sort_order = 999;
    res = await apiFetch('/api/admin/user-fields', { method: 'POST', body: JSON.stringify(payload) });
  }

  if (res.ok) {
    showToast('저장되었습니다', 'success');
    closeModal();
    loadFields();
  } else {
    const err = await res.json();
    showToast(err.detail || '저장 실패', 'error');
  }
}

async function deleteField(fieldId, label) {
  if (!confirm(`"${label}" 필드를 삭제하시겠습니까? 관련 데이터도 모두 삭제됩니다.`)) return;
  const res = await apiFetch(`/api/admin/user-fields/${fieldId}`, { method: 'DELETE' });
  if (res.ok || res.status === 204) {
    showToast('삭제되었습니다', 'success');
    loadFields();
  } else {
    const err = await res.json();
    showToast(err.detail || '삭제 실패', 'error');
  }
}
