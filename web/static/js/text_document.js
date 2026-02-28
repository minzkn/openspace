// SPDX-License-Identifier: MIT
// Copyright (c) 2026 JAEHYUK CHO
/* ============================================================
   text_document.js — CodeMirror 6 텍스트 편집기
   ============================================================ */

(function () {
  'use strict';

  // ---- Language mapping ----
  const LANG_MAP = {
    javascript: () => CM.javascript(),
    python:     () => CM.python(),
    html:       () => CM.html(),
    css:        () => CM.css(),
    markdown:   () => CM.markdown(),
    xml:        () => CM.xml(),
    sql:        () => CM.sql(),
    json:       () => CM.json(),
  };

  // ---- File extension mapping for download ----
  const EXT_MAP = {
    plaintext: '.txt', javascript: '.js', python: '.py', html: '.html',
    css: '.css', markdown: '.md', xml: '.xml', sql: '.sql', json: '.json',
  };

  // ---- State ----
  let editor = null;
  let currentVersion = DOC_DATA.version;
  let saveTimer = null;
  let ws = null;
  let wsReconnectTimer = null;
  let dirty = false;
  let saving = false;

  // Compartments for dynamic reconfiguration
  const langCompartment = new CM.Compartment();
  const readonlyCompartment = new CM.Compartment();

  // ---- Init ----
  document.addEventListener('DOMContentLoaded', () => {
    initEditor();
    initWebSocket();

    // 언어 셀렉터 초기값
    const langSelect = document.getElementById('lang-select');
    if (langSelect) langSelect.value = DOC_DATA.language;

    updateSaveStatus('ready');
  });

  function initEditor() {
    const isReadonly = DOC_DATA.status === 'CLOSED' && !IS_ADMIN;
    const langExt = LANG_MAP[DOC_DATA.language];

    const extensions = [
      CM.lineNumbers(),
      CM.highlightActiveLineGutter(),
      CM.highlightSpecialChars(),
      CM.history(),
      CM.foldGutter(),
      CM.drawSelection(),
      CM.dropCursor(),
      CM.indentOnInput(),
      CM.syntaxHighlighting(CM.defaultHighlightStyle, { fallback: true }),
      CM.bracketMatching(),
      CM.closeBrackets(),
      CM.autocompletion(),
      CM.rectangularSelection(),
      CM.crosshairCursor(),
      CM.highlightActiveLine(),
      CM.highlightSelectionMatches(),
      CM.keymap.of([
        ...CM.closeBracketsKeymap,
        ...CM.defaultKeymap,
        ...CM.searchKeymap,
        ...CM.historyKeymap,
        ...CM.foldKeymap,
        ...CM.completionKeymap,
        CM.indentWithTab,
        { key: 'Mod-s', run: () => { saveNow(); return true; } },
        { key: 'Mod-f', run: () => { CM.openSearchPanel(editor); return true; } },
      ]),
      CM.indentUnit.of('  '),
      langCompartment.of(langExt ? langExt() : []),
      readonlyCompartment.of(CM.EditorState.readOnly.of(isReadonly)),
      CM.EditorView.updateListener.of(update => {
        if (update.docChanged) {
          onDocChanged();
        }
        if (update.selectionSet || update.docChanged) {
          updateCursorInfo(update.state);
        }
      }),
      CM.EditorView.theme({
        '&': { height: '100%', fontSize: '14px' },
        '.cm-scroller': { overflow: 'auto', fontFamily: "'Consolas', 'Monaco', 'Courier New', monospace" },
        '.cm-content': { minHeight: '200px' },
      }),
    ];

    editor = new CM.EditorView({
      state: CM.EditorState.create({
        doc: DOC_DATA.content,
        extensions,
      }),
      parent: document.getElementById('editor-container'),
    });
  }

  function updateCursorInfo(state) {
    const pos = state.selection.main.head;
    const line = state.doc.lineAt(pos);
    const el = document.getElementById('cursor-info');
    if (el) el.textContent = `줄 ${line.number}, 열 ${pos - line.from + 1}`;
  }

  // ---- Dirty tracking + auto-save ----
  function onDocChanged() {
    dirty = true;
    updateSaveStatus('modified');
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveNow, 2000);
  }

  async function saveNow() {
    if (!dirty || saving) return;
    if (DOC_DATA.status === 'CLOSED' && !IS_ADMIN) return;

    saving = true;
    updateSaveStatus('saving');

    const content = editor.state.doc.toString();
    try {
      const res = await apiFetch(`/api/text-documents/${DOC_DATA.id}/save`, {
        method: 'POST',
        body: JSON.stringify({
          content: content,
          base_version: currentVersion,
        }),
      });

      if (res.ok) {
        const data = await res.json();
        currentVersion = data.version;
        dirty = false;
        updateSaveStatus('saved');
        updateVersionInfo(data.version);
      } else if (res.status === 409) {
        // 버전 충돌 — 서버 최신 버전 수락
        const err = await res.json();
        const detail = err.detail;
        currentVersion = detail.server_version;
        replaceContent(detail.content);
        dirty = false;
        updateSaveStatus('conflict-resolved');
        updateVersionInfo(detail.server_version);
        showToast('다른 사용자가 수정한 내용으로 갱신되었습니다', 'warning');
      } else if (res.status === 403) {
        updateSaveStatus('readonly');
        showToast('문서가 마감되어 저장할 수 없습니다', 'error');
      } else {
        const err = await res.json();
        updateSaveStatus('error');
        showToast(err.detail || '저장 실패', 'error');
      }
    } catch (e) {
      updateSaveStatus('error');
      console.error('Save error:', e);
    } finally {
      saving = false;
    }
  }

  function replaceContent(newContent) {
    if (!editor) return;
    const currentContent = editor.state.doc.toString();
    if (currentContent === newContent) return;

    // 커서 위치 보존
    const pos = editor.state.selection.main.head;
    editor.dispatch({
      changes: { from: 0, to: editor.state.doc.length, insert: newContent },
      selection: { anchor: Math.min(pos, newContent.length) },
    });
  }

  // ---- Status display ----
  const STATUS_TEXT = {
    ready: '준비',
    modified: '수정됨',
    saving: '저장 중...',
    saved: '저장됨',
    error: '저장 오류',
    readonly: '읽기 전용',
    'conflict-resolved': '충돌 해결됨',
  };

  function updateSaveStatus(status) {
    const el = document.getElementById('save-status');
    if (!el) return;
    el.textContent = STATUS_TEXT[status] || status;
    el.className = 'doc-save-status status-' + status;
  }

  function updateVersionInfo(version) {
    const el = document.getElementById('version-info');
    if (el) el.textContent = `v${version}`;
  }

  // ---- Language change ----
  window.changeLang = async function (lang) {
    const langExt = LANG_MAP[lang];
    editor.dispatch({
      effects: langCompartment.reconfigure(langExt ? langExt() : []),
    });

    // 서버에 언어 변경 저장 (admin만)
    if (IS_ADMIN) {
      await apiFetch(`/api/admin/text-documents/${DOC_DATA.id}`, {
        method: 'PATCH',
        body: JSON.stringify({ language: lang }),
      });
      DOC_DATA.language = lang;
    }
  };

  // ---- Download ----
  window.downloadDoc = function () {
    const content = editor.state.doc.toString();
    const ext = EXT_MAP[DOC_DATA.language] || '.txt';
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = DOC_DATA.title + ext;
    a.click();
    URL.revokeObjectURL(url);
  };

  // ---- Close/Reopen ----
  window.toggleDocStatus = async function () {
    if (!IS_ADMIN) return;
    const action = DOC_DATA.status === 'OPEN' ? 'close' : 'reopen';
    if (action === 'close' && !confirm('이 문서를 마감하시겠습니까?')) return;

    const res = await apiFetch(`/api/admin/text-documents/${DOC_DATA.id}/${action}`, { method: 'POST' });
    if (res.ok) {
      const data = await res.json();
      DOC_DATA.status = data.status;

      // 배지 업데이트
      const badge = document.querySelector('.doc-info .badge');
      if (badge) {
        badge.textContent = data.status;
        badge.className = `badge ${data.status.toLowerCase()}`;
      }

      // 버튼 텍스트 업데이트
      const btn = document.getElementById('toggle-status-btn');
      if (btn) btn.textContent = data.status === 'OPEN' ? '마감' : '재개';

      // 읽기 전용 상태 변경
      const isReadonly = data.status === 'CLOSED' && !IS_ADMIN;
      editor.dispatch({
        effects: readonlyCompartment.reconfigure(CM.EditorState.readOnly.of(isReadonly)),
      });

      showToast(data.status === 'CLOSED' ? '마감되었습니다' : '재개되었습니다', 'success');
    } else {
      const err = await res.json();
      showToast(err.detail || '상태 변경 실패', 'error');
    }
  };

  // ---- WebSocket ----
  function initWebSocket() {
    const proto = location.protocol === 'https:' ? 'wss:' : 'ws:';
    const url = `${proto}//${location.host}/ws/text-documents/${DOC_DATA.id}`;

    setConnState('connecting');

    ws = new WebSocket(url);

    ws.onopen = () => {
      setConnState('connected');
      clearTimeout(wsReconnectTimer);
    };

    ws.onmessage = (event) => {
      let msg;
      try { msg = JSON.parse(event.data); } catch { return; }

      switch (msg.type) {
        case 'connected':
          break;
        case 'pong':
          break;
        case 'doc_updated':
          handleRemoteUpdate(msg);
          break;
        case 'cursor':
          // 다른 사용자 커서 (향후 확장)
          break;
      }
    };

    ws.onclose = () => {
      setConnState('disconnected');
      scheduleReconnect();
    };

    ws.onerror = () => {
      setConnState('disconnected');
    };

    // 주기적 ping
    setInterval(() => {
      if (ws && ws.readyState === WebSocket.OPEN) {
        ws.send(JSON.stringify({ type: 'ping' }));
      }
    }, 30000);
  }

  function handleRemoteUpdate(msg) {
    if (msg.updated_by === CURRENT_USER) return;

    currentVersion = msg.version;
    replaceContent(msg.content);
    dirty = false;
    clearTimeout(saveTimer);
    updateSaveStatus('saved');
    updateVersionInfo(msg.version);
  }

  function setConnState(state) {
    const dot = document.getElementById('conn-dot');
    if (!dot) return;
    dot.className = 'conn-dot ' + state;
  }

  function scheduleReconnect() {
    clearTimeout(wsReconnectTimer);
    wsReconnectTimer = setTimeout(() => {
      initWebSocket();
    }, 3000);
  }

})();
