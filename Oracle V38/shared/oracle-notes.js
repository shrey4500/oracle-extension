// oracle-notes.js — Notes/Scratchpad management + All Tasks view
// Exports: OracleNotes namespace via window.OracleNotes

(function () {
  'use strict';

  const { escapeHtml, WEBHOOK_URL, createAuthenticatedPayload, formatDate, state } = window.Oracle;

  // Helper: convert plain URLs in escaped text to clickable links with word-break
  function linkifyText(escapedText) {
    return escapedText.replace(/(https?:\/\/[^\s<>&]+)/g, '<a href="$1" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">$1</a>');
  }

  // ============================================
  // Notes CRUD
  // ============================================

  async function loadNotes() {
    const container = document.querySelector('.scratchpad-container');
    if (!container) return;
    const emptyState = container.querySelector('.empty-state');
    const notesList = document.querySelector('.notes-list');
    if (emptyState) emptyState.style.display = 'none';
    if (notesList) notesList.style.display = 'none';
    if (typeof showLoader === 'function') showLoader(container);

    try {
      if (!state.isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({ action: 'list_notes', timestamp: new Date().toISOString() }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);
      const data = await response.json();
      let notesData = [];
      if (Array.isArray(data) && data.length > 0 && data[0].notes) notesData = data[0].notes;
      else if (Array.isArray(data)) notesData = data;
      else if (data.notes) notesData = data.notes;

      state.allNotes = notesData;
      if (typeof hideLoader === 'function') hideLoader(container);
      if (state.allNotes.length > 0) displayNotes(state.allNotes);
      else showEmptyNotesState();
      if (typeof updateTabCounts === 'function') updateTabCounts();
    } catch (error) {
      console.error('Error loading notes:', error);
      if (typeof hideLoader === 'function') hideLoader(container);
      if (emptyState) { emptyState.style.display = 'flex'; emptyState.innerHTML = '<h3>Error: ' + error.message + '</h3>'; }
    }
  }

  function displayNotes(notes) {
    const container = document.querySelector('.scratchpad-container');
    if (!container) return;
    if (typeof hideLoader === 'function') hideLoader(container);
    let notesList = document.querySelector('.notes-list');
    const emptyState = container.querySelector('.empty-state');
    if (!notes || notes.length === 0) { showEmptyNotesState(); return; }
    if (emptyState) emptyState.style.display = 'none';
    if (!notesList) {
      notesList = document.createElement('div');
      notesList.className = 'notes-list';
      const form = document.getElementById('noteForm');
      if (form) form.parentNode.insertBefore(notesList, form.nextSibling);
    }

    const sorted = [...notes].sort((a, b) => new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at));
    notesList.innerHTML = sorted.map((note, i) => {
      const raw = note.description || '';
      const max = 150;
      const needsTrunc = raw.length > max;
      const trunc = needsTrunc ? raw.substring(0, max) + '...' : raw;
      const display = trunc.split('\n').map(l => escapeHtml(l)).join('<br>');
      return `<div class="note-item" style="animation-delay:${i * 0.1}s" data-note-id="${note.id}">
        <div class="note-actions">
          <button class="note-action-btn note-edit" data-note-id="${note.id}" title="Edit">✏️</button>
          <button class="note-action-btn note-delete delete" data-note-id="${note.id}" title="Delete">🗑️</button>
        </div>
        <div class="note-title">${escapeHtml(note.title || 'Untitled')}</div>
        ${display ? `<div class="note-description">${display}</div>` : ''}
        <div class="note-meta">
          <span class="note-date">${formatDate(note.created_at)}</span>
          ${needsTrunc ? `<span class="note-view-more" data-note-id="${note.id}">View more</span>` : ''}
        </div>
      </div>`;
    }).join('');
    notesList.style.display = 'flex';
    addNoteEventListeners();
  }

  function addNoteEventListeners() {
    document.querySelectorAll('.note-edit').forEach(el => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const note = state.allNotes.find(n => n.id == el.dataset.noteId);
        if (note) showNoteForm(note);
      });
    });
    document.querySelectorAll('.note-delete').forEach(el => {
      el.addEventListener('click', async (e) => {
        e.stopPropagation();
        if (confirm('Delete this note?')) {
          try {
            await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(createAuthenticatedPayload({ action: 'note_deleted', note_id: el.dataset.noteId, timestamp: new Date().toISOString() })) });
            await loadNotes();
          } catch (e) { console.error('Error deleting note:', e); }
        }
      });
    });
    document.querySelectorAll('.note-view-more').forEach(el => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const note = state.allNotes.find(n => n.id == el.dataset.noteId);
        if (note) showNoteViewer(note);
      });
    });
  }

  function showEmptyNotesState() {
    const container = document.querySelector('.scratchpad-container');
    if (!container) return;
    if (typeof hideLoader === 'function') hideLoader(container);
    const notesList = document.querySelector('.notes-list');
    const emptyState = container.querySelector('.empty-state');
    if (notesList) notesList.style.display = 'none';
    if (emptyState) { emptyState.style.display = 'flex'; emptyState.innerHTML = '<h3>No notes yet</h3><p>Click "New Note" to create your first note</p>'; }
  }

  // ============================================
  // Note Form (create/edit)
  // ============================================

  function showNoteForm(note = null) {
    state.currentEditingNoteId = note?.id || null;
    const noteTitle = document.getElementById('noteTitle');
    const noteDescription = document.getElementById('noteDescription');
    const noteForm = document.getElementById('noteForm');
    const saveNoteBtn = document.getElementById('saveNoteBtn');
    if (!noteForm || !noteTitle || !noteDescription) return;

    noteTitle.value = note?.title || '';
    noteDescription.value = note?.description || '';
    noteForm.classList.add('active');
    if (saveNoteBtn) saveNoteBtn.textContent = note ? 'Update Note' : 'Save Note';

    if (note) {
      setTimeout(() => { noteDescription.focus(); noteDescription.setSelectionRange(0, 0); noteDescription.scrollTop = 0; }, 10);
    } else {
      noteTitle.focus();
    }

    const notesList = document.querySelector('.notes-list');
    if (notesList) notesList.style.display = 'none';
    noteDescription.addEventListener('keydown', handleBulletedList);
  }

  function hideNoteForm() {
    const noteForm = document.getElementById('noteForm');
    const noteTitle = document.getElementById('noteTitle');
    const noteDescription = document.getElementById('noteDescription');
    if (noteForm) noteForm.classList.remove('active');
    if (noteTitle) noteTitle.value = '';
    if (noteDescription) { noteDescription.value = ''; noteDescription.removeEventListener('keydown', handleBulletedList); }
    state.currentEditingNoteId = null;
    const notesList = document.querySelector('.notes-list');
    if (notesList) notesList.style.display = 'flex';
  }

  async function saveNote() {
    const noteTitle = document.getElementById('noteTitle');
    const noteDescription = document.getElementById('noteDescription');
    const saveNoteBtn = document.getElementById('saveNoteBtn');
    const title = noteTitle?.value?.trim();
    const description = noteDescription?.value?.trim();
    if (!title) { alert('Please enter a note title'); return; }
    if (!state.isAuthenticated) { alert('Please log in'); return; }
    if (saveNoteBtn) { saveNoteBtn.disabled = true; saveNoteBtn.textContent = 'Saving...'; }
    try {
      const action = state.currentEditingNoteId ? 'note_edited' : 'note_created';
      const payload = { action, title, description, timestamp: new Date().toISOString(), source: 'oracle-chrome-extension' };
      if (state.currentEditingNoteId) payload.note_id = state.currentEditingNoteId;
      const response = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload(payload)) });
      if (response.ok) { hideNoteForm(); await loadNotes(); }
      else throw new Error('Failed to save note');
    } catch (e) { alert('Error saving note: ' + e.message); }
    if (saveNoteBtn) { saveNoteBtn.disabled = false; saveNoteBtn.textContent = 'Save Note'; }
  }

  function handleBulletedList(e) {
    if (e.key !== 'Enter' || e.shiftKey) return;
    e.preventDefault();
    const ta = e.target, cp = ta.selectionStart;
    const before = ta.value.substring(0, cp), after = ta.value.substring(cp);
    const lines = before.split('\n'), curr = lines[lines.length - 1];
    const bm = curr.match(/^([•\-]\s)/);
    if (bm) {
      if (curr.trim() === '•' || curr.trim() === '-') {
        const nb = before.substring(0, before.length - curr.length);
        ta.value = nb + '\n' + after;
        ta.selectionStart = ta.selectionEnd = nb.length + 1;
      } else {
        ta.value = before + '\n• ' + after;
        ta.selectionStart = ta.selectionEnd = cp + 3;
      }
    } else if (curr.trim().length > 0) {
      ta.value = before + '\n• ' + after;
      ta.selectionStart = ta.selectionEnd = cp + 3;
    } else {
      ta.value = before + '\n' + after;
      ta.selectionStart = ta.selectionEnd = cp + 1;
    }
  }

  // ============================================
  // Note Viewer (read-only slider)
  // ============================================

  function showNoteViewer(note) {
    document.querySelectorAll('.note-viewer-overlay').forEach(s => s.remove());
    const overlay = document.createElement('div');
    overlay.className = 'note-viewer-overlay';

    // Position overlay to cover col3 only (like transcript slider)
    const col3Rect = window.Oracle.getCol3Rect();
    if (col3Rect) {
      overlay.style.cssText = `position:fixed;top:${col3Rect.top}px;left:${col3Rect.left}px;width:${col3Rect.width}px;bottom:0;z-index:10000;display:flex;flex-direction:column;animation:fadeIn 0.15s ease-out;`;
    } else {
      overlay.style.cssText = 'position:fixed;top:0;right:0;bottom:0;width:400px;z-index:10000;display:flex;flex-direction:column;animation:fadeIn 0.15s ease-out;';
    }

    const isDark = document.body.classList.contains('dark-mode');
    const slider = document.createElement('div');
    slider.className = 'note-viewer-slider';
    slider.style.cssText = `width:100%;height:100%;background:${isDark?'#1f2940':'white'};box-shadow:-4px 0 20px rgba(0,0,0,${isDark?'0.4':'0.15'});display:flex;flex-direction:column;animation:slideInRight 0.3s ease-out;border-radius:12px;border:1px solid ${isDark?'rgba(255,255,255,0.08)':'rgba(225,232,237,0.6)'};overflow:hidden;`;
    slider.innerHTML = `
      <div style="padding:20px;border-bottom:1px solid ${isDark?'rgba(255,255,255,0.1)':'rgba(225,232,237,0.5)'};display:flex;align-items:center;justify-content:space-between;flex-shrink:0;">
        <div style="display:flex;align-items:center;gap:12px;">
          <div style="width:40px;height:40px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:20px;">📝</div>
          <div>
            <div style="font-weight:600;font-size:16px;color:${isDark?'#e8e8e8':'#2c3e50'};">${escapeHtml(note.title || 'Untitled Note')}</div>
            <div style="font-size:12px;color:${isDark?'#888':'#7f8c8d'};">${formatDate(note.updated_at || note.created_at)}</div>
          </div>
        </div>
        <div style="display:flex;gap:8px;">
          <button class="note-edit-btn" style="background:rgba(102,126,234,0.1);border:none;color:#667eea;width:36px;height:36px;border-radius:8px;cursor:pointer;font-size:16px;">✏️</button>
          <button class="note-close-btn" style="background:rgba(231,76,60,0.1);border:none;color:#e74c3c;width:36px;height:36px;border-radius:8px;cursor:pointer;font-size:18px;">×</button>
        </div>
      </div>
      <div style="flex:1;overflow-y:auto;padding:20px;">
        <div style="white-space:pre-wrap;font-size:14px;line-height:1.6;color:${isDark?'#b0b0b0':'#5d6d7e'};word-break:break-word;overflow-wrap:break-word;">${linkifyText(escapeHtml(note.description || 'No content'))}</div>
      </div>`;

    overlay.appendChild(slider);
    document.body.appendChild(overlay);

    const close = () => { overlay.style.animation='fadeOut 0.2s ease-out'; slider.style.animation='slideOutRight 0.3s ease-out'; document.removeEventListener('keydown', kh); setTimeout(()=>{ overlay.remove(); window.Oracle.collapseCol3AfterSlider(); },250); };
    const kh = (e) => { if (e.key === 'Escape') { e.preventDefault(); close(); } };
    document.addEventListener('keydown', kh);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    slider.querySelector('.note-close-btn').addEventListener('click', close);
    slider.querySelector('.note-edit-btn').addEventListener('click', () => { document.removeEventListener('keydown', kh); overlay.remove(); window.Oracle.collapseCol3AfterSlider(); showNoteForm(note); });
  }

  // ============================================
  // All Tasks (completed tasks view)
  // ============================================

  async function loadAllTasks() {
    const container = document.querySelector('.alltasks-container');
    if (!container) return;
    const emptyState = container.querySelector('.empty-state');
    const tasksList = container.querySelector('.alltasks-list');
    if (emptyState) emptyState.style.display = 'none';
    if (tasksList) tasksList.style.display = 'none';
    if (typeof showLoader === 'function') showLoader(container);

    try {
      if (!state.isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({ action: 'list_all_todos', timestamp: new Date().toISOString() }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);
      const rt = await response.text();
      let data = [];
      if (rt && rt.trim()) { try { data = JSON.parse(rt); } catch { data = []; } }
      const allItems = Array.isArray(data) ? data : (data?.todos || []);
      state.allCompletedTasks = allItems.sort((a, b) => new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at));
      if (typeof hideLoader === 'function') hideLoader(container);
      if (state.allCompletedTasks.length > 0) displayAllTasks(state.allCompletedTasks);
      else showEmptyAllTasksState();
    } catch (error) {
      console.error('Error loading all tasks:', error);
      if (typeof hideLoader === 'function') hideLoader(container);
      if (emptyState) { emptyState.style.display = 'flex'; emptyState.innerHTML = '<h3>Error: ' + escapeHtml(error.message) + '</h3>'; }
    }
  }

  function displayAllTasks(tasks) {
    const container = document.querySelector('.alltasks-container');
    if (!container) return;
    const emptyState = container.querySelector('.empty-state');
    const tasksList = container.querySelector('.alltasks-list');
    if (emptyState) emptyState.style.display = 'none';
    if (!tasks || !tasks.length) { showEmptyAllTasksState(); return; }
    if (!tasksList) return;

    // These functions are defined in newtab.js and depend on heavy DOM builders
    // Delegate to them via global scope
    if (typeof window._displayAllTasksImpl === 'function') {
      window._displayAllTasksImpl(tasks);
    } else {
      // Fallback: simple list
      tasksList.innerHTML = tasks.map(t => `<div class="todo-item" data-todo-id="${t.id}" style="padding:12px;border-bottom:1px solid rgba(225,232,237,0.5);">
        <div style="font-weight:600;font-size:13px;">${escapeHtml(t.task_title || t.message || 'Task')}</div>
        <div style="font-size:11px;color:#7f8c8d;margin-top:4px;">${formatDate(t.updated_at || t.created_at)}</div>
      </div>`).join('');
      tasksList.style.display = 'flex';
    }
  }

  function showEmptyAllTasksState() {
    const container = document.querySelector('.alltasks-container');
    if (!container) return;
    const tasksList = container.querySelector('.alltasks-list');
    const emptyState = container.querySelector('.empty-state');
    if (tasksList) tasksList.style.display = 'none';
    if (emptyState) { emptyState.style.display = 'flex'; emptyState.innerHTML = '<h3>No completed tasks</h3><p>Tasks marked done in the last 24 hours will appear here</p>'; }
  }

  async function searchAllTasks(query) {
    const container = document.querySelector('.alltasks-container');
    if (!container || !query.trim()) return;
    const emptyState = container.querySelector('.empty-state');
    const tasksList = container.querySelector('.alltasks-list');
    if (emptyState) emptyState.style.display = 'none';
    if (tasksList) tasksList.style.display = 'none';
    if (typeof showLoader === 'function') showLoader(container);

    try {
      if (!state.isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({ action: 'search_tasks', query: query.trim(), timestamp: new Date().toISOString() }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);
      const rt = await response.text();
      let data = [];
      if (rt && rt.trim()) { try { data = JSON.parse(rt); } catch { data = []; } }
      const results = Array.isArray(data) ? data : (data?.todos || data?.results || []);
      if (typeof hideLoader === 'function') hideLoader(container);
      if (results.length > 0) displayAllTasks(results);
      else if (emptyState) { emptyState.style.display = 'flex'; emptyState.innerHTML = `<h3>No results</h3><p>No tasks found matching "${escapeHtml(query)}"</p>`; }
    } catch (error) {
      console.error('Error searching tasks:', error);
      if (typeof hideLoader === 'function') hideLoader(container);
      if (emptyState) { emptyState.style.display = 'flex'; emptyState.innerHTML = '<h3>Search error</h3><p>' + escapeHtml(error.message) + '</p>'; }
    }
  }

  function setupAllTasksSearch() {
    const searchBtn = document.getElementById('allTasksSearchBtn');
    const searchClose = document.getElementById('allTasksSearchClose');
    const searchInput = document.getElementById('allTasksSearchInput');
    const doSearch = () => { const q = searchInput?.value?.trim(); if (q) searchAllTasks(q); };
    searchBtn?.addEventListener('click', doSearch);
    searchInput?.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); doSearch(); } });
    searchClose?.addEventListener('click', () => {
      const bar = document.getElementById('allTasksSearch');
      if (bar) bar.style.display = 'none';
      if (searchInput) searchInput.value = '';
      loadAllTasks();
    });
  }

  // ============================================
  // Setup notes button handlers
  // ============================================
  function setupNoteButtons() {
    const addNoteBtn = document.getElementById('addNoteBtn');
    const refreshNotesBtn = document.getElementById('refreshNotesBtn');
    const saveNoteBtn = document.getElementById('saveNoteBtn');
    const cancelNoteBtn = document.getElementById('cancelNoteBtn');

    addNoteBtn?.addEventListener('click', () => {
      const activeTab = document.querySelector('.col3-tab.active');
      if (activeTab?.dataset?.col3 === 'alltasks') {
        const searchBar = document.getElementById('allTasksSearch');
        if (searchBar) {
          const vis = searchBar.style.display !== 'none';
          searchBar.style.display = vis ? 'none' : 'block';
          if (!vis) document.getElementById('allTasksSearchInput')?.focus();
        }
      } else {
        showNoteForm();
      }
    });
    cancelNoteBtn?.addEventListener('click', hideNoteForm);
    saveNoteBtn?.addEventListener('click', saveNote);
    refreshNotesBtn?.addEventListener('click', loadNotes);
  }

  // ============================================
  // EXPORT
  // ============================================
  window.OracleNotes = {
    loadNotes,
    displayNotes,
    showNoteForm,
    hideNoteForm,
    saveNote,
    showNoteViewer,
    showEmptyNotesState,
    loadAllTasks,
    displayAllTasks,
    showEmptyAllTasksState,
    searchAllTasks,
    setupAllTasksSearch,
    setupNoteButtons,
  };

})();
