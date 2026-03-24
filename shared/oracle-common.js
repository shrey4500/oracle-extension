// oracle-common.js — Shared constants, state, auth, utilities, read state
// Used by both newtab.js and sidepanel.js

(function () {
  'use strict';

  // ============================================
  // CONSTANTS
  // ============================================
  const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';
  const AUTH_URL = 'https://n8n-kqq5.onrender.com/webhook/e6bcd2c3-c714-46c7-94b8-8aeb9831429c';
  const CHAT_WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4';
  const STORAGE_KEY = 'oracle_user_data';
  const READ_TASKS_KEY = 'oracle_read_tasks';
  const TASK_TIMESTAMPS_KEY = 'oracle_task_timestamps';
  const ABLY_REFRESH_THROTTLE_MS = 30000;

  // ============================================
  // SHARED MUTABLE STATE
  // ============================================
  const state = {
    allTodos: [],
    allFyiItems: [],
    allCalendarItems: [],
    allCompletedTasks: [],
    allBookmarks: [],
    allNotes: [],
    allDocuments: [],
    userData: null,
    isAuthenticated: false,
    readTaskIds: new Set(),
    isInitialLoad: true,
    previousTaskTimestamps: new Map(),
    previousMeetingIds: new Set(),
    previousMeetingTimestamps: new Map(),
    modifiedMeetingDates: new Set(),
    activeTagFilters: [],
    recentlyCompletedIds: new Set(),
    pendingActionUpdates: [],
    pendingFyiUpdates: [],
    pendingActionData: null,
    pendingFyiData: null,
    lastAblyRefreshTime: 0,
    isTranscriptSliderOpen: false,
    isChatSliderOpen: false,
    isEditModeActive: false,
    currentEditingNoteId: null,
  };

  // ============================================
  // RECENTLY COMPLETED TRACKING
  // ============================================
  function addRecentlyCompleted(taskId) {
    const idStr = String(taskId);
    state.recentlyCompletedIds.add(idStr);
    setTimeout(() => {
      state.recentlyCompletedIds.delete(idStr);
    }, 120000);
  }

  function isRecentlyCompleted(taskId) {
    return state.recentlyCompletedIds.has(String(taskId));
  }

  // ============================================
  // READ STATE MANAGEMENT
  // ============================================
  function loadReadState() {
    try {
      const savedReadTasks = localStorage.getItem(READ_TASKS_KEY);
      if (savedReadTasks) {
        state.readTaskIds = new Set(JSON.parse(savedReadTasks));
      } else {
        state.readTaskIds = new Set();
      }

      const savedTimestamps = localStorage.getItem(TASK_TIMESTAMPS_KEY);
      if (savedTimestamps) {
        state.previousTaskTimestamps = new Map(Object.entries(JSON.parse(savedTimestamps)));
      } else {
        state.previousTaskTimestamps = new Map();
      }

      state.isInitialLoad = (state.readTaskIds.size === 0 && state.previousTaskTimestamps.size === 0);
    } catch (e) {
      console.error('Error loading read state:', e);
      state.readTaskIds = new Set();
      state.previousTaskTimestamps = new Map();
      state.isInitialLoad = true;
    }
  }

  function saveReadState() {
    try {
      localStorage.setItem(READ_TASKS_KEY, JSON.stringify([...state.readTaskIds]));
      const timestampsObj = Object.fromEntries(state.previousTaskTimestamps);
      localStorage.setItem(TASK_TIMESTAMPS_KEY, JSON.stringify(timestampsObj));
    } catch (e) {
      console.error('Error saving read state:', e);
    }
  }

  function markTaskAsRead(todoId) {
    const idStr = String(todoId);
    if (!state.readTaskIds.has(idStr)) {
      state.readTaskIds.add(idStr);
      saveReadState();
    }
    // Always update UI — find ALL DOM elements for this task and ensure unread is removed
    const selectors = [
      `.todo-item[data-todo-id="${todoId}"]`,
      `.task-group-task-item[data-todo-id="${todoId}"]`,
      `.slack-channel-task-item[data-todo-id="${todoId}"]`,
      `.document-task-item[data-todo-id="${todoId}"]`
    ];
    selectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        el.classList.remove('unread');
        checkParentGroupReadState(el);
      });
    });
  }

  // Check if all tasks in a parent group are now read; if so, remove group's unread class
  // Also update the group's task count display
  function checkParentGroupReadState(taskElement) {
    const group = taskElement.closest('.task-group, .slack-channel-group, .slack-channel-accordion-item, .document-accordion-item, .fyi-tag-group, .tag-group');
    if (!group) return;
    const unreadChildren = group.querySelectorAll('.todo-item.unread, .task-group-task-item.unread, .slack-channel-task-item.unread, .slack-channel-task-full.unread, .document-task-item.unread');
    if (unreadChildren.length === 0) {
      group.classList.remove('unread');
      // Also check if this is inside a Slack Channels accordion
      const slackAccordion = group.closest('.slack-channels-accordion');
      if (slackAccordion) {
        const anyUnreadChannels = slackAccordion.querySelectorAll('.slack-channel-accordion-item.unread');
        if (anyUnreadChannels.length === 0) {
          slackAccordion.classList.remove('has-unread');
        }
      }
      // Also check if this is inside a Documents accordion
      const docsAccordion = group.closest('.documents-accordion');
      if (docsAccordion) {
        const anyUnreadDocs = docsAccordion.querySelectorAll('.document-accordion-item.unread');
        if (anyUnreadDocs.length === 0) {
          docsAccordion.classList.remove('has-unread');
        }
      }
    }
    // Update the group's visible task count (subtract completed/removed tasks)
    const countEl = group.querySelector('.task-group-count, .slack-channel-item-meta');
    if (countEl) {
      const totalTasks = group.querySelectorAll('.todo-item:not(.completing):not([style*="display: none"]), .task-group-task-item:not(.completing):not([style*="display: none"]), .slack-channel-task-item:not(.completing):not([style*="display: none"]), .document-task-item:not(.completing):not([style*="display: none"])');
      // Only update if count element has the standard format
      const currentText = countEl.textContent;
      const countMatch = currentText.match(/^(\d+)\s+(action item|item|task|update)/);
      if (countMatch) {
        const label = countMatch[2];
        countEl.textContent = currentText.replace(/^\d+/, totalTasks.length);
      }
    }
  }

  function markTaskAsUnread(todoId) {
    const idStr = String(todoId);
    state.readTaskIds.delete(idStr);
    saveReadState();
  }

  function isTaskUnread(todoId) {
    if (state.isInitialLoad) return false;
    return !state.readTaskIds.has(String(todoId));
  }

  function markAllCurrentTasksAsRead() {
    const allCurrentTasks = [...state.allTodos, ...state.allFyiItems];
    allCurrentTasks.forEach(task => {
      const idStr = String(task.id);
      state.readTaskIds.add(idStr);
      if (task.updated_at) {
        state.previousTaskTimestamps.set(idStr, task.updated_at);
      }
    });
    state.isInitialLoad = false;
    saveReadState();
  }

  function updateLastRefreshTime() {
    state.lastAblyRefreshTime = Date.now();
  }

  // ============================================
  // AUTH
  // ============================================
  async function initAuth() {
    try {
      const r = await chrome.storage.local.get([STORAGE_KEY]);
      if (r[STORAGE_KEY] && r[STORAGE_KEY].userId) {
        state.userData = r[STORAGE_KEY];
        state.isAuthenticated = true;
        return true;
      }
      return false;
    } catch {
      return false;
    }
  }

  async function login(email, password) {
    try {
      const res = await fetch(AUTH_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'login', email, password, timestamp: new Date().toISOString() })
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      let data;
      try { data = await res.json(); } catch { const t = await res.text(); data = !isNaN(t) ? parseInt(t.trim()) : null; }
      let userId = typeof data === 'number' ? data
        : typeof data === 'string' && !isNaN(data) ? parseInt(data)
        : Array.isArray(data) && data[0] ? (data[0].id || data[0].user_id || data[0])
        : data?.id || data?.user_id || data?.userId;
      if (!userId) throw new Error('No user ID');
      state.userData = { userId, email, loginTime: new Date().toISOString() };
      await chrome.storage.local.set({ [STORAGE_KEY]: state.userData });
      state.isAuthenticated = true;
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  async function logout() {
    await chrome.storage.local.remove([STORAGE_KEY]);
    state.userData = null;
    state.isAuthenticated = false;
  }

  function createAuthenticatedPayload(basePayload) {
    if (!state.isAuthenticated || !state.userData) throw new Error('Not authenticated');
    return { ...basePayload, user_id: state.userData.userId, authenticated: true };
  }

  // ============================================
  // UTILITIES
  // ============================================
  function escapeHtml(text) {
    return (text || '').toString()
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function sortTodos(todos) {
    return [...todos].sort((a, b) => {
      const aD = a.due_by ? new Date(a.due_by).getTime() : null;
      const bD = b.due_by ? new Date(b.due_by).getTime() : null;
      if (aD && bD) return aD - bD;
      if (aD && !bD) return -1;
      if (!aD && bD) return 1;
      const aU = a.updated_at ? new Date(a.updated_at).getTime() : null;
      const bU = b.updated_at ? new Date(b.updated_at).getTime() : null;
      if (aU && bU) return bU - aU;
      if (aU && !bU) return -1;
      if (!aU && bU) return 1;
      return new Date(b.created_at) - new Date(a.created_at);
    });
  }

  function formatDate(dateString) {
    if (!dateString) return '';
    const diff = Date.now() - new Date(dateString);
    const m = Math.floor(diff / 60000);
    const h = Math.floor(diff / 3600000);
    const days = Math.floor(diff / 86400000);
    if (m < 1) return 'Just now';
    if (m < 60) return m + 'm ago';
    if (h < 24) return h + 'h ago';
    if (days < 7) return days + 'd ago';
    if (days < 30) return Math.floor(days / 7) + 'w ago';
    return Math.floor(days / 30) + 'mo ago';
  }

  function formatDueBy(dueDateString) {
    if (!dueDateString) return null;
    const diff = new Date(dueDateString) - new Date();
    if (diff < 0) {
      const h = Math.floor(Math.abs(diff) / 3600000);
      const d = Math.floor(Math.abs(diff) / 86400000);
      return h < 24 ? `Overdue ${h}h` : `Overdue ${d}d`;
    }
    const mins = Math.floor(diff / 60000);
    const hours = Math.floor(diff / 3600000);
    const days = Math.floor(diff / 86400000);
    if (mins < 60) return `Due ${mins}m`;
    if (hours < 24) return `Due ${hours}h ${mins % 60}m`;
    return `Due ${days}d ${hours % 24}h`;
  }

  function isValidTag(tag) {
    if (!tag || typeof tag !== 'string') return false;
    const t = tag.trim().toLowerCase();
    return t !== '' && t !== 'null' && t !== 'undefined' && t !== 'none' && t !== '0';
  }

  function showLoader(container) {
    let l = container.querySelector('.loading-overlay');
    if (!l) {
      l = document.createElement('div');
      l.className = 'loading-overlay';
      l.innerHTML = '<div class="spinner"></div>';
      container.appendChild(l);
    }
    l.style.display = 'flex';
  }

  function hideLoader(container) {
    const l = container.querySelector('.loading-overlay');
    if (l) l.style.display = 'none';
  }

  function getDateKey(dateTime) {
    const d = new Date(dateTime);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function formatDisplayName(name) {
    if (!name) return 'Unknown';
    if ((name.includes('.') || name.includes('_')) && !name.includes(' ')) {
      return name.split(/[._]/).map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(' ');
    }
    if (name.includes(' ')) {
      return name.split(' ').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
    }
    return name.charAt(0).toUpperCase() + name.slice(1);
  }

  function extractOrganiser(participantText) {
    if (!participantText) return '';
    const names = participantText.split(',').map(n => n.trim()).filter(Boolean);
    return names.length > 0 ? formatDisplayName(names[0]) : '';
  }

  // ============================================
  // LOGIN SCREEN (shared between newtab & sidepanel)
  // ============================================
  function showLoginScreen(onSuccess) {
    const c = document.querySelector('.container');
    if (c) c.style.display = 'none';
    document.body.insertAdjacentHTML('beforeend', `
      <div class="login-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);display:flex;align-items:center;justify-content:center;z-index:10000;">
        <div style="background:rgba(255,255,255,0.95);backdrop-filter:blur(20px);border-radius:20px;padding:30px;box-shadow:0 20px 40px rgba(0,0,0,0.2);max-width:320px;width:90%;">
          <div style="text-align:center;margin-bottom:24px;">
            <div style="width:50px;height:50px;margin:0 auto 12px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:24px;color:white;font-weight:bold;">∞</div>
            <h2 style="margin:0 0 6px;color:#2c3e50;font-size:20px;">Welcome to Oracle</h2>
            <p style="margin:0;color:#7f8c8d;font-size:13px;">Sign in to continue</p>
          </div>
          <form id="oracleLoginForm" style="display:flex;flex-direction:column;gap:16px;">
            <input type="email" id="loginEmail" required placeholder="Email" style="width:100%;padding:12px;border:2px solid rgba(225,232,237,0.8);border-radius:10px;font-size:14px;box-sizing:border-box;">
            <input type="password" id="loginPassword" required placeholder="Password" style="width:100%;padding:12px;border:2px solid rgba(225,232,237,0.8);border-radius:10px;font-size:14px;box-sizing:border-box;">
            <button type="submit" id="loginSubmitBtn" style="width:100%;background:linear-gradient(45deg,#667eea,#764ba2);color:white;border:none;padding:12px;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;">Sign In</button>
            <div id="loginError" style="display:none;background:rgba(231,76,60,0.1);border:1px solid rgba(231,76,60,0.3);color:#e74c3c;padding:10px;border-radius:6px;font-size:13px;text-align:center;"></div>
          </form>
        </div>
      </div>`);
    document.getElementById('oracleLoginForm').addEventListener('submit', async (e) => {
      e.preventDefault();
      const btn = document.getElementById('loginSubmitBtn');
      const err = document.getElementById('loginError');
      btn.disabled = true; btn.textContent = 'Signing in...'; err.style.display = 'none';
      const result = await login(
        document.getElementById('loginEmail').value.trim(),
        document.getElementById('loginPassword').value.trim()
      );
      if (result.success) {
        document.querySelector('.login-overlay').remove();
        document.querySelector('.container').style.display = 'flex';
        if (onSuccess) onSuccess();
      } else {
        err.textContent = result.error; err.style.display = 'block';
      }
      btn.disabled = false; btn.textContent = 'Sign In';
    });
  }

  // ============================================
  // DARK MODE HELPERS
  // ============================================
  function isDarkMode() {
    return document.body.classList.contains('dark-mode');
  }

  function initDarkMode() {
    if (localStorage.getItem('oracle_dark_mode') === 'true') {
      document.body.classList.add('dark-mode');
    }
  }

  function toggleDarkMode() {
    document.body.classList.toggle('dark-mode');
    localStorage.setItem('oracle_dark_mode', document.body.classList.contains('dark-mode'));
  }

  // ============================================
  // COL3 SLIDER HELPERS
  // Expand col3 when a slider opens (if collapsed), collapse on close
  // ============================================
  let _col3WasCollapsed = false;

  function expandCol3ForSlider() {
    const col3 = document.getElementById('col3');
    const layout = document.querySelector('.three-column-layout');
    const toggle = document.getElementById('sidebarToggle');
    if (!col3 || !layout) return;

    _col3WasCollapsed = col3.classList.contains('col3-hidden');
    if (_col3WasCollapsed) {
      col3.classList.remove('col3-hidden');
      layout.classList.remove('col3-collapsed');
      if (toggle) toggle.classList.remove('collapsed');
    }
  }

  function collapseCol3AfterSlider() {
    if (!_col3WasCollapsed) return;
    const col3 = document.getElementById('col3');
    const layout = document.querySelector('.three-column-layout');
    const toggle = document.getElementById('sidebarToggle');
    if (!col3 || !layout) return;

    col3.classList.add('col3-hidden');
    layout.classList.add('col3-collapsed');
    if (toggle) toggle.classList.add('collapsed');
    _col3WasCollapsed = false;
  }

  function getCol3Rect() {
    expandCol3ForSlider();
    // Force reflow so getBoundingClientRect returns correct values
    const col3 = document.getElementById('col3');
    if (col3) col3.offsetHeight; // trigger reflow
    return col3 ? col3.getBoundingClientRect() : null;
  }

  // ============================================
  // EXPORT
  // ============================================
  window.Oracle = {
    // Constants
    WEBHOOK_URL,
    AUTH_URL,
    CHAT_WEBHOOK_URL,
    STORAGE_KEY,
    READ_TASKS_KEY,
    TASK_TIMESTAMPS_KEY,
    ABLY_REFRESH_THROTTLE_MS,

    // State
    state,

    // Auth
    initAuth,
    login,
    logout,
    createAuthenticatedPayload,

    // Utilities
    escapeHtml,
    sortTodos,
    formatDate,
    formatDueBy,
    isValidTag,
    showLoader,
    hideLoader,
    getDateKey,
    formatDisplayName,
    extractOrganiser,

    // Read state
    loadReadState,
    saveReadState,
    markTaskAsRead,
    checkParentGroupReadState,
    markTaskAsUnread,
    isTaskUnread,
    markAllCurrentTasksAsRead,
    updateLastRefreshTime,

    // Recently completed
    addRecentlyCompleted,
    isRecentlyCompleted,

    // Login
    showLoginScreen,

    // Dark mode
    isDarkMode,
    initDarkMode,
    toggleDarkMode,

    // Col3 slider helpers
    expandCol3ForSlider,
    collapseCol3AfterSlider,
    getCol3Rect,
  };

})();
