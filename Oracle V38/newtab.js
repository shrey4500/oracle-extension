// Complete newtab.js with Scratchpad, Bookmarks, and Todos functionality
const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';
const AUTH_URL = 'https://n8n-kqq5.onrender.com/webhook/e6bcd2c3-c714-46c7-94b8-8aeb9831429c';
const STORAGE_KEY = 'oracle_user_data';
const READ_TASKS_KEY = 'oracle_read_tasks';
const TASK_TIMESTAMPS_KEY = 'oracle_task_timestamps';
let currentResponse = '';
let allTodos = [];
let allFyiItems = [];
let allCalendarItems = [];
let allBookmarks = [];
let allNotes = [];
let allDocuments = [];
let userData = null;
let isAuthenticated = false;
let currentEditingNoteId = null;
let readTaskIds = new Set(); // Track which tasks have been read
let isInitialLoad = true; // True until first data load completes
let previousTaskTimestamps = new Map(); // Track task updated_at timestamps to detect modifications
let modifiedMeetingDates = new Set(); // Track dates with modified/new meetings from Ably updates
let previousMeetingIds = new Set(); // Track meeting IDs from previous load to detect new/modified meetings
let previousMeetingTimestamps = new Map(); // Track meeting due_by + status to detect meaningful changes

// Convert Zoom web URLs to zoommtg:// protocol to open desktop app directly
function convertZoomToDeepLink(url) {
  try {
    const urlObj = new URL(url);
    if (!urlObj.hostname.includes('zoom.us') && !urlObj.hostname.includes('zoom.com')) return url;
    const pathMatch = urlObj.pathname.match(/\/j\/(\d+)/);
    if (!pathMatch) return url;
    const meetingId = pathMatch[1];
    const pwd = urlObj.searchParams.get('pwd') || '';
    return `zoommtg://zoom.us/join?confno=${meetingId}${pwd ? '&pwd=' + pwd : ''}`;
  } catch (e) {
    return url;
  }
}

// Pending updates for feed-like UX (LinkedIn/Instagram style)
let pendingActionUpdates = []; // Queued updates for Action tab list
let pendingFyiUpdates = []; // Queued updates for FYI tab list
let pendingActionData = null; // Full fetched data waiting to be displayed
let pendingFyiData = null; // Full fetched data waiting to be displayed

// Track recently completed task IDs to prevent them reappearing via pending updates
// Each entry: { id, timestamp } - auto-expires after 120 seconds
let recentlyCompletedIds = new Set();

function addRecentlyCompleted(taskId) {
  const idStr = String(taskId);
  recentlyCompletedIds.add(idStr);
  // Auto-remove after 120 seconds (enough time for backend to propagate)
  setTimeout(() => {
    recentlyCompletedIds.delete(idStr);
    console.log(`🗑️ Expired recently-completed tracking for task ${idStr}`);
  }, 120000);
  console.log(`✅ Tracking recently-completed task ${idStr} (${recentlyCompletedIds.size} total)`);
}

function isRecentlyCompleted(taskId) {
  return recentlyCompletedIds.has(String(taskId));
}

// Check if a parent group/accordion is now empty after a task was removed; if so, remove the group
function checkEmptyParentGroup(removedTaskElement) {
  if (!removedTaskElement) return;
  // Find the parent group before the element is removed
  const group = removedTaskElement.closest('.task-group, .slack-channel-accordion-item, .document-accordion-item, .fyi-tag-group');
  if (!group) return;
  // Use setTimeout so this runs AFTER the task is removed from DOM
  setTimeout(() => {
    const remainingTasks = group.querySelectorAll('.todo-item:not(.completing), .task-group-task-item:not(.completing), .slack-channel-task-item:not(.completing), .document-task-item:not(.completing)');
    if (remainingTasks.length === 0) {
      group.classList.add('completing');
      setTimeout(() => {
        group.remove();
        // If this was inside Slack Channels accordion, check if accordion is now empty
        // (the group may already be removed, so we check the DOM directly)
        document.querySelectorAll('.slack-channels-accordion').forEach(accordion => {
          const remainingChannels = accordion.querySelectorAll('.slack-channel-accordion-item');
          if (remainingChannels.length === 0) accordion.remove();
        });
      }, 400);
    }
  }, 50);
}

// Read state management functions - now persisted to localStorage

// ============================================
// SHARED COMPONENT ALIASES (Phase 7 migration)
// ============================================
const { loadReadState, saveReadState, markTaskAsRead, isTaskUnread, isValidTag, sortTodos, formatDate, formatDueBy, escapeHtml } = window.Oracle;
const { isMeetingLink, isDriveLink, isSlackLink, getSlackChannelUrl, extractDriveFileId, getCleanDriveFileUrl } = window.OracleIcons;
const { groupTasksByDriveFile, groupTasksBySlackChannel, groupTasksByTag, extractFileNameFromTask } = window.OracleGrouping;

function markTaskAsUnread(todoId) { readTaskIds.delete(String(todoId)); saveReadState(); }
function markAllCurrentTasksAsRead() {
  [...allTodos, ...allFyiItems, ...allCalendarItems].forEach(task => {
    readTaskIds.add(String(task.id));
    if (task.updated_at) previousTaskTimestamps.set(String(task.id), task.updated_at);
  });
  saveReadState();
}

// Multi-select state
let isMultiSelectMode = false;
let selectedTodoIds = new Set();

// Track Drive file IDs that have task groups in Action tab
let actionTabDriveFileIds = new Set();

// Track Slack task IDs that are shown in Action tab
let actionTabSlackTaskIds = new Set();

// Pending refresh flag for when tab becomes inactive
let pendingRefreshFlag = false;

// Ably refresh throttling
let lastAblyRefreshTime = 0;
const ABLY_REFRESH_THROTTLE_MS = 30000;

// Update refresh timestamp - call this on ANY data refresh (manual or Ably-triggered)
function updateLastRefreshTime() {
  lastAblyRefreshTime = Date.now();
  console.log(`⏱️ Refresh timestamp updated: ${new Date(lastAblyRefreshTime).toLocaleTimeString()}`);
}

// Keyboard navigation state
let keyboardSelectedTaskId = null;
let keyboardSelectedNoteId = null;
let currentKeyboardColumn = null; // 'action', 'fyi', or 'notes'

// Interaction state flags (to prevent auto-refresh during user interaction)
let isTranscriptSliderOpen = false;
let isChatSliderOpen = false;
let isEditModeActive = false;

// Tag filter state
let activeTagFilters = [];
let currentTodoFilter = 'starred'; // Track current filter for re-rendering

// Function to update tab counts and column counts
function updateTabCounts() {
  const todoCount = document.getElementById('todoCount');
  const fyiCount = document.getElementById('fyiCount');
  const scratchpadCount = document.getElementById('scratchpadCount');
  const bookmarkCount = document.getElementById('bookmarkCount');
  const documentsCount = document.getElementById('documentsCount');

  // Column counts for 3-column layout
  const actionColumnCount = document.getElementById('actionColumnCount');
  const fyiColumnCount = document.getElementById('fyiColumnCount');
  const scratchpadColumnCount = document.getElementById('scratchpadColumnCount');
  const bookmarkColumnCount = document.getElementById('bookmarkColumnCount');
  const documentsColumnCount = document.getElementById('documentsColumnCount');

  const meetingCount = allCalendarItems ? allCalendarItems.length : allTodos.filter(t => isMeetingLink(t.message_link) && t.status === 0).length;
  const starredCount = allTodos.filter(t => t.starred === 1 && !isMeetingLink(t.message_link)).length;
  const fyiItemsCount = allFyiItems.length > 0
    ? allFyiItems.filter(t => !isMeetingLink(t.message_link) && !isDriveLink(t.message_link)).length
    : allTodos.filter(t => t.starred === 0 && t.status === 0 && !isMeetingLink(t.message_link) && !isDriveLink(t.message_link)).length;

  if (todoCount) {
    todoCount.textContent = starredCount > 0 ? ` (${starredCount})` : '';
  }
  if (actionColumnCount) {
    actionColumnCount.textContent = starredCount;
  }

  if (fyiCount) {
    fyiCount.textContent = fyiItemsCount > 0 ? ` (${fyiItemsCount})` : '';
  }
  if (fyiColumnCount) {
    fyiColumnCount.textContent = fyiItemsCount;
  }

  if (scratchpadCount) {
    scratchpadCount.textContent = allNotes.length > 0 ? ` (${allNotes.length})` : '';
  }
  if (scratchpadColumnCount) {
    scratchpadColumnCount.textContent = allNotes.length;
  }

  if (bookmarkCount) {
    bookmarkCount.textContent = allBookmarks.length > 0 ? ` (${allBookmarks.length})` : '';
  }
  if (bookmarkColumnCount) {
    bookmarkColumnCount.textContent = allBookmarks.length;
  }

  if (documentsCount) {
    documentsCount.textContent = allDocuments.length > 0 ? ` (${allDocuments.length})` : '';
  }
  if (documentsColumnCount) {
    documentsColumnCount.textContent = allDocuments.length;
  }
}

let isCommandKeyPressed = false;

// Track Command key state
document.addEventListener('keydown', (e) => {
  if (e.key === 'Meta' || e.key === 'Control') {
    isCommandKeyPressed = true;
  }

  // Check if any input/textarea is focused or slider is open
  const activeElement = document.activeElement;
  const isInputFocused = activeElement && (
    activeElement.tagName === 'INPUT' ||
    activeElement.tagName === 'TEXTAREA' ||
    activeElement.isContentEditable
  );
  const isSliderOpen = document.querySelector('.transcript-slider-overlay') || document.querySelector('.note-viewer-overlay');
  const isNoteFormActive = document.querySelector('.note-form.active');

  // Command+Enter: Bulk mark selected todos as done (works anywhere on home screen)
  if ((e.metaKey || e.ctrlKey) && e.key === 'Enter' && !isSliderOpen && !isNoteFormActive) {
    if (selectedTodoIds.size > 0) {
      e.preventDefault();
      e.stopPropagation(); // Prevent any other handlers from firing
      // Clear keyboard selection to prevent slider from opening
      clearKeyboardSelection();
      handleBulkMarkDone();
      return;
    }
    // If keyboard selection active, mark that task as done
    if (keyboardSelectedTaskId && (currentKeyboardColumn === 'action' || currentKeyboardColumn === 'fyi')) {
      e.preventDefault();
      e.stopPropagation();
      markKeyboardSelectedTaskDone();
      return;
    }
  }

  // Escape - Clear all selections or close note form
  if (e.key === 'Escape') {
    if (isNoteFormActive) {
      e.preventDefault();
      if (typeof window.hideNoteForm === 'function') {
        window.hideNoteForm();
      }
      return;
    }
    if (!isInputFocused && !isSliderOpen) {
      e.preventDefault();
      clearKeyboardSelection();
      clearMultiSelection();
      return;
    }
  }

  // Cmd/Ctrl + Shift + Space - Open Quick Chat slider (Oracle)
  if ((e.metaKey || e.ctrlKey) && e.shiftKey && e.code === 'Space') {
    e.preventDefault();
    showChatSlider();
    return;
  }

  // Skip other shortcuts if input is focused or slider is open
  if (isInputFocused || isSliderOpen || isNoteFormActive) return;

  // Press "N" - Open New Message slider (only if not in notes column)
  if ((e.key === 'n' || e.key === 'N') && currentKeyboardColumn !== 'notes') {
    e.preventDefault();
    showNewMessageSlider();
    return;
  }

  // Press "F" - Toggle Fullscreen
  if (e.key === 'f' || e.key === 'F') {
    e.preventDefault();
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => console.error('Fullscreen error:', err));
    } else {
      document.exitFullscreen().catch(err => console.error('Exit fullscreen error:', err));
    }
    return;
  }

  // Press "1" - Select first task in Action column
  if (e.key === '1') {
    e.preventDefault();
    selectFirstTaskInColumn('action');
    return;
  }

  // Press "2" - Select first task in FYI column
  if (e.key === '2') {
    e.preventDefault();
    selectFirstTaskInColumn('fyi');
    return;
  }

  // Press "3" - Select first note in Notes column
  if (e.key === '3') {
    e.preventDefault();
    selectFirstNoteInColumn();
    return;
  }

  // Notes column shortcuts when notes column is active
  if (currentKeyboardColumn === 'notes') {
    // N - Create new note
    if (e.key === 'n' || e.key === 'N') {
      e.preventDefault();
      if (typeof window.showNoteForm === 'function') {
        window.showNoteForm();
      }
      return;
    }

    // E - Edit selected note
    if ((e.key === 'e' || e.key === 'E') && keyboardSelectedNoteId) {
      e.preventDefault();
      const note = (typeof allNotes !== 'undefined' ? allNotes : []).find(n => n.id == keyboardSelectedNoteId);
      if (note && typeof window.showNoteForm === 'function') {
        window.showNoteForm(note);
      }
      return;
    }

    // Enter - Open note viewer
    if (e.key === 'Enter' && keyboardSelectedNoteId) {
      e.preventDefault();
      const note = (typeof allNotes !== 'undefined' ? allNotes : []).find(n => n.id == keyboardSelectedNoteId);
      if (note && typeof window.showNoteViewer === 'function') {
        window.showNoteViewer(note);
      }
      return;
    }

    // Arrow navigation for notes
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      moveNoteSelection('down');
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      moveNoteSelection('up');
      return;
    }
  }

  // Only handle task navigation keys if we have a keyboard selection in action/fyi
  // and we're not inside a text input or contenteditable
  const activeEl = document.activeElement;
  const isInInput = activeEl && (activeEl.tagName === 'INPUT' || activeEl.tagName === 'TEXTAREA' || activeEl.isContentEditable || activeEl.closest('[contenteditable]'));
  if (keyboardSelectedTaskId && (currentKeyboardColumn === 'action' || currentKeyboardColumn === 'fyi') && !isInInput) {
    // Enter - Open transcript slider for selected task (but not if Cmd/Ctrl is held)
    if (e.key === 'Enter' && !e.metaKey && !e.ctrlKey) {
      e.preventDefault();
      e.stopPropagation();
      if (typeof window.showTranscriptSlider === 'function') {
        window.showTranscriptSlider(keyboardSelectedTaskId);
      }
      return;
    }

    // Down Arrow - Move to next task
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (e.shiftKey) {
        // Shift+Down - Add to multi-selection
        extendSelectionDown();
      } else {
        // Just move selection down
        moveKeyboardSelection('down');
      }
      return;
    }

    // Up Arrow - Move to previous task
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (e.shiftKey) {
        // Shift+Up - Add to multi-selection
        extendSelectionUp();
      } else {
        // Just move selection up
        moveKeyboardSelection('up');
      }
      return;
    }
  }
});

// Keyboard navigation helper functions
function selectFirstTaskInColumn(column) {
  clearKeyboardSelection();
  clearMultiSelection();

  currentKeyboardColumn = column;

  const columnId = column === 'action' ? 'actionColumn' : 'fyiColumn';
  const columnEl = document.getElementById(columnId);
  if (!columnEl) return;

  // Find the first todo-item that is NOT inside an accordion (meetings, documents, or task-group)
  const allTasks = Array.from(columnEl.querySelectorAll('.todo-item'));
  const firstTask = allTasks.find(task =>
    !task.closest('.meetings-accordion') &&
    !task.closest('.documents-accordion') &&
    !task.closest('.task-group')
  );
  if (!firstTask) return;

  keyboardSelectedTaskId = firstTask.dataset.todoId;
  highlightKeyboardSelectedTask();

  // Scroll task into view
  firstTask.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function moveKeyboardSelection(direction) {
  if (!keyboardSelectedTaskId || !currentKeyboardColumn) return;

  const columnId = currentKeyboardColumn === 'action' ? 'actionColumn' : 'fyiColumn';
  const columnEl = document.getElementById(columnId);
  if (!columnEl) return;

  // Get only tasks that are NOT inside accordions
  const tasks = Array.from(columnEl.querySelectorAll('.todo-item')).filter(task =>
    !task.closest('.meetings-accordion') &&
    !task.closest('.documents-accordion') &&
    !task.closest('.task-group')
  );
  const currentIndex = tasks.findIndex(t => t.dataset.todoId === keyboardSelectedTaskId);

  if (currentIndex === -1) return;

  let newIndex;
  if (direction === 'down') {
    newIndex = currentIndex < tasks.length - 1 ? currentIndex + 1 : currentIndex;
  } else {
    newIndex = currentIndex > 0 ? currentIndex - 1 : 0;
  }

  if (newIndex !== currentIndex) {
    clearKeyboardSelection();
    keyboardSelectedTaskId = tasks[newIndex].dataset.todoId;
    highlightKeyboardSelectedTask();
    tasks[newIndex].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function extendSelectionDown() {
  if (!keyboardSelectedTaskId || !currentKeyboardColumn) return;

  const columnId = currentKeyboardColumn === 'action' ? 'actionColumn' : 'fyiColumn';
  const columnEl = document.getElementById(columnId);
  if (!columnEl) return;

  const tasks = Array.from(columnEl.querySelectorAll('.todo-item')).filter(task =>
    !task.closest('.meetings-accordion') &&
    !task.closest('.documents-accordion') &&
    !task.closest('.task-group')
  );
  const currentIndex = tasks.findIndex(t => t.dataset.todoId === keyboardSelectedTaskId);

  if (currentIndex === -1 || currentIndex >= tasks.length - 1) return;

  // Add current to multi-selection if not already
  if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
    isMultiSelectMode = true;
    toggleTodoSelection(keyboardSelectedTaskId);
  }

  // Move to next and add it too
  const nextTask = tasks[currentIndex + 1];
  keyboardSelectedTaskId = nextTask.dataset.todoId;

  if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
    toggleTodoSelection(keyboardSelectedTaskId);
  }

  highlightKeyboardSelectedTask();
  nextTask.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function extendSelectionUp() {
  if (!keyboardSelectedTaskId || !currentKeyboardColumn) return;

  const columnId = currentKeyboardColumn === 'action' ? 'actionColumn' : 'fyiColumn';
  const columnEl = document.getElementById(columnId);
  if (!columnEl) return;

  const tasks = Array.from(columnEl.querySelectorAll('.todo-item')).filter(task =>
    !task.closest('.meetings-accordion') &&
    !task.closest('.documents-accordion') &&
    !task.closest('.task-group')
  );
  const currentIndex = tasks.findIndex(t => t.dataset.todoId === keyboardSelectedTaskId);

  if (currentIndex === -1 || currentIndex <= 0) return;

  // Add current to multi-selection if not already
  if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
    isMultiSelectMode = true;
    toggleTodoSelection(keyboardSelectedTaskId);
  }

  // Move to previous and add it too
  const prevTask = tasks[currentIndex - 1];
  keyboardSelectedTaskId = prevTask.dataset.todoId;

  if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
    toggleTodoSelection(keyboardSelectedTaskId);
  }

  highlightKeyboardSelectedTask();
  prevTask.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function highlightKeyboardSelectedTask() {
  // Remove previous highlight
  document.querySelectorAll('.todo-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });

  if (!keyboardSelectedTaskId) return;

  // Add highlight to current
  const task = document.querySelector(`.todo-item[data-todo-id="${keyboardSelectedTaskId}"]`);
  if (task) {
    task.classList.add('keyboard-selected');
  }
}

function clearKeyboardSelection() {
  keyboardSelectedTaskId = null;
  keyboardSelectedNoteId = null;
  document.querySelectorAll('.todo-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });
  document.querySelectorAll('.note-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });
}

// Note selection functions
function selectFirstNoteInColumn() {
  clearKeyboardSelection();
  clearMultiSelection();

  currentKeyboardColumn = 'notes';

  // Make sure Notes tab is active
  const notesTab = document.getElementById('notesTab');
  if (notesTab && !notesTab.classList.contains('active')) {
    notesTab.click();
  }

  const notesContent = document.getElementById('notesContent');
  if (!notesContent) return;

  const firstNote = notesContent.querySelector('.note-item');
  if (!firstNote) return;

  keyboardSelectedNoteId = firstNote.dataset.noteId;
  highlightKeyboardSelectedNote();

  firstNote.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function moveNoteSelection(direction) {
  if (!keyboardSelectedNoteId || currentKeyboardColumn !== 'notes') return;

  const notesContent = document.getElementById('notesContent');
  if (!notesContent) return;

  const notes = Array.from(notesContent.querySelectorAll('.note-item'));
  const currentIndex = notes.findIndex(n => n.dataset.noteId === keyboardSelectedNoteId);

  if (currentIndex === -1) return;

  let newIndex;
  if (direction === 'down') {
    newIndex = currentIndex < notes.length - 1 ? currentIndex + 1 : currentIndex;
  } else {
    newIndex = currentIndex > 0 ? currentIndex - 1 : 0;
  }

  if (newIndex !== currentIndex) {
    keyboardSelectedNoteId = notes[newIndex].dataset.noteId;
    highlightKeyboardSelectedNote();
    notes[newIndex].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function highlightKeyboardSelectedNote() {
  // Remove previous highlight from notes
  document.querySelectorAll('.note-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });

  if (!keyboardSelectedNoteId) return;

  // Add highlight to current note
  const note = document.querySelector(`.note-item[data-note-id="${keyboardSelectedNoteId}"]`);
  if (note) {
    note.classList.add('keyboard-selected');
  }
}

async function markKeyboardSelectedTaskDone() {
  if (!keyboardSelectedTaskId) return;

  const taskId = keyboardSelectedTaskId;
  const task = document.querySelector(`.todo-item[data-todo-id="${taskId}"]`);
  if (!task) return;

  const checkbox = task.querySelector('.todo-checkbox');
  if (checkbox) {
    checkbox.classList.add('checked');
    checkbox.innerHTML = '✓';
  }
  task.classList.add('completing');

  // Move selection to next task immediately
  moveKeyboardSelection('down');

  // Remove from DOM after animation
  setTimeout(() => {
    task.remove();
  }, 400);

  // Update local arrays immediately (optimistic update)
  allTodos = allTodos.filter(t => t.id != taskId);
  allFyiItems = allFyiItems.filter(t => t.id != taskId);

  // Reset read state so if task comes back, it will be unread (yellow)
  markTaskAsUnread(taskId);

  // Update counts
  if (typeof updateTabCounts === 'function') {
    updateTabCounts();
  }

  // Send to backend in background
  if (typeof window.updateTodoField === 'function') {
    window.updateTodoField(taskId, 'status', 1).catch(err => {
      console.error('Error updating task status:', err);
    });
  }
}

document.addEventListener('keyup', (e) => {
  if (e.key === 'Meta' || e.key === 'Control') {
    isCommandKeyPressed = false;
  }
});

// Clear selection when clicking outside
document.addEventListener('click', (e) => {
  // Don't clear selection if clicking on:
  // - A todo item
  // - The bulk action button
  // - The transcript slider overlay or its contents
  // - The attachment preview modal
  // - While holding Command/Ctrl key
  if (!e.target.closest('.todo-item') &&
    !e.target.closest('.bulk-action-btn') &&
    !e.target.closest('.transcript-slider-overlay') &&
    !e.target.closest('.attachment-preview-modal') &&
    !e.target.closest('.due-by-menu') &&
    !isCommandKeyPressed) {
    clearMultiSelection();
  }
});

async function initAuth() {
  try {
    const result = await chrome.storage.local.get([STORAGE_KEY]);
    if (result[STORAGE_KEY] && result[STORAGE_KEY].userId) {
      userData = result[STORAGE_KEY];
      isAuthenticated = true;
      // Sync to shared component state
      if (window.Oracle && window.Oracle.state) {
        window.Oracle.state.isAuthenticated = true;
        window.Oracle.state.userData = userData;
      }
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error initializing auth:', error);
    return false;
  }
}

async function login(email, password) {
  try {
    const response = await fetch(AUTH_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'login', email, password, timestamp: new Date().toISOString() })
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    let data;
    try {
      data = await response.json();
    } catch {
      const text = await response.text();
      data = !isNaN(text) ? parseInt(text.trim()) : null;
    }
    let userId = typeof data === 'number' ? data : typeof data === 'string' && !isNaN(data) ? parseInt(data) : Array.isArray(data) && data[0] ? data[0].id || data[0].user_id || data[0] : data?.id || data?.user_id || data?.userId;
    if (!userId) throw new Error('No user ID received');
    userData = { userId, email, loginTime: new Date().toISOString() };
    await chrome.storage.local.set({ [STORAGE_KEY]: userData });
    isAuthenticated = true;
    // Sync to shared component state
    if (window.Oracle && window.Oracle.state) {
      window.Oracle.state.isAuthenticated = true;
      window.Oracle.state.userData = userData;
    }
    return { success: true, userData };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function logout() {
  await chrome.storage.local.remove([STORAGE_KEY]);
  userData = null;
  isAuthenticated = false;
  return true;
}

function createAuthenticatedPayload(basePayload) {
  if (!isAuthenticated || !userData) throw new Error('Not authenticated');
  return { ...basePayload, user_id: userData.userId, authenticated: true };
}

function showLoader(container) {
  if (!container) return;
  const loader = document.createElement('div');
  loader.className = 'loading-overlay';
  loader.innerHTML = '<div class="spinner"></div>';
  container.style.position = 'relative';
  container.appendChild(loader);
}

function hideLoader(container) {
  if (!container) return;
  const loader = container.querySelector('.loading-overlay');
  if (loader) loader.remove();
}

function showLoginScreen() {
  const container = document.querySelector('.container');
  if (container) container.style.display = 'none';
  document.body.insertAdjacentHTML('beforeend', `<div class="login-overlay" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); display: flex; align-items: center; justify-content: center; z-index: 10000;"><div class="login-container" style="background: rgba(255,255,255,0.95); backdrop-filter: blur(20px); border-radius: 20px; padding: 40px; box-shadow: 0 20px 40px rgba(0,0,0,0.2); max-width: 450px; width: 90%;"><div style="text-align: center; margin-bottom: 30px;"><div style="width: 60px; height: 60px; margin: 0 auto 15px; background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 15px; display: flex; align-items: center; justify-content: center; font-size: 28px; color: white; font-weight: bold;">∞</div><h2 style="margin: 0 0 8px 0; color: #2c3e50; font-size: 24px; font-weight: 600;">Welcome to Oracle</h2><p style="margin: 0; color: #7f8c8d; font-size: 14px;">Sign in to access your tasks and AI assistant</p></div><form id="oracleLoginForm" style="display: flex; flex-direction: column; gap: 20px;"><input type="email" id="loginEmail" required placeholder="Email Address" style="width: 100%; padding: 16px; border: 2px solid rgba(225,232,237,0.8); border-radius: 12px; font-size: 14px; box-sizing: border-box;"><input type="password" id="loginPassword" required placeholder="Password" style="width: 100%; padding: 16px; border: 2px solid rgba(225,232,237,0.8); border-radius: 12px; font-size: 14px; box-sizing: border-box;"><button type="submit" id="loginSubmitBtn" style="width: 100%; background: linear-gradient(45deg, #667eea, #764ba2); color: white; border: none; padding: 16px; border-radius: 12px; font-size: 16px; font-weight: 600; cursor: pointer;">Sign In</button><div id="loginError" style="display: none; background: rgba(231,76,60,0.1); border: 1px solid rgba(231,76,60,0.3); color: #e74c3c; padding: 12px; border-radius: 8px; font-size: 14px; text-align: center;"></div></form></div></div>`);
  document.getElementById('oracleLoginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const submitBtn = document.getElementById('loginSubmitBtn');
    const errorDiv = document.getElementById('loginError');
    submitBtn.disabled = true;
    submitBtn.textContent = 'Signing in...';
    errorDiv.style.display = 'none';
    const result = await login(document.getElementById('loginEmail').value.trim(), document.getElementById('loginPassword').value.trim());
    if (result.success) {
      document.querySelector('.login-overlay').remove();
      document.querySelector('.container').style.display = 'flex';
      initializeMainInterface();
      setupProfileOverlay();
    } else {
      errorDiv.textContent = result.error;
      errorDiv.style.display = 'block';
    }
    submitBtn.disabled = false;
    submitBtn.textContent = 'Sign In';
  });
}

function initializeMainInterface() {
  const headerRight = document.querySelector('.header-right');
  if (headerRight && !document.getElementById('logoutBtn')) {
    const logoutBtn = document.createElement('button');
    logoutBtn.id = 'logoutBtn';
    logoutBtn.innerHTML = '⎋';
    logoutBtn.title = 'Logout';
    logoutBtn.style.cssText = 'width: 36px; height: 36px; background: rgba(231,76,60,0.1); border: 1px solid rgba(231,76,60,0.3); color: #e74c3c; border-radius: 50%; cursor: pointer; font-size: 14px; display: flex; align-items: center; justify-content: center; transition: all 0.3s;';
    logoutBtn.onmouseover = () => { logoutBtn.style.background = '#e74c3c'; logoutBtn.style.color = 'white'; };
    logoutBtn.onmouseout = () => { logoutBtn.style.background = 'rgba(231,76,60,0.1)'; logoutBtn.style.color = '#e74c3c'; };
    logoutBtn.onclick = async () => { await logout(); location.reload() };
    // Insert as first child so it appears leftmost
    headerRight.insertBefore(logoutBtn, headerRight.firstChild);
  }

  const todoTab = document.getElementById('todoTab');
  const fyiTab = document.getElementById('fyiTab');
  const scratchpadTab = document.getElementById('scratchpadTab');
  const notesTab = document.getElementById('notesTab');
  const bookmarkTab = document.getElementById('bookmarkTab');
  const documentsTab = document.getElementById('documentsTab');
  const helpTab = document.getElementById('helpTab');

  // For 3-column layout - content containers
  const notesContent = document.getElementById('notesContent');
  const scratchpadContent = document.getElementById('scratchpadContent');
  const bookmarkContent = document.getElementById('bookmarkContent');
  const documentsContent = document.getElementById('documentsContent');

  // Legacy single-column content references (for backwards compatibility)
  const todoContent = document.getElementById('todoContent');
  const fyiContent = document.getElementById('fyiContent');
  const helpContent = document.getElementById('helpContent');

  const learnRadio = document.getElementById('learnRadio');
  const answerRadio = document.getElementById('answerRadio');
  const learnSection = document.getElementById('learnSection');
  const answerSection = document.getElementById('answerSection');
  const learnInput = document.getElementById('learnInput');
  const answerInput = document.getElementById('answerInput');
  const askBtn = document.getElementById('askBtn');
  const responseSection = document.getElementById('responseSection');
  const responseContent = document.getElementById('responseContent');
  const copyBtn = document.getElementById('copyBtn');

  // Check if we're in 3-column layout
  const isThreeColumnLayout = document.querySelector('.three-column-layout') !== null;

  // Col3 tab switching for 3-column layout
  function setupCol3Tabs() {
    const col3Tabs = document.querySelectorAll('.col3-tab');
    const col3TabSlider = document.getElementById('col3TabSlider');
    const addNoteBtn = document.getElementById('addNoteBtn');
    const refreshCol3Btn = document.getElementById('refreshCol3Btn');

    col3Tabs.forEach((tab, index) => {
      tab.addEventListener('click', () => {
        // Update active tab
        col3Tabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        // Move slider - 4 tabs now
        if (col3TabSlider) {
          col3TabSlider.style.width = `calc(25% - 2px)`;
          col3TabSlider.style.left = `calc(${index * 25}% + 3px)`;
        }

        // Switch content
        const targetId = tab.dataset.col3;
        document.querySelectorAll('.col3-content').forEach(content => {
          content.classList.remove('active');
        });

        if (targetId === 'notes') {
          notesContent?.classList.add('active');
          if (addNoteBtn) addNoteBtn.style.display = 'inline-flex';
          if (addNoteBtn) addNoteBtn.title = 'New Note';
          if (addNoteBtn) addNoteBtn.textContent = '+';
          loadNotes();
        } else if (targetId === 'bookmarks') {
          bookmarkContent?.classList.add('active');
          if (addNoteBtn) addNoteBtn.style.display = 'none';
          loadBookmarks();
        } else if (targetId === 'documents') {
          documentsContent?.classList.add('active');
          if (addNoteBtn) addNoteBtn.style.display = 'none';
          loadDocuments();
        } else if (targetId === 'alltasks') {
          const allTasksContent = document.getElementById('allTasksContent');
          allTasksContent?.classList.add('active');
          // Show search button in place of + button
          if (addNoteBtn) {
            addNoteBtn.style.display = 'inline-flex';
            addNoteBtn.title = 'Search Tasks';
            addNoteBtn.textContent = '🔍';
          }
          loadAllTasks();
        }
      });
    });

    // Refresh button for col3 - refreshes current active content
    refreshCol3Btn?.addEventListener('click', () => {
      const activeTab = document.querySelector('.col3-tab.active');
      const targetId = activeTab?.dataset?.col3;
      if (targetId === 'notes') loadNotes();
      else if (targetId === 'bookmarks') loadBookmarks();
      else if (targetId === 'documents') loadDocuments();
      else if (targetId === 'alltasks') loadAllTasks();
    });
  }

  if (isThreeColumnLayout) {
    setupCol3Tabs();
  }

  // Column refresh buttons for 3-column layout
  document.getElementById('refreshActionBtn')?.addEventListener('click', () => {
    clearPendingUpdates('action');
    loadTodos('starred');
  });
  document.getElementById('refreshFyiBtn')?.addEventListener('click', () => {
    clearPendingUpdates('fyi');
    loadFYI();
  });

  // Legacy tab behavior for single-column layout
  todoTab?.addEventListener('click', () => {
    if (!isThreeColumnLayout) {
      todoTab.classList.add('active');
      fyiTab?.classList.remove('active');
      scratchpadTab?.classList.remove('active');
      bookmarkTab?.classList.remove('active');
      documentsTab?.classList.remove('active');
      helpTab?.classList.remove('active');
      todoContent?.classList.add('active');
      fyiContent?.classList.remove('active');
      scratchpadContent?.classList.remove('active');
      bookmarkContent?.classList.remove('active');
      documentsContent?.classList.remove('active');
      helpContent?.classList.remove('active');
    }
    loadTodos('starred');
  });

  fyiTab?.addEventListener('click', () => {
    if (!isThreeColumnLayout) {
      fyiTab.classList.add('active');
      todoTab?.classList.remove('active');
      scratchpadTab?.classList.remove('active');
      bookmarkTab?.classList.remove('active');
      documentsTab?.classList.remove('active');
      helpTab?.classList.remove('active');
      fyiContent?.classList.add('active');
      todoContent?.classList.remove('active');
      scratchpadContent?.classList.remove('active');
      bookmarkContent?.classList.remove('active');
      documentsContent?.classList.remove('active');
      helpContent?.classList.remove('active');
    }
    loadFYI();
  });

  scratchpadTab?.addEventListener('click', () => {
    if (!isThreeColumnLayout) {
      scratchpadTab.classList.add('active');
      todoTab?.classList.remove('active');
      fyiTab?.classList.remove('active');
      bookmarkTab?.classList.remove('active');
      documentsTab?.classList.remove('active');
      helpTab?.classList.remove('active');
      scratchpadContent?.classList.add('active');
      todoContent?.classList.remove('active');
      fyiContent?.classList.remove('active');
      bookmarkContent?.classList.remove('active');
      documentsContent?.classList.remove('active');
      helpContent?.classList.remove('active');
    }
    hideNoteForm();
    loadNotes();
  });

  bookmarkTab?.addEventListener('click', () => {
    if (!isThreeColumnLayout) {
      bookmarkTab.classList.add('active');
      todoTab?.classList.remove('active');
      fyiTab?.classList.remove('active');
      scratchpadTab?.classList.remove('active');
      documentsTab?.classList.remove('active');
      helpTab?.classList.remove('active');
      bookmarkContent?.classList.add('active');
      todoContent?.classList.remove('active');
      fyiContent?.classList.remove('active');
      scratchpadContent?.classList.remove('active');
      documentsContent?.classList.remove('active');
      helpContent?.classList.remove('active');
    }
    loadBookmarks();
  });

  documentsTab?.addEventListener('click', () => {
    if (!isThreeColumnLayout) {
      documentsTab.classList.add('active');
      todoTab?.classList.remove('active');
      fyiTab?.classList.remove('active');
      scratchpadTab?.classList.remove('active');
      bookmarkTab?.classList.remove('active');
      helpTab?.classList.remove('active');
      documentsContent?.classList.add('active');
      todoContent?.classList.remove('active');
      fyiContent?.classList.remove('active');
      scratchpadContent?.classList.remove('active');
      bookmarkContent?.classList.remove('active');
      helpContent?.classList.remove('active');
    }
    loadDocuments();
  });

  helpTab?.addEventListener('click', () => {
    helpTab.classList.add('active');
    todoTab?.classList.remove('active');
    fyiTab?.classList.remove('active');
    scratchpadTab?.classList.remove('active');
    bookmarkTab?.classList.remove('active');
    documentsTab?.classList.remove('active');
    helpContent?.classList.add('active');
    todoContent?.classList.remove('active');
    fyiContent?.classList.remove('active');
    scratchpadContent?.classList.remove('active');
    bookmarkContent?.classList.remove('active');
    documentsContent?.classList.remove('active');
  });

  const switchMode = () => {
    if (!learnSection || !answerSection) return;
    if (learnRadio?.checked) {
      learnSection.classList.add('active');
      answerSection.classList.remove('active');
    } else {
      answerSection.classList.add('active');
      learnSection.classList.remove('active');
    }
  };

  learnRadio?.addEventListener('change', switchMode);
  answerRadio?.addEventListener('change', switchMode);

  learnInput?.addEventListener('input', () => {
    const counter = learnInput.closest('.input-wrapper').querySelector('.char-counter');
    if (counter) counter.textContent = learnInput.value.length + ' / 5000';
  });

  answerInput?.addEventListener('input', () => {
    const counter = answerInput.closest('.input-wrapper').querySelector('.char-counter');
    if (counter) counter.textContent = answerInput.value.length + ' / 1000';
  });

  askBtn?.addEventListener('click', async () => {
    const mode = learnRadio?.checked ? 'learn' : 'answer';
    const input = mode === 'learn' ? learnInput.value.trim() : answerInput.value.trim();
    if (!input) { showResponse('Please enter text', false); return }
    if (!isAuthenticated) { showResponse('Please log in', false); return }
    askBtn.disabled = true;
    askBtn.textContent = mode === 'learn' ? 'Teaching...' : 'Thinking...';
    try {
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          timestamp: new Date().toISOString(),
          source: 'oracle-chrome-extension-newtab',
          action: mode,
          [mode === 'learn' ? 'contentToLearn' : 'query']: input
        }))
      });
      if (response.ok) {
        const data = await response.json();
        const responseText = (Array.isArray(data) && data[0]?.message) || data.message || data.answer || 'Success!';
        showResponse(responseText, true);
      } else {
        throw new Error('Request failed: ' + response.status);
      }
    } catch (error) {
      showResponse('Error: ' + error.message, false);
    }
    askBtn.disabled = false;
    askBtn.textContent = 'Submit';
  });

  function showResponse(message, isSuccess) {
    currentResponse = message;
    if (responseContent) {
      responseContent.textContent = message;
      responseContent.style.color = isSuccess ? '#27ae60' : '#e74c3c';
    }
    if (responseSection) responseSection.classList.add('show');
  }

  copyBtn?.addEventListener('click', async () => {
    if (currentResponse) {
      await navigator.clipboard.writeText(currentResponse);
      copyBtn.textContent = 'Copied!';
      copyBtn.classList.add('copied');
      setTimeout(() => {
        copyBtn.textContent = 'Copy';
        copyBtn.classList.remove('copied');
      }, 2000);
    }
  });

  // Scratchpad Functions — delegated to OracleNotes shared component
  const { loadNotes, showNoteForm, hideNoteForm } = window.OracleNotes;
  window.OracleNotes.setupNoteButtons();

  // Todo Functions
  async function loadTodos(filter = 'starred') {
    currentTodoFilter = filter; // Track current filter
    activeTagFilters = []; // Reset tag filters on new load
    const container = document.querySelector('.todos-container');
    const emptyState = container.querySelector('.empty-state');
    const todosList = container.querySelector('.todos-list');
    if (emptyState) emptyState.style.display = 'none';
    if (todosList) todosList.style.display = 'none';
    showLoader(container);

    // Capture initial load state BEFORE async operations to avoid race condition with loadFYI
    const wasInitialLoad = isInitialLoad;

    try {
      if (!isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'list_todos',
          filter: 'all',
          timestamp: new Date().toISOString()
        }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);

      // Handle potentially empty responses
      const responseText = await response.text();
      let data = [];
      if (responseText && responseText.trim()) {
        try {
          data = JSON.parse(responseText);
        } catch (parseError) {
          console.warn('Failed to parse todos response:', parseError);
          data = [];
        }
      }

      allTodos = Array.isArray(data) ? data : (data?.todos || data?.results || []);
      if (!allTodos) allTodos = [];
      updateLastRefreshTime(); // Update timestamp to prevent immediate Ably refresh

      // On initial load (fresh install), mark all todos as read
      // On refresh (saved state exists), just update timestamps for existing tasks
      // Use captured wasInitialLoad to avoid race condition
      if (wasInitialLoad) {
        // Fresh install - mark everything as read
        allTodos.forEach(task => {
          const idStr = String(task.id);
          readTaskIds.add(idStr);
          if (task.updated_at) {
            previousTaskTimestamps.set(idStr, task.updated_at);
          }
        });
        // Don't set isInitialLoad = false here - let loadFYI do it after both are done
        // But DO save state so tasks are marked as read
        saveReadState();
        console.log('📋 Todos initial load (fresh install): marked', allTodos.length, 'tasks as read');
      } else {
        // Refresh with existing state - check for genuinely new or updated tasks
        allTodos.forEach(task => {
          const idStr = String(task.id);
          // Check if this is a genuinely new task we've never seen
          const isNewTask = !previousTaskTimestamps.has(idStr);

          if (isNewTask) {
            // New task - DON'T add to readTaskIds, let it appear unread
            console.log(`🆕 New task detected: ${task.id}`);
          }

          // Check for updates on existing tasks
          if (task.updated_at) {
            const prevTimestamp = previousTaskTimestamps.get(idStr);
            if (prevTimestamp && prevTimestamp !== task.updated_at) {
              // Task was updated - mark as unread
              readTaskIds.delete(idStr);
              console.log(`📝 Task ${task.id} updated, marking unread`);
            }
            previousTaskTimestamps.set(idStr, task.updated_at);
          }
        });
        saveReadState();
        console.log('📋 Todos refresh: processed', allTodos.length, 'tasks');
      }

      hideLoader(container);
      if (allTodos.length > 0) displayTodos(allTodos, filter);
      else showEmptyTodosState(filter);
      updateBadge();
      updateTabCounts();
    } catch (error) {
      console.error('Error loading todos:', error);
      hideLoader(container);
      if (emptyState) {
        emptyState.style.display = 'flex';
        emptyState.innerHTML = '<h3>Error: ' + error.message + '</h3>';
      }
    }
  }

  // Animated version of loadTodos - no loader, just smooth transition
  async function loadTodosAnimated(filter = 'starred') {
    currentTodoFilter = filter;
    activeTagFilters = [];
    try {
      if (!isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'list_todos',
          filter: 'all',
          timestamp: new Date().toISOString()
        }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);

      // Handle potentially empty responses
      const responseText = await response.text();
      let data = [];
      if (responseText && responseText.trim()) {
        try {
          data = JSON.parse(responseText);
        } catch (parseError) {
          console.warn('Failed to parse todos response:', parseError);
          data = [];
        }
      }

      const newTodos = Array.isArray(data) ? data : (data?.todos || data?.results || []);
      if (!newTodos || !Array.isArray(newTodos)) { console.warn('Invalid todos data'); return; }

      console.log(`🔄 loadTodosAnimated: Got ${newTodos.length} todos, previously had ${allTodos.length}`);
      console.log(`📊 State: readTaskIds.size=${readTaskIds.size}, previousTaskTimestamps.size=${previousTaskTimestamps.size}, isInitialLoad=${isInitialLoad}`);

      // Find new items that weren't in the old list
      const oldIds = new Set(allTodos.map(t => t.id));
      const newItems = newTodos.filter(t => !oldIds.has(t.id));

      // Find updated items (existing tasks with changed updated_at timestamp)
      const updatedItems = newTodos.filter(t => {
        const idStr = String(t.id);
        if (!oldIds.has(t.id)) return false; // Skip new items
        const prevTimestamp = previousTaskTimestamps.get(idStr);
        const isUpdated = prevTimestamp && t.updated_at && prevTimestamp !== t.updated_at;
        if (isUpdated) {
          console.log(`📝 Task ${t.id} timestamp changed: ${prevTimestamp} → ${t.updated_at}`);
        }
        return isUpdated;
      });

      console.log(`✨ New items: ${newItems.length}, Updated items: ${updatedItems.length}`);

      // Mark updated tasks as unread by removing them from readTaskIds
      updatedItems.forEach(t => {
        const idStr = String(t.id);
        readTaskIds.delete(idStr);
        console.log(`📝 Task ${t.id} marked as unread (updated_at changed)`);
      });

      // Update timestamps for all tasks
      newTodos.forEach(t => {
        if (t.updated_at) {
          previousTaskTimestamps.set(String(t.id), t.updated_at);
        }
      });

      // Save state after processing
      if (updatedItems.length > 0 || newItems.length > 0) {
        saveReadState();
      }

      allTodos = newTodos;
      updateLastRefreshTime(); // Update timestamp to prevent immediate Ably refresh

      // Combine new and updated items for highlighting
      const itemsToHighlight = [...newItems.map(t => t.id), ...updatedItems.map(t => t.id)];

      if (allTodos.length > 0) {
        displayTodosAnimated(allTodos, filter, itemsToHighlight);
      } else {
        showEmptyTodosState(filter);
      }
      updateBadge();
      updateTabCounts();
    } catch (error) {
      console.error('Error loading todos:', error);
    }
  }

  // Display todos with animation for new items
  function displayTodosAnimated(todos, filter, newItemIds = []) {
    // Just call displayTodos but mark new items for animation
    displayTodos(todos, filter, newItemIds);
  }

  // FYI Functions - loads items with starred=0 and status=0
  async function loadFYI() {
    const container = document.querySelector('.fyi-container');
    const emptyState = container.querySelector('.empty-state');
    let fyiList = container.querySelector('.fyi-list');
    if (emptyState) emptyState.style.display = 'none';
    if (fyiList) fyiList.style.display = 'none';
    showLoader(container);

    // Capture initial load state BEFORE async operations to avoid race condition with loadTodos
    const wasInitialLoad = isInitialLoad;

    try {
      if (!isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'list_todos',
          filter: 'all',
          starred: 0,
          status: 0,
          timestamp: new Date().toISOString()
        }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);

      // Handle potentially empty responses
      const responseText = await response.text();
      let data = [];
      if (responseText && responseText.trim()) {
        try {
          data = JSON.parse(responseText);
        } catch (parseError) {
          console.warn('Failed to parse FYI response:', parseError);
          data = [];
        }
      }

      const allItems = Array.isArray(data) ? data : (data?.todos || []);
      // Filter for FYI items: starred=0 and status=0, excluding meeting links
      allFyiItems = allItems.filter(t => t.starred === 0 && t.status === 0 && !isMeetingLink(t.message_link));
      updateLastRefreshTime(); // Update timestamp to prevent immediate Ably refresh

      // On initial load (fresh install), mark all FYI items as read
      // On refresh (saved state exists), handle new/updated tasks
      // Use captured wasInitialLoad to avoid race condition
      if (wasInitialLoad) {
        // Fresh install - mark everything as read
        allFyiItems.forEach(task => {
          const idStr = String(task.id);
          readTaskIds.add(idStr);
          if (task.updated_at) {
            previousTaskTimestamps.set(idStr, task.updated_at);
          }
        });
        isInitialLoad = false; // Now safe to set this
        window.Oracle.state.isInitialLoad = false;
        saveReadState(); // Save the initial state
        console.log('✅ Initial load complete (fresh install). Total read tasks:', readTaskIds.size);
      } else {
        // Refresh with existing state - check for new/updated tasks
        allFyiItems.forEach(task => {
          const idStr = String(task.id);
          if (!readTaskIds.has(idStr) && !previousTaskTimestamps.has(idStr)) {
            // Genuinely new task we've never seen - leave as unread
            console.log(`🆕 New FYI task detected on refresh: ${task.id}`);
          }
          // Check for updates
          if (task.updated_at) {
            const prevTimestamp = previousTaskTimestamps.get(idStr);
            if (prevTimestamp && prevTimestamp !== task.updated_at) {
              readTaskIds.delete(idStr);
              console.log(`📝 FYI Task ${task.id} updated, marking unread`);
            }
            previousTaskTimestamps.set(idStr, task.updated_at);
          }
        });
        saveReadState();
        console.log('📋 FYI refresh: processed', allFyiItems.length, 'items');
      }

      hideLoader(container);
      if (allFyiItems.length > 0) {
        displayFYI(allFyiItems, container);
      } else {
        showEmptyFYIState(container);
      }
      updateTabCounts();
    } catch (error) {
      console.error('Error loading FYI items:', error);
      hideLoader(container);
      if (emptyState) {
        emptyState.style.display = 'flex';
        emptyState.innerHTML = '<h3>Error: ' + error.message + '</h3>';
      }
    }
  }

  // Animated version of loadFYI - no loader, smooth transition
  async function loadFYIAnimated() {
    const container = document.querySelector('.fyi-container');
    try {
      if (!isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'list_todos',
          filter: 'all',
          starred: 0,
          status: 0,
          timestamp: new Date().toISOString()
        }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);

      // Handle potentially empty responses
      const responseText = await response.text();
      let data = [];
      if (responseText && responseText.trim()) {
        try {
          data = JSON.parse(responseText);
        } catch (parseError) {
          console.warn('Failed to parse FYI animated response:', parseError);
          data = [];
        }
      }

      const allItems = Array.isArray(data) ? data : (data?.todos || []);

      // Find new items
      const oldIds = new Set(allFyiItems.map(t => t.id));
      const newFyiItems = allItems.filter(t => t.starred === 0 && t.status === 0 && !isMeetingLink(t.message_link));
      const newItemIds = newFyiItems.filter(t => !oldIds.has(t.id)).map(t => t.id);

      // Find updated items (existing tasks with changed updated_at timestamp)
      const updatedItems = newFyiItems.filter(t => {
        const idStr = String(t.id);
        if (!oldIds.has(t.id)) return false; // Skip new items
        const prevTimestamp = previousTaskTimestamps.get(idStr);
        return prevTimestamp && t.updated_at && prevTimestamp !== t.updated_at;
      });

      // Mark updated tasks as unread by removing them from readTaskIds
      updatedItems.forEach(t => {
        const idStr = String(t.id);
        readTaskIds.delete(idStr);
        console.log(`📝 FYI Task ${t.id} marked as unread (updated_at changed)`);
      });

      // Update timestamps for all FYI tasks
      newFyiItems.forEach(t => {
        if (t.updated_at) {
          previousTaskTimestamps.set(String(t.id), t.updated_at);
        }
      });

      // Save state after processing
      if (updatedItems.length > 0 || newItemIds.length > 0) {
        saveReadState();
      }

      allFyiItems = newFyiItems;
      updateLastRefreshTime(); // Update timestamp to prevent immediate Ably refresh

      // Combine new and updated items for highlighting
      const itemsToHighlight = [...newItemIds, ...updatedItems.map(t => t.id)];

      if (allFyiItems.length > 0) {
        displayFYI(allFyiItems, container, itemsToHighlight);
      } else {
        showEmptyFYIState(container);
      }
      updateTabCounts();
    } catch (error) {
      console.error('Error loading FYI items:', error);
    }
  }

  function displayFYI(items, container, newItemIds = []) {
    let fyiList = container.querySelector('.fyi-list');
    const emptyState = container.querySelector('.empty-state');

    if (!fyiList) {
      fyiList = document.createElement('div');
      fyiList.className = 'fyi-list todos-list';
      container.appendChild(fyiList);
    }

    if (emptyState) emptyState.style.display = 'none';
    fyiList.style.display = 'flex';

    // Separate drive items from non-drive items
    const driveItems = items.filter(item => isDriveLink(item.message_link));
    const nonDriveItems = items.filter(item => !isDriveLink(item.message_link));

    // Build documents accordion (excluding files already in Action tab)
    const documentsAccordionHtml = buildDocumentsAccordion(driveItems, actionTabDriveFileIds);

    // Filter out tasks that are already shown in Action tab
    const nonDriveItemsForFyi = nonDriveItems.filter(item => !actionTabSlackTaskIds.has(item.id));

    // Build Slack channels accordion (with filtered tasks)
    const slackChannelsAccordionHtml = buildSlackChannelsAccordion(nonDriveItemsForFyi);

    // Get tasks that are not grouped in Slack accordion
    const { slackGroups: fyiSlackGroups, nonSlackTasks: remainingNonDriveItems } = groupTasksBySlackChannel(nonDriveItemsForFyi);

    // Group remaining tasks by tags
    const { tagGroups: fyiTagGroups, untaggedTasks: itemsForIndividualDisplay } = groupTasksByTag(remainingNonDriveItems);

    // Remove existing accordions and tag groups
    let existingDocsAccordion = container.querySelector('.documents-accordion');
    if (existingDocsAccordion) existingDocsAccordion.remove();
    // Also clean from slot
    const docsSlotClean = document.getElementById('documentsAccordionSlot');
    if (docsSlotClean) {
      const existingSlotAccordion = docsSlotClean.querySelector('.documents-accordion');
      if (existingSlotAccordion) existingSlotAccordion.remove();
    }

    let existingSlackAccordion = container.querySelector('.slack-channels-accordion');
    if (existingSlackAccordion) existingSlackAccordion.remove();

    // Remove existing tag groups in FYI
    container.querySelectorAll('.fyi-tag-group').forEach(el => el.remove());

    // Insert documents accordion in the slot above FYI column content
    const documentsSlot = document.getElementById('documentsAccordionSlot');
    if (documentsSlot) {
      documentsSlot.innerHTML = documentsAccordionHtml || '';
      if (documentsAccordionHtml) {
        setupDocumentsAccordion(documentsSlot);
      }
    } else {
      // Fallback: insert inside container
      if (documentsAccordionHtml) {
        fyiList.insertAdjacentHTML('beforebegin', documentsAccordionHtml);
        setupDocumentsAccordion(container);
      }
    }

    // Insert Slack channels accordion after documents accordion (if any) and before fyi list
    if (slackChannelsAccordionHtml) {
      fyiList.insertAdjacentHTML('beforebegin', slackChannelsAccordionHtml);
      setupSlackChannelsAccordion(container);
    }

    // Build and insert tag group HTML for FYI
    let tagGroupsHtml = '';
    Object.values(fyiTagGroups).forEach(group => {
      tagGroupsHtml += buildFyiTagGroupHtml(group, newItemIds);
    });

    if (tagGroupsHtml) {
      fyiList.insertAdjacentHTML('beforebegin', tagGroupsHtml);
      setupFyiTagGroups(container);
    }

    // Use same sorting as Action tab for remaining items
    const sortedItems = sortTodos(itemsForIndividualDisplay);

    fyiList.innerHTML = `
      ${sortedItems.map((item, index) => {
      const messageLink = item.message_link || '';
      const secondaryLinks = item.secondary_links || [];
      const isNewItem = newItemIds.includes(item.id);

      // Use image files for icons
      const slackIconUrl = chrome.runtime.getURL('icon-slack.png');
      const gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
      const driveIconUrl = chrome.runtime.getURL('icon-drive.png');
      const freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
      const freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
      const freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
      const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
      const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
      const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
      const zoomIconUrl = chrome.runtime.getURL('icon-zoom.png');
      const googleMeetIconUrl = chrome.runtime.getURL('icon-google-meet.png');
      const googleCalendarIconUrl = chrome.runtime.getURL('icon-google-calendar.png');

      const linkIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

      let sourceIcon = linkIcon;
      let sourceTitle = 'View source';
      let isImageIcon = false;
      let iconUrl = '';

      if (messageLink.includes('zoom.us') || messageLink.includes('zoom.com')) {
        iconUrl = zoomIconUrl;
        sourceTitle = 'Open in Zoom';
        isImageIcon = true;
      } else if (messageLink.includes('meet.google.com')) {
        iconUrl = googleMeetIconUrl;
        sourceTitle = 'Open in Google Meet';
        isImageIcon = true;
      } else if (messageLink.includes('calendar.google.com')) {
        iconUrl = googleCalendarIconUrl;
        sourceTitle = 'Open in Google Calendar';
        isImageIcon = true;
      } else if (messageLink.includes('mail.google.com')) {
        iconUrl = gmailIconUrl;
        sourceTitle = 'Open in Gmail';
        isImageIcon = true;
      } else if (messageLink.includes('slack.com') || messageLink.includes('app.slack.com')) {
        iconUrl = slackIconUrl;
        sourceTitle = 'Open in Slack';
        isImageIcon = true;
      } else if (messageLink.includes('freshdesk.com')) {
        iconUrl = freshdeskIconUrl;
        sourceTitle = 'Open in Freshdesk';
        isImageIcon = true;
      } else if (messageLink.includes('freshrelease.com')) {
        iconUrl = freshreleaseIconUrl;
        sourceTitle = 'Open in Freshrelease';
        isImageIcon = true;
      } else if (messageLink.includes('freshservice.com')) {
        iconUrl = freshserviceIconUrl;
        sourceTitle = 'Open in Freshservice';
        isImageIcon = true;
      } else if (messageLink.includes('docs.google.com/document')) {
        iconUrl = googleDocsIconUrl;
        sourceTitle = 'Open in Google Docs';
        isImageIcon = true;
      } else if (messageLink.includes('docs.google.com/spreadsheets') || messageLink.includes('sheets.google.com')) {
        iconUrl = googleSheetsIconUrl;
        sourceTitle = 'Open in Google Sheets';
        isImageIcon = true;
      } else if (messageLink.includes('docs.google.com/presentation') || messageLink.includes('slides.google.com')) {
        iconUrl = googleSlidesIconUrl;
        sourceTitle = 'Open in Google Slides';
        isImageIcon = true;
      } else if (messageLink.includes('drive.google.com')) {
        iconUrl = driveIconUrl;
        sourceTitle = 'Open in Google Drive';
        isImageIcon = true;
      }

      const sourceIconHtml = isImageIcon
        ? '<img src="' + iconUrl + '" alt="' + sourceTitle + '" style="width: 14px; height: 14px; object-fit: contain;">'
        : sourceIcon;

      // Generate secondary links HTML
      const secondaryLinksHtml = secondaryLinks.filter(l => l != null).map(link => {
        let secondaryIconHtml = linkIcon;
        let secondaryTitle = 'View link';

        if (link.includes('zoom.us') || link.includes('zoom.com')) {
          secondaryIconHtml = '<img src="' + zoomIconUrl + '" alt="Zoom" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Zoom';
        } else if (link.includes('meet.google.com')) {
          secondaryIconHtml = '<img src="' + googleMeetIconUrl + '" alt="Google Meet" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Google Meet';
        } else if (link.includes('calendar.google.com')) {
          secondaryIconHtml = '<img src="' + googleCalendarIconUrl + '" alt="Google Calendar" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Google Calendar';
        } else if (link.includes('freshdesk.com')) {
          secondaryIconHtml = '<img src="' + freshdeskIconUrl + '" alt="Freshdesk" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Freshdesk';
        } else if (link.includes('freshrelease.com')) {
          secondaryIconHtml = '<img src="' + freshreleaseIconUrl + '" alt="Freshrelease" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Freshrelease';
        } else if (link.includes('freshservice.com')) {
          secondaryIconHtml = '<img src="' + freshserviceIconUrl + '" alt="Freshservice" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Freshservice';
        } else if (link.includes('docs.google.com/document')) {
          secondaryIconHtml = '<img src="' + googleDocsIconUrl + '" alt="Google Docs" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Google Docs';
        } else if (link.includes('docs.google.com/spreadsheets') || link.includes('sheets.google.com')) {
          secondaryIconHtml = '<img src="' + googleSheetsIconUrl + '" alt="Google Sheets" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Google Sheets';
        } else if (link.includes('docs.google.com/presentation') || link.includes('slides.google.com')) {
          secondaryIconHtml = '<img src="' + googleSlidesIconUrl + '" alt="Google Slides" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Google Slides';
        } else if (link.includes('drive.google.com')) {
          secondaryIconHtml = '<img src="' + driveIconUrl + '" alt="Google Drive" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Google Drive';
        } else if (link.includes('mail.google.com')) {
          secondaryIconHtml = '<img src="' + gmailIconUrl + '" alt="Gmail" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Gmail';
        } else if (link.includes('slack.com') || link.includes('app.slack.com')) {
          secondaryIconHtml = '<img src="' + slackIconUrl + '" alt="Slack" style="width: 14px; height: 14px; object-fit: contain;">';
          secondaryTitle = 'Open in Slack';
        }

        return '<span style="color: #bdc3c7; margin: 0 4px;">|</span><a href="' + link + '" target="_blank" class="todo-source" title="' + secondaryTitle + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + secondaryIconHtml + '</a>';
      }).join('');

      const dueByHtml = item.due_by ? '<span class="todo-due ' + (new Date(item.due_by) < new Date() ? 'overdue' : '') + '">' + formatDueBy(item.due_by) + '</span>' : '';

      // Handle descriptions - always show 2 lines with View more
      const taskNameRaw = item.task_name || '';
      const taskNameEscaped = escapeHtml(taskNameRaw);
      const maxLength = 180;
      const hasTitle = !!item.task_title;
      const needsViewMore = taskNameRaw.length > maxLength || (hasTitle && taskNameRaw.length > 60);
      let taskNameHtml;

      if (needsViewMore) {
        const truncated = escapeHtml(taskNameRaw.substring(0, maxLength));
        taskNameHtml = '<span class="todo-text-content" data-full-text="' + taskNameEscaped + '">' + taskNameEscaped + '</span><span class="view-more-inline" data-todo-id="' + item.id + '">View more</span>';
      } else {
        taskNameHtml = '<span class="todo-text-content">' + taskNameEscaped + '</span>';
      }

      // Render tags for each item (filter out invalid tags like "null")
      const itemTags = (item.tags || []).filter(isValidTag);
      const tagsHtml = itemTags.length > 0 ? `
          <div class="todo-tags" style="display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px;">
            ${itemTags.map(tag => '<span class="todo-tag" data-tag="' + escapeHtml(tag) + '" style="display: inline-block; background: rgba(102, 126, 234, 0.1); color: #667eea; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 500; cursor: pointer; border: 1px solid rgba(102, 126, 234, 0.3); transition: all 0.2s;">' + escapeHtml(tag) + '</span>').join('')}
          </div>
        ` : '';

      // Participant text (e.g., channel name)
      const participantHtml = item.participant_text ? '<div class="todo-participant">' + escapeHtml(item.participant_text) + '</div>' : '';

      // Check if task is unread
      const isUnread = isTaskUnread(item.id, item.updated_at);

      return `<div class="todo-item fyi-item ${isNewItem ? 'appearing' : ''} ${isUnread ? 'unread' : ''}" style="animation-delay: ${index * 0.05}s" data-todo-id="${item.id}">
          <div class="todo-left-actions">
            <div class="todo-checkbox" data-todo-id="${item.id}"></div>
          </div>
          <div class="todo-content">
            ${item.task_title ? '<div class="todo-title">' + escapeHtml(item.task_title) + '</div>' : ''}
            <div class="todo-text${needsViewMore ? ' truncated' : ''}">${taskNameHtml}</div>
            ${participantHtml}
            <div class="todo-meta">
              <span class="todo-date">${formatDate(item.updated_at || item.created_at)}</span>
              ${dueByHtml}
              ${messageLink ? '<a href="' + messageLink + '" target="_blank" class="todo-source" title="' + sourceTitle + '" style="padding: 6px;">' + sourceIconHtml + '</a>' : ''}
              ${secondaryLinksHtml}
            </div>
            ${tagsHtml}
          </div>
          <div class="todo-actions">
            <div class="todo-clock" data-todo-id="${item.id}" title="Remind me in">🕐</div>
          </div>
        </div>`;
    }).join('')}
    `;

    // Add event listeners - make whole item clickable for transcript or multi-select
    fyiList.querySelectorAll('.todo-item').forEach(item => {
      const todoId = item.dataset.todoId;
      item.addEventListener('click', (e) => {
        if (e.target.closest('.todo-checkbox') ||
          e.target.closest('.todo-clock') ||
          e.target.closest('.todo-source') ||
          e.target.closest('.todo-tag') ||
          e.target.closest('.view-more-inline') ||
          e.target.closest('a')) {
          return;
        }

        // Command/Ctrl + click on the item itself triggers multi-select
        if (e.metaKey || e.ctrlKey) {
          e.preventDefault();
          isMultiSelectMode = true;
          toggleTodoSelection(todoId);
          return;
        }

        showTranscriptSlider(todoId);
      });
    });

    // Checkbox click - with multi-select support (same as Action tab)
    fyiList.querySelectorAll('.todo-checkbox').forEach(checkbox => {
      checkbox.addEventListener('click', async (e) => {
        e.stopPropagation();
        const todoId = checkbox.dataset.todoId;
        // Check if Command/Ctrl key is pressed for multi-select
        if (e.metaKey || e.ctrlKey) {
          isMultiSelectMode = true;
          toggleTodoSelection(todoId);
        } else if (selectedTodoIds.size > 0) {
          // If already in multi-select mode, continue selecting
          toggleTodoSelection(todoId);
        } else {
          // Normal single-click behavior - mark as done with immediate removal
          const todoItem = checkbox.closest('.todo-item');
          checkbox.classList.add('checked');
          checkbox.innerHTML = '✓';
          todoItem.classList.add('completing');

          // Check if parent group becomes empty
          checkEmptyParentGroup(todoItem);

          // Remove from DOM after animation
          setTimeout(() => {
            todoItem.remove();
          }, 400);

          // Update local array immediately (optimistic update)
          allFyiItems = allFyiItems.filter(t => t.id != todoId);

          // Reset read state so if task comes back, it will be unread (yellow)
          markTaskAsUnread(todoId);

          // Update counts
          if (typeof updateTabCounts === 'function') {
            updateTabCounts();
          }

          // Send to backend in background
          updateTodoField(todoId, 'status', 1).catch(err => {
            console.error('Error updating FYI status:', err);
          });
        }
      });
    });

    // Clock/Reminder click - show due by menu
    fyiList.querySelectorAll('.todo-clock').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        btn.style.transform = 'scale(0.9)';
        setTimeout(() => btn.style.transform = 'scale(1)', 150);
        showDueByMenu(btn, btn.dataset.todoId);
      });
    });

    // View more inline - opens transcript
    addFyiViewMoreListeners(fyiList);
  }

  function addFyiViewMoreListeners(fyiList) {
    fyiList.querySelectorAll('.view-more-inline:not(.view-less)').forEach(btn => {
      // Skip if already has listener
      if (btn.dataset.listenerAttached) return;
      btn.dataset.listenerAttached = 'true';

      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const todoText = btn.closest('.todo-text');
        if (!todoText) return;
        const todoTextContent = todoText.querySelector('.todo-text-content');
        if (!todoTextContent || !todoTextContent.dataset.fullText) return;

        const fullText = todoTextContent.dataset.fullText;
        const todoId = btn.dataset.todoId;

        // Mark task as read when View more is clicked
        markTaskAsRead(todoId);

        // Expand - show full text with View less
        todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline view-less" data-todo-id="' + todoId + '">View less</span>';
        todoText.classList.remove('truncated');
        todoText.classList.add('expanded');

        // Attach listener to View less button
        const viewLessBtn = todoText.querySelector('.view-less');
        if (viewLessBtn) {
          viewLessBtn.addEventListener('click', (ev) => {
            ev.stopPropagation();
            const maxLength = 180;
            const truncated = fullText.substring(0, maxLength);
            todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline" data-todo-id="' + todoId + '">View more</span>';
            todoText.classList.remove('expanded');
            todoText.classList.add('truncated');
            // Re-initialize listeners
            addFyiViewMoreListeners(fyiList);
          });
        }
      });
    });
  }

  // Build HTML for FYI tag groups
  function buildFyiTagGroupHtml(group, newItemIds = []) {
    // Skip groups with null, undefined, or empty tag names
    if (!group.tagName || (typeof group.tagName === 'string' && group.tagName.trim() === '')) {
      return '';
    }

    const taskCount = group.tasks.length;
    if (taskCount === 0) return ''; // Skip empty groups

    const tagName = group.tagName;
    const hasUnread = group.tasks.some(t => isTaskUnread(t.id));

    // Sort tasks by date
    const sortedTasks = [...group.tasks].sort((a, b) =>
      new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at)
    );

    // Create a comma-separated list of task IDs
    const taskIds = sortedTasks.map(t => t.id).join(',');

    return `
      <div class="fyi-tag-group tag-group task-group ${hasUnread ? 'unread' : ''}" data-tag-name="${escapeHtml(tagName)}" data-task-ids="${taskIds}">
        <div class="task-group-header">
          <div class="task-group-checkbox" data-task-ids="${taskIds}" title="Mark all as done"></div>
          <div class="task-group-icon" style="background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 6px; width: 28px; height: 28px; display: flex; align-items: center; justify-content: center;">
            <span style="font-size: 14px;">🏷️</span>
          </div>
          <div class="task-group-info">
            <div class="task-group-title" title="${escapeHtml(tagName)}">${escapeHtml(tagName.length > 60 ? tagName.substring(0, 60) + '...' : tagName)}</div>
            <div class="task-group-meta">
              <span class="task-group-count">${taskCount} item${taskCount > 1 ? 's' : ''}</span>
              <span class="task-group-date">${formatDate(group.latestUpdate)}</span>
            </div>
          </div>
          <div class="task-group-actions">
            <span class="task-group-chevron">▼</span>
          </div>
        </div>
        <div class="task-group-tasks" style="display: none;">
          ${sortedTasks.map((task, index) => {
      const isUnread = isTaskUnread(task.id);
      const isNew = newItemIds.includes(task.id);
      const taskMessageLink = task.message_link || '';
      const secondaryLinks = task.secondary_links || [];

      // Handle long descriptions
      const taskName = task.task_name || '';
      const taskNameEscaped = escapeHtml(taskName);
      const maxLength = 180;
      const hasTaskTitle = !!task.task_title; const needsViewMore = taskName.length > maxLength || (hasTaskTitle && taskName.length > 60);
      let taskNameHtml;

      if (needsViewMore) {
        const truncated = escapeHtml(taskName.substring(0, maxLength));
        taskNameHtml = `<span class="todo-text-content" data-full-text="${taskNameEscaped}">${taskNameEscaped}</span><span class="view-more-inline" data-todo-id="${task.id}">View more</span>`;
      } else {
        taskNameHtml = `<span class="todo-text-content">${taskNameEscaped}</span>`;
      }

      // Get icon URLs
      const slackIconUrl = chrome.runtime.getURL('icon-slack.png');
      const gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
      const driveIconUrl = chrome.runtime.getURL('icon-drive.png');
      const freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
      const freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
      const freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
      const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
      const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
      const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
      const linkIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

      const getIconForLink = (link) => {
        if (link.includes('slack.com') || link.includes('app.slack.com')) {
          return { icon: '<img src="' + slackIconUrl + '" alt="Slack" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Slack' };
        } else if (link.includes('freshrelease.com')) {
          return { icon: '<img src="' + freshreleaseIconUrl + '" alt="Freshrelease" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshrelease' };
        } else if (link.includes('freshdesk.com')) {
          return { icon: '<img src="' + freshdeskIconUrl + '" alt="Freshdesk" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshdesk' };
        } else if (link.includes('freshservice.com')) {
          return { icon: '<img src="' + freshserviceIconUrl + '" alt="Freshservice" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshservice' };
        } else if (link.includes('mail.google.com')) {
          return { icon: '<img src="' + gmailIconUrl + '" alt="Gmail" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Gmail' };
        } else if (link.includes('docs.google.com/document')) {
          return { icon: '<img src="' + googleDocsIconUrl + '" alt="Google Docs" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Docs' };
        } else if (link.includes('docs.google.com/spreadsheets') || link.includes('sheets.google.com')) {
          return { icon: '<img src="' + googleSheetsIconUrl + '" alt="Google Sheets" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Sheets' };
        } else if (link.includes('docs.google.com/presentation') || link.includes('slides.google.com')) {
          return { icon: '<img src="' + googleSlidesIconUrl + '" alt="Google Slides" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Slides' };
        } else if (link.includes('drive.google.com')) {
          return { icon: '<img src="' + driveIconUrl + '" alt="Google Drive" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Drive' };
        }
        return { icon: linkIcon, title: 'View link' };
      };

      let sourceIconHtml = '';
      if (taskMessageLink) {
        const primaryIcon = getIconForLink(taskMessageLink);
        sourceIconHtml = '<a href="' + taskMessageLink + '" target="_blank" class="todo-source" title="' + primaryIcon.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + primaryIcon.icon + '</a>';
      }

      const secondaryLinksHtml = secondaryLinks.filter(l => l != null).map(link => {
        const iconData = getIconForLink(link);
        return '<span style="color: #bdc3c7; margin: 0 4px;">|</span><a href="' + link + '" target="_blank" class="todo-source" title="' + iconData.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + iconData.icon + '</a>';
      }).join('');

      // Render other tags (excluding the group's tag and empty tags)
      const todoTags = (task.tags || []).filter(t => t !== tagName && isValidTag(t));
      const tagsHtml = todoTags.length > 0 ? `
              <div class="todo-tags" style="display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px;">
                ${todoTags.map(tag => {
        const isActive = activeTagFilters.includes(tag);
        return `<span class="todo-tag ${isActive ? 'active' : ''}" data-tag="${escapeHtml(tag)}" style="display: inline-block; background: ${isActive ? 'linear-gradient(45deg, #667eea, #764ba2)' : 'rgba(102, 126, 234, 0.1)'}; color: ${isActive ? 'white' : '#667eea'}; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 500; cursor: pointer; border: 1px solid ${isActive ? 'transparent' : 'rgba(102, 126, 234, 0.3)'}; transition: all 0.2s;">${escapeHtml(tag)}</span>`;
      }).join('')}
              </div>
            ` : '';

      return `
              <div class="task-group-task-item todo-item ${isUnread ? 'unread' : ''} ${isNew ? 'appearing' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                <div class="todo-content" style="flex: 1; min-width: 0;">
                  <div class="todo-main-row" style="display: flex; align-items: flex-start; gap: 8px;">
                    <div class="todo-checkbox" data-todo-id="${task.id}"></div>
                    <div style="flex: 1; min-width: 0;">
                      ${task.task_title ? `<div class="todo-title" style="font-weight: 600; font-size: 13px; color: var(--text-primary); margin-bottom: 2px;">${escapeHtml(task.task_title)}</div>` : ''}
                      <div class="todo-text ${needsViewMore ? 'truncated' : ''}">${taskNameHtml}</div>
                      ${task.participant_text ? `<span class="todo-participant" style="display: block; font-size: 11px; color: var(--text-muted); margin-top: 2px;">${escapeHtml(task.participant_text)}</span>` : ''}
                    </div>
                  </div>
                  <div class="todo-meta" style="display: flex; align-items: center; gap: 6px; margin-top: 4px; margin-left: 24px;">
                    <span class="todo-date" style="font-size: 11px; color: var(--text-light);">${formatDate(task.updated_at || task.created_at)}</span>
                    <div class="todo-sources" style="display: flex; align-items: center; gap: 2px;">
                      ${sourceIconHtml}
                      ${secondaryLinksHtml}
                    </div>
                  </div>
                  ${tagsHtml}
                </div>
              </div>
            `;
    }).join('')}
        </div>
      </div>
    `;
  }

  // Setup event handlers for FYI tag groups
  function setupFyiTagGroups(container) {
    const tagGroups = container.querySelectorAll('.fyi-tag-group');

    tagGroups.forEach(group => {
      const header = group.querySelector('.task-group-header');
      const tasksContainer = group.querySelector('.task-group-tasks');
      const chevron = group.querySelector('.task-group-chevron');
      const groupCheckbox = group.querySelector('.task-group-checkbox');

      if (header && tasksContainer && chevron) {
        header.addEventListener('click', (e) => {
          if (e.target.closest('.task-group-checkbox')) return;

          const isExpanded = tasksContainer.style.display !== 'none';
          tasksContainer.style.display = isExpanded ? 'none' : 'block';
          chevron.style.transform = isExpanded ? 'rotate(0deg)' : 'rotate(180deg)';
          group.classList.toggle('expanded', !isExpanded);
        });
      }

      if (groupCheckbox) {
        console.log('Setting up FYI tag group checkbox handler');
        groupCheckbox.addEventListener('click', async (e) => {
          console.log('FYI tag group checkbox clicked');
          e.stopPropagation();
          e.preventDefault();
          const taskIds = groupCheckbox.getAttribute('data-task-ids').split(',').map(id => parseInt(id));
          console.log('Task IDs to mark as done:', taskIds);

          // Visual feedback with animation
          groupCheckbox.classList.add('checked');
          groupCheckbox.innerHTML = '✓';
          group.classList.add('completing');

          // Mark all tasks as done in background (don't wait)
          const promises = taskIds.map(id =>
            updateTodoField(id, 'status', 1).catch(err => console.error('Error marking task done:', err))
          );
          Promise.all(promises);

          // Remove the group after animation
          setTimeout(() => {
            group.remove();
            // Update FYI items array
            allFyiItems = allFyiItems.filter(item => !taskIds.includes(item.id));
            updateTabCounts();
          }, 400);
        });
      } else {
        console.log('FYI tag group checkbox NOT found');
      }

      // Setup individual task checkboxes
      const taskCheckboxes = group.querySelectorAll('.task-group-task-item .todo-checkbox');
      taskCheckboxes.forEach(checkbox => {
        checkbox.addEventListener('click', async (e) => {
          e.stopPropagation();
          const todoId = checkbox.getAttribute('data-todo-id');
          const taskItem = checkbox.closest('.task-group-task-item');

          // Visual feedback with animation
          checkbox.classList.add('checked');
          checkbox.innerHTML = '✓';
          taskItem.classList.add('completing');

          // Start API call in background (don't wait)
          updateTodoField(parseInt(todoId), 'status', 1).catch(err =>
            console.error('Error marking task done:', err)
          );

          // Remove item after animation
          setTimeout(() => {
            taskItem.remove();
            // Update FYI items array
            allFyiItems = allFyiItems.filter(item => item.id != todoId);

            // Update group count or remove group if empty
            const remainingTasks = group.querySelectorAll('.task-group-task-item').length;
            if (remainingTasks === 0) {
              group.remove();
            } else {
              const countEl = group.querySelector('.task-group-count');
              if (countEl) {
                countEl.textContent = `${remainingTasks} item${remainingTasks > 1 ? 's' : ''}`;
              }
            }
            updateTabCounts();
          }, 400);
        });
      });

      // Setup click handler for opening transcript slider on task items
      const taskItems = group.querySelectorAll('.task-group-task-item');
      taskItems.forEach(taskItem => {
        taskItem.addEventListener('click', (e) => {
          // Don't trigger if clicking on checkbox, view-more, or links
          if (e.target.closest('.todo-checkbox') ||
            e.target.closest('.view-more-inline') ||
            e.target.closest('a')) {
            return;
          }

          const todoId = taskItem.dataset.todoId;
          if (todoId) {
            showTranscriptSlider(todoId);
          }
        });
      });

      // Setup view more listeners for this group
      group.querySelectorAll('.view-more-inline:not(.view-less)').forEach(btn => {
        if (btn.dataset.listenerAttached) return;
        btn.dataset.listenerAttached = 'true';

        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          const todoText = btn.closest('.todo-text');
          if (!todoText) return;
          const todoTextContent = todoText.querySelector('.todo-text-content');
          if (!todoTextContent || !todoTextContent.dataset.fullText) return;

          const fullText = todoTextContent.dataset.fullText;
          const todoId = btn.dataset.todoId;

          markTaskAsRead(todoId);

          todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline view-less" data-todo-id="' + todoId + '">View less</span>';
          todoText.classList.remove('truncated');
          todoText.classList.add('expanded');

          const viewLessBtn = todoText.querySelector('.view-less');
          if (viewLessBtn) {
            viewLessBtn.addEventListener('click', (ev) => {
              ev.stopPropagation();
              const maxLength = 180;
              const truncated = fullText.substring(0, maxLength);
              todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline" data-todo-id="' + todoId + '">View more</span>';
              todoText.classList.remove('expanded');
              todoText.classList.add('truncated');
              setupFyiTagGroups(container);
            });
          }
        });
      });
    });
  }

  function showEmptyFYIState(container) {
    const emptyState = container.querySelector('.empty-state');
    const fyiList = container.querySelector('.fyi-list');
    if (fyiList) fyiList.style.display = 'none';
    if (emptyState) {
      emptyState.style.display = 'flex';
      emptyState.innerHTML = '<h3>No FYI items</h3><p>Items with starred=0 and status=0 will appear here</p>';
    }
  }

  async function updateTodoField(todoId, field, value) {
    // Track tasks being marked as done to prevent reappearing via pending updates
    if (field === 'status' && value === 1) {
      addRecentlyCompleted(todoId);
    }
    try {
      const payload = createAuthenticatedPayload({
        action: 'update_todo',
        todo_id: todoId,
        [field]: value,
        timestamp: new Date().toISOString()
      });
      await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      updateTabCounts();
    } catch (error) {
      console.error('Error updating todo:', error);
    }
  }

  // ===== All Tasks Functions (status=1, completed in last 24 hours) =====
  let allCompletedTasks = [];

  // All Tasks — delegated to OracleNotes (with local displayAllTasks impl for DOM builders)
  const { loadAllTasks, searchAllTasks, setupAllTasksSearch: _setupATS } = window.OracleNotes;
  // Expose displayAllTasks impl that uses local DOM builders
  window._displayAllTasksImpl = displayAllTasksLocal;
  _setupATS();

  function displayAllTasksLocal(tasks) {
    // Sync completed tasks to local variable so showTranscriptSlider can find them
    allCompletedTasks = tasks;
    const container = document.querySelector('.alltasks-container');
    if (!container) return;
    const tasksList = container.querySelector('.alltasks-list');
    const emptyState = container.querySelector('.empty-state');
    if (emptyState) emptyState.style.display = 'none';
    if (!tasks || !tasks.length || !tasksList) return;
    const meetingTasks = tasks.filter(t => isMeetingLink(t.message_link)).map(t => (!t.due_by && t.updated_at) ? { ...t, due_by: t.updated_at } : t);
    const nonMeetingTasks = tasks.filter(t => !isMeetingLink(t.message_link));
    const meetingsHtml = buildMeetingsAccordion(meetingTasks);
    const { driveGroups, nonDriveTasks } = groupTasksByDriveFile(nonMeetingTasks);
    const { slackGroups, nonSlackTasks } = groupTasksBySlackChannel(nonDriveTasks);
    const { tagGroups, untaggedTasks } = groupTasksByTag(nonSlackTasks);
    const items = [];
    Object.values(driveGroups).forEach(g => items.push({ type: 'group', data: g, ts: new Date(g.latestUpdate) }));
    Object.values(slackGroups).forEach(g => items.push({ type: 'slack-group', data: g, ts: new Date(g.latestUpdate) }));
    Object.values(tagGroups).forEach(g => items.push({ type: 'tag-group', data: g, ts: new Date(g.latestUpdate) }));
    untaggedTasks.forEach(t => items.push({ type: 'task', data: t, ts: new Date(t.updated_at || t.created_at) }));
    items.sort((a, b) => b.ts - a.ts);
    let html = meetingsHtml || '';
    items.forEach((it, idx) => {
      if (it.type === 'group') html += buildTaskGroupHtml(it.data, []);
      else if (it.type === 'slack-group') html += buildSlackChannelGroupHtml(it.data, []);
      else if (it.type === 'tag-group') html += buildTagGroupHtml(it.data, []);
      else html += buildSingleTodoHtml(it.data, idx, [], 'all');
    });
    tasksList.innerHTML = html; tasksList.style.display = 'flex';
    tasksList.querySelectorAll('.todo-checkbox').forEach(cb => { const b = document.createElement('button'); b.className = 'alltask-reactivate'; b.dataset.taskId = cb.dataset.todoId; b.title = 'Mark as active'; b.textContent = '↩'; cb.replaceWith(b); });
    tasksList.querySelectorAll('.meeting-checkbox').forEach(cb => { const b = document.createElement('button'); b.className = 'alltask-reactivate'; b.dataset.taskId = cb.dataset.todoId; b.textContent = '↩'; b.style.cssText = 'width:24px;height:24px;min-width:24px;background:rgba(46,204,113,0.1);border:1px solid rgba(46,204,113,0.3);color:#27ae60;border-radius:6px;cursor:pointer;font-size:12px;display:flex;align-items:center;justify-content:center;'; cb.replaceWith(b); });
    tasksList.querySelectorAll('.document-group-checkbox,.task-group-checkbox,.slack-group-checkbox,.tag-group-checkbox').forEach(cb => cb.style.display = 'none');
    tasksList.querySelectorAll('.todo-clock').forEach(el => el.style.display = 'none');
    tasksList.querySelectorAll('.unread').forEach(el => el.classList.remove('unread'));
    if (meetingsHtml) setupMeetingsAccordion(tasksList);
    setupTaskGroups(tasksList); setupSlackChannelGroups(tasksList); setupTagGroups(tasksList);
    addViewMoreListeners();
    tasksList.querySelectorAll('.alltask-reactivate').forEach(btn => { btn.addEventListener('click', async (e) => { e.stopPropagation(); const id = parseInt(btn.dataset.taskId); btn.textContent = '...'; btn.disabled = true; try { await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({ action: 'activate_todo', todo_id: id, timestamp: new Date().toISOString() })) }); const it = btn.closest('.todo-item'); if (it) { it.style.transition = 'opacity 0.3s,transform 0.3s'; it.style.opacity = '0'; it.style.transform = 'translateX(20px)'; setTimeout(() => { it.remove(); if (!tasksList.querySelectorAll('.todo-item').length) { const es = container.querySelector('.empty-state'); if (es) { es.style.display = 'flex'; es.innerHTML = '<h3>No completed tasks</h3>'; } } }, 300); } showToastNotification('Task reactivated!'); } catch (er) { btn.textContent = '↩'; btn.disabled = false; } }); });
    tasksList.querySelectorAll('.todo-item').forEach(item => { if (item.closest('.task-group')) return; item.addEventListener('click', (e) => { if (e.target.closest('.alltask-reactivate,.todo-source,.todo-tag,.view-more-inline,a')) return; const id = item.dataset.todoId; if (id) showTranscriptSlider(id); }); });
  }

  // Bookmark Functions
  async function loadBookmarks() {
    const container = document.querySelector('.bookmarks-container');
    const emptyState = container.querySelector('.empty-state');
    const bookmarksList = document.querySelector('.bookmarks-list');
    if (emptyState) emptyState.style.display = 'none';
    if (bookmarksList) bookmarksList.style.display = 'none';
    showLoader(container);
    try {
      if (!isAuthenticated) throw new Error('Not authenticated');
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'list_bookmarks',
          timestamp: new Date().toISOString()
        }))
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);
      const data = await response.json();
      let bookmarksData = [];
      if (Array.isArray(data) && data.length > 0 && data[0].bookmarks) {
        bookmarksData = data[0].bookmarks;
      } else if (Array.isArray(data)) {
        bookmarksData = data;
      } else if (data.bookmarks) {
        bookmarksData = data.bookmarks;
      }
      // Sort bookmarks by created_at in descending order (newest first) before storing
      allBookmarks = bookmarksData.sort((a, b) => {
        const dateA = new Date(a.created_at);
        const dateB = new Date(b.created_at);
        return dateB - dateA;
      });
      hideLoader(container);
      if (allBookmarks.length > 0) displayBookmarks(allBookmarks);
      else showEmptyBookmarksState();
      updateTabCounts();
    } catch (error) {
      console.error('Error loading bookmarks:', error);
      hideLoader(container);
      if (emptyState) {
        emptyState.style.display = 'flex';
        emptyState.innerHTML = '<h3>Error: ' + error.message + '</h3>';
      }
    }
  }

  function displayBookmarks(bookmarks) {
    const container = document.querySelector('.bookmarks-container');
    hideLoader(container);
    let bookmarksList = document.querySelector('.bookmarks-list');
    const emptyState = container.querySelector('.empty-state');
    if (bookmarks.length > 0) {
      if (emptyState) emptyState.style.display = 'none';
      if (!bookmarksList) {
        bookmarksList = document.createElement('div');
        bookmarksList.className = 'bookmarks-list';
        container.appendChild(bookmarksList);
      }
      bookmarksList.innerHTML = bookmarks.map((bookmark, index) => {
        return `<div class="bookmark-item" style="animation-delay: ${index * 0.1}s" data-bookmark-id="${bookmark.id}" data-bookmark-title="${escapeHtml(bookmark.title || 'Untitled')}" data-bookmark-url="${bookmark.url || ''}">
          <div class="bookmark-content">
            <div class="bookmark-title" data-bookmark-id="${bookmark.id}">${escapeHtml(bookmark.title || 'Untitled')}</div>
            ${bookmark.url ? `<a href="${bookmark.url}" target="_blank" class="bookmark-url">${bookmark.url}</a>` : ''}
            <div class="bookmark-meta">
              <span class="bookmark-date">${formatDate(bookmark.created_at)}</span>
              <a href="${bookmark.url || '#'}" target="_blank" class="bookmark-link">🔗 Open</a>
            </div>
          </div>
          <div class="bookmark-actions">
            <div class="bookmark-copy" data-bookmark-id="${bookmark.id}" title="Copy">📋</div>
            <div class="bookmark-delete" data-bookmark-id="${bookmark.id}" title="Delete">🗑️</div>
          </div>
        </div>`;
      }).join('');
      bookmarksList.style.display = 'flex';
      addBookmarkEventListeners();
    } else {
      showEmptyBookmarksState();
    }
  }

  function addBookmarkEventListeners() {
    // Track selected bookmarks
    let selectedBookmarks = new Set();
    let isCommandPressed = false;

    // Track Command/Ctrl key
    document.addEventListener('keydown', (e) => {
      if (e.metaKey || e.ctrlKey) isCommandPressed = true;
    });

    document.addEventListener('keyup', (e) => {
      if (!e.metaKey && !e.ctrlKey) isCommandPressed = false;
    });

    // Click outside to clear selection
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.bookmark-item') && !e.target.closest('.bookmark-copy')) {
        clearBookmarkSelection();
      }
    });

    // Delete bookmark
    document.querySelectorAll('.bookmark-delete').forEach(el => {
      el.addEventListener('click', async (e) => {
        e.stopPropagation();
        const bookmarkItem = el.closest('.bookmark-item');
        const bookmarkId = el.dataset.bookmarkId;

        if (isCommandPressed) {
          // Multi-select mode
          toggleBookmarkSelection(bookmarkItem);
        } else {
          // Single delete
          if (confirm('Delete this bookmark?')) {
            try {
              await fetch(WEBHOOK_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(createAuthenticatedPayload({
                  action: 'delete_bookmark',
                  bookmark_id: bookmarkId,
                  timestamp: new Date().toISOString()
                }))
              });
              await loadBookmarks();
            } catch (error) {
              console.error('Error deleting bookmark:', error);
            }
          }
        }
      });
    });

    // Copy single bookmark
    document.querySelectorAll('.bookmark-copy').forEach(el => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const bookmarkItem = el.closest('.bookmark-item');

        if (isCommandPressed) {
          // Multi-select mode
          toggleBookmarkSelection(bookmarkItem);
        } else {
          // Single copy
          copySingleBookmark(bookmarkItem);
        }
      });
    });

    // Bulk copy button
    const bulkCopyBtn = document.getElementById('bulkCopyBookmarksBtn');
    if (bulkCopyBtn) {
      bulkCopyBtn.addEventListener('click', () => {
        copySelectedBookmarks();
      });
    }

    // Bulk delete button
    const bulkDeleteBtn = document.getElementById('bulkDeleteBookmarksBtn');
    if (bulkDeleteBtn) {
      bulkDeleteBtn.addEventListener('click', () => {
        deleteSelectedBookmarks();
      });
    }

    function toggleBookmarkSelection(bookmarkItem) {
      const id = bookmarkItem.dataset.bookmarkId;

      if (selectedBookmarks.has(id)) {
        selectedBookmarks.delete(id);
        bookmarkItem.classList.remove('selected');
      } else {
        selectedBookmarks.add(id);
        bookmarkItem.classList.add('selected');
      }

      updateBulkCopyButton();
    }

    function clearBookmarkSelection() {
      selectedBookmarks.clear();
      document.querySelectorAll('.bookmark-item.selected').forEach(item => {
        item.classList.remove('selected');
      });
      updateBulkCopyButton();
    }

    function updateBulkCopyButton() {
      const bulkCopyBtn = document.getElementById('bulkCopyBookmarksBtn');
      const bulkDeleteBtn = document.getElementById('bulkDeleteBookmarksBtn');
      const copyCountSpan = bulkCopyBtn?.querySelector('.bulk-count');
      const deleteCountSpan = bulkDeleteBtn?.querySelector('.bulk-count');

      if (selectedBookmarks.size > 0) {
        bulkCopyBtn?.classList.add('show');
        bulkDeleteBtn?.classList.add('show');
        if (copyCountSpan) copyCountSpan.textContent = selectedBookmarks.size;
        if (deleteCountSpan) deleteCountSpan.textContent = selectedBookmarks.size;
      } else {
        bulkCopyBtn?.classList.remove('show');
        bulkDeleteBtn?.classList.remove('show');
      }
    }

    function copySingleBookmark(bookmarkItem) {
      const title = bookmarkItem.dataset.bookmarkTitle;
      const url = bookmarkItem.dataset.bookmarkUrl;

      const text = `${title}\n${url}`;

      navigator.clipboard.writeText(text).then(() => {
        showCopyToast('Bookmark copied!');
      }).catch(err => {
        console.error('Copy failed:', err);
      });
    }

    function copySelectedBookmarks() {
      const bookmarks = [];

      selectedBookmarks.forEach(id => {
        const item = document.querySelector(`.bookmark-item[data-bookmark-id="${id}"]`);
        if (item) {
          const title = item.dataset.bookmarkTitle;
          const url = item.dataset.bookmarkUrl;
          bookmarks.push(`${title}\n${url}`);
        }
      });

      const text = bookmarks.join('\n\n');

      navigator.clipboard.writeText(text).then(() => {
        showCopyToast(`${selectedBookmarks.size} bookmarks copied!`);
        clearBookmarkSelection();
      }).catch(err => {
        console.error('Copy failed:', err);
      });
    }

    async function deleteSelectedBookmarks() {
      const count = selectedBookmarks.size;

      if (!confirm(`Delete ${count} selected bookmark${count > 1 ? 's' : ''}?`)) {
        return;
      }

      try {
        // Delete all selected bookmarks
        const deletePromises = Array.from(selectedBookmarks).map(id =>
          fetch(WEBHOOK_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(createAuthenticatedPayload({
              action: 'delete_bookmark',
              bookmark_id: id,
              timestamp: new Date().toISOString()
            }))
          })
        );

        await Promise.all(deletePromises);

        showCopyToast(`${count} bookmark${count > 1 ? 's' : ''} deleted!`);
        clearBookmarkSelection();
        await loadBookmarks();
      } catch (error) {
        console.error('Error deleting bookmarks:', error);
        alert('Failed to delete some bookmarks');
      }
    }

    function showCopyToast(message) {
      const toast = document.createElement('div');
      toast.style.cssText = 'position: fixed; top: 20px; right: 20px; background: linear-gradient(45deg, #667eea, #764ba2); color: white; padding: 12px 24px; border-radius: 8px; font-size: 14px; font-weight: 600; box-shadow: 0 4px 12px rgba(0,0,0,0.2); z-index: 10000; animation: slideDown 0.3s ease-out;';
      toast.textContent = message;
      document.body.appendChild(toast);

      setTimeout(() => {
        toast.style.animation = 'slideUp 0.3s ease-out';
        setTimeout(() => toast.remove(), 300);
      }, 2000);
    }
  }

  function showEmptyBookmarksState() {
    const container = document.querySelector('.bookmarks-container');
    const emptyState = container.querySelector('.empty-state');
    const bookmarksList = document.querySelector('.bookmarks-list');
    hideLoader(container);
    if (bookmarksList) bookmarksList.style.display = 'none';
    if (emptyState) {
      emptyState.style.display = 'flex';
      emptyState.innerHTML = '<h3>No bookmarks found</h3><p>Right-click on any page and select Oracle > Bookmark to save</p>';
    }
  }

  function updateBadge() {
    const now = new Date();
    const overdueTodos = allTodos.filter(t => t.status === 0 && t.due_by && new Date(t.due_by) < now);
    const overdueCount = overdueTodos.length;
    if (overdueCount > 0) {
      chrome.action.setBadgeText({ text: overdueCount.toString() });
      chrome.action.setBadgeBackgroundColor({ color: "#e74c3c" });
    } else {
      const starredCount = allTodos.filter(t => t.starred === 1).length;
      if (starredCount > 0) {
        chrome.action.setBadgeText({ text: starredCount.toString() });
        chrome.action.setBadgeBackgroundColor({ color: "#667eea" });
      } else {
        chrome.action.setBadgeText({ text: "" });
      }
    }
  }

  function sortTodos(todos) {
    return todos.sort((a, b) => {
      // 0. Unread items always come first
      const aUnread = isTaskUnread(a.id);
      const bUnread = isTaskUnread(b.id);
      if (aUnread && !bUnread) return -1;
      if (!aUnread && bUnread) return 1;

      // 1. Sort by due_by if present (closest to farthest)
      const aDue = a.due_by ? new Date(a.due_by).getTime() : null;
      const bDue = b.due_by ? new Date(b.due_by).getTime() : null;
      if (aDue && bDue) return aDue - bDue;
      if (aDue && !bDue) return -1;
      if (!aDue && bDue) return 1;

      // 2. If no due_by, sort by updated_at (most recent first)
      const aUpdated = a.updated_at ? new Date(a.updated_at).getTime() : null;
      const bUpdated = b.updated_at ? new Date(b.updated_at).getTime() : null;
      if (aUpdated && bUpdated) return bUpdated - aUpdated;
      if (aUpdated && !bUpdated) return -1;
      if (!aUpdated && bUpdated) return 1;

      // 3. If no updated_at, sort by created_at (most recent first)
      return new Date(b.created_at) - new Date(a.created_at);
    });
  }

  function displayTodos(todos, currentFilter, newItemIds = []) {
    const container = document.querySelector('.todos-container');
    hideLoader(container);
    currentTodoFilter = currentFilter; // Store current filter

    // Separate meeting items from non-meeting items
    const meetingTodos = todos.filter(t => isMeetingLink(t.message_link) && t.status === 0);
    const nonMeetingTodos = todos.filter(t => !isMeetingLink(t.message_link));

    // Store meeting items globally for accordion
    allCalendarItems = meetingTodos;

    const starredCount = nonMeetingTodos.filter(t => t.starred === 1).length;
    const activeCount = nonMeetingTodos.filter(t => t.status === 0).length;
    document.querySelectorAll('.filter-count').forEach(el => {
      const parent = el.closest('.filter-btn');
      if (parent) {
        const filter = parent.dataset.filter;
        if (filter === 'starred') el.textContent = starredCount;
        if (filter === 'active') el.textContent = activeCount;
        if (filter === 'all') el.textContent = nonMeetingTodos.length;
      }
    });
    updateFilterSlider();

    // Apply status filter first (on non-meeting todos)
    let filteredTodos = nonMeetingTodos;
    if (currentFilter === 'starred') filteredTodos = nonMeetingTodos.filter(t => t.starred === 1);
    else if (currentFilter === 'active') filteredTodos = nonMeetingTodos.filter(t => t.status === 0);

    // Store base filtered count before tag filtering
    const baseFilteredCount = filteredTodos.length;

    // Apply tag filters
    if (activeTagFilters.length > 0) {
      filteredTodos = filteredTodos.filter(todo => {
        const todoTags = todo.tags || [];
        return activeTagFilters.every(filterTag => todoTags.includes(filterTag));
      });
    }

    const sortedTodos = sortTodos(filteredTodos);
    let todosList = container.querySelector('.todos-list');
    const emptyState = container.querySelector('.empty-state');

    // Build filter chips HTML for active tag filters
    let filterChipsHtml = '';
    if (activeTagFilters.length > 0) {
      filterChipsHtml = `
        <div class="active-filters" style="margin-bottom: 12px; display: flex; flex-wrap: wrap; gap: 6px; align-items: center; padding: 0 4px;">
          <span style="font-size: 12px; color: #7f8c8d; font-weight: 500;">Filters:</span>
          ${activeTagFilters.map(tag => `
            <span class="filter-chip" data-tag="${escapeHtml(tag)}" style="display: inline-flex; align-items: center; gap: 4px; background: linear-gradient(45deg, #667eea, #764ba2); color: white; padding: 4px 10px; border-radius: 6px; font-size: 11px; font-weight: 500; cursor: pointer; transition: all 0.2s;">
              ${escapeHtml(tag)}
              <span style="font-size: 14px; opacity: 0.8; margin-left: 2px;">×</span>
            </span>
          `).join('')}
          <button class="clear-filters-btn" style="background: rgba(231,76,60,0.1); border: 1px solid rgba(231,76,60,0.3); color: #e74c3c; padding: 4px 10px; border-radius: 6px; font-size: 11px; cursor: pointer; font-weight: 500; transition: all 0.2s;">Clear all</button>
        </div>
      `;
    }

    // Update the starred tasks header with filter count
    const headerH3 = container.querySelector('.todos-header h3');
    if (headerH3) {
      const filterLabel = currentFilter === 'starred' ? 'Starred Tasks' : (currentFilter === 'active' ? 'Active Tasks' : 'All Tasks');
      if (activeTagFilters.length > 0) {
        headerH3.textContent = `${filterLabel} (${sortedTodos.length} / ${baseFilteredCount})`;
      } else {
        headerH3.textContent = `${filterLabel} (${baseFilteredCount})`;
      }
    }

    // Build meetings accordion HTML
    const meetingsAccordionHtml = buildMeetingsAccordion(meetingTodos);

    if (sortedTodos.length > 0 || activeTagFilters.length > 0 || meetingTodos.length > 0) {
      if (emptyState) emptyState.style.display = 'none';
      if (!todosList) {
        todosList = document.createElement('div');
        todosList.className = 'todos-list';
        container.appendChild(todosList);
      }

      // Remove and re-add meetings accordion in the slot above column content
      const meetingsSlot = document.getElementById('meetingsAccordionSlot');
      if (meetingsSlot) {
        meetingsSlot.innerHTML = meetingsAccordionHtml || '';
        if (meetingsAccordionHtml) {
          setupMeetingsAccordion(meetingsSlot);
        }
      } else {
        // Fallback: insert inside container
        let existingAccordion = container.querySelector('.meetings-accordion');
        if (existingAccordion) existingAccordion.remove();
        if (meetingsAccordionHtml) {
          todosList.insertAdjacentHTML('beforebegin', meetingsAccordionHtml);
          setupMeetingsAccordion(container);
        }
      }

      // Insert filter chips before the todos list
      let existingFilters = container.querySelector('.active-filters');
      if (existingFilters) existingFilters.remove();
      if (filterChipsHtml) {
        todosList.insertAdjacentHTML('beforebegin', filterChipsHtml);
        // Attach filter chip event listeners
        attachTagFilterListeners(container);
      }

      if (sortedTodos.length > 0) {
        // Group Drive/Docs tasks by file ID
        const { driveGroups, nonDriveTasks } = groupTasksByDriveFile(sortedTodos);

        // Group Slack tasks by channel (from non-drive tasks)
        const { slackGroups, nonSlackTasks } = groupTasksBySlackChannel(nonDriveTasks);

        // Group remaining tasks by tags
        const { tagGroups, untaggedTasks } = groupTasksByTag(nonSlackTasks);

        // Update global sets for FYI filtering
        actionTabDriveFileIds = new Set(Object.keys(driveGroups));
        // Store task IDs shown in Action tab Slack groups (not channel names)
        actionTabSlackTaskIds = new Set();
        Object.values(slackGroups).forEach(group => {
          group.tasks.forEach(task => actionTabSlackTaskIds.add(task.id));
        });

        // Create a unified list of items (both groups and individual tasks) for sorting
        const unifiedItems = [];

        // Add drive task groups with their latest timestamp
        Object.values(driveGroups).forEach(group => {
          unifiedItems.push({
            type: 'group',
            data: group,
            sortTimestamp: new Date(group.latestUpdate)
          });
        });

        // Add Slack channel groups
        Object.values(slackGroups).forEach(group => {
          unifiedItems.push({
            type: 'slack-group',
            data: group,
            sortTimestamp: new Date(group.latestUpdate)
          });
        });

        // Add tag groups
        Object.values(tagGroups).forEach(group => {
          unifiedItems.push({
            type: 'tag-group',
            data: group,
            sortTimestamp: new Date(group.latestUpdate)
          });
        });

        // Add remaining untagged tasks
        untaggedTasks.forEach(task => {
          unifiedItems.push({
            type: 'task',
            data: task,
            sortTimestamp: new Date(task.updated_at || task.created_at)
          });
        });

        // Sort: unread items first, then by timestamp descending
        unifiedItems.sort((a, b) => {
          // Check if items have unread tasks
          const aHasUnread = a.type === 'task'
            ? isTaskUnread(a.data.id)
            : (a.type === 'group' || a.type === 'slack-group' || a.type === 'tag-group')
              ? a.data.tasks.some(t => isTaskUnread(t.id))
              : false;
          const bHasUnread = b.type === 'task'
            ? isTaskUnread(b.data.id)
            : (b.type === 'group' || b.type === 'slack-group' || b.type === 'tag-group')
              ? b.data.tasks.some(t => isTaskUnread(t.id))
              : false;

          // Unread items come first
          if (aHasUnread && !bHasUnread) return -1;
          if (!aHasUnread && bHasUnread) return 1;

          // Then sort by timestamp
          return b.sortTimestamp - a.sortTimestamp;
        });

        // Build HTML in sorted order
        let combinedHtml = '';
        unifiedItems.forEach((item, index) => {
          if (item.type === 'group') {
            combinedHtml += buildTaskGroupHtml(item.data, newItemIds);
          } else if (item.type === 'slack-group') {
            combinedHtml += buildSlackChannelGroupHtml(item.data, newItemIds);
          } else if (item.type === 'tag-group') {
            combinedHtml += buildTagGroupHtml(item.data, newItemIds);
          } else {
            const todo = item.data;
            combinedHtml += buildSingleTodoHtml(todo, index, newItemIds, currentFilter);
          }
        });

        todosList.innerHTML = combinedHtml;
        todosList.style.display = 'flex';
        addTodoEventListeners();
        addViewMoreListeners();
        // Attach tag click listeners
        attachTodoTagListeners();
        // Setup task group event handlers
        setupTaskGroups(todosList);
        // Setup Slack channel group event handlers
        setupSlackChannelGroups(todosList);
        // Setup tag group event handlers
        setupTagGroups(todosList);
      } else {
        // No results after tag filtering
        todosList.innerHTML = `
          <div class="empty-state" style="display: flex;">
            <h3>No tasks match these filters</h3>
            <p>Try removing some filters to see more tasks</p>
          </div>
        `;
        todosList.style.display = 'flex';
      }
    } else {
      showEmptyTodosState(currentFilter);
    }
  }

  // Helper function to build a single todo item HTML
  function buildSingleTodoHtml(todo, index, newItemIds = [], currentFilter = 'starred') {
    const messageLink = todo.message_link || '';
    const secondaryLinks = todo.secondary_links || [];

    // Use image files for icons
    const slackIconUrl = chrome.runtime.getURL('icon-slack.png');
    const gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
    const driveIconUrl = chrome.runtime.getURL('icon-drive.png');
    const freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
    const freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
    const freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
    const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
    const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
    const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');

    const linkIcon = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

    // Image icon URLs
    const zoomIconUrl = chrome.runtime.getURL('icon-zoom.png');
    const googleMeetIconUrl = chrome.runtime.getURL('icon-google-meet.png');
    const googleCalendarIconUrl = chrome.runtime.getURL('icon-google-calendar.png');

    let sourceIcon = linkIcon;
    let sourceTitle = 'View source';
    let isImageIcon = false;
    let iconUrl = '';

    if (messageLink.includes('zoom.us') || messageLink.includes('zoom.com')) {
      iconUrl = zoomIconUrl;
      sourceTitle = 'Open in Zoom';
      isImageIcon = true;
    } else if (messageLink.includes('meet.google.com')) {
      iconUrl = googleMeetIconUrl;
      sourceTitle = 'Open in Google Meet';
      isImageIcon = true;
    } else if (messageLink.includes('calendar.google.com')) {
      iconUrl = googleCalendarIconUrl;
      sourceTitle = 'Open in Google Calendar';
      isImageIcon = true;
    } else if (messageLink.includes('mail.google.com')) {
      iconUrl = gmailIconUrl;
      sourceTitle = 'Open in Gmail';
      isImageIcon = true;
    } else if (messageLink.includes('slack.com') || messageLink.includes('app.slack.com')) {
      iconUrl = slackIconUrl;
      sourceTitle = 'Open in Slack';
      isImageIcon = true;
    } else if (messageLink.includes('freshdesk.com')) {
      iconUrl = freshdeskIconUrl;
      sourceTitle = 'Open in Freshdesk';
      isImageIcon = true;
    } else if (messageLink.includes('freshrelease.com')) {
      iconUrl = freshreleaseIconUrl;
      sourceTitle = 'Open in Freshrelease';
      isImageIcon = true;
    } else if (messageLink.includes('freshservice.com')) {
      iconUrl = freshserviceIconUrl;
      sourceTitle = 'Open in Freshservice';
      isImageIcon = true;
    } else if (messageLink.includes('docs.google.com/document')) {
      iconUrl = googleDocsIconUrl;
      sourceTitle = 'Open in Google Docs';
      isImageIcon = true;
    } else if (messageLink.includes('docs.google.com/spreadsheets') || messageLink.includes('sheets.google.com')) {
      iconUrl = googleSheetsIconUrl;
      sourceTitle = 'Open in Google Sheets';
      isImageIcon = true;
    } else if (messageLink.includes('docs.google.com/presentation') || messageLink.includes('slides.google.com')) {
      iconUrl = googleSlidesIconUrl;
      sourceTitle = 'Open in Google Slides';
      isImageIcon = true;
    } else if (messageLink.includes('drive.google.com')) {
      iconUrl = driveIconUrl;
      sourceTitle = 'Open in Google Drive';
      isImageIcon = true;
    }

    const sourceIconHtml = isImageIcon
      ? `<img src="${iconUrl}" alt="${sourceTitle}" style="width: 14px; height: 14px; object-fit: contain;">`
      : sourceIcon;

    // Generate secondary links HTML
    const secondaryLinksHtml = secondaryLinks.filter(l => l != null).map(link => {
      let secondaryIconHtml = linkIcon;
      let secondaryTitle = 'View link';

      if (link.includes('zoom.us') || link.includes('zoom.com')) {
        secondaryIconHtml = `<img src="${zoomIconUrl}" alt="Zoom" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Zoom';
      } else if (link.includes('meet.google.com')) {
        secondaryIconHtml = `<img src="${googleMeetIconUrl}" alt="Google Meet" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Google Meet';
      } else if (link.includes('calendar.google.com')) {
        secondaryIconHtml = `<img src="${googleCalendarIconUrl}" alt="Google Calendar" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Google Calendar';
      } else if (link.includes('freshdesk.com')) {
        secondaryIconHtml = `<img src="${freshdeskIconUrl}" alt="Freshdesk" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Freshdesk';
      } else if (link.includes('freshrelease.com')) {
        secondaryIconHtml = `<img src="${freshreleaseIconUrl}" alt="Freshrelease" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Freshrelease';
      } else if (link.includes('freshservice.com')) {
        secondaryIconHtml = `<img src="${freshserviceIconUrl}" alt="Freshservice" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Freshservice';
      } else if (link.includes('docs.google.com/document')) {
        secondaryIconHtml = `<img src="${googleDocsIconUrl}" alt="Google Docs" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Google Docs';
      } else if (link.includes('docs.google.com/spreadsheets') || link.includes('sheets.google.com')) {
        secondaryIconHtml = `<img src="${googleSheetsIconUrl}" alt="Google Sheets" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Google Sheets';
      } else if (link.includes('docs.google.com/presentation') || link.includes('slides.google.com')) {
        secondaryIconHtml = `<img src="${googleSlidesIconUrl}" alt="Google Slides" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Google Slides';
      } else if (link.includes('drive.google.com')) {
        secondaryIconHtml = `<img src="${driveIconUrl}" alt="Google Drive" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Google Drive';
      } else if (link.includes('mail.google.com')) {
        secondaryIconHtml = `<img src="${gmailIconUrl}" alt="Gmail" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Gmail';
      } else if (link.includes('slack.com') || link.includes('app.slack.com')) {
        secondaryIconHtml = `<img src="${slackIconUrl}" alt="Slack" style="width: 14px; height: 14px; object-fit: contain;">`;
        secondaryTitle = 'Open in Slack';
      }

      return `<span style="color: #bdc3c7; margin: 0 4px;">|</span><a href="${link}" target="_blank" class="todo-source" title="${secondaryTitle}" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">${secondaryIconHtml}</a>`;
    }).join('');

    const dueByHtml = todo.due_by ? `<span class="todo-due ${new Date(todo.due_by) < new Date() ? 'overdue' : ''}">${formatDueBy(todo.due_by)}</span>` : '';
    const showClock = currentFilter === 'starred' && todo.starred === 1;

    // Handle descriptions - always show with View more when needed
    const taskNameRaw = todo.task_name || todo.task_title || 'Untitled Task';
    const taskNameEscaped = escapeHtml(taskNameRaw);
    const maxLength = 180;
    const hasTitle = !!todo.task_title;
    const needsViewMore = taskNameRaw.length > maxLength || (hasTitle && taskNameRaw.length > 60);
    let taskNameHtml;

    if (needsViewMore) {
      const truncated = escapeHtml(taskNameRaw.substring(0, maxLength));
      taskNameHtml = `<span class="todo-text-content" data-full-text="${taskNameEscaped}">${taskNameEscaped}</span><span class="view-more-inline" data-todo-id="${todo.id}">View more</span>`;
    } else {
      taskNameHtml = `<span class="todo-text-content">${taskNameEscaped}</span>`;
    }














    // Render tags for each todo (filter out invalid tags like "null")
    const todoTags = (todo.tags || []).filter(isValidTag);
    const tagsHtml = todoTags.length > 0 ? `
            <div class="todo-tags" style="display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px;">
              ${todoTags.map(tag => {
      const isActive = activeTagFilters.includes(tag);
      return `<span class="todo-tag ${isActive ? 'active' : ''}" data-tag="${escapeHtml(tag)}" style="display: inline-block; background: ${isActive ? 'linear-gradient(45deg, #667eea, #764ba2)' : 'rgba(102, 126, 234, 0.1)'}; color: ${isActive ? 'white' : '#667eea'}; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 500; cursor: pointer; border: 1px solid ${isActive ? 'transparent' : 'rgba(102, 126, 234, 0.3)'}; transition: all 0.2s;">${escapeHtml(tag)}</span>`;
    }).join('')}
            </div>
          ` : '';


    // Participant text (e.g., channel name)
    const participantHtml = todo.participant_text ? `<div class="todo-participant">${escapeHtml(todo.participant_text)}</div>` : '';

    // Check if task is unread
    const isUnread = isTaskUnread(todo.id, todo.updated_at);

    return `<div class="todo-item ${todo.status === 1 ? 'completed' : ''} ${newItemIds.includes(todo.id) ? 'appearing' : ''} ${isUnread ? 'unread' : ''}" style="animation-delay: ${index * 0.05}s" data-todo-id="${todo.id}">
            <div class="todo-left-actions">
              <div class="todo-checkbox ${todo.status === 1 ? 'checked' : ''}" data-todo-id="${todo.id}">${todo.status === 1 ? '✓' : ''}</div>
            </div>
            <div class="todo-content">
              ${todo.task_title ? `<div class="todo-title">${escapeHtml(todo.task_title)}</div>` : ''}
              <div class="todo-text${needsViewMore ? ' truncated' : ''}">${taskNameHtml}</div>
              ${participantHtml}
              <div class="todo-meta">
                <span class="todo-date">${formatDate(todo.updated_at || todo.created_at)}</span>
                ${dueByHtml}
                ${messageLink ? `<a href="${messageLink}" target="_blank" class="todo-source" title="${sourceTitle}" style="padding: 6px;">${sourceIconHtml}</a>` : ''}
                ${secondaryLinksHtml}
              </div>
              ${tagsHtml}
            </div>
            <div class="todo-actions">
              ${showClock ? '<div class="todo-clock" data-todo-id="' + todo.id + '" title="Remind me in">🕐</div>' : ''}
            </div>
          </div>`;
  }

  // Tag filtering functions
  function toggleTagFilter(tag) {
    const index = activeTagFilters.indexOf(tag);
    if (index > -1) {
      activeTagFilters.splice(index, 1);
    } else {
      activeTagFilters.push(tag);
    }
    // Re-render with current filters
    displayTodos(allTodos, currentTodoFilter);
  }

  function clearAllTagFilters() {
    activeTagFilters = [];
    displayTodos(allTodos, currentTodoFilter);
  }

  function attachTagFilterListeners(container) {
    // Filter chip click listeners (to remove filter)
    container.querySelectorAll('.filter-chip').forEach(chip => {
      chip.onclick = (e) => {
        e.stopPropagation();
        toggleTagFilter(chip.dataset.tag);
      };
    });

    // Clear all button
    const clearBtn = container.querySelector('.clear-filters-btn');
    if (clearBtn) {
      clearBtn.onclick = (e) => {
        e.stopPropagation();
        clearAllTagFilters();
      };
    }
  }

  function attachTodoTagListeners() {
    document.querySelectorAll('.todo-tag').forEach(tagEl => {
      tagEl.onclick = (e) => {
        e.stopPropagation();
        toggleTagFilter(tagEl.dataset.tag);
      };
    });
  }

  // Meetings Accordion Functions
  function formatMeetingTime(dueBy) {
    if (!dueBy) return '';
    const due = new Date(dueBy);
    const now = new Date();
    const diffMs = due - now;

    // If overdue
    if (diffMs < 0) {
      return 'Overdue';
    }

    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const dayAfterTomorrow = new Date(today);
    dayAfterTomorrow.setDate(dayAfterTomorrow.getDate() + 2);

    const dueDate = new Date(due.getFullYear(), due.getMonth(), due.getDate());

    // Format time as "11 AM" or "2:30 PM"
    let hours = due.getHours();
    const minutes = due.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    const timeStr = minutes === 0 ? `${hours} ${ampm}` : `${hours}:${minutes.toString().padStart(2, '0')} ${ampm}`;

    if (dueDate.getTime() === today.getTime()) {
      return `Today ${timeStr}`;
    } else if (dueDate.getTime() === tomorrow.getTime()) {
      return `Tomorrow ${timeStr}`;
    } else {
      // Format as "Jan 15, 11 AM"
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${months[due.getMonth()]} ${due.getDate()}, ${timeStr}`;
    }
  }

  function buildMeetingsAccordion(meetings) {
    if (!meetings || meetings.length === 0) return '';

    // Sort meetings by due_by
    const sortedMeetings = [...meetings].sort((a, b) => {
      if (!a.due_by && !b.due_by) return 0;
      if (!a.due_by) return 1;
      if (!b.due_by) return -1;
      return new Date(a.due_by) - new Date(b.due_by);
    });

    // Group meetings by date
    const meetingsByDate = {};
    sortedMeetings.forEach(meeting => {
      const dateKey = meeting.due_by ? getDateKey(meeting.due_by) : 'no-date';
      if (!meetingsByDate[dateKey]) {
        meetingsByDate[dateKey] = [];
      }
      meetingsByDate[dateKey].push(meeting);
    });

    // Get sorted date keys (only dates with meetings)
    const dateKeys = Object.keys(meetingsByDate).filter(key => key !== 'no-date').sort();
    if (meetingsByDate['no-date'] && meetingsByDate['no-date'].length > 0) {
      dateKeys.push('no-date'); // Add no-date at the end if exists
    }

    // Check for meaningful meeting changes (new, rescheduled, cancelled)
    const currentMeetingIds = new Set(meetings.map(m => String(m.id)));
    const newOrModifiedMeetingIds = new Set();

    // Build snapshot of meaningful fields: due_by (time/date) + status (cancelled)
    // Ignores updated_at which changes on every sync
    const currentMeetingSnapshots = new Map();
    meetings.forEach(m => {
      currentMeetingSnapshots.set(String(m.id), `${m.due_by || ''}|${m.status ?? ''}`);
    });

    // Detect: new meetings, time/date changes, cancellations (status changes)
    if (!isInitialLoad && previousMeetingIds.size > 0) {
      currentMeetingIds.forEach(id => {
        if (!previousMeetingIds.has(id)) {
          // New meeting
          console.log('New meeting detected:', id);
          newOrModifiedMeetingIds.add(id);
        } else {
          const prevSnapshot = previousMeetingTimestamps.get(id);
          const currSnapshot = currentMeetingSnapshots.get(id);
          if (prevSnapshot !== currSnapshot) {
            // Time/date changed or status changed (e.g. cancelled)
            console.log('Meeting changed (time/status):', id, 'was:', prevSnapshot, 'now:', currSnapshot);
            newOrModifiedMeetingIds.add(id);
          }
        }
      });
    }

    console.log('newOrModifiedMeetingIds:', [...newOrModifiedMeetingIds]);
    console.log('isInitialLoad:', isInitialLoad);

    // Reset modifiedMeetingDates to only include currently-changed meeting dates
    modifiedMeetingDates.clear();

    // Mark dates that have new/changed meetings
    const updatedDateKeys = new Set();
    sortedMeetings.forEach(meeting => {
      if (newOrModifiedMeetingIds.has(String(meeting.id))) {
        const dateKey = meeting.due_by ? getDateKey(meeting.due_by) : 'no-date';
        updatedDateKeys.add(dateKey);
        modifiedMeetingDates.add(dateKey);
      }
    });

    console.log('modifiedMeetingDates:', [...modifiedMeetingDates]);

    // Update previous meeting IDs and snapshots for next comparison
    previousMeetingIds = currentMeetingIds;
    previousMeetingTimestamps = currentMeetingSnapshots;

    const hasAnyUpdates = modifiedMeetingDates.size > 0 && !isInitialLoad;
    console.log('hasAnyUpdates:', hasAnyUpdates);

    const googleCalendarIconUrl = chrome.runtime.getURL('icon-google-calendar.png');
    const zoomIconUrl = chrome.runtime.getURL('icon-zoom.png');
    const googleMeetIconUrl = chrome.runtime.getURL('icon-google-meet.png');

    // Build date tabs HTML
    const dateTabsHtml = dateKeys.map((dateKey, index) => {
      const meetingsOnDate = meetingsByDate[dateKey];
      const dateInfo = formatDateTabLabel(dateKey);
      const hasUpdates = modifiedMeetingDates.has(dateKey);
      const isActive = index === 0;

      return `
        <div class="meeting-date-tab ${isActive ? 'active' : ''} ${hasUpdates ? 'has-updates' : ''}" 
             data-date-key="${dateKey}">
          <span class="meeting-date-day">${dateInfo.dayName}</span>
          <span class="meeting-date-label">${dateInfo.dateLabel}</span>
          <span class="meeting-date-count">${meetingsOnDate.length}</span>
        </div>
      `;
    }).join('');

    // Build meeting content sections for each date
    const dateContentSections = dateKeys.map((dateKey, index) => {
      const meetingsOnDate = meetingsByDate[dateKey];
      const isActive = index === 0;

      // Sort meetings within date: unread first, then by time
      const sortedMeetingsOnDate = [...meetingsOnDate].sort((a, b) => {
        const aUnread = newOrModifiedMeetingIds.has(String(a.id));
        const bUnread = newOrModifiedMeetingIds.has(String(b.id));
        if (aUnread && !bUnread) return -1;
        if (!aUnread && bUnread) return 1;
        // Then by due_by time
        if (a.due_by && b.due_by) return new Date(a.due_by) - new Date(b.due_by);
        return 0;
      });

      const meetingItems = sortedMeetingsOnDate.map(meeting => {
        const isUnread = newOrModifiedMeetingIds.has(String(meeting.id));
        return buildMeetingItemHtml(meeting, googleCalendarIconUrl, zoomIconUrl, googleMeetIconUrl, isUnread);
      }).join('');

      return `
        <div class="meetings-date-content ${isActive ? 'active' : ''}" data-date-key="${dateKey}">
          <div class="meetings-list">
            ${meetingItems}
          </div>
        </div>
      `;
    }).join('');

    return `
      <div class="meetings-accordion ${hasAnyUpdates ? 'has-updates' : ''}">
        <div class="meetings-accordion-header">
          <div class="meetings-accordion-title">
            <span>📅</span>
            <span>Meetings</span>
            <span class="meetings-accordion-count">${meetings.length}</span>
          </div>
          <span class="meetings-accordion-chevron">▼</span>
        </div>
        <div class="meetings-accordion-content">
          <div class="meetings-date-tabs-container">
            <div class="meetings-date-tabs">
              ${dateTabsHtml}
            </div>
          </div>
          ${dateContentSections}
        </div>
      </div>
    `;
  }

  // Helper function to get consistent date key from datetime
  function getDateKey(dateTime) {
    if (!dateTime) return 'no-date';
    const date = new Date(dateTime);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  }

  // Helper function to format date tab label
  function formatDateTabLabel(dateKey) {
    if (dateKey === 'no-date') {
      return { dayName: 'TBD', dateLabel: 'No Date' };
    }

    const [year, month, day] = dateKey.split('-').map(Number);
    const date = new Date(year, month - 1, day);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    // Get ordinal suffix
    const getOrdinal = (n) => {
      const s = ['th', 'st', 'nd', 'rd'];
      const v = n % 100;
      return n + (s[(v - 20) % 10] || s[v] || s[0]);
    };

    const dayName = dayNames[date.getDay()];
    const dateNum = getOrdinal(date.getDate());
    const monthName = monthNames[date.getMonth()];

    // Check if it's today or tomorrow
    if (date.getTime() === today.getTime()) {
      return { dayName: 'Today', dateLabel: `${dateNum} ${monthName}` };
    } else if (date.getTime() === tomorrow.getTime()) {
      return { dayName: 'Tomorrow', dateLabel: `${dateNum} ${monthName}` };
    }

    return { dayName: dayName, dateLabel: `${dateNum} ${monthName}` };
  }

  // Format display name: "john.doe" or "john_doe" → "John Doe", already proper names pass through
  function formatDisplayName(name) {
    if (!name) return 'Unknown';
    // If name contains dots or underscores but no spaces, split and titlecase
    if ((name.includes('.') || name.includes('_')) && !name.includes(' ')) {
      return name
        .split(/[._]/)
        .map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
        .join(' ');
    }
    // If already has spaces, just ensure each word is capitalised
    if (name.includes(' ')) {
      return name
        .split(' ')
        .map(part => part.charAt(0).toUpperCase() + part.slice(1))
        .join(' ');
    }
    // Single word — capitalise first letter
    return name.charAt(0).toUpperCase() + name.slice(1);
  }

  // Extract organiser name from participant_text field
  function extractOrganiser(participantText) {
    if (!participantText) return '';
    // participant_text contains comma-separated names; first one is typically the organiser
    const names = participantText.split(',').map(n => n.trim()).filter(Boolean);
    return names.length > 0 ? formatDisplayName(names[0]) : '';
  }

  // Show meeting detail slider with participants, conference link, date/time, room
  async function showMeetingDetailSlider(meetingId) {
    const meeting = allCalendarItems.find(t => String(t.id) === String(meetingId)) || allCompletedTasks.find(t => String(t.id) === String(meetingId));
    if (!meeting) return;

    markTaskAsRead(meetingId);
    isTranscriptSliderOpen = true;

    // Remove any existing slider
    document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
    document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());

    const isDarkMode = document.body.classList.contains('dark-mode');

    // V37: Fixed overlay covering col3's screen position
    const col3El = document.getElementById('col3');
    const col3Rect = window.Oracle.getCol3Rect();

    const overlay = document.createElement('div');
    overlay.className = 'transcript-slider-overlay';
    if (col3Rect) {
      overlay.style.cssText = `position: fixed; top: ${col3Rect.top}px; left: ${col3Rect.left}px; width: ${col3Rect.width}px; bottom: 0; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;`;
    } else {
      overlay.style.cssText = 'position: fixed; top: 0; right: 0; bottom: 0; width: 400px; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;';
    }


    const slider = document.createElement('div');
    slider.className = 'transcript-slider';
    slider.style.cssText = `width: 100%; height: 100%; background: ${isDarkMode ? '#1f2940' : 'white'}; box-shadow: -4px 0 20px rgba(0,0,0,${isDarkMode ? '0.4' : '0.15'}); display: flex; flex-direction: column; animation: slideInRight 0.3s ease-out; border-radius: 12px; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};`;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `padding: 20px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0;`;
    header.innerHTML = `
      <div style="display: flex; align-items: center; gap: 12px;">
        <div style="width: 40px; height: 40px; background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 20px;">📅</div>
        <div style="flex: 1; min-width: 0;">
          <div style="font-weight: 600; font-size: 16px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Meeting Details</div>
          <div class="meeting-detail-subtitle" style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">Loading...</div>
        </div>
      </div>
      <button class="transcript-close-btn" style="background: rgba(231,76,60,0.1); border: none; color: #e74c3c; width: 36px; height: 36px; border-radius: 8px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">×</button>
    `;
    slider.appendChild(header);

    // Title section
    const titleSection = document.createElement('div');
    titleSection.style.cssText = `padding: 14px 20px; background: ${isDarkMode ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'}; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; flex-shrink: 0;`;
    titleSection.innerHTML = `<div style="font-weight: 600; font-size: 15px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(meeting.task_title || meeting.task_name || 'Meeting')}</div>`;
    slider.appendChild(titleSection);

    // Content container with loading
    const contentContainer = document.createElement('div');
    contentContainer.style.cssText = 'flex: 1; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; gap: 20px;';
    contentContainer.innerHTML = `
      <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 16px;">
        <div class="spinner" style="width: 40px; height: 40px; border: 4px solid rgba(102,126,234,0.2); border-top-color: #667eea; border-radius: 50%; animation: spin 0.8s linear infinite;"></div>
        <div style="font-size: 14px; color: #7f8c8d;">Loading meeting details...</div>
      </div>
    `;
    slider.appendChild(contentContainer);

    overlay.appendChild(slider);
    document.body.appendChild(overlay);

    // Close handler
    const closeSlider = () => {
      isTranscriptSliderOpen = false;
      slider.style.animation = 'slideOutRight 0.3s ease-out';
      document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());
      setTimeout(() => { overlay.remove(); window.Oracle.collapseCol3AfterSlider(); }, 300);
    };
    slider.querySelector('.transcript-close-btn').addEventListener('click', closeSlider);
    // V37: Click outside disabled — only Escape closes
    document.addEventListener('keydown', function escHandler(e) {
      if (e.key === 'Escape') { closeSlider(); document.removeEventListener('keydown', escHandler); }
    });

    // Fetch meeting details from n8n
    try {
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'fetch_task_details',
          todo_id: meetingId,
          message_link: meeting.message_link,
          timestamp: new Date().toISOString(),
          source: 'oracle-chrome-extension-newtab'
        }))
      });

      if (!response.ok) throw new Error('Failed to fetch meeting details');
      const responseText = await response.text();
      let data = {};
      if (responseText && responseText.trim()) {
        try { data = JSON.parse(responseText); } catch (e) { data = {}; }
      }

      const responseData = Array.isArray(data) ? data[0] : data;
      const meetingDetails = responseData?.meeting_details || responseData || {};

      // Update subtitle
      const subtitleEl = slider.querySelector('.meeting-detail-subtitle');
      if (subtitleEl) subtitleEl.textContent = meetingDetails.date_time ? 'Scheduled' : 'Details loaded';

      // Build detail content
      let detailHtml = '';

      // 1. Date & Time
      const dateTime = meetingDetails.date_time || meeting.due_by;
      if (dateTime) {
        const dt = new Date(dateTime);
        const dateStr = dt.toLocaleDateString([], { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
        const timeStr = dt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const endTime = meetingDetails.end_time ? new Date(meetingDetails.end_time).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
        detailHtml += `
          <div style="display: flex; align-items: flex-start; gap: 12px;">
            <div style="width: 36px; height: 36px; background: rgba(102,126,234,0.1); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">🕐</div>
            <div>
              <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${dateStr}</div>
              <div style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'}; margin-top: 2px;">${timeStr}${endTime ? ' – ' + endTime : ''}</div>
            </div>
          </div>
        `;
      }

      // 2. Conference Link
      const confLink = meetingDetails.conference_link || meeting.message_link || '';
      if (confLink) {
        let confLabel = 'Join Meeting';
        let confIconHtml = '<div style="width: 36px; height: 36px; background: rgba(46,204,113,0.1); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">🔗</div>';
        try {
          if (confLink.includes('zoom')) {
            confLabel = 'Join Zoom Meeting';
            const iconUrl = chrome.runtime.getURL('icon-zoom.png');
            confIconHtml = `<div style="width: 36px; height: 36px; background: rgba(46,204,113,0.1); border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;"><img src="${iconUrl}" style="width: 22px; height: 22px; object-fit: contain;"></div>`;
          } else if (confLink.includes('meet.google')) {
            confLabel = 'Join Google Meet';
            const iconUrl = chrome.runtime.getURL('icon-google-meet.png');
            confIconHtml = `<div style="width: 36px; height: 36px; background: rgba(46,204,113,0.1); border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;"><img src="${iconUrl}" style="width: 22px; height: 22px; object-fit: contain;"></div>`;
          } else if (confLink.includes('teams')) {
            confLabel = 'Join Teams Meeting';
          }
        } catch (e) { /* chrome.runtime may not be available */ }
        detailHtml += `
          <div style="display: flex; align-items: flex-start; gap: 12px;">
            ${confIconHtml}
            <div style="flex: 1; min-width: 0;">
              <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Conference</div>
              <a href="${confLink}" target="_blank" style="font-size: 12px; color: #667eea; text-decoration: underline; word-break: break-all; display: block; margin-top: 2px;">${confLabel}</a>
            </div>
          </div>
        `;
      }

      // 3. Meeting Room
      const room = meetingDetails.meeting_room || meetingDetails.location;
      if (room) {
        detailHtml += `
          <div style="display: flex; align-items: flex-start; gap: 12px;">
            <div style="width: 36px; height: 36px; background: rgba(241,196,15,0.1); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">📍</div>
            <div>
              <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Location / Room</div>
              <div style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'}; margin-top: 2px;">${escapeHtml(room)}</div>
            </div>
          </div>
        `;
      }

      // 4. Participants
      const participants = meetingDetails.participants || [];
      if (participants.length > 0) {
        const participantRows = participants.map(p => {
          const rawName = p.name || p.email?.split('@')[0] || 'Unknown';
          const name = formatDisplayName(rawName);
          const email = p.email || '';
          const status = p.status || p.response_status || 'unknown';
          const isOrganiser = p.is_organiser || p.organizer || false;
          let statusIcon = '❓';
          let statusColor = isDarkMode ? '#888' : '#95a5a6';
          let statusLabel = status;
          if (status === 'accepted' || status === 'yes') { statusIcon = '✅'; statusColor = '#27ae60'; statusLabel = 'Accepted'; }
          else if (status === 'declined' || status === 'no') { statusIcon = '❌'; statusColor = '#e74c3c'; statusLabel = 'Declined'; }
          else if (status === 'tentative' || status === 'maybe') { statusIcon = '❔'; statusColor = '#f39c12'; statusLabel = 'Tentative'; }
          else if (status === 'needsAction' || status === 'pending' || status === 'awaiting') { statusIcon = '⏳'; statusColor = isDarkMode ? '#888' : '#95a5a6'; statusLabel = 'Pending'; }
          const initial = name.charAt(0).toUpperCase();
          return `
            <div style="display: flex; align-items: center; gap: 10px; padding: 8px 0; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(225,232,237,0.4)'};">
              <div style="width: 28px; height: 28px; background: linear-gradient(135deg, #667eea, #764ba2); border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: 11px; font-weight: 600; flex-shrink: 0;">${initial}</div>
              <div style="flex: 1; min-width: 0;">
                <div style="font-weight: 600; font-size: 12px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(name)}${isOrganiser ? ' <span style="font-size: 10px; color: #667eea; font-weight: 500;">(Organiser)</span>' : ''}</div>
                ${email ? `<div style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'}; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escapeHtml(email)}</div>` : ''}
              </div>
              <div style="display: flex; align-items: center; gap: 4px; flex-shrink: 0;">
                <span style="font-size: 12px;">${statusIcon}</span>
                <span style="font-size: 11px; color: ${statusColor}; font-weight: 500;">${statusLabel}</span>
              </div>
            </div>
          `;
        }).join('');

        detailHtml += `
          <div style="display: flex; align-items: flex-start; gap: 12px;">
            <div style="width: 36px; height: 36px; background: rgba(155,89,182,0.1); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">👥</div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; margin-bottom: 8px;">Participants (${participants.length})</div>
              <div>${participantRows}</div>
            </div>
          </div>
        `;
      }

      if (!detailHtml) {
        detailHtml = `<div style="text-align: center; padding: 40px 20px; color: ${isDarkMode ? '#888' : '#95a5a6'};">
          <div style="font-size: 32px; margin-bottom: 12px;">📅</div>
          <div style="font-size: 14px;">No additional meeting details available.</div>
        </div>`;
      }

      contentContainer.innerHTML = detailHtml;

    } catch (error) {
      console.error('Error fetching meeting details:', error);
      contentContainer.innerHTML = `<div style="text-align: center; padding: 40px 20px; color: #e74c3c;">
        <div style="font-size: 32px; margin-bottom: 12px;">⚠️</div>
        <div style="font-size: 14px;">Failed to load meeting details</div>
        <div style="font-size: 12px; margin-top: 4px; color: ${isDarkMode ? '#888' : '#95a5a6'};">${escapeHtml(error.message)}</div>
      </div>`;
    }
  }

  // Helper function to build individual meeting item HTML
  function buildMeetingItemHtml(meeting, googleCalendarIconUrl, zoomIconUrl, googleMeetIconUrl, isUnread = false) {
    const messageLink = meeting.message_link || '';
    const secondaryLinks = meeting.secondary_links || [];

    // Determine primary button icon and title based on message_link
    let primaryIconUrl = '';
    let primaryTitle = 'Open';
    let primaryLink = messageLink;

    if (messageLink.includes('zoom.us') || messageLink.includes('zoom.com')) {
      primaryIconUrl = zoomIconUrl;
      primaryTitle = 'Join';
      primaryLink = convertZoomToDeepLink(messageLink);
    } else if (messageLink.includes('meet.google.com')) {
      primaryIconUrl = googleMeetIconUrl;
      primaryTitle = 'Join';
    } else if (messageLink.includes('teams.microsoft.com') || messageLink.includes('teams.live.com')) {
      primaryTitle = 'Join';
    } else if (messageLink.includes('calendar.google.com') || messageLink.includes('google.com/calendar')) {
      primaryIconUrl = googleCalendarIconUrl;
      primaryTitle = 'Open Calendar';
    }

    // Fallback to calendar icon if no specific icon detected
    if (!primaryIconUrl) {
      primaryIconUrl = googleCalendarIconUrl;
      primaryTitle = 'Open';
    }

    // Build secondary link icons (shown to the LEFT of primary button)
    let secondaryIconsHtml = '';

    secondaryLinks.forEach(link => {
      if (!link) return;

      let secIconUrl = '';
      let secTitle = '';

      if (link.includes('meet.google.com')) {
        secIconUrl = googleMeetIconUrl;
        secTitle = 'Join Google Meet';
      } else if (link.includes('zoom.us') || link.includes('zoom.com')) {
        secIconUrl = zoomIconUrl;
        secTitle = 'Join Zoom';
      } else if (link.includes('calendar.google.com') || link.includes('google.com/calendar')) {
        // Only show calendar icon if primary is NOT calendar
        if (!messageLink.includes('calendar.google.com') && !messageLink.includes('google.com/calendar')) {
          secIconUrl = googleCalendarIconUrl;
          secTitle = 'Open in Google Calendar';
        }
      }

      if (secIconUrl) {
        const secHref = (secIconUrl === zoomIconUrl) ? convertZoomToDeepLink(link) : link;
        secondaryIconsHtml += `
          <a href="${secHref}" target="_blank" class="meeting-secondary-btn" title="${secTitle}" style="display: flex; align-items: center; justify-content: center; width: 32px; height: 32px; background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.2); border-radius: 6px; transition: all 0.2s;">
            <img src="${secIconUrl}" alt="${secTitle}" style="width: 18px; height: 18px;">
          </a>
        `;
      }
    });

    const timeText = formatMeetingTime(meeting.due_by);
    const isOverdue = meeting.due_by && new Date(meeting.due_by) < new Date();

    const title = meeting.task_title || meeting.task_name || 'Meeting';
    const truncatedTitle = title.length > 60 ? title.substring(0, 60) + '...' : title;

    // Extract organiser from participant_text (first name listed)
    const organiserName = extractOrganiser(meeting.participant_text);
    const timeWithOrganiser = timeText ? (organiserName ? `${timeText} | ${organiserName}` : timeText) : '';

    return `
      <div class="meeting-item ${isUnread ? 'unread' : ''}" data-meeting-id="${meeting.id}" data-todo-id="${meeting.id}">
        <div class="meeting-checkbox" data-todo-id="${meeting.id}" title="Mark as done"></div>
        <div class="meeting-info">
          <div class="meeting-title" title="${escapeHtml(title)}">${escapeHtml(truncatedTitle)}</div>
          ${timeWithOrganiser ? `<div class="meeting-time ${isOverdue ? 'overdue' : ''}">${escapeHtml(timeWithOrganiser)}</div>` : ''}
        </div>
        <div class="meeting-actions" style="display: flex; align-items: center; gap: 8px;">
          ${secondaryIconsHtml}
          <a href="${primaryLink}" target="_blank" class="meeting-join-btn meeting-mark-read" title="${primaryTitle}" style="display: flex; align-items: center; justify-content: center; width: 32px; height: 32px; background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.2); border-radius: 6px; transition: all 0.2s;">
            <img src="${primaryIconUrl}" alt="${primaryTitle}" style="width: 18px; height: 18px;">
          </a>
        </div>
      </div>
    `;
  }

  function setupMeetingsAccordion(container) {
    const accordion = container.querySelector('.meetings-accordion');
    if (!accordion) return;

    const header = accordion.querySelector('.meetings-accordion-header');
    const content = accordion.querySelector('.meetings-accordion-content');

    // Toggle accordion on header click
    header.addEventListener('click', () => {
      accordion.classList.toggle('open');

      // When accordion is opened, remove the has-updates class from accordion title
      if (accordion.classList.contains('open')) {
        accordion.classList.remove('has-updates');
      }
    });

    // Setup date tab switching
    const dateTabs = accordion.querySelectorAll('.meeting-date-tab');
    const dateContents = accordion.querySelectorAll('.meetings-date-content');

    dateTabs.forEach(tab => {
      tab.addEventListener('click', (e) => {
        e.stopPropagation();
        const dateKey = tab.dataset.dateKey;

        // Update active tab
        dateTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        // Show corresponding content
        dateContents.forEach(content => {
          if (content.dataset.dateKey === dateKey) {
            content.classList.add('active');
          } else {
            content.classList.remove('active');
          }
        });

        // Remove has-updates class from this date tab when clicked
        tab.classList.remove('has-updates');
        modifiedMeetingDates.delete(dateKey);

        // If no more modified dates, remove has-updates from accordion
        if (modifiedMeetingDates.size === 0) {
          accordion.classList.remove('has-updates');
        }
      });
    });

    // Setup click on date content area to clear updates for that date
    dateContents.forEach(content => {
      content.addEventListener('click', (e) => {
        const dateKey = content.dataset.dateKey;
        const dateTab = accordion.querySelector(`.meeting-date-tab[data-date-key="${dateKey}"]`);
        if (dateTab) {
          dateTab.classList.remove('has-updates');
          modifiedMeetingDates.delete(dateKey);

          // If no more modified dates, remove has-updates from accordion
          if (modifiedMeetingDates.size === 0) {
            accordion.classList.remove('has-updates');
          }
        }
      });
    });

    // Setup click handlers for meeting items to open detail slider
    accordion.querySelectorAll('.meeting-item').forEach(meetingItem => {
      // Click on the meeting item opens meeting detail slider
      meetingItem.addEventListener('click', (e) => {
        // Don't trigger if clicking checkbox, links, or join buttons
        if (e.target.closest('.meeting-checkbox')) return;
        if (e.target.closest('a')) return;
        if (e.target.closest('.meeting-join-btn')) return;
        if (e.target.closest('.meeting-secondary-btn')) return;
        meetingItem.classList.remove('unread');
        const meetingId = meetingItem.dataset.meetingId || meetingItem.dataset.todoId;
        if (meetingId) showMeetingDetailSlider(meetingId);
      });

      // Click on calendar/join button marks as read
      const joinBtn = meetingItem.querySelector('.meeting-join-btn');
      if (joinBtn) {
        joinBtn.addEventListener('click', () => {
          meetingItem.classList.remove('unread');
        });
      }
    });

    // Meeting checkbox click - mark as done (with multi-select support)
    accordion.querySelectorAll('.meeting-checkbox').forEach(checkbox => {
      checkbox.addEventListener('click', async (e) => {
        e.stopPropagation();
        const todoId = checkbox.dataset.todoId;
        const meetingItem = checkbox.closest('.meeting-item');
        const dateContent = meetingItem.closest('.meetings-date-content');
        const dateKey = dateContent ? dateContent.dataset.dateKey : null;

        // Check if Command/Ctrl key is pressed for multi-select
        if (e.metaKey || e.ctrlKey) {
          isMultiSelectMode = true;
          toggleTodoSelection(todoId, meetingItem);
        } else if (selectedTodoIds.size > 0) {
          // If already in multi-select mode, continue selecting
          toggleTodoSelection(todoId, meetingItem);
        } else {
          // Single item mark as done
          // Immediate visual feedback with animation
          checkbox.innerHTML = '✓';
          checkbox.style.background = 'linear-gradient(45deg, #667eea, #764ba2)';
          checkbox.style.borderColor = 'transparent';
          checkbox.style.color = 'white';
          meetingItem.classList.add('completing');

          // Start API call in background (don't wait for it)
          updateTodoField(todoId, 'status', 1).catch(error => {
            console.error('Error marking meeting done:', error);
          });

          // Remove item after animation completes
          setTimeout(() => {
            meetingItem.remove();

            // Update count for this date tab
            if (dateKey) {
              const dateTab = accordion.querySelector(`.meeting-date-tab[data-date-key="${dateKey}"]`);
              if (dateTab) {
                const countEl = dateTab.querySelector('.meeting-date-count');
                if (countEl) {
                  const currentCount = parseInt(countEl.textContent) - 1;
                  countEl.textContent = currentCount;

                  // Remove date tab and content if no meetings left for this date
                  if (currentCount === 0) {
                    dateTab.remove();
                    if (dateContent) dateContent.remove();

                    // Activate next available date tab
                    const remainingTabs = accordion.querySelectorAll('.meeting-date-tab');
                    const remainingContents = accordion.querySelectorAll('.meetings-date-content');
                    if (remainingTabs.length > 0) {
                      remainingTabs[0].classList.add('active');
                      if (remainingContents.length > 0) {
                        remainingContents[0].classList.add('active');
                      }
                    }
                  }
                }
              }
            }

            // Update total count
            const totalCountEl = accordion.querySelector('.meetings-accordion-count');
            if (totalCountEl) {
              const currentTotal = parseInt(totalCountEl.textContent) - 1;
              totalCountEl.textContent = currentTotal;

              // Remove accordion if no meetings left
              if (currentTotal === 0) {
                accordion.remove();
              }
            }

            updateTabCounts();
          }, 400);
        }
      });
    });
  }

  // Build Documents Accordion for FYI tab
  function buildDocumentsAccordion(driveTasks, excludeFileIds = new Set()) {
    if (!driveTasks || driveTasks.length === 0) return '';

    // Group by file ID and exclude files that have groups in Action tab
    const { driveGroups } = groupTasksByDriveFile(driveTasks);

    // Filter out files that exist in Action tab
    const filteredGroups = Object.values(driveGroups).filter(group => !excludeFileIds.has(group.fileId));

    if (filteredGroups.length === 0) return '';

    // Sort: groups with unread tasks first, then by latest update
    filteredGroups.sort((a, b) => {
      const aHasUnread = a.tasks.some(task => isTaskUnread(task.id));
      const bHasUnread = b.tasks.some(task => isTaskUnread(task.id));
      if (aHasUnread && !bHasUnread) return -1;
      if (!aHasUnread && bHasUnread) return 1;
      return new Date(b.latestUpdate) - new Date(a.latestUpdate);
    });

    const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
    const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
    const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
    const driveIconUrl = chrome.runtime.getURL('icon-drive.png');

    const documentItems = filteredGroups.map(group => {
      const fileUrl = group.fileUrl;
      let iconUrl = driveIconUrl;
      let docType = 'Document';

      if (fileUrl.includes('docs.google.com/document')) {
        iconUrl = googleDocsIconUrl;
        docType = 'Doc';
      } else if (fileUrl.includes('docs.google.com/spreadsheets') || fileUrl.includes('sheets.google.com')) {
        iconUrl = googleSheetsIconUrl;
        docType = 'Sheet';
      } else if (fileUrl.includes('docs.google.com/presentation') || fileUrl.includes('slides.google.com')) {
        iconUrl = googleSlidesIconUrl;
        docType = 'Slides';
      }

      const taskCount = group.tasks.length;
      const title = group.groupTitle || 'Document';
      const truncatedTitle = title.length > 50 ? title.substring(0, 50) + '...' : title;

      // Documents accordion: never mark as unread (documents refresh silently)
      const hasUnreadTask = false;

      // Create a comma-separated list of task IDs for bulk mark done
      const taskIds = group.tasks.map(t => t.id).join(',');

      return `
        <div class="document-accordion-item" data-file-id="${group.fileId}" data-task-ids="${taskIds}" data-file-url="${fileUrl}">
          <div class="document-item-header">
            <div class="document-group-checkbox" data-task-ids="${taskIds}" title="Mark all tasks as done"></div>
            <img src="${iconUrl}" alt="${docType}" style="width: 20px; height: 20px; flex-shrink: 0;">
            <div class="document-item-info">
              <div class="document-item-title" title="${escapeHtml(title)}">${escapeHtml(truncatedTitle)}</div>
              <div class="document-item-meta">${taskCount} update${taskCount > 1 ? 's' : ''} · ${formatDate(group.latestUpdate)}</div>
            </div>
            <button class="document-learn-btn" data-file-url="${fileUrl}" title="Learn about this document">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"></path>
                <path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"></path>
              </svg>
            </button>
            <a href="${fileUrl}" target="_blank" class="document-open-btn" title="Open document">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path>
                <polyline points="15 3 21 3 21 9"></polyline>
                <line x1="10" y1="14" x2="21" y2="3"></line>
              </svg>
            </a>
          </div>
          <div class="document-tasks-list" style="display: none;">
            ${group.tasks.map(task => {
        return `
              <div class="document-task-item todo-item" data-todo-id="${task.id}">
                <div class="document-task-checkbox" data-todo-id="${task.id}" title="Mark as done"></div>
                <div class="document-task-content">
                  <div class="document-task-text">${escapeHtml(task.task_name?.substring(0, 100) || '')}${task.task_name?.length > 100 ? '...' : ''}</div>
                  <div class="document-task-meta">${formatDate(task.updated_at || task.created_at)}</div>
                </div>
              </div>
            `}).join('')}
          </div>
        </div>
      `;
    }).join('');

    const totalDocs = filteredGroups.length;

    return `
      <div class="documents-accordion">
        <div class="documents-accordion-header">
          <div class="documents-accordion-title">
            <span>📄</span>
            <span>Documents</span>
            <span class="documents-accordion-count">${totalDocs}</span>
          </div>
          <span class="documents-accordion-chevron">▼</span>
        </div>
        <div class="documents-accordion-content">
          <div class="documents-accordion-list">
            ${documentItems}
          </div>
        </div>
      </div>
    `;
  }

  function setupDocumentsAccordion(container) {
    const accordion = container.querySelector('.documents-accordion');
    if (!accordion) return;

    const header = accordion.querySelector('.documents-accordion-header');

    // Toggle main accordion on header click
    header.addEventListener('click', () => {
      accordion.classList.toggle('open');
    });

    // Toggle individual document items
    accordion.querySelectorAll('.document-accordion-item').forEach(item => {
      const itemHeader = item.querySelector('.document-item-header');
      const tasksList = item.querySelector('.document-tasks-list');
      const groupCheckbox = item.querySelector('.document-group-checkbox');

      // Group checkbox click - mark all tasks in this document as done
      if (groupCheckbox) {
        groupCheckbox.addEventListener('click', async (e) => {
          e.stopPropagation();
          const taskIds = groupCheckbox.dataset.taskIds?.split(',').filter(id => id) || [];
          if (taskIds.length === 0) return;

          // Visual feedback
          groupCheckbox.classList.add('checked');
          groupCheckbox.innerHTML = '✓';
          item.classList.add('completing');

          // Mark all tasks as done in background
          const promises = taskIds.map(id =>
            updateTodoField(id, 'status', 1).catch(err => console.error('Error marking task done:', err))
          );

          // Don't wait - start removal animation immediately
          Promise.all(promises);

          // Remove the entire document item after animation
          setTimeout(() => {
            item.remove();
            // Update accordion count
            const countEl = accordion.querySelector('.documents-accordion-count');
            const currentCount = parseInt(countEl.textContent) - 1;
            countEl.textContent = currentCount;
            if (currentCount === 0) {
              accordion.remove();
            }
            updateTabCounts();
          }, 400);
        });
      }

      itemHeader.addEventListener('click', (e) => {
        // Don't toggle if clicking the open button, learn button, or checkbox
        if (e.target.closest('.document-open-btn')) return;
        if (e.target.closest('.document-learn-btn')) return;
        if (e.target.closest('.document-group-checkbox')) return;

        e.stopPropagation();
        item.classList.toggle('expanded');
        tasksList.style.display = item.classList.contains('expanded') ? 'block' : 'none';

        // When expanding, mark all tasks in this group as read
        if (item.classList.contains('expanded')) {
          const taskIds = item.dataset.taskIds?.split(',').filter(id => id) || [];
          taskIds.forEach(id => markTaskAsRead(id));
          item.classList.remove('unread');
        }
      });

      // Learn button click - send file URL to webhook
      const learnBtn = item.querySelector('.document-learn-btn');
      if (learnBtn) {
        learnBtn.addEventListener('click', async (e) => {
          e.stopPropagation();
          const fileUrl = learnBtn.dataset.fileUrl;
          if (!fileUrl) return;

          // Visual feedback - show loading state
          const originalContent = learnBtn.innerHTML;
          learnBtn.innerHTML = '<div class="spinner" style="width:12px;height:12px;border:2px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;"></div>';
          learnBtn.disabled = true;

          try {
            const response = await fetch(WEBHOOK_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(createAuthenticatedPayload({
                action: 'learn_document',
                file_url: fileUrl,
                timestamp: new Date().toISOString()
              }))
            });

            if (response.ok) {
              // Show success feedback
              learnBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#27ae60" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>';
              setTimeout(() => {
                learnBtn.innerHTML = originalContent;
                learnBtn.disabled = false;
              }, 2000);
              showToastNotification('Document sent for learning!');
            } else {
              throw new Error('Request failed');
            }
          } catch (error) {
            console.error('Error sending document for learning:', error);
            learnBtn.innerHTML = originalContent;
            learnBtn.disabled = false;
            showToastNotification('Failed to send document');
          }
        });
      }
    });

    // Document task checkbox click - mark as done
    accordion.querySelectorAll('.document-task-checkbox').forEach(checkbox => {
      checkbox.addEventListener('click', async (e) => {
        e.stopPropagation();
        const todoId = checkbox.dataset.todoId;
        const taskItem = checkbox.closest('.document-task-item');
        const docItem = checkbox.closest('.document-accordion-item');

        // Immediate visual feedback with animation
        checkbox.innerHTML = '✓';
        checkbox.classList.add('checked');
        taskItem.classList.add('completing');

        // Start API call in background (don't wait for it)
        updateTodoField(todoId, 'status', 1).catch(error => {
          console.error('Error marking document task done:', error);
        });

        // Remove item after animation completes
        setTimeout(() => {
          taskItem.remove();
          // Update count or remove document item if empty
          const remainingTasks = docItem.querySelectorAll('.document-task-item').length;
          if (remainingTasks === 0) {
            docItem.remove();
            // Update accordion count
            const countEl = accordion.querySelector('.documents-accordion-count');
            const currentCount = parseInt(countEl.textContent) - 1;
            countEl.textContent = currentCount;
            if (currentCount === 0) {
              accordion.remove();
            }
          } else {
            // Update the meta text
            const metaEl = docItem.querySelector('.document-item-meta');
            if (metaEl) {
              metaEl.textContent = `${remainingTasks} update${remainingTasks > 1 ? 's' : ''} · ${metaEl.textContent.split('·')[1]?.trim() || ''}`;
            }
          }
          updateTabCounts();
        }, 400);
      });
    });

    // Make document task items clickable to open transcript
    accordion.querySelectorAll('.document-task-item').forEach(taskItem => {
      taskItem.addEventListener('click', (e) => {
        // Don't open transcript if clicking checkbox
        if (e.target.closest('.document-task-checkbox')) return;
        e.stopPropagation();
        const todoId = taskItem.dataset.todoId;
        if (todoId) {
          showTranscriptSlider(todoId);
        }
      });
    });
  }

  // Build task group HTML for Action tab (multiple tasks under same document)
  function buildTaskGroupHtml(group, newItemIds = []) {
    const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
    const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
    const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
    const driveIconUrl = chrome.runtime.getURL('icon-drive.png');

    const fileUrl = group.fileUrl;
    let iconUrl = driveIconUrl;
    let docType = 'Document';

    if (fileUrl.includes('docs.google.com/document')) {
      iconUrl = googleDocsIconUrl;
      docType = 'Doc';
    } else if (fileUrl.includes('docs.google.com/spreadsheets') || fileUrl.includes('sheets.google.com')) {
      iconUrl = googleSheetsIconUrl;
      docType = 'Sheet';
    } else if (fileUrl.includes('docs.google.com/presentation') || fileUrl.includes('slides.google.com')) {
      iconUrl = googleSlidesIconUrl;
      docType = 'Slides';
    }

    const taskCount = group.tasks.length;
    const title = group.groupTitle || 'Document';
    const hasUnread = group.tasks.some(t => isTaskUnread(t.id));

    // Sort tasks by date
    const sortedTasks = [...group.tasks].sort((a, b) =>
      new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at)
    );

    // Create a comma-separated list of task IDs for bulk mark done
    const taskIds = sortedTasks.map(t => t.id).join(',');

    // Get icon URLs for secondary links
    const slackIconUrlLocal = chrome.runtime.getURL('icon-slack.png');
    const gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
    const freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
    const freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
    const freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
    const linkIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

    // Helper to get icon for a link (same as Slack group)
    const getIconForLink = (link) => {
      if (link.includes('slack.com') || link.includes('app.slack.com')) {
        return { icon: '<img src="' + slackIconUrlLocal + '" alt="Slack" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Slack' };
      } else if (link.includes('freshrelease.com')) {
        return { icon: '<img src="' + freshreleaseIconUrl + '" alt="Freshrelease" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshrelease' };
      } else if (link.includes('freshdesk.com')) {
        return { icon: '<img src="' + freshdeskIconUrl + '" alt="Freshdesk" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshdesk' };
      } else if (link.includes('freshservice.com')) {
        return { icon: '<img src="' + freshserviceIconUrl + '" alt="Freshservice" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshservice' };
      } else if (link.includes('mail.google.com')) {
        return { icon: '<img src="' + gmailIconUrl + '" alt="Gmail" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Gmail' };
      } else if (link.includes('docs.google.com/document')) {
        return { icon: '<img src="' + googleDocsIconUrl + '" alt="Google Docs" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Docs' };
      } else if (link.includes('docs.google.com/spreadsheets') || link.includes('sheets.google.com')) {
        return { icon: '<img src="' + googleSheetsIconUrl + '" alt="Google Sheets" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Sheets' };
      } else if (link.includes('docs.google.com/presentation') || link.includes('slides.google.com')) {
        return { icon: '<img src="' + googleSlidesIconUrl + '" alt="Google Slides" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Slides' };
      } else if (link.includes('drive.google.com')) {
        return { icon: '<img src="' + driveIconUrl + '" alt="Google Drive" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Drive' };
      }
      return { icon: linkIcon, title: 'View link' };
    };

    return `
      <div class="task-group ${hasUnread ? 'unread' : ''}" data-file-id="${group.fileId}" data-task-ids="${taskIds}" data-file-url="${fileUrl}">
        <div class="task-group-header">
          <div class="task-group-checkbox" data-task-ids="${taskIds}" title="Mark all tasks as done"></div>
          <div class="task-group-icon">
            <img src="${iconUrl}" alt="${docType}" style="width: 24px; height: 24px;">
          </div>
          <div class="task-group-info">
            <div class="task-group-title" title="${escapeHtml(title)}">${escapeHtml(title.length > 60 ? title.substring(0, 60) + '...' : title)}</div>
            <div class="task-group-meta">
              <span class="task-group-count">${taskCount} action item${taskCount > 1 ? 's' : ''}</span>
              <span class="task-group-date">${formatDate(group.latestUpdate)}</span>
            </div>
          </div>
          <div class="task-group-actions">
            <button class="task-group-learn-btn" data-file-url="${fileUrl}" title="Learn about this document">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"></path>
                <path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"></path>
              </svg>
            </button>
            <a href="${fileUrl}" target="_blank" class="task-group-open-btn" title="Open document">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path>
                <polyline points="15 3 21 3 21 9"></polyline>
                <line x1="10" y1="14" x2="21" y2="3"></line>
              </svg>
            </a>
            <span class="task-group-chevron">▼</span>
          </div>
        </div>
        <div class="task-group-tasks" style="display: none;">
          ${sortedTasks.map((task, index) => {
      const isUnread = isTaskUnread(task.id);
      const isNew = newItemIds.includes(task.id);
      const taskMessageLink = task.message_link || '';
      const secondaryLinks = task.secondary_links || [];

      // Handle long descriptions (same as Slack group)
      const taskName = task.task_name || '';
      const taskNameEscaped = escapeHtml(taskName);
      const maxLength = 150;
      const hasTaskTitle = !!task.task_title; const needsViewMore = taskName.length > maxLength || (hasTaskTitle && taskName.length > 60);
      let taskNameHtml;

      if (needsViewMore) {
        const truncated = escapeHtml(taskName.substring(0, maxLength));
        taskNameHtml = '<span class="todo-text-content" data-full-text="' + taskNameEscaped + '">' + taskNameEscaped + '</span><span class="view-more-inline" data-todo-id="' + task.id + '">View more</span>';
      } else {
        taskNameHtml = '<span class="todo-text-content">' + taskNameEscaped + '</span>';
      }

      // Render tags (filter out invalid tags like "null")
      const todoTags = (task.tags || []).filter(isValidTag);
      const tagsHtml = todoTags.length > 0 ? '<div class="todo-tags" style="display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px;">' + todoTags.map(tag => {
        const isActive = activeTagFilters.includes(tag);
        return '<span class="todo-tag ' + (isActive ? 'active' : '') + '" data-tag="' + escapeHtml(tag) + '" style="display: inline-block; background: ' + (isActive ? 'linear-gradient(45deg, #667eea, #764ba2)' : 'rgba(102, 126, 234, 0.1)') + '; color: ' + (isActive ? 'white' : '#667eea') + '; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 500; cursor: pointer; border: 1px solid ' + (isActive ? 'transparent' : 'rgba(102, 126, 234, 0.3)') + '; transition: all 0.2s;">' + escapeHtml(tag) + '</span>';
      }).join('') + '</div>' : '';

      // Generate primary source icon
      let sourceIconHtml = '';
      if (taskMessageLink) {
        const primaryIcon = getIconForLink(taskMessageLink);
        sourceIconHtml = '<a href="' + taskMessageLink + '" target="_blank" class="todo-source" title="' + primaryIcon.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + primaryIcon.icon + '</a>';
      }

      // Generate secondary links HTML
      const secondaryLinksHtml = secondaryLinks.filter(l => l != null).map(link => {
        const iconData = getIconForLink(link);
        return '<span style="color: #bdc3c7; margin: 0 4px;">|</span><a href="' + link + '" target="_blank" class="todo-source" title="' + iconData.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + iconData.icon + '</a>';
      }).join('');

      return `
              <div class="task-group-task-item todo-item ${isUnread ? 'unread' : ''} ${isNew ? 'appearing' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                <div class="todo-left-actions">
                  <div class="todo-checkbox" data-todo-id="${task.id}"></div>
                </div>
                <div class="todo-content" style="flex: 1; min-width: 0;">
                  ${task.task_title ? `<div class="todo-title">${escapeHtml(task.task_title)}</div>` : ''}
                  <div class="todo-text${needsViewMore ? ' truncated' : ''}">${taskNameHtml}</div>
                  <div class="todo-meta">
                    <span class="todo-date">${formatDate(task.updated_at || task.created_at)}</span>
                    ${sourceIconHtml}
                    ${secondaryLinksHtml}
                  </div>
                  ${tagsHtml}
                </div>
              </div>
            `;
    }).join('')}
        </div>
      </div>
    `;
  }

  function setupTaskGroups(container) {
    container.querySelectorAll('.task-group:not(.slack-channel-group):not(.tag-group)').forEach(group => {
      const header = group.querySelector('.task-group-header');
      const tasksContainer = group.querySelector('.task-group-tasks');
      const chevron = group.querySelector('.task-group-chevron');
      const groupCheckbox = group.querySelector('.task-group-checkbox');

      // Group checkbox click - mark all tasks as done
      if (groupCheckbox) {
        groupCheckbox.addEventListener('click', async (e) => {
          e.stopPropagation();
          const taskIds = groupCheckbox.dataset.taskIds?.split(',').filter(id => id) || [];
          if (taskIds.length === 0) return;

          // Visual feedback
          groupCheckbox.classList.add('checked');
          groupCheckbox.innerHTML = '✓';
          group.classList.add('completing');

          // Mark all tasks as done in background
          const promises = taskIds.map(id =>
            updateTodoField(id, 'status', 1).catch(err => console.error('Error marking task done:', err))
          );

          // Don't wait - start removal animation immediately
          Promise.all(promises);

          // Remove the entire group after animation
          setTimeout(() => {
            group.remove();
            updateTabCounts();
          }, 400);
        });
      }

      header.addEventListener('click', (e) => {
        // Don't toggle if clicking the open button, learn button, or checkbox
        if (e.target.closest('.task-group-open-btn')) return;
        if (e.target.closest('.task-group-learn-btn')) return;
        if (e.target.closest('.task-group-checkbox')) return;

        group.classList.toggle('expanded');
        tasksContainer.style.display = group.classList.contains('expanded') ? 'block' : 'none';
        chevron.style.transform = group.classList.contains('expanded') ? 'rotate(180deg)' : '';

        // When expanding, mark all tasks in this group as read
        if (group.classList.contains('expanded')) {
          const taskIds = group.dataset.taskIds?.split(',').filter(id => id) || [];
          taskIds.forEach(id => markTaskAsRead(id));
          group.classList.remove('unread');
        }
      });

      // Learn button click - send file URL to webhook
      const learnBtn = group.querySelector('.task-group-learn-btn');
      if (learnBtn) {
        learnBtn.addEventListener('click', async (e) => {
          e.stopPropagation();
          const fileUrl = learnBtn.dataset.fileUrl;
          if (!fileUrl) return;

          // Visual feedback - show loading state
          const originalContent = learnBtn.innerHTML;
          learnBtn.innerHTML = '<div class="spinner" style="width:12px;height:12px;border:2px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;"></div>';
          learnBtn.disabled = true;

          try {
            const response = await fetch(WEBHOOK_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(createAuthenticatedPayload({
                action: 'learn_document',
                file_url: fileUrl,
                timestamp: new Date().toISOString()
              }))
            });

            if (response.ok) {
              // Show success feedback
              learnBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#27ae60" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>';
              setTimeout(() => {
                learnBtn.innerHTML = originalContent;
                learnBtn.disabled = false;
              }, 2000);
              showToastNotification('Document sent for learning!');
            } else {
              throw new Error('Request failed');
            }
          } catch (error) {
            console.error('Error sending document for learning:', error);
            learnBtn.innerHTML = originalContent;
            learnBtn.disabled = false;
            showToastNotification('Failed to send document');
          }
        });
      }

      // Task item clicks
      group.querySelectorAll('.task-group-task-item').forEach(taskItem => {
        if (taskItem.dataset.groupHandlerAttached) return;
        taskItem.dataset.groupHandlerAttached = 'true';
        const checkbox = taskItem.querySelector('.todo-checkbox');
        const todoId = taskItem.dataset.todoId;

        // Checkbox click
        if (checkbox) {
          checkbox.addEventListener('click', async (e) => {
            e.stopPropagation();
            checkbox.classList.add('checked');
            checkbox.innerHTML = '✓';
            taskItem.classList.add('completing');

            try {
              await updateTodoField(todoId, 'status', 1);
              setTimeout(() => taskItem.remove(), 400);

              // Update group count or remove group if empty
              const remainingTasks = group.querySelectorAll('.task-group-task-item').length - 1;
              if (remainingTasks === 0) {
                group.remove();
              } else {
                const countEl = group.querySelector('.task-group-count');
                if (countEl) {
                  countEl.textContent = `${remainingTasks} action item${remainingTasks > 1 ? 's' : ''}`;
                }
              }
              updateTabCounts();
            } catch (err) {
              console.error('Error marking task done:', err);
              checkbox.classList.remove('checked');
              checkbox.innerHTML = '';
              taskItem.classList.remove('completing');
            }
          });
        }

        // Task item click (open transcript)
        taskItem.addEventListener('click', (e) => {
          if (e.target.closest('.todo-checkbox')) return;
          if (e.target.closest('.todo-source')) return; // Don't open transcript when clicking source link
          if (e.target.closest('a')) return; // Don't open transcript when clicking any link
          showTranscriptSlider(todoId);
        });

        // Source link click - stop propagation to prevent slider from opening
        taskItem.querySelectorAll('.todo-source').forEach(sourceLink => {
          sourceLink.addEventListener('click', (e) => {
            e.stopPropagation();
          });
        });
      });
    });
  }

  // Build Slack channel group HTML for Action tab (multiple tasks under same channel)
  function buildSlackChannelGroupHtml(group, newItemIds = []) {
    const slackIconUrl = chrome.runtime.getURL('icon-slack.png');

    const taskCount = group.tasks.length;
    const channelName = group.channelName || 'Slack Channel';
    const hasUnread = group.tasks.some(t => isTaskUnread(t.id));

    // Sort tasks by date
    const sortedTasks = [...group.tasks].sort((a, b) =>
      new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at)
    );

    // Get the first task's message link for the channel link
    const channelLink = sortedTasks[0]?.message_link || '';
    const slackChannelUrl = getSlackChannelUrl(channelLink);

    // Create a comma-separated list of task IDs for bulk mark done
    const taskIds = sortedTasks.map(t => t.id).join(',');

    return `
      <div class="slack-channel-group task-group ${hasUnread ? 'unread' : ''}" data-channel-name="${escapeHtml(channelName)}" data-task-ids="${taskIds}">
        <div class="task-group-header">
          <div class="task-group-checkbox" data-task-ids="${taskIds}" title="Mark all tasks as done"></div>
          <div class="task-group-icon">
            ${slackChannelUrl
        ? `<a href="${escapeHtml(slackChannelUrl)}" target="_blank" title="Open Slack channel" style="display: inline-flex;"><img src="${slackIconUrl}" alt="Slack" style="width: 24px; height: 24px; cursor: pointer;"></a>`
        : `<img src="${slackIconUrl}" alt="Slack" style="width: 24px; height: 24px;">`
      }
          </div>
          <div class="task-group-info">
            <div class="task-group-title" title="${escapeHtml(channelName)}">${slackChannelUrl
        ? `<a href="${escapeHtml(slackChannelUrl)}" target="_blank" style="color: inherit; text-decoration: none;">${escapeHtml(channelName.length > 60 ? channelName.substring(0, 60) + '...' : channelName)}</a>`
        : escapeHtml(channelName.length > 60 ? channelName.substring(0, 60) + '...' : channelName)}</div>
            <div class="task-group-meta">
              <span class="task-group-count">${taskCount} action item${taskCount > 1 ? 's' : ''}</span>
              <span class="task-group-date">${formatDate(group.latestUpdate)}</span>
            </div>
          </div>
          <div class="task-group-actions">
            <span class="task-group-chevron">▼</span>
          </div>
        </div>
        <div class="task-group-tasks" style="display: none;">
          ${sortedTasks.map((task, index) => {
          const isUnread = isTaskUnread(task.id);
          const isNew = newItemIds.includes(task.id);
          const taskMessageLink = task.message_link || '';
          const secondaryLinks = task.secondary_links || [];

          // Handle long descriptions
          const taskName = task.task_name || '';
          const taskNameEscaped = escapeHtml(taskName);
          const maxLength = 150;
          const hasTaskTitle = !!task.task_title; const needsViewMore = taskName.length > maxLength || (hasTaskTitle && taskName.length > 60);
          let taskNameHtml;

          if (needsViewMore) {
            const truncated = escapeHtml(taskName.substring(0, maxLength));
            taskNameHtml = `<span class="todo-text-content" data-full-text="${taskNameEscaped}">${taskNameEscaped}</span><span class="view-more-inline" data-todo-id="${task.id}">View more</span>`;
          } else {
            taskNameHtml = `<span class="todo-text-content">${taskNameEscaped}</span>`;
          }

          // Render tags (filter out invalid tags like "null")
          const todoTags = (task.tags || []).filter(isValidTag);
          const tagsHtml = todoTags.length > 0 ? `
              <div class="todo-tags" style="display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px;">
                ${todoTags.map(tag => {
            const isActive = activeTagFilters.includes(tag);
            return `<span class="todo-tag ${isActive ? 'active' : ''}" data-tag="${escapeHtml(tag)}" style="display: inline-block; background: ${isActive ? 'linear-gradient(45deg, #667eea, #764ba2)' : 'rgba(102, 126, 234, 0.1)'}; color: ${isActive ? 'white' : '#667eea'}; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 500; cursor: pointer; border: 1px solid ${isActive ? 'transparent' : 'rgba(102, 126, 234, 0.3)'}; transition: all 0.2s;">${escapeHtml(tag)}</span>`;
          }).join('')}
              </div>
            ` : '';

          // Get all icon URLs
          const slackIconUrlLocal = chrome.runtime.getURL('icon-slack.png');
          const gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
          const driveIconUrl = chrome.runtime.getURL('icon-drive.png');
          const freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
          const freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
          const freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
          const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
          const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
          const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
          const linkIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

          // Helper to get icon for a link
          const getIconForLink = (link) => {
            if (link.includes('slack.com') || link.includes('app.slack.com')) {
              return { icon: '<img src="' + slackIconUrlLocal + '" alt="Slack" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Slack' };
            } else if (link.includes('freshrelease.com')) {
              return { icon: '<img src="' + freshreleaseIconUrl + '" alt="Freshrelease" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshrelease' };
            } else if (link.includes('freshdesk.com')) {
              return { icon: '<img src="' + freshdeskIconUrl + '" alt="Freshdesk" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshdesk' };
            } else if (link.includes('freshservice.com')) {
              return { icon: '<img src="' + freshserviceIconUrl + '" alt="Freshservice" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshservice' };
            } else if (link.includes('mail.google.com')) {
              return { icon: '<img src="' + gmailIconUrl + '" alt="Gmail" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Gmail' };
            } else if (link.includes('docs.google.com/document')) {
              return { icon: '<img src="' + googleDocsIconUrl + '" alt="Google Docs" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Docs' };
            } else if (link.includes('docs.google.com/spreadsheets') || link.includes('sheets.google.com')) {
              return { icon: '<img src="' + googleSheetsIconUrl + '" alt="Google Sheets" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Sheets' };
            } else if (link.includes('docs.google.com/presentation') || link.includes('slides.google.com')) {
              return { icon: '<img src="' + googleSlidesIconUrl + '" alt="Google Slides" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Slides' };
            } else if (link.includes('drive.google.com')) {
              return { icon: '<img src="' + driveIconUrl + '" alt="Google Drive" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Drive' };
            }
            return { icon: linkIcon, title: 'View link' };
          };

          // Generate primary source icon
          let sourceIconHtml = '';
          if (taskMessageLink) {
            const primaryIcon = getIconForLink(taskMessageLink);
            sourceIconHtml = '<a href="' + taskMessageLink + '" target="_blank" class="todo-source" title="' + primaryIcon.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + primaryIcon.icon + '</a>';
          }

          // Generate secondary links HTML
          const secondaryLinksHtml = secondaryLinks.filter(l => l != null).map(link => {
            const iconData = getIconForLink(link);
            return '<span style="color: #bdc3c7; margin: 0 4px;">|</span><a href="' + link + '" target="_blank" class="todo-source" title="' + iconData.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + iconData.icon + '</a>';
          }).join('');

          return `
              <div class="task-group-task-item slack-channel-task-full todo-item ${isUnread ? 'unread' : ''} ${isNew ? 'appearing' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                <div class="todo-left-actions">
                  <div class="todo-checkbox" data-todo-id="${task.id}"></div>
                </div>
                <div class="todo-content" style="flex: 1; min-width: 0;">
                  ${task.task_title ? `<div class="todo-title">${escapeHtml(task.task_title)}</div>` : ''}
                  <div class="todo-text${needsViewMore ? ' truncated' : ''}">${taskNameHtml}</div>
                  <div class="todo-meta">
                    <span class="todo-date">${formatDate(task.updated_at || task.created_at)}</span>
                    ${sourceIconHtml}
                    ${secondaryLinksHtml}
                  </div>
                  ${tagsHtml}
                </div>
              </div>
            `;
        }).join('')}
        </div>
      </div>
    `;
  }

  // Build HTML for tag-based task groups
  function buildTagGroupHtml(group, newItemIds = []) {
    // Skip groups with null, undefined, or empty tag names
    if (!group.tagName || (typeof group.tagName === 'string' && group.tagName.trim() === '')) {
      return '';
    }

    const taskCount = group.tasks.length;
    if (taskCount === 0) return ''; // Skip empty groups

    const tagName = group.tagName;
    const hasUnread = group.tasks.some(t => isTaskUnread(t.id));

    // Sort tasks by date
    const sortedTasks = [...group.tasks].sort((a, b) =>
      new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at)
    );

    // Create a comma-separated list of task IDs for bulk mark done
    const taskIds = sortedTasks.map(t => t.id).join(',');

    return `
      <div class="tag-group task-group ${hasUnread ? 'unread' : ''}" data-tag-name="${escapeHtml(tagName)}" data-task-ids="${taskIds}">
        <div class="task-group-header">
          <div class="task-group-checkbox" data-task-ids="${taskIds}" title="Mark all tasks as done"></div>
          <div class="task-group-icon" style="background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 6px; width: 28px; height: 28px; display: flex; align-items: center; justify-content: center;">
            <span style="font-size: 14px;">🏷️</span>
          </div>
          <div class="task-group-info">
            <div class="task-group-title" title="${escapeHtml(tagName)}">${escapeHtml(tagName.length > 60 ? tagName.substring(0, 60) + '...' : tagName)}</div>
            <div class="task-group-meta">
              <span class="task-group-count">${taskCount} action item${taskCount > 1 ? 's' : ''}</span>
              <span class="task-group-date">${formatDate(group.latestUpdate)}</span>
            </div>
          </div>
          <div class="task-group-actions">
            <span class="task-group-chevron">▼</span>
          </div>
        </div>
        <div class="task-group-tasks" style="display: none;">
          ${sortedTasks.map((task, index) => {
      const isUnread = isTaskUnread(task.id);
      const isNew = newItemIds.includes(task.id);
      const taskMessageLink = task.message_link || '';
      const secondaryLinks = task.secondary_links || [];

      // Handle long descriptions
      const taskName = task.task_name || '';
      const taskNameEscaped = escapeHtml(taskName);
      const maxLength = 180;
      const hasTaskTitle = !!task.task_title; const needsViewMore = taskName.length > maxLength || (hasTaskTitle && taskName.length > 60);
      let taskNameHtml;

      if (needsViewMore) {
        const truncated = escapeHtml(taskName.substring(0, maxLength));
        taskNameHtml = `<span class="todo-text-content" data-full-text="${taskNameEscaped}">${taskNameEscaped}</span><span class="view-more-inline" data-todo-id="${task.id}">View more</span>`;
      } else {
        taskNameHtml = `<span class="todo-text-content">${taskNameEscaped}</span>`;
      }

      // Get all icon URLs - Tag Group
      const slackIconUrl = chrome.runtime.getURL('icon-slack.png');
      const gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
      const driveIconUrl = chrome.runtime.getURL('icon-drive.png');
      const freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
      const freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
      const freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
      const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
      const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
      const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
      const linkIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

      // Helper to get icon for a link
      const getIconForLink = (link) => {
        if (link.includes('slack.com') || link.includes('app.slack.com')) {
          return { icon: '<img src="' + slackIconUrl + '" alt="Slack" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Slack' };
        } else if (link.includes('freshrelease.com')) {
          return { icon: '<img src="' + freshreleaseIconUrl + '" alt="Freshrelease" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshrelease' };
        } else if (link.includes('freshdesk.com')) {
          return { icon: '<img src="' + freshdeskIconUrl + '" alt="Freshdesk" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshdesk' };
        } else if (link.includes('freshservice.com')) {
          return { icon: '<img src="' + freshserviceIconUrl + '" alt="Freshservice" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshservice' };
        } else if (link.includes('mail.google.com')) {
          return { icon: '<img src="' + gmailIconUrl + '" alt="Gmail" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Gmail' };
        } else if (link.includes('docs.google.com/document')) {
          return { icon: '<img src="' + googleDocsIconUrl + '" alt="Google Docs" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Docs' };
        } else if (link.includes('docs.google.com/spreadsheets') || link.includes('sheets.google.com')) {
          return { icon: '<img src="' + googleSheetsIconUrl + '" alt="Google Sheets" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Sheets' };
        } else if (link.includes('docs.google.com/presentation') || link.includes('slides.google.com')) {
          return { icon: '<img src="' + googleSlidesIconUrl + '" alt="Google Slides" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Slides' };
        } else if (link.includes('drive.google.com')) {
          return { icon: '<img src="' + driveIconUrl + '" alt="Google Drive" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Drive' };
        }
        return { icon: linkIcon, title: 'View link' };
      };

      // Generate primary source icon
      let sourceIconHtml = '';
      if (taskMessageLink) {
        const primaryIcon = getIconForLink(taskMessageLink);
        sourceIconHtml = '<a href="' + taskMessageLink + '" target="_blank" class="todo-source" title="' + primaryIcon.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + primaryIcon.icon + '</a>';
      }

      // Generate secondary links HTML
      const secondaryLinksHtml = secondaryLinks.filter(l => l != null).map(link => {
        const iconData = getIconForLink(link);
        return '<span style="color: #bdc3c7; margin: 0 4px;">|</span><a href="' + link + '" target="_blank" class="todo-source" title="' + iconData.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + iconData.icon + '</a>';
      }).join('');

      // Render other tags (excluding the group's tag and empty tags)
      const todoTags = (task.tags || []).filter(t => t !== tagName && isValidTag(t));
      const tagsHtml = todoTags.length > 0 ? `
              <div class="todo-tags" style="display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px;">
                ${todoTags.map(tag => {
        const isActive = activeTagFilters.includes(tag);
        return `<span class="todo-tag ${isActive ? 'active' : ''}" data-tag="${escapeHtml(tag)}" style="display: inline-block; background: ${isActive ? 'linear-gradient(45deg, #667eea, #764ba2)' : 'rgba(102, 126, 234, 0.1)'}; color: ${isActive ? 'white' : '#667eea'}; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 500; cursor: pointer; border: 1px solid ${isActive ? 'transparent' : 'rgba(102, 126, 234, 0.3)'}; transition: all 0.2s;">${escapeHtml(tag)}</span>`;
      }).join('')}
              </div>
            ` : '';

      return `
              <div class="task-group-task-item todo-item ${isUnread ? 'unread' : ''} ${isNew ? 'appearing' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                <div class="todo-content" style="flex: 1; min-width: 0;">
                  <div class="todo-main-row" style="display: flex; align-items: flex-start; gap: 8px;">
                    <div class="todo-checkbox" data-todo-id="${task.id}"></div>
                    <div style="flex: 1; min-width: 0;">
                      ${task.task_title ? `<div class="todo-title" style="font-weight: 600; font-size: 13px; color: var(--text-primary); margin-bottom: 2px;">${escapeHtml(task.task_title)}</div>` : ''}
                      <div class="todo-text ${needsViewMore ? 'truncated' : ''}">${taskNameHtml}</div>
                      ${task.participant_text ? `<span class="todo-participant" style="display: block; font-size: 11px; color: var(--text-muted); margin-top: 2px;">${escapeHtml(task.participant_text)}</span>` : ''}
                    </div>
                  </div>
                  <div class="todo-meta" style="display: flex; align-items: center; gap: 6px; margin-top: 4px; margin-left: 24px;">
                    <span class="todo-date" style="font-size: 11px; color: var(--text-light);">${formatDate(task.updated_at || task.created_at)}</span>
                    <div class="todo-sources" style="display: flex; align-items: center; gap: 2px;">
                      ${sourceIconHtml}
                      ${secondaryLinksHtml}
                    </div>
                  </div>
                  ${tagsHtml}
                </div>
              </div>
            `;
    }).join('')}
        </div>
      </div>
    `;
  }

  // Setup event handlers for tag groups (similar to setupSlackChannelGroups)
  function setupTagGroups(container) {
    const tagGroups = container.querySelectorAll('.tag-group');
    console.log('setupTagGroups called, found groups:', tagGroups.length);

    tagGroups.forEach((group, idx) => {
      const header = group.querySelector('.task-group-header');
      const tasksContainer = group.querySelector('.task-group-tasks');
      const chevron = group.querySelector('.task-group-chevron');
      const groupCheckbox = group.querySelector('.task-group-checkbox');

      console.log(`Tag group ${idx}: header=${!!header}, tasksContainer=${!!tasksContainer}, chevron=${!!chevron}`);

      if (header && tasksContainer && chevron) {
        // Toggle expand/collapse
        header.addEventListener('click', (e) => {
          // Don't toggle if clicking on checkbox
          if (e.target.closest('.task-group-checkbox')) return;

          const isExpanded = tasksContainer.style.display !== 'none';
          tasksContainer.style.display = isExpanded ? 'none' : 'block';
          chevron.style.transform = isExpanded ? 'rotate(0deg)' : 'rotate(180deg)';
          group.classList.toggle('expanded', !isExpanded);
          console.log(`Tag group toggled: now ${!isExpanded ? 'expanded' : 'collapsed'}`);

          // When expanding, mark all tasks in this group as read
          if (!isExpanded) {
            const taskIds = group.dataset.taskIds?.split(',').filter(id => id) || [];
            taskIds.forEach(id => markTaskAsRead(id));
            group.classList.remove('unread');
          }
        });
      }

      // Group checkbox - mark all tasks as done
      if (groupCheckbox) {
        groupCheckbox.addEventListener('click', async (e) => {
          e.stopPropagation();
          const taskIdsAttr = groupCheckbox.getAttribute('data-task-ids');
          const taskIds = taskIdsAttr ? taskIdsAttr.split(',').filter(id => id).map(id => parseInt(id)) : [];
          if (taskIds.length === 0) return;

          // Visual feedback
          groupCheckbox.classList.add('checked');
          groupCheckbox.innerHTML = '✓';
          group.classList.add('completing');

          // Update local arrays immediately (optimistic update)
          taskIds.forEach(taskId => {
            allTodos = allTodos.filter(t => t.id != taskId);
            allFyiItems = allFyiItems.filter(t => t.id != taskId);
            markTaskAsUnread(taskId); // Reset read state so if task comes back, it will be unread
          });

          // Mark all tasks as done in background
          const promises = taskIds.map(id =>
            updateTodoField(id, 'status', 1).catch(err => console.error('Error marking task done:', err))
          );

          // Don't wait - start removal animation immediately
          Promise.all(promises);

          // Remove the entire group after animation
          setTimeout(() => {
            group.remove();
            updateTabCounts();
          }, 400);
        });
      }

      // Setup individual task checkboxes within the group
      const taskCheckboxes = group.querySelectorAll('.task-group-task-item .todo-checkbox');
      taskCheckboxes.forEach(checkbox => {
        checkbox.addEventListener('click', async (e) => {
          e.stopPropagation();
          const todoId = checkbox.getAttribute('data-todo-id');
          const taskItem = checkbox.closest('.task-group-task-item');

          // Visual feedback
          checkbox.classList.add('checked');
          checkbox.innerHTML = '✓';
          if (taskItem) {
            taskItem.classList.add('completing');
          }

          // Update local arrays immediately (optimistic update)
          allTodos = allTodos.filter(t => t.id != todoId);
          allFyiItems = allFyiItems.filter(t => t.id != todoId);
          markTaskAsUnread(todoId);

          try {
            await updateTodoField(parseInt(todoId), 'status', 1);

            // Remove task item after animation
            setTimeout(() => {
              if (taskItem) taskItem.remove();

              // Update group count or remove group if empty
              const remainingTasks = group.querySelectorAll('.task-group-task-item').length;
              if (remainingTasks === 0) {
                group.remove();
              } else {
                const countEl = group.querySelector('.task-group-count');
                if (countEl) {
                  countEl.textContent = `${remainingTasks} action item${remainingTasks > 1 ? 's' : ''}`;
                }
              }
              updateTabCounts();
            }, 400);
          } catch (err) {
            console.error('Error marking task done:', err);
            checkbox.classList.remove('checked');
            checkbox.innerHTML = '';
            if (taskItem) taskItem.classList.remove('completing');
          }
        });
      });
    });
  }

  // Setup event handlers for Slack channel groups (similar to setupTaskGroups)
  function setupSlackChannelGroups(container) {
    const slackGroups = container.querySelectorAll('.slack-channel-group');
    console.log('setupSlackChannelGroups called, found groups:', slackGroups.length);

    slackGroups.forEach((group, idx) => {
      const header = group.querySelector('.task-group-header');
      const tasksContainer = group.querySelector('.task-group-tasks');
      const chevron = group.querySelector('.task-group-chevron');
      const groupCheckbox = group.querySelector('.task-group-checkbox');

      console.log(`Setting up slack group ${idx}:`, { header: !!header, tasksContainer: !!tasksContainer, chevron: !!chevron });

      if (!header || !tasksContainer) {
        console.error('Missing header or tasks container for slack group', idx);
        return;
      }

      // Group checkbox click - mark all tasks as done
      if (groupCheckbox) {
        groupCheckbox.addEventListener('click', async (e) => {
          e.stopPropagation();
          const taskIds = groupCheckbox.dataset.taskIds?.split(',').filter(id => id) || [];
          if (taskIds.length === 0) return;

          // Visual feedback
          groupCheckbox.classList.add('checked');
          groupCheckbox.innerHTML = '✓';
          group.classList.add('completing');

          // Mark all tasks as done in background
          const promises = taskIds.map(id =>
            updateTodoField(id, 'status', 1).catch(err => console.error('Error marking task done:', err))
          );

          // Don't wait - start removal animation immediately
          Promise.all(promises);

          // Remove the entire group after animation
          setTimeout(() => {
            group.remove();
            updateTabCounts();
          }, 400);
        });
      }

      header.addEventListener('click', (e) => {
        console.log('Slack channel header clicked');
        // Don't toggle if clicking the open button or checkbox
        if (e.target.closest('.task-group-open-btn')) return;
        if (e.target.closest('.task-group-checkbox')) return;

        group.classList.toggle('expanded');
        tasksContainer.style.display = group.classList.contains('expanded') ? 'block' : 'none';
        if (chevron) {
          chevron.style.transform = group.classList.contains('expanded') ? 'rotate(180deg)' : '';
        }

        // When expanding, mark all tasks in this group as read
        if (group.classList.contains('expanded')) {
          const taskIds = group.dataset.taskIds?.split(',').filter(id => id) || [];
          taskIds.forEach(id => markTaskAsRead(id));
          group.classList.remove('unread');
        }
      });

      // Task item clicks
      group.querySelectorAll('.task-group-task-item').forEach(taskItem => {
        if (taskItem.dataset.groupHandlerAttached) return;
        taskItem.dataset.groupHandlerAttached = 'true';
        const checkbox = taskItem.querySelector('.todo-checkbox');
        const todoId = taskItem.dataset.todoId;

        // Checkbox click
        if (checkbox) {
          checkbox.addEventListener('click', async (e) => {
            e.stopPropagation();
            checkbox.classList.add('checked');
            checkbox.innerHTML = '✓';
            taskItem.classList.add('completing');

            try {
              await updateTodoField(todoId, 'status', 1);
              setTimeout(() => taskItem.remove(), 400);

              // Update group count or remove group if empty
              const remainingTasks = group.querySelectorAll('.task-group-task-item').length - 1;
              if (remainingTasks === 0) {
                group.remove();
              } else {
                const countEl = group.querySelector('.task-group-count');
                if (countEl) {
                  countEl.textContent = `${remainingTasks} action item${remainingTasks > 1 ? 's' : ''}`;
                }
              }
              updateTabCounts();
            } catch (err) {
              console.error('Error marking task done:', err);
              checkbox.classList.remove('checked');
              checkbox.innerHTML = '';
              taskItem.classList.remove('completing');
            }
          });
        }

        // Task item click (open transcript)
        taskItem.addEventListener('click', (e) => {
          console.log('Slack task item clicked, todoId:', todoId);
          if (e.target.closest('.todo-checkbox')) return;
          if (e.target.closest('.todo-source')) return;
          if (e.target.closest('.todo-tag')) return;
          if (e.target.closest('.view-more-inline')) return;
          showTranscriptSlider(todoId);
        });

        // Source link click - stop propagation
        const sourceLink = taskItem.querySelector('.todo-source');
        if (sourceLink) {
          sourceLink.addEventListener('click', (e) => {
            e.stopPropagation();
          });
        }

        // Tag click handlers
        taskItem.querySelectorAll('.todo-tag').forEach(tagEl => {
          tagEl.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleTagFilter(tagEl.dataset.tag);
          });
        });

        // View more/less click - use event delegation for dynamic content
        taskItem.addEventListener('click', (e) => {
          const viewMoreBtn = e.target.closest('.view-more-inline');
          if (viewMoreBtn) {
            e.stopPropagation();
            const textContent = taskItem.querySelector('.todo-text-content');
            const todoText = taskItem.querySelector('.todo-text');
            if (!textContent || !textContent.dataset.fullText) return;

            const fullText = textContent.dataset.fullText;
            const isExpanded = viewMoreBtn.classList.contains('view-less');

            if (isExpanded) {
              // Collapse - show truncated with View more
              const maxLength = 150;
              const truncated = fullText.substring(0, maxLength);
              todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline" data-todo-id="' + todoId + '">View more</span>';
              if (todoText) {
                todoText.classList.remove('expanded');
                todoText.classList.add('truncated');
              }
            } else {
              // Expand - show full text with View less
              todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline view-less" data-todo-id="' + todoId + '">View less</span>';
              if (todoText) {
                todoText.classList.remove('truncated');
                todoText.classList.add('expanded');
              }
              // Mark task as read
              markTaskAsRead(todoId);
            }
          }
        });
      });
    });
  }

  // Build Slack Channels Accordion for FYI tab
  function buildSlackChannelsAccordion(tasks) {
    if (!tasks || tasks.length === 0) return '';

    // Group by channel
    const { slackGroups } = groupTasksBySlackChannel(tasks);

    // Tasks are already filtered before calling this function
    const filteredGroups = Object.values(slackGroups);

    if (filteredGroups.length === 0) return '';

    // Sort by latest update
    filteredGroups.sort((a, b) => new Date(b.latestUpdate) - new Date(a.latestUpdate));

    const slackIconUrl = chrome.runtime.getURL('icon-slack.png');

    const channelItems = filteredGroups.map(group => {
      const taskCount = group.tasks.length;
      const channelName = group.channelName || 'Slack Channel';
      const truncatedName = channelName.length > 50 ? channelName.substring(0, 50) + '...' : channelName;

      // Check if any task in this channel is unread
      const hasUnreadTask = group.tasks.some(task => isTaskUnread(task.id, task.updated_at));

      // Create a comma-separated list of task IDs for bulk mark done
      const taskIds = group.tasks.map(t => t.id).join(',');

      // Get Slack channel URL from first task's message link
      const sortedTasks = [...group.tasks].sort((a, b) => new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at));
      const slackChannelUrl = getSlackChannelUrl(sortedTasks[0]?.message_link || '');

      return `
        <div class="slack-channel-accordion-item ${hasUnreadTask ? 'unread' : ''}" data-channel-name="${escapeHtml(channelName)}" data-task-ids="${taskIds}">
          <div class="slack-channel-item-header">
            <div class="slack-channel-group-checkbox" data-task-ids="${taskIds}" title="Mark all tasks as done"></div>
            ${slackChannelUrl
          ? `<a href="${escapeHtml(slackChannelUrl)}" target="_blank" title="Open Slack channel" style="display: inline-flex; flex-shrink: 0;"><img src="${slackIconUrl}" alt="Slack" style="width: 20px; height: 20px; cursor: pointer;"></a>`
          : `<img src="${slackIconUrl}" alt="Slack" style="width: 20px; height: 20px; flex-shrink: 0;">`
        }
            <div class="slack-channel-item-info">
              <div class="slack-channel-item-title" title="${escapeHtml(channelName)}">${slackChannelUrl
          ? `<a href="${escapeHtml(slackChannelUrl)}" target="_blank" style="color: inherit; text-decoration: none;">${escapeHtml(truncatedName)}</a>`
          : escapeHtml(truncatedName)}</div>
              <div class="slack-channel-item-meta">${taskCount} update${taskCount > 1 ? 's' : ''} · ${formatDate(group.latestUpdate)}</div>
            </div>
          </div>
          <div class="slack-channel-tasks-list" style="display: none;">
            ${group.tasks.map(task => {
            const taskName = task.task_name || '';
            const taskNameEscaped = escapeHtml(taskName);
            const maxLength = 180;
            const hasTaskTitle = !!task.task_title; const isLong = taskName.length > maxLength || (hasTaskTitle && taskName.length > 60);
            const taskMessageLink = task.message_link || '';
            const secondaryLinks = task.secondary_links || [];
            const taskUnread = isTaskUnread(task.id, task.updated_at);

            // Get all icon URLs
            const gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
            const driveIconUrl = chrome.runtime.getURL('icon-drive.png');
            const freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
            const freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
            const freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
            const googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
            const googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
            const googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
            const linkIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

            // Helper to get icon for a link
            const getIconForLink = (link) => {
              if (link.includes('slack.com') || link.includes('app.slack.com')) {
                return { icon: '<img src="' + slackIconUrl + '" alt="Slack" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Slack' };
              } else if (link.includes('freshrelease.com')) {
                return { icon: '<img src="' + freshreleaseIconUrl + '" alt="Freshrelease" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshrelease' };
              } else if (link.includes('freshdesk.com')) {
                return { icon: '<img src="' + freshdeskIconUrl + '" alt="Freshdesk" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshdesk' };
              } else if (link.includes('freshservice.com')) {
                return { icon: '<img src="' + freshserviceIconUrl + '" alt="Freshservice" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Freshservice' };
              } else if (link.includes('mail.google.com')) {
                return { icon: '<img src="' + gmailIconUrl + '" alt="Gmail" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Gmail' };
              } else if (link.includes('docs.google.com/document')) {
                return { icon: '<img src="' + googleDocsIconUrl + '" alt="Google Docs" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Docs' };
              } else if (link.includes('docs.google.com/spreadsheets') || link.includes('sheets.google.com')) {
                return { icon: '<img src="' + googleSheetsIconUrl + '" alt="Google Sheets" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Sheets' };
              } else if (link.includes('docs.google.com/presentation') || link.includes('slides.google.com')) {
                return { icon: '<img src="' + googleSlidesIconUrl + '" alt="Google Slides" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Slides' };
              } else if (link.includes('drive.google.com')) {
                return { icon: '<img src="' + driveIconUrl + '" alt="Google Drive" style="width: 14px; height: 14px; object-fit: contain;">', title: 'Open in Google Drive' };
              }
              return { icon: linkIcon, title: 'View link' };
            };

            // Generate primary source icon
            let sourceIconHtml = '';
            if (taskMessageLink) {
              const primaryIcon = getIconForLink(taskMessageLink);
              sourceIconHtml = '<a href="' + taskMessageLink + '" target="_blank" class="todo-source" title="' + primaryIcon.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s; margin-left: 8px;">' + primaryIcon.icon + '</a>';
            }

            // Generate secondary links HTML
            const secondaryLinksHtml = secondaryLinks.filter(l => l != null).map(link => {
              const iconData = getIconForLink(link);
              return '<span style="color: #bdc3c7; margin: 0 4px;">|</span><a href="' + link + '" target="_blank" class="todo-source" title="' + iconData.title + '" style="display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 4px; background: rgba(102, 126, 234, 0.08); transition: all 0.2s;">' + iconData.icon + '</a>';
            }).join('');

            // Handle View more for long text
            let taskTextHtml;
            if (isLong) {
              taskTextHtml = '<span class="todo-text-content" data-full-text="' + taskNameEscaped + '">' + taskNameEscaped + '</span><span class="view-more-inline" data-todo-id="' + task.id + '">View more</span>';
            } else {
              taskTextHtml = '<span class="todo-text-content">' + taskNameEscaped + '</span>';
            }

            return `
              <div class="slack-channel-task-item todo-item ${taskUnread ? 'unread' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                <div class="slack-channel-task-checkbox" data-todo-id="${task.id}" title="Mark as done"></div>
                <div class="slack-channel-task-content" style="flex: 1; min-width: 0;">
                  ${task.task_title ? `<div class="todo-title" style="font-weight: 600; font-size: 12px; color: var(--text-primary); margin-bottom: 2px;">${escapeHtml(task.task_title)}</div>` : ''}
                  <div class="slack-channel-task-text todo-text${isLong ? ' truncated' : ''}">${taskTextHtml}</div>
                  <div class="slack-channel-task-meta" style="display: flex; align-items: center;">${formatDate(task.updated_at || task.created_at)}${sourceIconHtml}${secondaryLinksHtml}</div>
                </div>
              </div>
            `}).join('')}
          </div>
        </div>
      `;
    }).join('');

    const totalChannels = filteredGroups.length;
    const hasAnyUnread = filteredGroups.some(group => group.tasks.some(t => isTaskUnread(t.id)));

    return `
      <div class="slack-channels-accordion ${hasAnyUnread ? 'has-unread' : ''}">
        <div class="slack-channels-accordion-header">
          <div class="slack-channels-accordion-title">
            <img src="${slackIconUrl}" alt="Slack" style="width: 16px; height: 16px;">
            <span>Slack Channels</span>
            <span class="slack-channels-accordion-count">${totalChannels}</span>
          </div>
          <span class="slack-channels-accordion-chevron">▼</span>
        </div>
        <div class="slack-channels-accordion-content">
          <div class="slack-channels-accordion-list">
            ${channelItems}
          </div>
        </div>
      </div>
    `;
  }

  function setupSlackChannelsAccordion(container) {
    const accordion = container.querySelector('.slack-channels-accordion');
    if (!accordion) return;

    const header = accordion.querySelector('.slack-channels-accordion-header');

    // Toggle main accordion on header click
    header.addEventListener('click', () => {
      accordion.classList.toggle('open');
    });

    // Toggle individual channel items
    accordion.querySelectorAll('.slack-channel-accordion-item').forEach(item => {
      const itemHeader = item.querySelector('.slack-channel-item-header');
      const tasksList = item.querySelector('.slack-channel-tasks-list');
      const groupCheckbox = item.querySelector('.slack-channel-group-checkbox');

      // Group checkbox click - mark all tasks in this channel as done
      if (groupCheckbox) {
        groupCheckbox.addEventListener('click', async (e) => {
          e.stopPropagation();
          const taskIds = groupCheckbox.dataset.taskIds?.split(',').filter(id => id) || [];
          if (taskIds.length === 0) return;

          // Visual feedback
          groupCheckbox.classList.add('checked');
          groupCheckbox.innerHTML = '✓';
          item.classList.add('completing');

          // Mark all tasks as done in background
          const promises = taskIds.map(id =>
            updateTodoField(id, 'status', 1).catch(err => console.error('Error marking task done:', err))
          );

          // Don't wait - start removal animation immediately
          Promise.all(promises);

          // Remove the entire channel item after animation
          setTimeout(() => {
            item.remove();
            // Update accordion count
            const countEl = accordion.querySelector('.slack-channels-accordion-count');
            const currentCount = parseInt(countEl.textContent) - 1;
            countEl.textContent = currentCount;
            if (currentCount === 0) {
              accordion.remove();
            }
            updateTabCounts();
          }, 400);
        });
      }

      itemHeader.addEventListener('click', (e) => {
        // Don't toggle if clicking the open button or checkbox
        if (e.target.closest('.slack-channel-open-btn')) return;
        if (e.target.closest('.slack-channel-group-checkbox')) return;

        e.stopPropagation();
        item.classList.toggle('expanded');
        tasksList.style.display = item.classList.contains('expanded') ? 'block' : 'none';

        // When expanding, mark all tasks in this group as read
        if (item.classList.contains('expanded')) {
          const taskIds = item.dataset.taskIds?.split(',').filter(id => id) || [];
          taskIds.forEach(id => markTaskAsRead(id));
          item.classList.remove('unread');
          // Check if entire accordion still has unread channels
          const anyUnreadChannels = accordion.querySelectorAll('.slack-channel-accordion-item.unread');
          if (anyUnreadChannels.length === 0) {
            accordion.classList.remove('has-unread');
          }
        }
      });
    });

    // Slack task checkbox click - mark as done
    accordion.querySelectorAll('.slack-channel-task-checkbox').forEach(checkbox => {
      checkbox.addEventListener('click', async (e) => {
        e.stopPropagation();
        const todoId = checkbox.dataset.todoId;
        const taskItem = checkbox.closest('.slack-channel-task-item');
        const channelItem = checkbox.closest('.slack-channel-accordion-item');

        // Immediate visual feedback with animation
        checkbox.innerHTML = '✓';
        checkbox.classList.add('checked');
        taskItem.classList.add('completing');

        // Start API call in background (don't wait for it)
        updateTodoField(todoId, 'status', 1).catch(error => {
          console.error('Error marking Slack task done:', error);
        });

        // Remove item after animation completes
        setTimeout(() => {
          taskItem.remove();
          // Update count or remove channel item if empty
          const remainingTasks = channelItem.querySelectorAll('.slack-channel-task-item').length;
          if (remainingTasks === 0) {
            channelItem.remove();
            // Update accordion count
            const countEl = accordion.querySelector('.slack-channels-accordion-count');
            const currentCount = parseInt(countEl.textContent) - 1;
            countEl.textContent = currentCount;
            if (currentCount === 0) {
              accordion.remove();
            }
          } else {
            // Update channel item meta
            const metaEl = channelItem.querySelector('.slack-channel-item-meta');
            if (metaEl) {
              const dateText = metaEl.textContent.split('·')[1]?.trim() || '';
              metaEl.textContent = `${remainingTasks} update${remainingTasks > 1 ? 's' : ''} · ${dateText}`;
            }
          }
          updateTabCounts();
        }, 400);
      });
    });

    // Task item click - open transcript slider
    accordion.querySelectorAll('.slack-channel-task-item').forEach(taskItem => {
      // Source link click - stop propagation
      taskItem.querySelectorAll('.todo-source').forEach(sourceLink => {
        sourceLink.addEventListener('click', (e) => {
          e.stopPropagation();
        });
      });

      const todoId = taskItem.dataset.todoId;

      taskItem.addEventListener('click', (e) => {
        if (e.target.closest('.slack-channel-task-checkbox')) return;
        if (e.target.closest('.todo-source')) return;

        // Handle view more/less clicks
        const viewMoreBtn = e.target.closest('.view-more-inline');
        if (viewMoreBtn) {
          e.stopPropagation();
          const textContent = taskItem.querySelector('.todo-text-content');
          const todoText = taskItem.querySelector('.todo-text, .slack-channel-task-text');
          if (!textContent || !textContent.dataset.fullText) return;

          const fullText = textContent.dataset.fullText;
          const isExpanded = viewMoreBtn.classList.contains('view-less');

          if (isExpanded) {
            // Collapse - show truncated with View more
            const maxLength = 180;
            const truncated = fullText.substring(0, maxLength);
            todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline" data-todo-id="' + todoId + '">View more</span>';
            if (todoText) {
              todoText.classList.remove('expanded');
              todoText.classList.add('truncated');
            }
          } else {
            // Expand - show full text with View less
            todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline view-less" data-todo-id="' + todoId + '">View less</span>';
            if (todoText) {
              todoText.classList.remove('truncated');
              todoText.classList.add('expanded');
            }
            // Mark task as read
            markTaskAsRead(todoId);
          }
          return;
        }

        if (todoId) {
          showTranscriptSlider(todoId);
        }
      });
    });
  }

  function updateFilterSlider() {
    const activeBtn = document.querySelector('.filter-btn.active');
    const slider = document.querySelector('.filter-slider');
    if (!activeBtn || !slider) return;
    const track = document.querySelector('.filter-track');
    const btns = Array.from(track.querySelectorAll('.filter-btn'));
    const index = btns.indexOf(activeBtn);
    const btnWidth = activeBtn.offsetWidth;
    let leftOffset = 0;
    for (let i = 0; i < index; i++) {
      leftOffset += btns[i].offsetWidth;
    }
    slider.style.left = leftOffset + 'px';
    slider.style.width = btnWidth + 'px';
  }

  function addTodoEventListeners() {
    document.querySelectorAll('.todo-item').forEach(item => {
      // Prevent duplicate click handlers
      if (item.dataset.clickHandlerAttached) return;
      item.dataset.clickHandlerAttached = 'true';

      const checkbox = item.querySelector('.todo-checkbox');
      const star = item.querySelector('.todo-star');
      const clock = item.querySelector('.todo-clock');
      const todoId = item.dataset.todoId;

      // Make whole item clickable for transcript or multi-select
      item.addEventListener('click', (e) => {
        // Don't trigger if clicking on interactive elements
        if (e.target.closest('.todo-checkbox') ||
          e.target.closest('.todo-star') ||
          e.target.closest('.todo-clock') ||
          e.target.closest('.todo-source') ||
          e.target.closest('.todo-tag') ||
          e.target.closest('.view-more-inline') ||
          e.target.closest('a')) {
          return;
        }

        // Command/Ctrl + click on the item itself triggers multi-select
        if (e.metaKey || e.ctrlKey) {
          e.preventDefault();
          isMultiSelectMode = true;
          toggleTodoSelection(todoId);
          return;
        }

        showTranscriptSlider(todoId);
      });

      checkbox?.addEventListener('click', (e) => {
        e.stopPropagation();
        // Check if Command/Ctrl key is pressed for multi-select
        if (e.metaKey || e.ctrlKey) {
          isMultiSelectMode = true;
          toggleTodoSelection(checkbox.dataset.todoId);
        } else if (selectedTodoIds.size > 0) {
          // If already in multi-select mode, continue selecting
          toggleTodoSelection(checkbox.dataset.todoId);
        } else {
          // Normal single-click behavior
          handleTodoAction('toggle-complete', checkbox.dataset.todoId, checkbox);
        }
      });
      star?.addEventListener('click', (e) => {
        e.stopPropagation();
        handleTodoAction('toggle-star', star.dataset.todoId, star);
      });
      clock?.addEventListener('click', (e) => {
        e.stopPropagation();
        clock.style.transform = 'scale(0.9)';
        setTimeout(() => clock.style.transform = 'scale(1)', 150);
        showDueByMenu(clock, clock.dataset.todoId);
      });
    });
  }

  function addViewMoreListeners() {
    document.querySelectorAll('.view-more-inline:not(.view-less)').forEach(btn => {
      // Skip if already has listener
      if (btn.dataset.listenerAttached) return;
      btn.dataset.listenerAttached = 'true';

      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const todoText = btn.closest('.todo-text');
        if (!todoText) return;
        const todoTextContent = todoText.querySelector('.todo-text-content');
        if (!todoTextContent || !todoTextContent.dataset.fullText) return;

        const fullText = todoTextContent.dataset.fullText;
        const todoId = btn.dataset.todoId;

        // Mark task as read when View more is clicked
        markTaskAsRead(todoId);

        // Expand - show full text with View less
        todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline view-less" data-todo-id="' + todoId + '">View less</span>';
        todoText.classList.remove('truncated');
        todoText.classList.add('expanded');

        // Attach listener to View less button
        const viewLessBtn = todoText.querySelector('.view-less');
        if (viewLessBtn) {
          viewLessBtn.addEventListener('click', (ev) => {
            ev.stopPropagation();
            const maxLength = 180;
            const truncated = fullText.substring(0, maxLength);
            todoText.innerHTML = '<span class="todo-text-content" data-full-text="' + escapeHtml(fullText) + '">' + escapeHtml(fullText) + '</span><span class="view-more-inline" data-todo-id="' + todoId + '">View more</span>';
            todoText.classList.remove('expanded');
            todoText.classList.add('truncated');
            // Re-initialize listeners
            addViewMoreListeners();
          });
        }
      });
    });
  }

  // Helper function to mark a task done and open the next task's slider
  function markTaskDoneAndOpenNext(currentTodoId, closeSliderFn) {
    // Find the next visible task from the DOM (not from arrays which may include filtered items)
    let nextTodoId = null;

    // Get all visible todo items in the Action column
    const actionColumn = document.getElementById('todosContainer');
    const fyiColumn = document.getElementById('fyiContainer');

    // Get visible tasks from both columns (excluding calendar/meeting items)
    const actionTasks = actionColumn ?
      Array.from(actionColumn.querySelectorAll('.todo-item[data-todo-id]')).filter(el => !el.closest('.meetings-list')) : [];
    const fyiTasks = fyiColumn ?
      Array.from(fyiColumn.querySelectorAll('.todo-item[data-todo-id]')).filter(el => !el.closest('.meetings-list')) : [];

    // Find current task in visible lists
    const currentActionIndex = actionTasks.findIndex(el => el.dataset.todoId == currentTodoId);
    const currentFyiIndex = fyiTasks.findIndex(el => el.dataset.todoId == currentTodoId);

    if (currentActionIndex !== -1) {
      // Current task is in Action column - get next visible task
      if (currentActionIndex < actionTasks.length - 1) {
        // Next task in Action
        nextTodoId = actionTasks[currentActionIndex + 1].dataset.todoId;
      } else if (actionTasks.length > 1) {
        // Was last in Action, go to previous
        nextTodoId = actionTasks[currentActionIndex - 1]?.dataset.todoId;
      } else if (fyiTasks.length > 0) {
        // No more Action tasks, go to first FYI
        nextTodoId = fyiTasks[0].dataset.todoId;
      }
    } else if (currentFyiIndex !== -1) {
      // Current task is in FYI column - get next visible task
      if (currentFyiIndex < fyiTasks.length - 1) {
        // Next task in FYI
        nextTodoId = fyiTasks[currentFyiIndex + 1].dataset.todoId;
      } else if (fyiTasks.length > 1) {
        // Was last in FYI, go to previous
        nextTodoId = fyiTasks[currentFyiIndex - 1]?.dataset.todoId;
      }
    } else {
      // Task not found in visible lists (might be from calendar or elsewhere)
      // Try to get the first visible task from Action or FYI
      if (actionTasks.length > 0) {
        nextTodoId = actionTasks[0].dataset.todoId;
      } else if (fyiTasks.length > 0) {
        nextTodoId = fyiTasks[0].dataset.todoId;
      }
    }

    console.log('markTaskDoneAndOpenNext:', {
      currentTodoId,
      nextTodoId,
      actionTasksCount: actionTasks.length,
      fyiTasksCount: fyiTasks.length
    });

    // Animate out the task in the list (if visible)
    const taskElement = document.querySelector(`.todo-item[data-todo-id="${currentTodoId}"]`);
    if (taskElement) {
      const checkbox = taskElement.querySelector('.todo-checkbox');
      if (checkbox) {
        checkbox.classList.add('checked');
        checkbox.innerHTML = '✓';
      }
      taskElement.classList.add('completing');

      // Remove from DOM after animation
      setTimeout(() => {
        taskElement.remove();
      }, 400);
    }

    // Update local arrays immediately (optimistic update)
    allTodos = allTodos.filter(t => t.id != currentTodoId);
    allFyiItems = allFyiItems.filter(t => t.id != currentTodoId);

    // Track as recently completed to prevent reappearing via pending updates
    addRecentlyCompleted(currentTodoId);

    // Reset read state so if task comes back, it will be unread (yellow)
    markTaskAsUnread(currentTodoId);

    // Update counts
    if (typeof updateTabCounts === 'function') {
      updateTabCounts();
    }

    // Send to backend in background
    updateTodoField(currentTodoId, 'status', 1).catch(err => {
      console.error('Error marking task done:', err);
    });

    // Close current slider and open next
    if (closeSliderFn) {
      closeSliderFn();
    }

    // Open next task's slider after a brief delay
    if (nextTodoId) {
      setTimeout(() => {
        showTranscriptSlider(nextTodoId);
      }, 300);
    }
  }

  async function showTranscriptSlider(todoId) {
    // Find the todo with this ID (check allTodos, allFyiItems, allCalendarItems, and allCompletedTasks)
    const todo = allTodos.find(t => t.id == todoId) || allFyiItems.find(t => t.id == todoId) || allCalendarItems.find(t => t.id == todoId) || allCompletedTasks.find(t => t.id == todoId);
    if (!todo) return;

    // Mark task as read when slider opens
    markTaskAsRead(todoId);

    // Set flag to prevent auto-refresh
    isTranscriptSliderOpen = true;

    // Remove any existing transcript slider (but NOT chat sliders)
    document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
    document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());

    // V37: Overlay on top of col3 using fixed positioning — cols 1&2 untouched
    const col3El = document.getElementById('col3');
    const threeColLayout = document.querySelector('.three-column-layout');
    const col3Rect = window.Oracle.getCol3Rect();

    const overlay = document.createElement('div');
    overlay.className = 'transcript-slider-overlay';

    if (col3Rect) {
      // Fixed overlay exactly covering col3's screen position
      overlay.style.cssText = `position: fixed; top: ${col3Rect.top}px; left: ${col3Rect.left}px; width: ${col3Rect.width}px; bottom: 0; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out; transition: width 0.3s ease, left 0.3s ease;`;
    } else {
      overlay.style.cssText = 'position: fixed; top: 0; right: 0; bottom: 0; width: 400px; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;';
    }

    // Check if dark mode is active
    const isDarkMode = document.body.classList.contains('dark-mode');
    let isExpanded = false;

    // Create slider panel
    const slider = document.createElement('div');
    slider.className = 'transcript-slider';
    slider.style.cssText = `width: 100%; height: 100%; background: ${isDarkMode ? '#1f2940' : 'white'}; box-shadow: -4px 0 20px rgba(0,0,0,${isDarkMode ? '0.4' : '0.15'}); display: flex; flex-direction: column; animation: slideInRight 0.3s ease-out; transition: width 0.3s ease; overflow-x: hidden; border-radius: 12px; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};`;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `padding: 20px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0;`;
    header.innerHTML = `
      <div style="display: flex; align-items: center; gap: 12px;">
        <div style="width: 40px; height: 40px; background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 20px;">💬</div>
        <div style="flex: 1; min-width: 0;">
          <div style="font-weight: 600; font-size: 16px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Conversation</div>
          <div class="transcript-message-count" style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">Loading...</div>
        </div>
      </div>
      <div style="display: flex; align-items: center; gap: 8px;">
        <button class="transcript-expand-btn" title="Expand" style="background: ${isDarkMode ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.1)'}; border: none; color: #667eea; width: 36px; height: 36px; border-radius: 8px; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">⬅</button>
        ${todo.message_link ? (() => { const _icon = (window.OracleIcons.getIconForLink || (() => ({icon:'🔗',title:'Open source'})))(todo.message_link); return `<a href="${escapeHtml(todo.message_link)}" target="_blank" title="${_icon.title}" style="display: flex; align-items: center; justify-content: center; width: 36px; height: 36px; border-radius: 8px; background: ${isDarkMode ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.1)'}; transition: all 0.2s; text-decoration: none;">${_icon.icon}</a>`; })() : ''}
        <button class="transcript-close-btn" style="background: rgba(231,76,60,0.1); border: none; color: #e74c3c; width: 36px; height: 36px; border-radius: 8px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">×</button>
      </div>
    `;
    slider.appendChild(header);

    // Task title section - will be updated with participants after fetch
    let titleSection = null;
    if (todo.task_title) {
      titleSection = document.createElement('div');
      titleSection.className = 'transcript-title-section';
      titleSection.style.cssText = `padding: 12px 20px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.1)' : 'rgba(102, 126, 234, 0.05)'}; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; flex-shrink: 0;`;
      titleSection.innerHTML = `
        <div style="font-weight: 600; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(todo.task_title)}</div>
        <div class="transcript-title-participants" style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'}; margin-top: 4px; display: none;"></div>
      `;
      slider.appendChild(titleSection);
    }

    // Messages container with loading spinner
    const messagesContainer = document.createElement('div');
    messagesContainer.className = 'transcript-messages-container';
    messagesContainer.style.cssText = 'flex: 1; overflow-y: auto; overflow-x: hidden; padding: 20px; display: flex; flex-direction: column; gap: 16px;';
    messagesContainer.innerHTML = `
      <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 16px;">
        <div class="spinner" style="width: 40px; height: 40px; border: 4px solid rgba(102, 126, 234, 0.2); border-top-color: #667eea; border-radius: 50%; animation: spin 0.8s linear infinite;"></div>
        <div style="font-size: 14px; color: #7f8c8d;">Loading conversation...</div>
      </div>
    `;
    slider.appendChild(messagesContainer);

    // Reply section
    const replySection = document.createElement('div');
    replySection.style.cssText = `padding: 16px 20px; border-top: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; background: ${isDarkMode ? 'rgba(26, 26, 46, 0.8)' : 'rgba(248,249,250,0.8)'}; flex-shrink: 0; position: relative;`;
    replySection.innerHTML = `
      <div class="transcript-attachments" style="display: none; margin-bottom: 10px; padding: 10px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.1)' : 'rgba(102, 126, 234, 0.05)'}; border-radius: 8px; border: 1px dashed rgba(102, 126, 234, 0.3);">
        <div class="attachments-list" style="display: flex; flex-wrap: wrap; gap: 8px;"></div>
      </div>
      <div class="transcript-oracle-overlay" style="display: none; position: absolute; bottom: 100%; left: 0; right: 0; height: 120%; min-height: 400px; background: ${isDarkMode ? '#1a1f2e' : '#f8f9fa'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(200,210,220,0.8)'}; border-bottom: none; border-radius: 12px 12px 0 0; box-shadow: 0 -4px 20px rgba(0,0,0,0.2); z-index: 100; flex-direction: column; overflow: hidden;">
        <div class="oracle-overlay-header" style="display: flex; align-items: center; justify-content: space-between; padding: 12px 16px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(200,210,220,0.6)'}; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.12)'};">
          <div style="display: flex; align-items: center; gap: 8px;">
            <div style="width: 24px; height: 24px; background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 6px; display: flex; align-items: center; justify-content: center; font-size: 14px; color: white;">∞</div>
            <span style="font-weight: 600; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Oracle Assistant</span>
          </div>
          <button class="oracle-overlay-close" style="background: transparent; border: none; color: ${isDarkMode ? '#888' : '#7f8c8d'}; cursor: pointer; font-size: 18px; padding: 4px 8px; border-radius: 4px; transition: all 0.2s;">×</button>
        </div>
        <div class="oracle-overlay-content" style="flex: 1; overflow-y: auto; padding: 16px; font-size: 14px; line-height: 1.6; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">
          <div class="oracle-loading" style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 12px;">
            <div class="spinner" style="width: 32px; height: 32px; border: 3px solid rgba(102, 126, 234, 0.2); border-top-color: #667eea; border-radius: 50%; animation: spin 0.8s linear infinite;"></div>
            <div style="color: ${isDarkMode ? '#888' : '#7f8c8d'};">Analyzing thread...</div>
          </div>
        </div>
      </div>
      <div class="transcript-recipients-panel" style="display: none; margin-bottom: 10px; padding: 12px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.1)' : 'rgba(102, 126, 234, 0.05)'}; border-radius: 10px; border: 1px solid ${isDarkMode ? 'rgba(102, 126, 234, 0.3)' : 'rgba(102, 126, 234, 0.15)'};">
        <div class="recipients-section" style="margin-bottom: 8px; position: relative;">
          <div style="font-size: 11px; font-weight: 600; color: ${isDarkMode ? '#b0b0b0' : '#5d6d7e'}; margin-bottom: 4px;">To</div>
          <div style="display: flex; flex-wrap: wrap; gap: 4px; align-items: center; padding: 6px 10px; background: ${isDarkMode ? '#16213e' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'}; border-radius: 8px; min-height: 36px;">
            <div class="recipients-to-chips" style="display: flex; flex-wrap: wrap; gap: 4px;"></div>
            <input type="text" class="recipients-to-input" placeholder="Type 3+ chars to search..." style="flex: 1; min-width: 120px; border: none; background: transparent; outline: none; font-size: 12px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; padding: 2px 0;">
          </div>
          <div class="recipients-to-results" style="display: none; position: absolute; bottom: 100%; left: 0; right: 0; margin-bottom: 4px; background: ${isDarkMode ? '#1f2940' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.15)' : 'rgba(225,232,237,0.6)'}; border-radius: 8px; box-shadow: 0 -4px 16px rgba(0,0,0,0.15); max-height: 180px; overflow-y: auto; z-index: 1000;"></div>
        </div>
        <div class="recipients-section" style="margin-bottom: 8px; position: relative;">
          <div style="font-size: 11px; font-weight: 600; color: ${isDarkMode ? '#b0b0b0' : '#5d6d7e'}; margin-bottom: 4px;">CC</div>
          <div style="display: flex; flex-wrap: wrap; gap: 4px; align-items: center; padding: 6px 10px; background: ${isDarkMode ? '#16213e' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'}; border-radius: 8px; min-height: 36px;">
            <div class="recipients-cc-chips" style="display: flex; flex-wrap: wrap; gap: 4px;"></div>
            <input type="text" class="recipients-cc-input" placeholder="Type 3+ chars to search..." style="flex: 1; min-width: 120px; border: none; background: transparent; outline: none; font-size: 12px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; padding: 2px 0;">
          </div>
          <div class="recipients-cc-results" style="display: none; position: absolute; bottom: 100%; left: 0; right: 0; margin-bottom: 4px; background: ${isDarkMode ? '#1f2940' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.15)' : 'rgba(225,232,237,0.6)'}; border-radius: 8px; box-shadow: 0 -4px 16px rgba(0,0,0,0.15); max-height: 180px; overflow-y: auto; z-index: 1000;"></div>
        </div>
        <div class="recipients-section" style="position: relative;">
          <div style="font-size: 11px; font-weight: 600; color: ${isDarkMode ? '#b0b0b0' : '#5d6d7e'}; margin-bottom: 4px;">BCC</div>
          <div style="display: flex; flex-wrap: wrap; gap: 4px; align-items: center; padding: 6px 10px; background: ${isDarkMode ? '#16213e' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'}; border-radius: 8px; min-height: 36px;">
            <div class="recipients-bcc-chips" style="display: flex; flex-wrap: wrap; gap: 4px;"></div>
            <input type="text" class="recipients-bcc-input" placeholder="Type 3+ chars to search..." style="flex: 1; min-width: 120px; border: none; background: transparent; outline: none; font-size: 12px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; padding: 2px 0;">
          </div>
          <div class="recipients-bcc-results" style="display: none; position: absolute; bottom: 100%; left: 0; right: 0; margin-bottom: 4px; background: ${isDarkMode ? '#1f2940' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.15)' : 'rgba(225,232,237,0.6)'}; border-radius: 8px; box-shadow: 0 -4px 16px rgba(0,0,0,0.15); max-height: 180px; overflow-y: auto; z-index: 1000;"></div>
        </div>
      </div>
      <div style="display: flex; gap: 12px; align-items: flex-end;">
        <div style="display: flex; flex-direction: column; gap: 6px; align-items: center;">
          <button class="transcript-mic-btn" title="Voice to text" style="background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.3); color: #667eea; width: 44px; height: 38px; border-radius: 10px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">
            🎤
          </button>
          <button class="transcript-recipients-btn" title="Manage Recipients" style="background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.3); color: #667eea; width: 44px; height: 38px; border-radius: 10px; cursor: pointer; font-size: 16px; display: none; align-items: center; justify-content: center; transition: all 0.2s;">
            👤
          </button>
          <input type="file" class="transcript-file-input" multiple style="display: none;">
          <button class="transcript-attach-btn" title="Attach files" style="background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.3); color: #667eea; width: 44px; height: 38px; border-radius: 10px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">
            📎
          </button>
        </div>
        <div class="transcript-reply-input" contenteditable="true" placeholder="Type your reply..." style="flex: 1; padding: 12px 16px; border: 2px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'}; border-radius: 12px; font-family: inherit; font-size: 14px; min-height: 80px; max-height: 50vh; overflow-y: auto; transition: all 0.2s; outline: none; background: ${isDarkMode ? '#16213e' : 'white'}; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.5;"></div>
        <div style="display: flex; flex-direction: column; gap: 6px; align-items: center;">
          <button type="button" class="transcript-oracle-btn" title="Ask Oracle Assistant" style="background: linear-gradient(45deg, #667eea, #764ba2); border: none; color: white; width: 34px; height: 34px; border-radius: 10px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.2s; padding: 0; font-size: 14px;">
            ∞
          </button>
          <button class="transcript-send-btn" title="Send (⌘+Enter)" style="background: linear-gradient(45deg, #667eea, #764ba2); color: white; border: none; width: 34px; height: 34px; border-radius: 10px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; justify-content: center;">
            ↗
          </button>
          <button class="transcript-done-btn" title="Mark as Done & Close" style="background: rgba(39,174,96,0.12); border: 1px solid rgba(39,174,96,0.3); color: #27ae60; width: 34px; height: 34px; border-radius: 10px; font-size: 16px; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; justify-content: center;">
            ✓
          </button>
        </div>
      </div>
      <div style="margin-top: 8px; font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">Press Enter to type, ⌘+Enter to send</div>
    `;
    slider.appendChild(replySection);

    overlay.appendChild(slider);
    // V37: Append to body with fixed positioning
    document.body.appendChild(overlay);

    // Ensure no element is focused when slider opens (allows arrow key scrolling)
    // Use setTimeout to ensure this runs after any auto-focus from contenteditable
    setTimeout(() => {
      const replyInput = slider.querySelector('.transcript-reply-input');
      if (replyInput) replyInput.blur();
      document.activeElement?.blur();
    }, 10);

    // Also prevent the contenteditable from auto-focusing
    const replyInputEl = slider.querySelector('.transcript-reply-input');
    if (replyInputEl) {
      replyInputEl.setAttribute('tabindex', '-1');
      // Re-enable tabindex after a moment so user can still tab to it
      setTimeout(() => {
        replyInputEl.setAttribute('tabindex', '0');
      }, 100);
    }

    // Setup close handlers
    const closeBtn = slider.querySelector('.transcript-close-btn');
    const expandBtn = slider.querySelector('.transcript-expand-btn');
    let sliderKeyHandler; // Declare here so closeSlider can access it

    // Toggle expand function (used by button and keyboard shortcut)
    const toggleExpand = () => {
      isExpanded = !isExpanded;
      if (isExpanded && col3Rect) {
        const expandedWidth = Math.round(col3Rect.width * 1.5);
        overlay.style.width = expandedWidth + 'px';
        overlay.style.left = (col3Rect.left - (expandedWidth - col3Rect.width)) + 'px';
        expandBtn.innerHTML = '➡';
        expandBtn.title = 'Collapse';
      } else if (col3Rect) {
        overlay.style.width = col3Rect.width + 'px';
        overlay.style.left = col3Rect.left + 'px';
        expandBtn.innerHTML = '⬅';
        expandBtn.title = 'Expand';
      }
    };

    const closeSlider = () => {
      isTranscriptSliderOpen = false;
      overlay.style.animation = 'fadeOut 0.2s ease-out';
      slider.style.animation = 'slideOutRight 0.3s ease-out';
      document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());
      if (sliderKeyHandler) {
        document.removeEventListener('keydown', sliderKeyHandler);
      }
      setTimeout(() => { overlay.remove(); window.Oracle.collapseCol3AfterSlider(); }, 250);
    };

    // Expand/collapse functionality
    expandBtn.addEventListener('click', toggleExpand);

    closeBtn.addEventListener('click', closeSlider);
    // V37: Click outside (on backdrop) to close
    // V37: Click outside disabled — only Escape closes

    // Comprehensive keyboard handler for slider
    sliderKeyHandler = (e) => {
      const replyInput = slider.querySelector('.transcript-reply-input');
      const messagesContainer = slider.querySelector('.transcript-messages-container');
      const isReplyFocused = document.activeElement === replyInput || replyInput.contains(document.activeElement);

      // Command+Enter behavior depends on focus
      if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
        e.preventDefault();

        if (isReplyFocused) {
          // Cmd+Enter while typing → send the reply
          const sendBtn = slider.querySelector('.transcript-send-btn');
          if (sendBtn && !sendBtn.disabled) {
            sendBtn.click();
          }
        } else {
          // Cmd+Enter while not typing → send reply if content exists, then mark done
          const replyContent = replyInput?.textContent?.trim() || '';
          if (replyContent) {
            const sendBtn = slider.querySelector('.transcript-send-btn');
            if (sendBtn && !sendBtn.disabled) {
              sendBtn.click();
            }
          }
          markTaskDoneAndOpenNext(todoId, closeSlider);
        }
        return;
      }

      // When reply IS focused - handle special cases first
      if (isReplyFocused) {
        // Escape - blur the reply input (unfocus it)
        if (e.key === 'Escape') {
          e.preventDefault();
          replyInput.blur();
          return;
        }

        // Allow all other keys to work normally in the contenteditable
        // (Enter for newline, Up/Down for cursor movement, etc.)
        return;
      }

      // Escape - close slider (when reply is NOT focused, unless attachment preview is open)
      if (e.key === 'Escape') {
        if (document.querySelector('.attachment-preview-modal')) {
          return;
        }
        closeSlider();
        document.removeEventListener('keydown', sliderKeyHandler);
        return;
      }

      // Shift+ArrowLeft - toggle expand/collapse (works regardless of focus)
      if (e.shiftKey && e.key === 'ArrowLeft') {
        e.preventDefault();
        toggleExpand();
        return;
      }

      // When reply is NOT focused
      // Enter - focus the reply input
      if (e.key === 'Enter' && !e.metaKey && !e.ctrlKey) {
        e.preventDefault();
        replyInput.focus();
        return;
      }

      // Up/Down arrows - scroll the transcript
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        messagesContainer.scrollBy({ top: -100, behavior: 'smooth' });
        return;
      }
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        messagesContainer.scrollBy({ top: 100, behavior: 'smooth' });
        return;
      }
    };
    document.addEventListener('keydown', sliderKeyHandler);

    // Determine if this is a Drive/Docs task
    const isDriveTask = isDriveLink(todo.message_link);

    // Fetch conversation from API
    try {
      const fetchAction = isDriveTask ? 'fetch_comment_details' : 'fetch_task_details';
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: fetchAction,
          todo_id: todoId,
          message_link: todo.message_link,
          timestamp: new Date().toISOString(),
          source: 'oracle-chrome-extension-newtab'
        }))
      });

      if (!response.ok) throw new Error('Failed to fetch conversation');

      // Read response as text first to handle empty/malformed responses
      const responseText = await response.text();
      let data = {};

      if (responseText && responseText.trim()) {
        try {
          data = JSON.parse(responseText);
        } catch (parseError) {
          console.warn('Failed to parse conversation response:', parseError);
          data = {};
        }
      }

      // Parse transcript messages from response
      let messages = [];

      // Handle both array and object response formats
      const responseData = Array.isArray(data) ? data[0] : data;
      const transcript = responseData?.transcript || [];

      // Extract participants if available (for Gmail threads)
      const participants = responseData?.participants || [];
      const toParticipants = responseData?.to_participants || [];
      const ccParticipants = responseData?.cc_participants || [];
      console.log('Participants extracted:', participants);

      if (transcript && transcript.length > 0) {
        try {
          const flatTranscript = transcript.flat();
          messages = flatTranscript.map(msgStr => {
            try {
              return typeof msgStr === 'string' ? JSON.parse(msgStr) : msgStr;
            } catch {
              return null;
            }
          }).filter(m => m !== null);
        } catch (e) {
          console.error('Error parsing transcript:', e);
        }
      }

      // Store messages and todo on slider for Oracle button access
      slider.transcriptMessages = messages;
      slider.transcriptTodo = todo;

      // Update message count
      const countEl = slider.querySelector('.transcript-message-count');
      if (countEl) {
        countEl.textContent = messages.length > 0
          ? `${messages.length} message${messages.length !== 1 ? 's' : ''}`
          : 'No messages';
      }

      // Show participants in title section (below task title)
      const titleParticipantsEl = slider.querySelector('.transcript-title-participants');
      if (participants.length > 0) {
        const allNames = participants.map(p => p.name || p.email?.split('@')[0] || 'Unknown');
        const fullParticipantNames = allNames.join(', ');

        let displayNames;
        if (allNames.length <= 3) {
          displayNames = fullParticipantNames;
        } else {
          const firstThree = allNames.slice(0, 3).join(', ');
          const othersCount = allNames.length - 3;
          displayNames = `${firstThree} +${othersCount} others`;
        }

        if (titleParticipantsEl) {
          titleParticipantsEl.textContent = displayNames;
          titleParticipantsEl.title = fullParticipantNames;
          titleParticipantsEl.style.display = 'block';
        }
      } else if (todo.participant_text && titleParticipantsEl) {
        // Fallback: use participant_text from todo for Slack/other sources
        titleParticipantsEl.textContent = todo.participant_text;
        titleParticipantsEl.title = todo.participant_text;
        titleParticipantsEl.style.display = 'block';
      }

      // Store participants for recipients panel (will be populated after recipients management is initialized)
      slider._participants = participants;
      slider._toParticipants = toParticipants;
      slider._ccParticipants = ccParticipants;

      // Check if this is an email/gmail task (has multiple messages or participants)
      const isEmailTask = participants.length > 0 || (messages.length > 1 && messages.some(m => m.format === 'html'));
      slider._isEmailTask = isEmailTask;

      // Render messages
      if (messages.length > 0) {
        messagesContainer.innerHTML = '';
        messages.forEach((msg, msgIndex) => {
          const isLastMessage = msgIndex === messages.length - 1;
          const isCollapsed = isEmailTask && !isLastMessage && messages.length > 1;

          const msgDiv = document.createElement('div');
          msgDiv.className = `transcript-message ${isCollapsed ? 'collapsed' : ''}`;
          msgDiv.style.cssText = 'display: flex; flex-direction: column; gap: 6px;';

          // Always use the actual message timestamp from the conversation thread.
          // Previously used todo.updated_at for the last message, but that reflects
          // sync/task-update time, not when the message was actually sent — causing
          // misleading "Just now" for old messages across Slack, Gmail, and Drive.
          let displayTime = formatTimeAgoFresh(msg.time || '');

          const msgHeader = document.createElement('div');
          msgHeader.style.cssText = `display: flex; align-items: center; gap: 8px; ${isCollapsed ? 'cursor: pointer;' : ''}`;
          msgHeader.innerHTML = `
            ${isCollapsed ? '<div class="collapse-indicator" style="color: #667eea; font-size: 10px; margin-right: -4px;">▶</div>' : ''}
            <div style="width: 32px; height: 32px; background: linear-gradient(135deg, #667eea, #764ba2); border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: 12px; font-weight: 600;">${(msg.message_from || 'U').charAt(0).toUpperCase()}</div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(msg.message_from || 'Unknown')}</div>
              <div style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${displayTime}</div>
            </div>
          `;

          const msgContent = document.createElement('div');
          msgContent.className = 'transcript-message-content';
          // Updated styling to support HTML tables and rich content
          msgContent.style.cssText = `margin-left: 40px; padding: 12px 16px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border-radius: 12px; border-top-left-radius: 4px; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.6; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word; overflow: visible; max-width: 100%; ${isCollapsed ? 'display: none;' : ''}`;

          // Format the message content (preserves HTML for emails, formats plain text for Slack)
          const htmlContent = msg.message_html || '';
          const formattedContent = formatMessageContent(msg.message || '');

          // If we have rich HTML from the parser, always render in iframe for proper containment
          if (htmlContent) {
            msgContent.dataset.emailIframe = 'true';
            msgContent.dataset.rawHtml = htmlContent;
          } else if (isComplexEmailHtml(msg.message || '')) {
            msgContent.dataset.emailIframe = 'true';
            msgContent.dataset.rawHtml = msg.message;
          } else {
            msgContent.innerHTML = formattedContent;
          }

          msgDiv.appendChild(msgHeader);
          msgDiv.appendChild(msgContent);

          // Create attachments container (also hidden when collapsed)
          let attachmentsDiv = null;

          // Render attachments if present
          if (msg.attachments && msg.attachments.length > 0) {
            // Filter out generic "link" type attachments with name "Attachment" - these are usually duplicates of inline URLs
            const validAttachments = msg.attachments.filter(att => {
              // Skip if it's a generic "Attachment" link type
              if (att.type === 'link' && (!att.name || att.name === 'Attachment')) {
                return false;
              }
              // Skip empty attachments
              if (!att.name && !att.url && !att.text) {
                return false;
              }
              return true;
            });

            if (validAttachments.length > 0) {
              attachmentsDiv = document.createElement('div');
              attachmentsDiv.className = 'transcript-message-attachments';
              attachmentsDiv.style.cssText = `margin-left: 40px; margin-top: 8px; display: flex; flex-wrap: wrap; gap: 8px; ${isCollapsed ? 'display: none;' : ''}`;

              validAttachments.forEach(attachment => {
                const attachEl = renderTranscriptAttachment(attachment);
                if (attachEl) {
                  attachmentsDiv.appendChild(attachEl);
                }
              });

              if (attachmentsDiv.children.length > 0) {
                msgDiv.appendChild(attachmentsDiv);
              }
            }
          }

          // Render reactions if present
          if (msg.reactions && msg.reactions.length > 0) {
            const reactionsDiv = document.createElement('div');
            reactionsDiv.className = 'transcript-message-reactions';
            reactionsDiv.style.cssText = `margin-left: 40px; margin-top: 4px; display: flex; flex-wrap: wrap; gap: 4px; ${isCollapsed ? 'display: none;' : ''}`;

            msg.reactions.forEach(reaction => {
              const emojiName = reaction.emoji || reaction.name || '';
              const count = reaction.count || 1;
              const users = reaction.users || [];

              // Convert common Slack emoji names to unicode
              const emojiUnicode = convertSlackEmoji(emojiName);
              const tooltipText = users.length > 0 ? users.join(', ') : `${count} reaction${count > 1 ? 's' : ''}`;

              const pill = document.createElement('span');
              pill.className = 'reaction-pill';
              pill.style.cssText = `display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border: 1px solid ${isDarkMode ? 'rgba(102, 126, 234, 0.25)' : 'rgba(102, 126, 234, 0.15)'}; color: ${isDarkMode ? '#b0b0b0' : '#555'}; cursor: default; position: relative;`;
              pill.innerHTML = `<span style="font-size: 14px;">${emojiUnicode}</span>${count > 1 ? `<span style="font-size: 11px; font-weight: 500;">${count}</span>` : ''}`;

              // Custom styled tooltip
              const tooltip = document.createElement('div');
              tooltip.style.cssText = `display: none; position: absolute; bottom: calc(100% + 6px); left: 50%; transform: translateX(-50%); padding: 6px 10px; border-radius: 8px; font-size: 11px; line-height: 1.4; white-space: nowrap; z-index: 10001; pointer-events: none; background: ${isDarkMode ? '#2a2a2a' : '#333'}; color: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.25);`;
              tooltip.textContent = tooltipText;
              pill.appendChild(tooltip);
              pill.addEventListener('mouseenter', () => tooltip.style.display = 'block');
              pill.addEventListener('mouseleave', () => tooltip.style.display = 'none');

              reactionsDiv.appendChild(pill);
            });

            msgDiv.appendChild(reactionsDiv);
          }

          // Render thread indicator if present
          if (msg.thread && msg.thread.reply_count > 0) {
            const threadDiv = document.createElement('div');
            threadDiv.className = 'transcript-message-thread';
            threadDiv.style.cssText = `margin-left: 40px; margin-top: 4px; display: flex; align-items: center; gap: 6px; ${isCollapsed ? 'display: none;' : ''}`;

            const replyUsers = msg.thread.reply_users || [];
            const replyCount = msg.thread.reply_count;
            const latestReply = msg.thread.latest_reply || '';

            // Mini avatars for reply users (max 3)
            let avatarsHtml = '';
            const showUsers = replyUsers.slice(0, 3);
            const avatarColors = ['#667eea', '#764ba2', '#f093fb', '#4facfe', '#43e97b'];
            showUsers.forEach((user, i) => {
              const initial = (typeof user === 'string' ? user : 'U').charAt(0).toUpperCase();
              avatarsHtml += `<div style="width: 20px; height: 20px; border-radius: 50%; background: ${avatarColors[i % avatarColors.length]}; display: flex; align-items: center; justify-content: center; color: white; font-size: 9px; font-weight: 600; margin-left: ${i > 0 ? '-4px' : '0'}; border: 1.5px solid ${isDarkMode ? '#1f2940' : 'white'}; position: relative; z-index: ${3 - i};">${initial}</div>`;
            });

            const latestReplyText = latestReply ? ` · Last reply ${formatTimeAgoFresh(latestReply)}` : '';

            threadDiv.innerHTML = `
              <div style="display: flex; align-items: center;">${avatarsHtml}</div>
              <span style="font-size: 12px; color: #667eea; font-weight: 600; cursor: default;">${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}</span>
              <span style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${latestReplyText}</span>
            `;

            msgDiv.appendChild(threadDiv);
          }

          // Add click handler for collapsed messages to expand
          if (isCollapsed) {
            msgHeader.addEventListener('click', () => {
              const indicator = msgHeader.querySelector('.collapse-indicator');
              const isCurrentlyCollapsed = msgContent.style.display === 'none';

              if (isCurrentlyCollapsed) {
                msgContent.style.display = 'block';
                if (attachmentsDiv) attachmentsDiv.style.display = 'flex';
                msgDiv.querySelectorAll('.transcript-message-reactions, .transcript-message-thread').forEach(el => el.style.display = 'flex');
                if (indicator) indicator.textContent = '▼';
                // Re-trigger iframe height calculation since it was hidden during initial render
                const iframe = msgContent.querySelector('iframe');
                if (iframe && iframe.contentDocument?.body) {
                  const resize = () => {
                    try {
                      const bodyH = iframe.contentDocument.body.scrollHeight;
                      const docH = iframe.contentDocument.documentElement.scrollHeight;
                      iframe.style.height = (Math.max(bodyH, docH) + 8) + 'px';
                    } catch(e) {}
                  };
                  resize();
                  setTimeout(resize, 100);
                  setTimeout(resize, 300);
                  setTimeout(resize, 600);
                }
              } else {
                msgContent.style.display = 'none';
                if (attachmentsDiv) attachmentsDiv.style.display = 'none';
                msgDiv.querySelectorAll('.transcript-message-reactions, .transcript-message-thread').forEach(el => el.style.display = 'none');
                if (indicator) indicator.textContent = '▶';
              }
            });
          }

          messagesContainer.appendChild(msgDiv);
        });

        // Scroll to show the top of the latest (last) message
        setTimeout(() => {
          const lastMessage = messagesContainer.lastElementChild;
          if (lastMessage) {
            lastMessage.scrollIntoView({ behavior: 'auto', block: 'start' });
          }
          // Render complex email HTML in iframes now that elements are in DOM
          messagesContainer.querySelectorAll('[data-email-iframe="true"]').forEach(el => {
            renderEmailInIframe(el.dataset.rawHtml, el, isDarkMode);
            delete el.dataset.emailIframe;
            delete el.dataset.rawHtml;
          });
        }, 100);
      } else {
        messagesContainer.innerHTML = `
          <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">
            <div style="font-size: 48px; opacity: 0.5;">💬</div>
            <div style="font-size: 14px;">No conversation history available</div>
          </div>
        `;
      }

    } catch (error) {
      console.error('Error fetching conversation:', error);
      messagesContainer.innerHTML = `
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 12px; color: #e74c3c;">
          <div style="font-size: 48px;">⚠️</div>
          <div style="font-size: 14px;">Failed to load conversation</div>
          <div style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">${escapeHtml(error.message)}</div>
        </div>
      `;
      const countEl = slider.querySelector('.transcript-message-count');
      if (countEl) countEl.textContent = 'Error loading';
    }

    // Setup reply functionality
    const replyInput = slider.querySelector('.transcript-reply-input');
    const sendBtn = slider.querySelector('.transcript-send-btn');
    const doneBtn = slider.querySelector('.transcript-done-btn');

    // Mark Done button handler — marks task as done and closes slider
    if (doneBtn) {
      doneBtn.addEventListener('click', () => {
        // Immediately close the slider
        closeSlider();

        // Immediately remove the task row from the DOM (Action or FYI list)
        const taskRow = document.querySelector(`.todo-item[data-todo-id="${todo.id}"]`);
        if (taskRow) {
          const checkbox = taskRow.querySelector('.todo-checkbox');
          if (checkbox) { checkbox.classList.add('checked'); checkbox.innerHTML = '✓'; }
          taskRow.classList.add('completing');
          // Check if parent group becomes empty after this removal
          checkEmptyParentGroup(taskRow);
          setTimeout(() => taskRow.remove(), 400);
        }

        // Also check task-group-task-item (tasks inside collapsed groups)
        const groupTaskRow = document.querySelector(`.task-group-task-item[data-todo-id="${todo.id}"]`);
        if (groupTaskRow && groupTaskRow !== taskRow) {
          groupTaskRow.classList.add('completing');
          checkEmptyParentGroup(groupTaskRow);
          setTimeout(() => {
            groupTaskRow.remove();
            // Update group count
            const parentGroup = groupTaskRow.closest('.task-group, .slack-channel-accordion-item');
            if (parentGroup) {
              const countEl = parentGroup.querySelector('.task-group-count, .slack-channel-item-count');
              const remaining = parentGroup.querySelectorAll('.task-group-task-item:not(.completing)').length;
              if (countEl && remaining > 0) {
                countEl.textContent = `${remaining} ${remaining === 1 ? 'item' : 'items'}`;
              }
            }
          }, 400);
        }

        // Also remove the group row if it exists (Documents column)
        const groupRow = document.querySelector(`.group-item[data-todo-id="${todo.id}"]`);
        if (groupRow) {
          groupRow.classList.add('completing');
          setTimeout(() => groupRow.remove(), 400);
        }

        // Update local arrays immediately (optimistic)
        allTodos = allTodos.filter(t => t.id != todo.id);
        allFyiItems = allFyiItems.filter(t => t.id != todo.id);

        // Track as recently completed so it doesn't reappear via pending updates
        addRecentlyCompleted(todo.id);
        markTaskAsUnread(todo.id);

        // Update counts
        if (typeof updateTabCounts === 'function') updateTabCounts();
        if (typeof window.updateBadge === 'function') window.updateBadge();

        // Send to backend in background (fire and forget)
        updateTodoField(todo.id, 'status', 1).catch(err => {
          console.error('Error marking task done from slider:', err);
        });
      });
      doneBtn.addEventListener('mouseenter', () => { doneBtn.style.background = 'rgba(39,174,96,0.22)'; doneBtn.style.transform = 'scale(1.05)'; });
      doneBtn.addEventListener('mouseleave', () => { doneBtn.style.background = 'rgba(39,174,96,0.12)'; doneBtn.style.transform = 'scale(1)'; });
    }

    // Placeholder behavior for contenteditable
    replyInput.addEventListener('focus', () => {
      if (replyInput.textContent.trim() === '') {
        replyInput.innerHTML = '';
      }
      replyInput.style.borderColor = '#667eea';
      replyInput.style.boxShadow = '0 0 0 3px rgba(102, 126, 234, 0.1)';
    });

    replyInput.addEventListener('blur', () => {
      replyInput.style.borderColor = isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)';
      replyInput.style.boxShadow = 'none';
    });

    // Oracle Assistant button handler
    const oracleBtn = slider.querySelector('.transcript-oracle-btn');
    const oracleOverlay = slider.querySelector('.transcript-oracle-overlay');
    const oracleCloseBtn = slider.querySelector('.oracle-overlay-close');
    const oracleContent = slider.querySelector('.oracle-overlay-content');

    if (oracleBtn && oracleOverlay) {
      // Track Oracle assistant state on the slider for follow-up messages
      let oracleSessionId = null;
      let oracleConversationHistory = []; // Track conversation for follow-ups
      let oracleFullResponseText = '';
      slider._isOracleActive = false; // Track if Oracle overlay is currently active

      // Close overlay handler
      oracleCloseBtn?.addEventListener('click', () => {
        oracleOverlay.style.display = 'none';
        slider._isOracleActive = false;
      });

      // Helper: Add copy buttons to Oracle response
      const addOracleCopyButtons = (rawText) => {
        const existingActions = oracleContent.querySelector('.oracle-copy-actions');
        if (existingActions) existingActions.remove();

        const linkSvg = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>';
        const copySvg = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';

        const actionsDiv = document.createElement('div');
        actionsDiv.className = 'oracle-copy-actions';
        actionsDiv.style.cssText = 'display:flex;gap:6px;margin-top:12px;padding-top:8px;border-top:1px solid ' + (isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)') + ';';

        const makeCopyBtn = (title, svg, textToCopy) => {
          const btn = document.createElement('button');
          btn.title = title;
          btn.innerHTML = svg;
          btn.style.cssText = 'background:' + (isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)') + ';border:1px solid rgba(102,126,234,0.2);color:#667eea;width:28px;height:28px;border-radius:6px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.2s;';
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            navigator.clipboard.writeText(textToCopy).then(() => {
              btn.innerHTML = '✓'; btn.style.color = '#27ae60';
              setTimeout(() => { btn.innerHTML = svg; btn.style.color = '#667eea'; }, 1500);
            });
          });
          return btn;
        };

        // Extract URLs from response
        const urls = [...new Set((rawText.match(/https?:\/\/[^\s\]\)]+/g) || []))];
        if (urls.length > 0) {
          actionsDiv.appendChild(makeCopyBtn('Copy links', linkSvg, urls.join('\n')));
        }
        actionsDiv.appendChild(makeCopyBtn('Copy message', copySvg, rawText));

        const responseDiv = oracleContent.querySelector('.oracle-response');
        if (responseDiv) responseDiv.appendChild(actionsDiv);
      };

      // Core function to send a message to Oracle assistant (used for initial + follow-ups)
      const sendToOracle = async (userMessage, isFollowUp) => {
        const messages = slider.transcriptMessages || [];
        const currentTodo = slider.transcriptTodo || {};

        slider._isOracleActive = true;

        if (!isFollowUp) {
          oracleOverlay.style.display = 'flex';
          oracleConversationHistory = [];
          oracleFullResponseText = '';
          // Clear any existing content (prevents duplicate spinners from template)
          oracleContent.innerHTML = '';
          // Ensure flex column layout for proper bubble alignment
          oracleContent.style.display = 'flex';
          oracleContent.style.flexDirection = 'column';
        } else {
          // For follow-ups, show the user's message as a right-aligned bubble
          const userBubble = document.createElement('div');
          userBubble.style.cssText = 'padding:8px 12px;background:' + (isDarkMode ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.1)') + ';border-radius:10px;margin-bottom:12px;margin-left:auto;font-size:13px;color:' + (isDarkMode ? '#e0e0e0' : '#2c3e50') + ';max-width:85%;text-align:right;';
          userBubble.textContent = userMessage;
          oracleContent.appendChild(userBubble);
        }

        // Show a single loading spinner
        const loadingDiv = document.createElement('div');
        loadingDiv.className = 'oracle-loading';
        loadingDiv.style.cssText = 'display:flex;flex-direction:column;align-items:center;justify-content:center;padding:20px;gap:12px;';
        loadingDiv.innerHTML = '<div class="spinner" style="width:32px;height:32px;border:3px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;"></div><div style="color:' + (isDarkMode ? '#888' : '#7f8c8d') + ';">' + (isFollowUp ? 'Thinking...' : 'Analyzing thread...') + '</div>';
        oracleContent.appendChild(loadingDiv);
        oracleContent.scrollTop = oracleContent.scrollHeight;

        // Prepare thread data
        const threadConversation = messages.map(msg => ({
          role: msg.user || msg.sender || msg.message_from || 'unknown',
          content: msg.text || msg.body || msg.message || '',
          timestamp: msg.time || msg.timestamp || msg.ts || ''
        }));

        let userId = window.Oracle?.state?.userData?.userId;
        if (!userId) {
          try {
            const r = await chrome.storage.local.get(['oracle_user_data']);
            userId = r?.oracle_user_data?.userId || null;
          } catch { userId = null; }
        }

        const sessionId = oracleSessionId || ('transcript_' + todoId + '_' + Date.now());
        oracleSessionId = sessionId;

        oracleConversationHistory.push({ role: 'user', content: userMessage });

        const payload = {
          message: userMessage,
          session_id: sessionId,
          conversation: isFollowUp ? oracleConversationHistory : threadConversation,
          timestamp: new Date().toISOString(),
          source: isFollowUp ? 'oracle-transcript-followup' : 'oracle-transcript',
          user_id: userId,
          context: {
            task_title: currentTodo.task_title || '',
            task_name: currentTodo.task_name || '',
            message_link: currentTodo.message_link || ''
          }
        };

        // Function to convert URLs to hyperlinks with icons
        const formatOracleResponse = (text) => {
          let formatted = escapeHtml(text);
          const urlPattern = /\[(https?:\/\/[^\]]+)\]|(?<!\[)(https?:\/\/[^\s\]<>\"]+)/g;
          formatted = formatted.replace(urlPattern, (match, bracketUrl, plainUrl) => {
            const url = bracketUrl || plainUrl;
            let icon = '🔗', title = 'Open link';
            const bgColor = isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)';
            if (url.includes('slack.com')) { icon = '<img src="' + chrome.runtime.getURL('icon-slack.png') + '" style="width:14px;height:14px;">'; title = 'Open in Slack'; }
            else if (url.includes('docs.google.com/document')) { icon = '<img src="' + chrome.runtime.getURL('icon-google-docs.png') + '" style="width:14px;height:14px;">'; title = 'Open Google Doc'; }
            else if (url.includes('docs.google.com/spreadsheets') || url.includes('sheets.google.com')) { icon = '<img src="' + chrome.runtime.getURL('icon-google-sheets.png') + '" style="width:14px;height:14px;">'; title = 'Open Google Sheet'; }
            else if (url.includes('drive.google.com')) { icon = '<img src="' + chrome.runtime.getURL('icon-drive.png') + '" style="width:14px;height:14px;">'; title = 'Open Google Drive'; }
            else if (url.includes('zoom.us') || url.includes('zoom.com')) { icon = '<img src="' + chrome.runtime.getURL('icon-zoom.png') + '" style="width:14px;height:14px;">'; title = 'Open Zoom'; }
            else if (url.includes('meet.google.com')) { icon = '<img src="' + chrome.runtime.getURL('icon-google-meet.png') + '" style="width:14px;height:14px;">'; title = 'Open Google Meet'; }
            else if (url.includes('github.com')) { icon = '🐙'; title = 'Open GitHub'; }
            else if (url.includes('freshworks.com') || url.includes('freshdesk.com')) { icon = '🍃'; title = 'Open Freshworks'; }
            return '<a href="' + url + '" target="_blank" title="' + title + '" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;background:' + bgColor + ';border-radius:4px;text-decoration:none;margin:0 2px;border:1px solid rgba(102,126,234,0.2);transition:all 0.2s;vertical-align:middle;">' + icon + '</a>';
          });
          formatted = formatted.replace(/\\n/g, '<br>').replace(/\n/g, '<br>');
          formatted = formatted.replace(/^- /gm, '• ');
          return formatted;
        };

        try {
          const ablyKey = window.ABLY_CHAT_API_KEY;
          if (ablyKey && typeof Ably !== 'undefined') {
            let fullResponseText = '', streamStarted = false;
            const ablyRealtime = new Ably.Realtime({ key: ablyKey });
            const chatChannel = ablyRealtime.channels.get(sessionId);

            await new Promise((resolve, reject) => { chatChannel.attach((err) => err ? reject(err) : resolve()); });

            const streamTimeout = setTimeout(() => {
              chatChannel.unsubscribe(); ablyRealtime.close();
              if (!streamStarted) {
                const loadEl = oracleContent.querySelector('.oracle-loading');
                if (loadEl) loadEl.innerHTML = '<div style="padding:16px;color:#e74c3c;text-align:center;"><span style="font-size:24px;">⏱</span><div>Response timed out</div></div>';
              }
            }, 120000);

            chatChannel.subscribe('token', (msg) => {
              if (!streamStarted) {
                streamStarted = true;
                const loadEl = oracleContent.querySelector('.oracle-loading');
                if (loadEl) loadEl.remove();
                const responseDiv = document.createElement('div');
                responseDiv.className = 'oracle-response';
                responseDiv.style.cssText = 'white-space:pre-wrap;word-wrap:break-word;margin-right:auto;max-width:95%;padding:10px 14px;background:' + (isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.04)') + ';border-radius:10px;border-top-left-radius:4px;margin-bottom:8px;';
                oracleContent.appendChild(responseDiv);
              }
              let d = msg.data;
              if (typeof d === 'string') { try { d = JSON.parse(d); } catch { d = { text: d }; } }
              fullResponseText += (d.text || d);
              const responseDiv = oracleContent.querySelector('.oracle-response:last-of-type');
              if (responseDiv) responseDiv.innerHTML = formatOracleResponse(fullResponseText);
              oracleContent.scrollTop = oracleContent.scrollHeight;
            });

            chatChannel.subscribe('done', (msg) => {
              clearTimeout(streamTimeout); chatChannel.unsubscribe(); ablyRealtime.close();
              let d = msg.data;
              if (typeof d === 'string') { try { d = JSON.parse(d); } catch { d = {}; } }
              if (d?.fullText) fullResponseText = d.fullText;
              if (!streamStarted) {
                const loadEl = oracleContent.querySelector('.oracle-loading');
                if (loadEl) loadEl.remove();
                const responseDiv = document.createElement('div');
                responseDiv.className = 'oracle-response';
                responseDiv.style.cssText = 'white-space:pre-wrap;word-wrap:break-word;margin-right:auto;max-width:95%;padding:10px 14px;background:' + (isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.04)') + ';border-radius:10px;border-top-left-radius:4px;margin-bottom:8px;';
                oracleContent.appendChild(responseDiv);
              }
              oracleFullResponseText = fullResponseText;
              oracleConversationHistory.push({ role: 'assistant', content: fullResponseText });
              const responseDiv = oracleContent.querySelector('.oracle-response:last-of-type');
              if (responseDiv) responseDiv.innerHTML = formatOracleResponse(fullResponseText);
              // Add copy buttons
              addOracleCopyButtons(fullResponseText);
              oracleContent.scrollTop = oracleContent.scrollHeight;
            });

            chatChannel.subscribe('error', (msg) => {
              clearTimeout(streamTimeout); chatChannel.unsubscribe(); ablyRealtime.close();
              let d = msg.data;
              if (typeof d === 'string') { try { d = JSON.parse(d); } catch { d = {}; } }
              const loadEl = oracleContent.querySelector('.oracle-loading');
              if (loadEl) loadEl.innerHTML = '<div style="padding:16px;color:#e74c3c;text-align:center;"><span style="font-size:24px;">⚠️</span><div>' + escapeHtml(d?.message || 'Stream error') + '</div></div>';
            });

            fetch('https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(payload)
            }).catch(err => console.error('Oracle transcript webhook error:', err));

          } else {
            const response = await fetch('https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(payload)
            });
            if (!response.ok) throw new Error('HTTP ' + response.status);
            const result = await response.text();
            let responseContent = result;
            try { const j = JSON.parse(result); responseContent = j.response || j.message || j.output || result; } catch {}
            const loadEl = oracleContent.querySelector('.oracle-loading');
            if (loadEl) loadEl.remove();
            oracleFullResponseText = responseContent;
            oracleConversationHistory.push({ role: 'assistant', content: responseContent });
            const responseDiv = document.createElement('div');
            responseDiv.className = 'oracle-response';
            responseDiv.style.cssText = 'white-space:pre-wrap;word-wrap:break-word;margin-right:auto;max-width:95%;padding:10px 14px;background:' + (isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.04)') + ';border-radius:10px;border-top-left-radius:4px;margin-bottom:8px;';
            responseDiv.innerHTML = formatOracleResponse(responseContent);
            oracleContent.appendChild(responseDiv);
            addOracleCopyButtons(responseContent);
          }
        } catch (error) {
          console.error('Oracle request failed:', error);
          const loadEl = oracleContent.querySelector('.oracle-loading');
          if (loadEl) loadEl.innerHTML = '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:12px;color:#e74c3c;"><span style="font-size:24px;">⚠️</span><div>Failed to get response from Oracle</div><div style="font-size:12px;color:' + (isDarkMode ? '#888' : '#95a5a6') + ';">' + escapeHtml(error.message) + '</div></div>';
        }
      };

      // Oracle button click - initial thread analysis
      oracleBtn.addEventListener('click', () => {
        const currentTodo = slider.transcriptTodo || {};
        sendToOracle('Analyze this thread and provide insights: ' + (currentTodo.task_title || currentTodo.task_name || 'Thread'), false);
      });

      // Expose sendToOracle on the slider so the send button can use it for follow-ups
      slider._sendToOracle = sendToOracle;

      // Hover effect for Oracle button
      oracleBtn.addEventListener('mouseenter', () => {
        oracleBtn.style.transform = 'scale(1.05)';
        oracleBtn.style.boxShadow = '0 2px 8px rgba(102, 126, 234, 0.4)';
      });
      oracleBtn.addEventListener('mouseleave', () => {
        oracleBtn.style.transform = 'scale(1)';
        oracleBtn.style.boxShadow = 'none';
      });
    }

    // Keyboard shortcuts for formatting + auto-list support
    replyInput.addEventListener('keydown', (e) => {
      // Cmd/Ctrl+Enter → send
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        const sendBtn = slider.querySelector('.transcript-send-btn');
        if (sendBtn && !sendBtn.disabled) sendBtn.click();
        return;
      }
      // Plain Enter → insert line break (not <div>)
      if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey && !e.shiftKey) {
        e.preventDefault();
        document.execCommand('insertLineBreak');
        return;
      }
      // Cmd/Ctrl shortcuts
      if (e.ctrlKey || e.metaKey) {
        switch (e.key.toLowerCase()) {
          case 'b':
            e.preventDefault();
            document.execCommand('bold', false, null);
            break;
          case 'i':
            e.preventDefault();
            document.execCommand('italic', false, null);
            break;
          case 'u':
            e.preventDefault();
            document.execCommand('underline', false, null);
            break;
        }
      }
    });

    // Auto-list: convert "1. " to ordered list, "- " to unordered list
    replyInput.addEventListener('input', (e) => {
      if (e.inputType !== 'insertText' || e.data !== ' ') return;
      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      const node = sel.anchorNode;
      if (!node || node.nodeType !== 3) return;

      // Get text before cursor in this text node
      const textBeforeCursor = node.textContent.substring(0, sel.anchorOffset);

      // Check for "1. " (or any number followed by dot and space) at start
      if (/^\d+\.\s$/.test(textBeforeCursor)) {
        node.textContent = '';
        document.execCommand('insertOrderedList', false, null);
      }
      // Check for "- " at start
      else if (/^-\s$/.test(textBeforeCursor)) {
        node.textContent = '';
        document.execCommand('insertUnorderedList', false, null);
      }
    });

    // URL paste on selected text: if text is selected and a URL is pasted, hyperlink it
    replyInput.addEventListener('paste', (e) => {
      const sel = window.getSelection();
      if (!sel.rangeCount) return;

      const pastedText = (e.clipboardData || window.clipboardData).getData('text').trim();
      const isUrl = /^https?:\/\/\S+$/i.test(pastedText);

      if (isUrl && !sel.isCollapsed) {
        // Text is selected and pasting a URL — create hyperlink
        e.preventDefault();
        document.execCommand('createLink', false, pastedText);
        // Style the link
        const link = sel.anchorNode.parentElement.closest('a') || sel.anchorNode.parentElement.querySelector('a:last-child');
        if (link) {
          link.style.color = '#667eea';
          link.style.textDecoration = 'underline';
          link.setAttribute('target', '_blank');
        }
      }
    });

    // @Mention functionality
    let mentionDropdown = null;
    let mentionSearchTimeout = null;
    let mentionSelectedIndex = 0;
    let mentionResults = [];
    let currentMentionQuery = '';
    let mentionedUsers = []; // Track all mentioned users for sending

    // Create mention dropdown element
    const createMentionDropdown = () => {
      if (mentionDropdown) return mentionDropdown;

      mentionDropdown = document.createElement('div');
      mentionDropdown.className = 'mention-dropdown';
      mentionDropdown.style.display = 'none';

      // Position relative to the reply section container
      replySection.style.position = 'relative';
      replySection.insertBefore(mentionDropdown, replySection.firstChild);

      return mentionDropdown;
    };

    // Show loading state in dropdown
    const showMentionLoading = () => {
      const dropdown = createMentionDropdown();
      dropdown.innerHTML = `
        <div class="mention-dropdown-header">Searching users...</div>
        <div class="mention-loading">
          <div class="spinner"></div>
          <span>Loading...</span>
        </div>
      `;
      dropdown.style.display = 'block';
    };

    // Show mention results
    const showMentionResults = (users) => {
      console.log('[showMentionResults] Called with', users.length, 'users');
      const dropdown = createMentionDropdown();
      mentionResults = users;
      mentionSelectedIndex = 0;

      // Force clear first
      while (dropdown.firstChild) {
        dropdown.removeChild(dropdown.firstChild);
      }

      if (users.length === 0) {
        const headerDiv = document.createElement('div');
        headerDiv.className = 'mention-dropdown-header';
        headerDiv.textContent = 'Users';
        dropdown.appendChild(headerDiv);
        
        const emptyDiv = document.createElement('div');
        emptyDiv.className = 'mention-empty';
        emptyDiv.textContent = `No users found for "${currentMentionQuery}"`;
        dropdown.appendChild(emptyDiv);
      } else {
        const headerDiv = document.createElement('div');
        headerDiv.className = 'mention-dropdown-header';
        headerDiv.textContent = `Users (${users.length})`;
        dropdown.appendChild(headerDiv);
        
        const listDiv = document.createElement('div');
        listDiv.className = 'mention-dropdown-list';
        
        users.forEach((user, index) => {
          const initials = (user.name || user.email || '?').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
          
          const item = document.createElement('div');
          item.className = `mention-item ${index === 0 ? 'selected' : ''}`;
          item.dataset.index = index;
          item.dataset.name = user.name || user.email || '';
          item.dataset.email = user.email || '';
          item.dataset.slackId = user.slack_id || '';
          
          item.innerHTML = `
            <div class="mention-avatar">${initials}</div>
            <div class="mention-info">
              <div class="mention-name">${escapeHtml(user.name || user.email || 'Unknown')}</div>
              ${user.email ? `<div class="mention-email">${escapeHtml(user.email)}</div>` : ''}
            </div>
          `;
          listDiv.appendChild(item);
        });
        
        dropdown.appendChild(listDiv);
      }
      dropdown.style.display = 'block';
      
      // Force a reflow
      void dropdown.offsetHeight;
      
      console.log('[showMentionResults] Done. display:', dropdown.style.display, 'childElementCount:', dropdown.childElementCount, 'offsetHeight:', dropdown.offsetHeight, 'firstChild text:', dropdown.firstChild?.textContent?.substring(0, 30));
    };

    // Use event delegation for mention item clicks (more reliable)
    createMentionDropdown().addEventListener('mousedown', (e) => {
      // Use mousedown instead of click to fire before blur
      const mentionItem = e.target.closest('.mention-item');
      if (mentionItem) {
        e.preventDefault();
        e.stopPropagation();

        const name = mentionItem.dataset.name;
        const email = mentionItem.dataset.email;
        const slackId = mentionItem.dataset.slackId || '';

        console.log('Mention clicked:', { name, email, slackId });
        insertMention(name, email, slackId);
      }
    });

    // Hide mention dropdown
    const hideMentionDropdown = () => {
      console.log('[hideMentionDropdown] Called. Stack:', new Error().stack?.split('\n')[2]?.trim());
      if (mentionDropdown) {
        mentionDropdown.style.display = 'none';
      }
      mentionResults = [];
      mentionSelectedIndex = 0;
      currentMentionQuery = '';
    };

    // Update selected item in dropdown
    const updateMentionSelection = () => {
      if (!mentionDropdown) return;

      const items = mentionDropdown.querySelectorAll('.mention-item');
      items.forEach((item, index) => {
        if (index === mentionSelectedIndex) {
          item.classList.add('selected');
          item.scrollIntoView({ block: 'nearest' });
        } else {
          item.classList.remove('selected');
        }
      });
    };

    // Insert mention into the reply input (DOM-based to preserve existing mentions)
    const insertMention = (name, email, slackId) => {
      console.log('insertMention called:', { name, email, slackId });

      // Add to mentioned users list
      if (email && !mentionedUsers.find(u => u.email === email)) {
        mentionedUsers.push({ name, email, slack_id: slackId });
      }

      // Create the mention span element
      const mentionSpan = document.createElement('span');
      mentionSpan.className = 'mention-tag';
      mentionSpan.contentEditable = 'false';
      mentionSpan.dataset.email = email || '';
      mentionSpan.dataset.slackId = slackId || '';
      mentionSpan.textContent = `@${name}`;

      // Find the @query text node to replace using DOM traversal
      // Walk backwards through child nodes to find the text node containing the @query
      const selection = window.getSelection();
      let replaced = false;

      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        let containerNode = range.endContainer;

        // Must be a text node
        if (containerNode.nodeType === Node.TEXT_NODE) {
          const textContent = containerNode.textContent;
          const cursorOffset = range.endOffset;

          // Search backwards from cursor for @ in this text node
          let atPos = -1;
          for (let i = cursorOffset - 1; i >= 0; i--) {
            if (textContent[i] === '@') {
              atPos = i;
              break;
            }
            if (textContent[i] === ' ' || textContent[i] === '\n') {
              break;
            }
          }

          if (atPos !== -1) {
            // Split: text before @, the @query part, and text after cursor
            const beforeText = textContent.substring(0, atPos);
            const afterText = textContent.substring(cursorOffset);

            const parent = containerNode.parentNode;

            // Build replacement: beforeText node, mention span, space + afterText node
            const frag = document.createDocumentFragment();
            if (beforeText) {
              frag.appendChild(document.createTextNode(beforeText));
            }
            frag.appendChild(mentionSpan);
            frag.appendChild(document.createTextNode('\u00A0' + afterText)); // nbsp + remaining text

            parent.replaceChild(frag, containerNode);
            replaced = true;
          }
        }
      }

      if (!replaced) {
        // Fallback: append mention at end
        replyInput.appendChild(mentionSpan);
        replyInput.appendChild(document.createTextNode('\u00A0'));
      }

      // Move cursor to end
      const newRange = document.createRange();
      const sel = window.getSelection();
      newRange.selectNodeContents(replyInput);
      newRange.collapse(false);
      sel.removeAllRanges();
      sel.addRange(newRange);

      console.log('Mention inserted successfully');

      // Set cooldown to prevent checkForMention from immediately re-triggering
      mentionInsertCooldown = true;
      setTimeout(() => { mentionInsertCooldown = false; }, 500);

      hideMentionDropdown();
      replyInput.focus();
    };

    // Helper function to get caret position in contenteditable
    const getCaretCharacterOffsetWithin = (element) => {
      let caretOffset = 0;
      const selection = window.getSelection();
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        const preCaretRange = range.cloneRange();
        preCaretRange.selectNodeContents(element);
        preCaretRange.setEnd(range.endContainer, range.endOffset);
        caretOffset = preCaretRange.toString().length;
      }
      return caretOffset;
    };

    // Search for users via webhook
    let mentionAbortController = null;
    let mentionInsertCooldown = false;
    const searchMentionUsers = async (query) => {
      // Cancel any in-flight search
      if (mentionAbortController) {
        mentionAbortController.abort();
        mentionAbortController = null;
      }
      currentMentionQuery = query;
      showMentionLoading();

      const controller = new AbortController();
      mentionAbortController = controller;
      const timeoutId = setTimeout(() => controller.abort(), 10000);

      try {
        const response = await fetch(WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(createAuthenticatedPayload({
            action: 'search_user',
            query: query,
            timestamp: new Date().toISOString()
          })),
          signal: controller.signal
        });

        clearTimeout(timeoutId);
        mentionAbortController = null;

        if (!response.ok) throw new Error('Search failed');

        // Read response as text first to debug
        const responseText = await response.text();
        console.log('Raw response:', responseText);

        let data = [];

        // Only parse if response is not empty
        if (responseText && responseText.trim()) {
          try {
            data = JSON.parse(responseText);
          } catch (e) {
            console.warn('Failed to parse response:', e, 'Raw:', responseText);
            data = [];
          }
        }

        // If data is still a string (double-encoded), parse again
        if (typeof data === 'string') {
          try {
            data = JSON.parse(data);
          } catch (e) {
            console.warn('Failed to parse inner string:', e);
            data = [];
          }
        }

        // Handle different response formats
        let users = [];
        if (Array.isArray(data)) {
          users = data;
        } else if (data && data.users && Array.isArray(data.users)) {
          users = data.users;
        } else if (data && data.results && Array.isArray(data.results)) {
          users = data.results;
        }

        // Map Supabase Members table fields to expected format
        users = users.map(user => ({
          name: user['Full Name'] || user.name || user.full_name || '',
          email: user.user_email_ID || user.email || '',
          slack_id: user.user_slack_ID || user.slack_id || ''
        })).filter(user => user.name || user.email); // Filter out empty entries

        console.log('Mapped users:', users); // Debug log
        console.log('Query match check:', { queryFromClosure: query, currentMentionQuery, abortControllerExists: !!mentionAbortController, dropdownDisplay: mentionDropdown?.style?.display });

        // Show results if this search wasn't superseded by a newer search
        // (If aborted, we'd be in the catch block, so reaching here means this is the latest completed search)
        currentMentionQuery = query;
        showMentionResults(users);
        console.log('showMentionResults called, dropdown display:', mentionDropdown?.style?.display, 'innerHTML length:', mentionDropdown?.innerHTML?.length);
        console.log('[CRITICAL CHECK] isConnected:', mentionDropdown?.isConnected, 'parentNode:', mentionDropdown?.parentNode?.className, 'document.contains:', document.contains(mentionDropdown));
        
        // Check if there's a DIFFERENT .mention-dropdown visible in the DOM
        const allDropdowns = document.querySelectorAll('.mention-dropdown');
        console.log('[CRITICAL CHECK] Total .mention-dropdown elements in DOM:', allDropdowns.length);
        allDropdowns.forEach((dd, i) => {
          console.log(`  dropdown[${i}]: display=${dd.style.display}, isConnected=${dd.isConnected}, isSameNode=${dd === mentionDropdown}, hasLoading=${!!dd.querySelector('.mention-loading')}, hasMentionItem=${!!dd.querySelector('.mention-item')}`);
        });
        
        // Debug: check if something hides the dropdown shortly after
        const dropdownRef = mentionDropdown;
        setTimeout(() => {
          console.log('[DEBUG 100ms after showMentionResults] display:', dropdownRef?.style?.display, 'has mention-item:', dropdownRef?.querySelector('.mention-item') ? 'YES' : 'NO', 'has mention-loading:', dropdownRef?.querySelector('.mention-loading') ? 'YES' : 'NO', 'offsetHeight:', dropdownRef?.offsetHeight, 'isConnected:', dropdownRef?.isConnected);
        }, 100);
        setTimeout(() => {
          console.log('[DEBUG 500ms after showMentionResults] display:', dropdownRef?.style?.display, 'has mention-item:', dropdownRef?.querySelector('.mention-item') ? 'YES' : 'NO', 'offsetHeight:', dropdownRef?.offsetHeight);
        }, 500);
      } catch (error) {
        clearTimeout(timeoutId);
        mentionAbortController = null;
        if (error.name === 'AbortError') {
          console.log('Mention search aborted/timed out for:', query);
        } else {
          console.error('Error searching users:', error);
        }
        if (query === currentMentionQuery) {
          hideMentionDropdown();
        }
      }
    };

    // Detect @mention pattern in input (DOM-aware to skip existing mention spans)
    const checkForMention = () => {
      if (mentionInsertCooldown) return;
      const selection = window.getSelection();
      if (!selection.rangeCount) return;

      const range = selection.getRangeAt(0);
      const containerNode = range.endContainer;

      // Only look for @mentions in text nodes (not inside mention-tag spans)
      if (containerNode.nodeType !== Node.TEXT_NODE) {
        console.log('[checkForMention] Non-text node detected, nodeType:', containerNode.nodeType, 'nodeName:', containerNode.nodeName, 'searchInFlight:', !!mentionAbortController);
        // Don't fully hide/reset if a search is in-flight - just hide the visual dropdown
        if (mentionAbortController) {
          if (mentionDropdown) mentionDropdown.style.display = 'none';
        } else {
          hideMentionDropdown();
        }
        return;
      }

      // If the text node is inside a mention-tag, ignore it
      if (containerNode.parentElement && containerNode.parentElement.classList && containerNode.parentElement.classList.contains('mention-tag')) {
        console.log('[checkForMention] Inside mention-tag, hiding');
        if (mentionAbortController) {
          if (mentionDropdown) mentionDropdown.style.display = 'none';
        } else {
          hideMentionDropdown();
        }
        return;
      }

      const textContent = containerNode.textContent;
      const cursorOffset = range.endOffset;

      // Find @ symbol before cursor in this text node
      let atPos = -1;
      for (let i = cursorOffset - 1; i >= 0; i--) {
        if (textContent[i] === '@') {
          atPos = i;
          break;
        }
        // Stop if we hit a space, nbsp, or newline before finding @
        if (textContent[i] === ' ' || textContent[i] === '\u00A0' || textContent[i] === '\n') {
          break;
        }
      }

      if (atPos !== -1) {
        const query = textContent.substring(atPos + 1, cursorOffset);

        // Only search if query has 3+ characters
        if (query.length >= 3 && !/\s/.test(query)) {
          // Skip if same query is already being searched or was just searched
          if (query === currentMentionQuery && mentionDropdown && mentionDropdown.style.display === 'block') {
            return;
          }
          // Debounce the search
          if (mentionSearchTimeout) {
            clearTimeout(mentionSearchTimeout);
          }
          mentionSearchTimeout = setTimeout(() => {
            searchMentionUsers(query);
          }, 400);
        } else if (query.length < 3) {
          hideMentionDropdown();
        }
      } else {
        hideMentionDropdown();
      }
    };

    // Listen for input changes
    replyInput.addEventListener('input', checkForMention);

    // Handle keyboard navigation in mention dropdown
    replyInput.addEventListener('keydown', (e) => {
      if (!mentionDropdown || mentionDropdown.style.display === 'none' || mentionResults.length === 0) {
        return;
      }

      switch (e.key) {
        case 'ArrowDown':
          e.preventDefault();
          mentionSelectedIndex = Math.min(mentionSelectedIndex + 1, mentionResults.length - 1);
          updateMentionSelection();
          break;
        case 'ArrowUp':
          e.preventDefault();
          mentionSelectedIndex = Math.max(mentionSelectedIndex - 1, 0);
          updateMentionSelection();
          break;
        case 'Enter':
          if (mentionResults.length > 0) {
            e.preventDefault();
            const selected = mentionResults[mentionSelectedIndex];
            insertMention(selected.name || selected.email, selected.email || '', selected.slack_id || '');
          }
          break;
        case 'Escape':
          e.preventDefault();
          hideMentionDropdown();
          break;
        case 'Tab':
          if (mentionResults.length > 0) {
            e.preventDefault();
            const selected = mentionResults[mentionSelectedIndex];
            insertMention(selected.name || selected.email, selected.email || '', selected.slack_id || '');
          }
          break;
      }
    }, true); // Use capture phase to handle before other handlers

    // Hide dropdown when clicking outside
    document.addEventListener('click', (e) => {
      if (mentionDropdown && !mentionDropdown.contains(e.target) && !replyInput.contains(e.target)) {
        hideMentionDropdown();
      }
    });

    // Function to get mentioned users from the reply input (for send)
    const getMentionedUsersFromInput = () => {
      const mentionTags = replyInput.querySelectorAll('.mention-tag');
      const users = [];
      mentionTags.forEach(tag => {
        const email = tag.dataset.email;
        const slackId = tag.dataset.slackId;
        const name = tag.textContent.replace('@', '');
        if (email || slackId) {
          users.push({ name, email, slack_id: slackId });
        }
      });
      return users;
    };

    // Function to convert reply text to Slack format (replace @Name with <@SLACK_ID>)
    const getSlackFormattedText = () => {
      let text = '';
      replyInput.childNodes.forEach(node => {
        if (node.nodeType === Node.TEXT_NODE) {
          text += node.textContent;
        } else if (node.classList && node.classList.contains('mention-tag')) {
          const slackId = node.dataset.slackId;
          if (slackId) {
            text += `<@${slackId}>`;
          } else {
            text += node.textContent; // Fallback to @Name if no slack ID
          }
        } else if (node.nodeName === 'BR') {
          text += '\n';
        } else {
          text += node.textContent || '';
        }
      });
      return text.trim();
    };

    // Attachment handling
    const attachBtn = slider.querySelector('.transcript-attach-btn');
    const fileInput = slider.querySelector('.transcript-file-input');
    const attachmentsContainer = slider.querySelector('.transcript-attachments');
    const attachmentsList = slider.querySelector('.attachments-list');
    let pendingAttachments = [];

    attachBtn.addEventListener('click', () => {
      fileInput.click();
    });

    attachBtn.addEventListener('mouseenter', () => {
      attachBtn.style.background = 'rgba(102, 126, 234, 0.2)';
    });

    attachBtn.addEventListener('mouseleave', () => {
      attachBtn.style.background = 'rgba(102, 126, 234, 0.1)';
    });

    // Voice-to-text (mic) button using Web Speech API
    const micBtn = slider.querySelector('.transcript-mic-btn');
    if (micBtn) {
      let recognition = null;
      let isRecording = false;

      const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

      if (!SpeechRecognition) {
        // Browser doesn't support Speech Recognition — hide the button
        micBtn.style.display = 'none';
      } else {
        micBtn.addEventListener('click', () => {
          if (isRecording && recognition) {
            // Stop recording
            recognition.stop();
            return;
          }

          recognition = new SpeechRecognition();
          recognition.continuous = true;
          recognition.interimResults = true;
          recognition.lang = 'en-US';

          // Track what's been finalized vs interim
          let finalTranscript = '';

          // Visual: switch to recording state
          isRecording = true;
          micBtn.innerHTML = '⏹';
          micBtn.style.background = 'rgba(231,76,60,0.15)';
          micBtn.style.borderColor = 'rgba(231,76,60,0.4)';
          micBtn.style.color = '#e74c3c';
          micBtn.title = 'Stop recording';
          // Add a pulsing animation
          micBtn.style.animation = 'pulse 1.5s ease-in-out infinite';

          // Capture existing text in the reply input so we append, not replace
          const existingText = replyInput.innerText.trim();

          recognition.onresult = (event) => {
            let interimTranscript = '';
            for (let i = event.resultIndex; i < event.results.length; i++) {
              const transcript = event.results[i][0].transcript;
              if (event.results[i].isFinal) {
                finalTranscript += transcript;
              } else {
                interimTranscript += transcript;
              }
            }
            // Show final + interim in the reply input
            const prefix = existingText ? existingText + ' ' : '';
            const interimHtml = interimTranscript ? '<span style="color:#95a5a6;font-style:italic;">' + escapeHtml(interimTranscript) + '</span>' : '';
            replyInput.innerHTML = escapeHtml(prefix + finalTranscript) + interimHtml;
          };

          recognition.onend = () => {
            isRecording = false;
            micBtn.innerHTML = '🎤';
            micBtn.style.background = 'rgba(102, 126, 234, 0.1)';
            micBtn.style.borderColor = 'rgba(102, 126, 234, 0.3)';
            micBtn.style.color = '#667eea';
            micBtn.title = 'Voice to text';
            micBtn.style.animation = '';
            // Finalize the text (remove interim spans)
            const prefix = existingText ? existingText + ' ' : '';
            if (finalTranscript) {
              replyInput.innerHTML = escapeHtml(prefix + finalTranscript);
              // Place cursor at end
              const range = document.createRange();
              range.selectNodeContents(replyInput);
              range.collapse(false);
              const sel = window.getSelection();
              sel.removeAllRanges();
              sel.addRange(range);
            }
            recognition = null;
          };

          recognition.onerror = (event) => {
            console.error('Speech recognition error:', event.error);
            isRecording = false;
            micBtn.innerHTML = '🎤';
            micBtn.style.background = 'rgba(102, 126, 234, 0.1)';
            micBtn.style.borderColor = 'rgba(102, 126, 234, 0.3)';
            micBtn.style.color = '#667eea';
            micBtn.title = 'Voice to text';
            micBtn.style.animation = '';
            if (event.error === 'not-allowed') {
              alert('Microphone access was denied. Please allow microphone permission for this extension.');
            }
            recognition = null;
          };

          recognition.start();
        });

        micBtn.addEventListener('mouseenter', () => {
          if (!isRecording) micBtn.style.background = 'rgba(102, 126, 234, 0.2)';
        });
        micBtn.addEventListener('mouseleave', () => {
          if (!isRecording) micBtn.style.background = 'rgba(102, 126, 234, 0.1)';
        });
      }
    }

    // Drag and drop support for the reply section
    const replyInputArea = replyInput.parentElement; // The flex container with input and buttons

    replySection.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
      replySection.style.background = 'rgba(102, 126, 234, 0.15)';
      replySection.style.borderColor = 'rgba(102, 126, 234, 0.5)';
    });

    replySection.addEventListener('dragleave', (e) => {
      e.preventDefault();
      e.stopPropagation();
      replySection.style.background = 'rgba(248,249,250,0.8)';
      replySection.style.borderColor = '';
    });

    replySection.addEventListener('drop', async (e) => {
      e.preventDefault();
      e.stopPropagation();
      replySection.style.background = 'rgba(248,249,250,0.8)';
      replySection.style.borderColor = '';

      const files = Array.from(e.dataTransfer.files);
      if (files.length === 0) return;

      for (const file of files) {
        // Convert file to base64
        const base64 = await fileToBase64(file);
        pendingAttachments.push({
          name: file.name,
          type: file.type,
          size: file.size,
          data: base64
        });

        // Add to UI
        const attachmentEl = document.createElement('div');
        attachmentEl.className = 'attachment-item';
        attachmentEl.style.cssText = 'display: flex; align-items: center; gap: 6px; background: white; padding: 6px 10px; border-radius: 6px; border: 1px solid rgba(225, 232, 237, 0.6); font-size: 12px;';

        const icon = getFileIcon(file.type);
        attachmentEl.innerHTML = `
          <span>${icon}</span>
          <span style="max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${escapeHtml(file.name)}</span>
          <span style="color: #95a5a6; font-size: 10px;">(${formatFileSize(file.size)})</span>
          <button class="remove-attachment" data-name="${escapeHtml(file.name)}" style="background: none; border: none; color: #e74c3c; cursor: pointer; font-size: 14px; padding: 0 4px;">×</button>
        `;
        attachmentsList.appendChild(attachmentEl);

        // Remove button handler
        attachmentEl.querySelector('.remove-attachment').addEventListener('click', (e) => {
          const name = e.target.dataset.name;
          pendingAttachments = pendingAttachments.filter(a => a.name !== name);
          attachmentEl.remove();
          if (pendingAttachments.length === 0) {
            attachmentsContainer.style.display = 'none';
          }
        });
      }

      attachmentsContainer.style.display = 'block';
    });

    fileInput.addEventListener('change', async (e) => {
      const files = Array.from(e.target.files);
      if (files.length === 0) return;

      for (const file of files) {
        // Convert file to base64
        const base64 = await fileToBase64(file);
        pendingAttachments.push({
          name: file.name,
          type: file.type,
          size: file.size,
          data: base64
        });

        // Add to UI
        const attachmentEl = document.createElement('div');
        attachmentEl.className = 'attachment-item';
        attachmentEl.style.cssText = 'display: flex; align-items: center; gap: 6px; background: white; padding: 6px 10px; border-radius: 6px; border: 1px solid rgba(225, 232, 237, 0.6); font-size: 12px;';

        const icon = getFileIcon(file.type);
        attachmentEl.innerHTML = `
          <span>${icon}</span>
          <span style="max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${escapeHtml(file.name)}</span>
          <span style="color: #95a5a6; font-size: 10px;">(${formatFileSize(file.size)})</span>
          <button class="remove-attachment" data-name="${escapeHtml(file.name)}" style="background: none; border: none; color: #e74c3c; cursor: pointer; font-size: 14px; padding: 0 4px;">×</button>
        `;
        attachmentsList.appendChild(attachmentEl);

        // Remove button handler
        attachmentEl.querySelector('.remove-attachment').addEventListener('click', (e) => {
          const name = e.target.dataset.name;
          pendingAttachments = pendingAttachments.filter(a => a.name !== name);
          attachmentEl.remove();
          if (pendingAttachments.length === 0) {
            attachmentsContainer.style.display = 'none';
          }
        });
      }

      attachmentsContainer.style.display = 'block';
      fileInput.value = ''; // Reset for next selection
    });

    // Helper function to convert file to base64
    function fileToBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
      });
    }

    // Helper function to get file icon
    function getFileIcon(mimeType) {
      if (mimeType.startsWith('image/')) return '🖼️';
      if (mimeType.includes('pdf')) return '📄';
      if (mimeType.includes('word') || mimeType.includes('document')) return '📝';
      if (mimeType.includes('sheet') || mimeType.includes('excel')) return '📊';
      if (mimeType.includes('presentation') || mimeType.includes('powerpoint')) return '📽️';
      if (mimeType.startsWith('video/')) return '🎬';
      if (mimeType.startsWith('audio/')) return '🎵';
      if (mimeType.includes('zip') || mimeType.includes('archive')) return '📦';
      return '📎';
    }

    // Helper function to format file size
    function formatFileSize(bytes) {
      if (bytes < 1024) return bytes + ' B';
      if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
      return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }

    // ========== RECIPIENTS MANAGEMENT ==========
    const recipientsBtn = slider.querySelector('.transcript-recipients-btn');
    const recipientsPanel = slider.querySelector('.transcript-recipients-panel');
    const ccInput = slider.querySelector('.recipients-cc-input');
    const bccInput = slider.querySelector('.recipients-bcc-input');
    const toInput = slider.querySelector('.recipients-to-input');
    const ccResults = slider.querySelector('.recipients-cc-results');
    const bccResults = slider.querySelector('.recipients-bcc-results');
    const toResults = slider.querySelector('.recipients-to-results');
    const toChipsContainer = slider.querySelector('.recipients-to-chips');
    const ccChipsContainer = slider.querySelector('.recipients-cc-chips');
    const bccChipsContainer = slider.querySelector('.recipients-bcc-chips');

    // Recipients state (will be populated from participants data)
    let recipientTo = [];
    let recipientCc = [];
    let recipientBcc = [];
    let recipientsPanelOpen = false;
    let recipientSearchTimeout = null;

    // Helper: create a chip element for a recipient
    function createRecipientChip(person, type, opts = {}) {
      const chip = document.createElement('span');
      const displayName = person.name || person.email?.split('@')[0] || 'Unknown';
      chip.style.cssText = `display: inline-flex; align-items: center; gap: 4px; background: ${isDarkMode ? 'rgba(102,126,234,0.25)' : 'linear-gradient(135deg, rgba(102,126,234,0.15), rgba(118,75,162,0.1))'}; color: ${isDarkMode ? '#a0aeff' : '#667eea'}; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 500; cursor: grab; transition: all 0.2s; white-space: nowrap;`;
      chip.title = person.email || '';
      chip.dataset.email = person.email || '';
      chip.dataset.name = displayName;
      chip.dataset.slackId = person.slack_id || '';
      chip.dataset.chipType = type;

      // Make chip draggable
      chip.draggable = true;
      chip.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('text/plain', JSON.stringify({
          name: person.name || displayName,
          email: person.email,
          slack_id: person.slack_id || '',
          fromType: type
        }));
        chip.style.opacity = '0.5';
      });
      chip.addEventListener('dragend', () => { chip.style.opacity = '1'; });

      if (opts.removable !== false) {
        chip.innerHTML = `${escapeHtml(displayName)} <span class="chip-remove" style="font-size: 12px; opacity: 0.7; margin-left: 2px;">×</span>`;
        chip.querySelector('.chip-remove').addEventListener('click', (e) => {
          e.stopPropagation();
          if (type === 'to') recipientTo = recipientTo.filter(r => r.email !== person.email);
          else if (type === 'cc') recipientCc = recipientCc.filter(r => r.email !== person.email);
          else if (type === 'bcc') recipientBcc = recipientBcc.filter(r => r.email !== person.email);
          chip.remove();
        });
      } else {
        chip.innerHTML = escapeHtml(displayName);
      }

      // Click on a To chip to add to CC
      if (type === 'to' && opts.clickToCC) {
        chip.addEventListener('click', () => {
          const email = chip.dataset.email;
          if (email && !recipientCc.some(r => r.email === email)) {
            recipientCc.push({ name: chip.dataset.name, email: email, slack_id: chip.dataset.slackId || '' });
            ccChipsContainer.appendChild(createRecipientChip({ name: chip.dataset.name, email, slack_id: chip.dataset.slackId || '' }, 'cc'));
          }
        });
        chip.title = (person.email || '') + ' — click to add to CC, or drag to move';
      }

      return chip;
    }

    // Setup drag-and-drop on recipient containers
    function setupDropZone(container, targetType) {
      if (!container) return;
      const parentDiv = container.parentElement; // The wrapper div with border

      parentDiv.addEventListener('dragover', (e) => {
        e.preventDefault();
        parentDiv.style.borderColor = '#667eea';
        parentDiv.style.background = isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)';
      });
      parentDiv.addEventListener('dragleave', () => {
        parentDiv.style.borderColor = isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)';
        parentDiv.style.background = isDarkMode ? '#16213e' : 'white';
      });
      parentDiv.addEventListener('drop', (e) => {
        e.preventDefault();
        parentDiv.style.borderColor = isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)';
        parentDiv.style.background = isDarkMode ? '#16213e' : 'white';

        try {
          const data = JSON.parse(e.dataTransfer.getData('text/plain'));
          const { name, email, slack_id, fromType } = data;
          if (!email || fromType === targetType) return;

          // Remove from source
          if (fromType === 'to') {
            recipientTo = recipientTo.filter(r => r.email !== email);
            toChipsContainer.querySelector(`[data-email="${email}"]`)?.remove();
          } else if (fromType === 'cc') {
            recipientCc = recipientCc.filter(r => r.email !== email);
            ccChipsContainer.querySelector(`[data-email="${email}"]`)?.remove();
          } else if (fromType === 'bcc') {
            recipientBcc = recipientBcc.filter(r => r.email !== email);
            bccChipsContainer.querySelector(`[data-email="${email}"]`)?.remove();
          }

          // Add to target (if not already there)
          const person = { name, email, slack_id: slack_id || '' };
          if (targetType === 'to') {
            if (!recipientTo.some(r => r.email === email)) {
              recipientTo.push(person);
              toChipsContainer.appendChild(createRecipientChip(person, 'to', { removable: true, clickToCC: true }));
            }
          } else if (targetType === 'cc') {
            if (!recipientCc.some(r => r.email === email)) {
              recipientCc.push(person);
              ccChipsContainer.appendChild(createRecipientChip(person, 'cc'));
            }
          } else if (targetType === 'bcc') {
            if (!recipientBcc.some(r => r.email === email)) {
              recipientBcc.push(person);
              bccChipsContainer.appendChild(createRecipientChip(person, 'bcc'));
            }
          }
        } catch (err) { /* ignore invalid drops */ }
      });
    }

    setupDropZone(toChipsContainer, 'to');
    setupDropZone(ccChipsContainer, 'cc');
    setupDropZone(bccChipsContainer, 'bcc');

    // Helper: search users for CC/BCC fields (match the new message gmail search pattern)
    async function searchRecipientsFor(query, resultsContainer, targetType) {
      if (query.length < 3) {
        resultsContainer.style.display = 'none';
        return;
      }
      resultsContainer.style.display = 'block';
      resultsContainer.innerHTML = '<div style="padding: 10px; text-align: center; font-size: 12px; color: #7f8c8d;">Searching...</div>';

      // Build exclude lists from current recipients
      const existingToEmails = recipientTo.map(r => r.email).filter(Boolean);
      const existingCcEmails = recipientCc.map(r => r.email).filter(Boolean);
      const existingBccEmails = recipientBcc.map(r => r.email).filter(Boolean);
      const isCC = targetType === 'cc';

      try {
        const response = await fetch(WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'search_user_new_gmail',
            query: query,
            platform: 'gmail',
            existing_to_ids: existingToEmails,
            existing_cc_ids: isCC ? existingCcEmails : existingBccEmails,
            exclude_ids: [...existingToEmails, ...existingCcEmails, ...existingBccEmails],
            timestamp: new Date().toISOString(),
            source: 'oracle-chrome-extension-newtab',
            user_id: userData?.userId,
            authenticated: true
          })
        });

        if (response.ok) {
          let data = await response.json();
          // Handle both array and object responses
          if (!Array.isArray(data)) {
            data = data.results || data.members || data.users || [];
          }

          // Normalize field names from API and filter to employees only (exclude channels for Gmail CC/BCC)
          data = data.filter(r => r.type === 'employee' || (!r.type && r.user_email_ID));
          data = data.map(r => ({
            name: r.Full_Name || r['Full Name'] || r.full_name || r.name || r.user_email_ID || 'Unknown',
            email: r.user_email_ID || r.email || '',
            slack_id: r.user_slack_ID || r.slack_id || r.id || ''
          })).filter(r => r.email); // Only show results with an email

          if (data.length === 0) {
            // No search results - hide dropdown unless user is typing an email
            if (query.includes('@')) {
              resultsContainer.innerHTML = `<div class="recipient-result" data-email="${escapeHtml(query)}" data-name="${escapeHtml(query.split('@')[0])}" style="padding: 8px 12px; cursor: pointer; font-size: 12px; transition: background 0.15s; color: ${isDarkMode ? '#a0aeff' : '#667eea'};">+ Add "${escapeHtml(query)}"</div>`;
            } else {
              resultsContainer.style.display = 'none';
              return;
            }
          } else {
            // Show results + always show "add manually" option at the bottom if query looks like email
            let manualAddHtml = '';
            if (query.includes('@')) {
              manualAddHtml = `<div class="recipient-result" data-email="${escapeHtml(query)}" data-name="${escapeHtml(query.split('@')[0])}" style="padding: 8px 12px; cursor: pointer; font-size: 12px; transition: background 0.15s; color: ${isDarkMode ? '#a0aeff' : '#667eea'}; border-top: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};">+ Add "${escapeHtml(query)}" manually</div>`;
            }
            resultsContainer.innerHTML = data.map(r => `
              <div class="recipient-result" data-email="${escapeHtml(r.email || '')}" data-name="${escapeHtml(r.name || r.email?.split('@')[0] || '')}" data-slack-id="${escapeHtml(r.slack_id || r.id || '')}" 
                   style="display: flex; align-items: center; gap: 8px; padding: 8px 12px; cursor: pointer; transition: background 0.15s;">
                <div style="width: 28px; height: 28px; border-radius: 50%; background: linear-gradient(45deg, #667eea, #764ba2); display: flex; align-items: center; justify-content: center; color: white; font-size: 11px; font-weight: 600;">${(r.name || r.email || 'U').charAt(0).toUpperCase()}</div>
                <div style="flex: 1; min-width: 0;">
                  <div style="font-weight: 600; font-size: 12px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(r.name || r.email?.split('@')[0] || '')}</div>
                  <div style="font-size: 10px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">${escapeHtml(r.email || '')}</div>
                </div>
              </div>
            `).join('') + manualAddHtml;
          }

          // Click handlers for results
          resultsContainer.querySelectorAll('.recipient-result').forEach(el => {
            el.addEventListener('mouseenter', () => { el.style.background = isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'; });
            el.addEventListener('mouseleave', () => { el.style.background = 'transparent'; });
            el.addEventListener('click', () => {
              const email = el.dataset.email;
              const name = el.dataset.name;
              const slackId = el.dataset.slackId || '';
              const person = { name, email, slack_id: slackId };

              if (targetType === 'to') {
                if (!recipientTo.some(r => r.email === email)) {
                  recipientTo.push(person);
                  toChipsContainer.appendChild(createRecipientChip(person, 'to', { removable: true, clickToCC: true }));
                }
                if (toInput) toInput.value = '';
              } else if (targetType === 'cc') {
                if (!recipientCc.some(r => r.email === email)) {
                  recipientCc.push(person);
                  ccChipsContainer.appendChild(createRecipientChip(person, 'cc'));
                }
                ccInput.value = '';
              } else if (targetType === 'bcc') {
                if (!recipientBcc.some(r => r.email === email)) {
                  recipientBcc.push(person);
                  bccChipsContainer.appendChild(createRecipientChip(person, 'bcc'));
                }
                bccInput.value = '';
              }
              resultsContainer.style.display = 'none';
            });
          });
        }
      } catch (err) {
        console.error('Recipient search error:', err);
        resultsContainer.style.display = 'none';
      }
    }

    // Toggle recipients panel
    if (recipientsBtn) {
      recipientsBtn.addEventListener('click', () => {
        recipientsPanelOpen = !recipientsPanelOpen;
        recipientsPanel.style.display = recipientsPanelOpen ? 'block' : 'none';
        recipientsBtn.style.background = recipientsPanelOpen
          ? 'linear-gradient(45deg, #667eea, #764ba2)'
          : 'rgba(102, 126, 234, 0.1)';
        recipientsBtn.style.color = recipientsPanelOpen ? 'white' : '#667eea';
        recipientsBtn.style.borderColor = recipientsPanelOpen ? 'transparent' : 'rgba(102, 126, 234, 0.3)';
      });
    }

    // To input search
    if (toInput) {
      toInput.addEventListener('input', () => {
        clearTimeout(recipientSearchTimeout);
        recipientSearchTimeout = setTimeout(() => {
          searchRecipientsFor(toInput.value.trim(), toResults, 'to');
        }, 300);
      });
      toInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          const val = toInput.value.trim();
          if (val && val.includes('@') && !recipientTo.some(r => r.email === val)) {
            recipientTo.push({ name: val.split('@')[0], email: val, slack_id: '' });
            toChipsContainer.appendChild(createRecipientChip({ name: val.split('@')[0], email: val }, 'to', { removable: true, clickToCC: true }));
            toInput.value = '';
            toResults.style.display = 'none';
          }
        }
      });
    }

    // CC input search
    if (ccInput) {
      ccInput.addEventListener('input', () => {
        clearTimeout(recipientSearchTimeout);
        recipientSearchTimeout = setTimeout(() => {
          searchRecipientsFor(ccInput.value.trim(), ccResults, 'cc');
        }, 300);
      });
      ccInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          const val = ccInput.value.trim();
          if (val && val.includes('@') && !recipientCc.some(r => r.email === val)) {
            recipientCc.push({ name: val.split('@')[0], email: val, slack_id: '' });
            ccChipsContainer.appendChild(createRecipientChip({ name: val.split('@')[0], email: val }, 'cc'));
            ccInput.value = '';
            ccResults.style.display = 'none';
          }
        }
      });
    }

    // BCC input search
    if (bccInput) {
      bccInput.addEventListener('input', () => {
        clearTimeout(recipientSearchTimeout);
        recipientSearchTimeout = setTimeout(() => {
          searchRecipientsFor(bccInput.value.trim(), bccResults, 'bcc');
        }, 300);
      });
      bccInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          const val = bccInput.value.trim();
          if (val && val.includes('@') && !recipientBcc.some(r => r.email === val)) {
            recipientBcc.push({ name: val.split('@')[0], email: val, slack_id: '' });
            bccChipsContainer.appendChild(createRecipientChip({ name: val.split('@')[0], email: val }, 'bcc'));
            bccInput.value = '';
            bccResults.style.display = 'none';
          }
        }
      });
    }

    // Hide results on click outside
    document.addEventListener('click', (e) => {
      if (ccResults && !ccResults.contains(e.target) && e.target !== ccInput) ccResults.style.display = 'none';
      if (bccResults && !bccResults.contains(e.target) && e.target !== bccInput) bccResults.style.display = 'none';
    });

    // Function to populate recipients from participants data
    function populateRecipients(participants) {
      if (!participants || participants.length === 0) return;

      const toParts = slider._toParticipants || [];
      const ccParts = slider._ccParticipants || [];

      // If we have separated to/cc lists, use them
      if (toParts.length > 0 || ccParts.length > 0) {
        // Populate To
        toChipsContainer.innerHTML = '';
        toParts.forEach(p => {
          recipientTo.push({ name: p.name || '', email: p.email || '', slack_id: p.slack_id || '' });
          toChipsContainer.appendChild(createRecipientChip(
            { name: p.name || '', email: p.email || '', slack_id: p.slack_id || '' },
            'to',
            { removable: true, clickToCC: true }
          ));
        });
        // Populate CC
        ccChipsContainer.innerHTML = '';
        ccParts.forEach(p => {
          recipientCc.push({ name: p.name || '', email: p.email || '', slack_id: p.slack_id || '' });
          ccChipsContainer.appendChild(createRecipientChip(
            { name: p.name || '', email: p.email || '', slack_id: p.slack_id || '' },
            'cc'
          ));
        });
      } else {
        // Fallback: all participants go to To (backward compatibility)
        toChipsContainer.innerHTML = '';
        participants.forEach(p => {
          recipientTo.push({ name: p.name || '', email: p.email || '', slack_id: p.slack_id || '' });
          toChipsContainer.appendChild(createRecipientChip(
            { name: p.name || '', email: p.email || '', slack_id: p.slack_id || '' },
            'to',
            { removable: true, clickToCC: true }
          ));
        });
      }
    }

    // Auto-add @mentioned users to CC when they are inserted via mention
    // Wrap the original insertMention to also add to CC
    const _origInsertMention = insertMention;
    const insertMentionWithCC = (name, email, slackId) => {
      _origInsertMention(name, email, slackId);
      // Auto-add to CC if email exists and not already in CC or To
      if (email && !recipientCc.some(r => r.email === email) && !recipientTo.some(r => r.email === email)) {
        recipientCc.push({ name, email, slack_id: slackId || '' });
        ccChipsContainer.appendChild(createRecipientChip({ name, email, slack_id: slackId || '' }, 'cc'));
      }
    };
    // Override the mention insertion in the reply input's event flow
    // We patch the replyInput's mention handler by overriding the dropdown click
    // (The actual insertMention is called from dropdown click handlers set up earlier,
    //  so we re-bind those to use our wrapper)
    if (mentionDropdown) {
      let isPatching = false;
      const observer = new MutationObserver(() => {
        if (isPatching) return; // Prevent infinite loop
        isPatching = true;
        mentionDropdown.querySelectorAll('.mention-item:not([data-cc-patched])').forEach(item => {
          item.setAttribute('data-cc-patched', 'true');
          item.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const name = item.dataset?.name || item.querySelector('.mention-name')?.textContent || '';
            const email = item.dataset?.email || item.querySelector('.mention-email')?.textContent || '';
            const slackId = item.dataset?.slackId || '';
            insertMentionWithCC(name, email, slackId);
          });
        });
        isPatching = false;
      });
      observer.observe(mentionDropdown, { childList: true, subtree: true });
    }

    // Now populate recipients from stored participants
    if (slider._participants && slider._participants.length > 0) {
      populateRecipients(slider._participants);
    }

    // Show recipients button only for email tasks (hidden by default in HTML)
    if (slider._isEmailTask) {
      if (recipientsBtn) recipientsBtn.style.display = 'flex';
    }
    // ========== END RECIPIENTS MANAGEMENT ==========

    sendBtn.addEventListener('click', async () => {
      // Check if Oracle overlay is active — if so, send as follow-up to Oracle instead of thread
      if (slider._isOracleActive && slider._sendToOracle) {
        const followUpText = replyInput.innerText.trim();
        if (!followUpText) return;
        replyInput.innerHTML = '';
        await slider._sendToOracle(followUpText, true);
        return;
      }
      // Get HTML content from contenteditable div
      const replyHtml = replyInput.innerHTML.trim();
      // Also get plain text for fallback
      const replyText = replyInput.innerText.trim();

      // Build email-safe HTML: convert mention-tags to styled spans, wrap in proper HTML
      const getEmailSafeHtml = () => {
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = replyHtml;
        // Convert mention-tag spans to styled inline spans for email clients
        tempDiv.querySelectorAll('.mention-tag').forEach(tag => {
          const styledSpan = document.createElement('span');
          styledSpan.style.cssText = 'color: #1a73e8; font-weight: 600;';
          styledSpan.textContent = tag.textContent;
          tag.replaceWith(styledSpan);
        });
        // Wrap in a basic email-safe container with inline styles
        return `<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #202124; line-height: 1.6;">${tempDiv.innerHTML}</div>`;
      };
      const emailSafeHtml = getEmailSafeHtml();
      // Get Slack-formatted text (converts @Name to <@SLACK_ID>)
      const replyTextSlack = getSlackFormattedText();

      if (!replyText && pendingAttachments.length === 0) return;

      sendBtn.disabled = true;
      sendBtn.innerHTML = '<span style="font-size: 14px;">⏳</span>';

      // Get mentioned users from the input
      const mentionedUsersInReply = getMentionedUsersFromInput();
      console.log('Mentioned users:', mentionedUsersInReply);
      console.log('Slack formatted text:', replyTextSlack);

      try {
        // Determine the reply action based on task type
        const replyAction = isDriveTask ? 'reply_drive_comment' : 'reply_to_thread';

        const payload = createAuthenticatedPayload({
          action: replyAction,
          todo_id: todoId,
          message_link: todo.message_link,
          reply_text: isDriveTask ? replyText : replyTextSlack, // Use plain text for Drive, Slack-formatted for others
          reply_text_plain: replyText, // Keep plain text as backup
          reply_html: emailSafeHtml,
          mentioned_users: mentionedUsersInReply.length > 0 ? mentionedUsersInReply : undefined,
          to_emails: recipientTo.length > 0 ? recipientTo.map(r => r.email).filter(Boolean) : undefined,
          cc_emails: recipientCc.length > 0 ? recipientCc.map(r => r.email).filter(Boolean) : undefined,
          bcc_emails: recipientBcc.length > 0 ? recipientBcc.map(r => r.email).filter(Boolean) : undefined,
          attachments: pendingAttachments.length > 0 ? pendingAttachments : undefined,
          timestamp: new Date().toISOString(),
          source: 'oracle-chrome-extension-newtab'
        });

        const response = await fetch(WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        if (response.ok) {
          sendBtn.style.background = 'linear-gradient(45deg, #27ae60, #2ecc71)';
          sendBtn.innerHTML = '<span style="font-size: 14px;">✓</span>';
          replyInput.innerHTML = '';

          // Clear mentioned users
          mentionedUsers = [];

          // Clear attachments
          pendingAttachments = [];
          attachmentsList.innerHTML = '';
          attachmentsContainer.style.display = 'none';

          // Close recipients panel
          if (recipientsPanelOpen) {
            recipientsPanelOpen = false;
            recipientsPanel.style.display = 'none';
            if (recipientsBtn) {
              recipientsBtn.style.background = 'rgba(102, 126, 234, 0.1)';
              recipientsBtn.style.color = '#667eea';
              recipientsBtn.style.borderColor = 'rgba(102, 126, 234, 0.3)';
            }
          }

          // Refresh conversation after successful reply
          setTimeout(async () => {
            sendBtn.style.background = 'linear-gradient(45deg, #667eea, #764ba2)';
            sendBtn.innerHTML = '<span style="font-size: 16px;">↗</span>';
            sendBtn.disabled = false;

            // Show loading state in messages container
            const messagesContainer = slider.querySelector('.transcript-messages-container');
            if (messagesContainer) {
              const currentContent = messagesContainer.innerHTML;
              messagesContainer.innerHTML += `
                <div class="refresh-indicator" style="display: flex; align-items: center; justify-content: center; padding: 12px; gap: 8px;">
                  <div class="spinner" style="width: 16px; height: 16px; border: 2px solid rgba(102, 126, 234, 0.2); border-top-color: #667eea; border-radius: 50%; animation: spin 0.8s linear infinite;"></div>
                  <span style="font-size: 12px; color: #7f8c8d;">Refreshing conversation...</span>
                </div>
              `;

              try {
                // Fetch updated conversation (use correct action based on task type)
                const refreshAction = isDriveTask ? 'fetch_comment_details' : 'fetch_task_details';
                const refreshResponse = await fetch(WEBHOOK_URL, {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify(createAuthenticatedPayload({
                    action: refreshAction,
                    todo_id: todoId,
                    message_link: todo.message_link,
                    timestamp: new Date().toISOString(),
                    source: 'oracle-chrome-extension-newtab'
                  }))
                });

                if (refreshResponse.ok) {
                  const data = await refreshResponse.json();

                  // Parse transcript messages from response
                  let messages = [];
                  const transcript = data.transcript || (Array.isArray(data) && data[0]?.transcript) || [];

                  if (transcript && transcript.length > 0) {
                    try {
                      const flatTranscript = transcript.flat();
                      messages = flatTranscript.map(msgStr => {
                        try {
                          return typeof msgStr === 'string' ? JSON.parse(msgStr) : msgStr;
                        } catch {
                          return null;
                        }
                      }).filter(m => m !== null);
                    } catch (e) {
                      console.error('Error parsing transcript:', e);
                    }
                  }

                  // Update message count
                  const countEl = slider.querySelector('.transcript-message-count');
                  if (countEl) {
                    countEl.textContent = messages.length > 0
                      ? `${messages.length} message${messages.length !== 1 ? 's' : ''}`
                      : 'No messages';
                  }

                  // Re-render messages
                  if (messages.length > 0) {
                    messagesContainer.innerHTML = '';
                    // Check current dark mode state for refresh
                    const refreshDarkMode = document.body.classList.contains('dark-mode');
                    messages.forEach(msg => {
                      const msgDiv = document.createElement('div');
                      msgDiv.style.cssText = 'display: flex; flex-direction: column; gap: 6px;';

                      const msgHeader = document.createElement('div');
                      msgHeader.style.cssText = 'display: flex; align-items: center; gap: 8px;';
                      msgHeader.innerHTML = `
                        <div style="width: 32px; height: 32px; background: linear-gradient(135deg, #667eea, #764ba2); border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: 12px; font-weight: 600;">${(msg.message_from || 'U').charAt(0).toUpperCase()}</div>
                        <div style="flex: 1; min-width: 0;">
                          <div style="font-weight: 600; font-size: 13px; color: ${refreshDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(msg.message_from || 'Unknown')}</div>
                          <div style="font-size: 11px; color: ${refreshDarkMode ? '#666' : '#95a5a6'};">${escapeHtml(formatTimeAgoFresh(msg.time || ''))}</div>
                        </div>
                      `;

                      const msgContent = document.createElement('div');
                      msgContent.style.cssText = `margin-left: 40px; padding: 12px 16px; background: ${refreshDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border-radius: 12px; border-top-left-radius: 4px; font-size: 14px; color: ${refreshDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.6; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word; overflow: visible; max-width: 100%;`;
                      const htmlContent = msg.message_html || '';
                      if (htmlContent) {
                        renderEmailInIframe(htmlContent, msgContent, refreshDarkMode);
                      } else if (isComplexEmailHtml(msg.message || '')) {
                        renderEmailInIframe(msg.message, msgContent, refreshDarkMode);
                      } else {
                        msgContent.innerHTML = formatMessageContent(msg.message || '');
                      }

                      msgDiv.appendChild(msgHeader);
                      msgDiv.appendChild(msgContent);

                      // Render attachments if present
                      if (msg.attachments && msg.attachments.length > 0) {
                        // Filter out generic "link" type attachments
                        const validAttachments = msg.attachments.filter(att => {
                          if (att.type === 'link' && (!att.name || att.name === 'Attachment')) {
                            return false;
                          }
                          if (!att.name && !att.url && !att.text) {
                            return false;
                          }
                          return true;
                        });

                        if (validAttachments.length > 0) {
                          const attachmentsDiv = document.createElement('div');
                          attachmentsDiv.style.cssText = 'margin-left: 40px; margin-top: 8px; display: flex; flex-wrap: wrap; gap: 8px;';

                          validAttachments.forEach(attachment => {
                            const attachEl = renderTranscriptAttachment(attachment);
                            if (attachEl) {
                              attachmentsDiv.appendChild(attachEl);
                            }
                          });

                          if (attachmentsDiv.children.length > 0) {
                            msgDiv.appendChild(attachmentsDiv);
                          }
                        }
                      }

                      messagesContainer.appendChild(msgDiv);
                    });

                    // Scroll to bottom to show new message
                    messagesContainer.scrollTop = messagesContainer.scrollHeight;
                  }
                }
              } catch (refreshError) {
                console.error('Error refreshing conversation:', refreshError);
                // Remove refresh indicator on error
                const indicator = messagesContainer.querySelector('.refresh-indicator');
                if (indicator) indicator.remove();
              }
            }
          }, 1500);
        } else {
          throw new Error('Failed to send reply');
        }
      } catch (error) {
        console.error('Error sending reply:', error);
        sendBtn.style.background = 'linear-gradient(45deg, #e74c3c, #c0392b)';
        sendBtn.innerHTML = '<span style="font-size: 14px;">✗</span>';
        setTimeout(() => {
          sendBtn.style.background = 'linear-gradient(45deg, #667eea, #764ba2)';
          sendBtn.innerHTML = '<span style="font-size: 16px;">↗</span>';
          sendBtn.disabled = false;
        }, 2000);
      }
    });

    // Enter key in textarea just adds newline (default behavior)
    // Only Shift+Enter or clicking Send button will submit
    // No keydown handler needed - textarea handles Enter naturally

    setTimeout(() => replyInput.focus(), 300);
  }

  function showDueByMenu(clockElement, todoId) {
    document.querySelectorAll('.due-by-menu').forEach(m => m.remove());
    const menu = document.createElement('div');
    menu.className = 'due-by-menu';
    menu.style.cssText = 'position: fixed; background: rgba(255,255,255,0.95); backdrop-filter: blur(10px); border: 1px solid rgba(225,232,237,0.6); border-radius: 10px; box-shadow: 0 4px 16px rgba(0,0,0,0.12); padding: 4px; z-index: 10000; min-width: 120px;';
    const header = document.createElement('div');
    header.textContent = 'Remind me in';
    header.style.cssText = 'padding: 8px 12px; font-size: 11px; font-weight: 600; color: #7f8c8d; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid rgba(225,232,237,0.4);';
    menu.appendChild(header);
    const options = [
      { label: '30 mins', minutes: 30 },
      { label: '1 hour', minutes: 60 },
      { label: '3 hours', minutes: 180 },
      { label: '6 hours', minutes: 360 },
      { label: '24 hours', minutes: 1440 },
      { label: 'Remove reminder', minutes: null, isRemove: true }
    ];
    options.forEach(opt => {
      const button = document.createElement('button');
      button.textContent = opt.label;
      button.style.cssText = 'width: 100%; padding: 10px 12px; border: none; background: transparent; text-align: left; cursor: pointer; border-radius: 6px; font-size: 13px; color: #2c3e50; transition: all 0.2s; font-weight: 500;' + (opt.isRemove ? ' border-top: 1px solid rgba(225,232,237,0.4); margin-top: 4px; color: #e74c3c;' : '');
      button.onmouseover = () => {
        button.style.background = opt.isRemove ? 'rgba(231,76,60,0.1)' : 'linear-gradient(135deg, rgba(102, 126, 234, 0.1), rgba(118, 75, 162, 0.1))';
        button.style.transform = 'translateX(2px)';
      };
      button.onmouseout = () => {
        button.style.background = 'transparent';
        button.style.transform = 'translateX(0)';
      };
      button.onclick = () => setDueBy(todoId, opt.minutes, menu);
      menu.appendChild(button);
    });
    const rect = clockElement.getBoundingClientRect();
    menu.style.top = (rect.bottom + 5) + 'px';
    menu.style.left = (rect.left - 120 + 28) + 'px';
    document.body.appendChild(menu);
    setTimeout(() => {
      document.addEventListener('click', function closeMenu(e) {
        if (!menu.contains(e.target) && e.target !== clockElement) {
          menu.remove();
          document.removeEventListener('click', closeMenu);
        }
      });
    }, 10);
  }

  async function setDueBy(todoId, minutes, menu) {
    menu.remove();
    const container = document.querySelector('.todos-container');
    showLoader(container);
    try {
      const dueByTime = minutes === null ? null : new Date(Date.now() + minutes * 60000).toISOString();
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'set_due_by',
          todo_id: todoId,
          due_by: dueByTime,
          timestamp: new Date().toISOString(),
          source: 'oracle-chrome-extension-newtab'
        }))
      });
      if (response.ok) {
        const activeFilter = document.querySelector('.filter-btn.active');
        await loadTodos(activeFilter?.dataset.filter || 'starred');
      } else {
        throw new Error('Failed to set due date');
      }
    } catch (error) {
      console.error('Error setting due date:', error);
      hideLoader(container);
      alert('Failed to set due date: ' + error.message);
    }
  }

  async function handleTodoAction(action, todoId, element) {
    try {
      element.style.transform = 'scale(0.9)';
      setTimeout(() => element.style.transform = '', 150);
      if (action === 'edit') {
        // Set flag to prevent auto-refresh
        isEditModeActive = true;

        const todoItem = element.closest('.todo-item');
        const contentDiv = todoItem.querySelector('.todo-content');
        const titleDiv = contentDiv.querySelector('.todo-title');
        const textDiv = contentDiv.querySelector('.todo-text');

        // Get original values
        const origTitle = titleDiv ? titleDiv.textContent.trim() : '';
        const origText = textDiv.textContent.trim();

        // Create edit form
        const editContainer = document.createElement('div');
        editContainer.style.cssText = 'width: 100%; display: flex; flex-direction: column; gap: 8px;';

        const titleInput = document.createElement('input');
        titleInput.type = 'text';
        titleInput.value = origTitle;
        titleInput.placeholder = 'Task Title (optional)';
        titleInput.style.cssText = 'width: 100%; padding: 8px; border: 2px solid #667eea; border-radius: 6px; font-size: 14px; font-family: inherit; font-weight: 600;';

        const textarea = document.createElement('textarea');
        textarea.value = origText;
        textarea.placeholder = 'Task Description';
        textarea.style.cssText = 'width: 100%; padding: 8px; border: 2px solid #667eea; border-radius: 6px; font-size: 13px; font-family: inherit; resize: vertical; min-height: 60px; line-height: 1.4;';

        editContainer.appendChild(titleInput);
        editContainer.appendChild(textarea);

        // Replace content with edit form
        if (titleDiv) titleDiv.style.display = 'none';
        textDiv.innerHTML = '';
        textDiv.appendChild(editContainer);

        titleInput.focus();
        textarea.style.height = 'auto';
        textarea.style.height = textarea.scrollHeight + 'px';

        element.innerHTML = '✓';
        element.style.background = 'linear-gradient(45deg, #27ae60, #2ecc71)';
        element.style.color = 'white';

        const save = async () => {
          isEditModeActive = false; // Reset flag when edit completes
          const newTitle = titleInput.value.trim();
          const newText = textarea.value.trim();

          if (!newText || (newTitle === origTitle && newText === origText)) {
            // Restore original
            if (titleDiv) {
              titleDiv.style.display = '';
              titleDiv.textContent = origTitle;
            }
            textDiv.textContent = origText;
            element.innerHTML = '✏️';
            element.style.background = '';
            element.style.color = '';
            return;
          }

          try {
            const response = await fetch(WEBHOOK_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(createAuthenticatedPayload({
                action: 'update_task_text',
                todo_id: todoId,
                new_text: newText,
                new_title: newTitle,
                timestamp: new Date().toISOString(),
                source: 'oracle-chrome-extension-newtab'
              }))
            });
            if (response.ok) {
              // Check which tab is active and reload accordingly
              const fyiTab = document.getElementById('fyiTab');
              if (fyiTab && fyiTab.classList.contains('active')) {
                await loadFYI();
              } else {
                const activeFilter = document.querySelector('.filter-btn.active');
                await loadTodos(activeFilter?.dataset.filter || 'starred');
              }
            } else {
              throw new Error('Failed to save');
            }
          } catch (error) {
            console.error('Error saving:', error);
            if (titleDiv) {
              titleDiv.style.display = '';
              titleDiv.textContent = origTitle;
            }
            textDiv.textContent = origText;
            element.innerHTML = '✏️';
            element.style.background = '';
            element.style.color = '';
          }
        };

        const handleClick = async (e) => {
          e.stopPropagation();
          await save();
        };
        element.addEventListener('click', handleClick);

        const handleKeyDown = async e => {
          if (e.key === 'Enter' && !e.shiftKey && e.target === titleInput) {
            e.preventDefault();
            textarea.focus();
          } else if (e.key === 'Enter' && !e.shiftKey && e.target === textarea) {
            e.preventDefault();
            await save();
          } else if (e.key === 'Escape') {
            isEditModeActive = false; // Reset flag on cancel
            if (titleDiv) {
              titleDiv.style.display = '';
              titleDiv.textContent = origTitle;
            }
            textDiv.textContent = origText;
            element.innerHTML = '✏️';
            element.style.background = '';
            element.style.color = '';
            element.removeEventListener('click', handleClick);
          }
        };

        titleInput.onkeydown = handleKeyDown;
        textarea.onkeydown = handleKeyDown;
        textarea.oninput = () => {
          textarea.style.height = 'auto';
          textarea.style.height = textarea.scrollHeight + 'px';
        };
        return;
      }
      let basePayload = {
        todo_id: todoId,
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-newtab'
      };
      if (action === 'toggle-complete') {
        const isChecked = element.classList.contains('checked');
        basePayload.action = 'toggle_todo';
        basePayload.status = isChecked ? 0 : 1;

        const todoItem = element.closest('.todo-item');
        const todoId = element.dataset.todoId;

        if (isChecked) {
          // Uncompleting - just update visually
          element.classList.remove('checked');
          element.innerHTML = '';
          todoItem.classList.remove('completed');
        } else {
          // Completing - play slide-out animation and remove immediately
          element.classList.add('checked');
          element.innerHTML = '✓';
          todoItem.classList.add('completed');
          todoItem.classList.add('completing');

          // Remove from DOM after animation
          setTimeout(() => {
            todoItem.remove();
          }, 400);

          // Update local arrays immediately (optimistic update)
          allTodos = allTodos.filter(t => t.id != todoId);
          allFyiItems = allFyiItems.filter(t => t.id != todoId);

          // Reset read state so if task comes back, it will be unread (yellow)
          markTaskAsUnread(todoId);

          // Update counts
          if (typeof updateTabCounts === 'function') {
            updateTabCounts();
          }

          // Send request to backend in background
          if (isAuthenticated && userData) {
            fetch(WEBHOOK_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(createAuthenticatedPayload(basePayload))
            }).catch(err => console.error('Error updating task:', err));
          }
          return; // Exit early
        }
      } else if (action === 'toggle-star') {
        const isStarred = element.classList.contains('active');
        basePayload.action = 'toggle_star';
        basePayload.starred = isStarred ? 0 : 1;
        if (isStarred) {
          element.classList.remove('active');
        } else {
          element.classList.add('active');
        }
      }
      if (isAuthenticated && userData) {
        // Send request without showing loader
        await fetch(WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(createAuthenticatedPayload(basePayload))
        });
        // Use animated refresh
        const currentFilter = document.querySelector('.filter-btn.active');
        await loadTodosAnimated(currentFilter?.dataset.filter || 'starred');
      }
    } catch (error) {
      console.error('Error:', error);
    }
  }

  function showEmptyTodosState(filter) {
    const container = document.querySelector('.todos-container');
    const emptyState = container.querySelector('.empty-state');
    const todosList = container.querySelector('.todos-list');
    hideLoader(container);
    if (todosList) todosList.style.display = 'none';
    if (emptyState) {
      emptyState.style.display = 'flex';
      let message = filter === 'starred' ? 'No starred todos' : filter === 'active' ? 'No active todos' : 'No todos found';
      emptyState.innerHTML = '<h3>' + message + '</h3><p>Click Refresh to check again</p>';
    }
  }

  function setupTodoList() {
    const refreshBtn = document.querySelector('.refresh-btn');
    const filterBtns = document.querySelectorAll('.filter-btn');
    refreshBtn?.addEventListener('click', () => {
      const activeFilter = document.querySelector('.filter-btn.active');
      loadTodos(activeFilter?.dataset.filter || 'starred');
    });
    filterBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        filterBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        updateFilterSlider();
        loadTodos(btn.dataset.filter);
      });
    });
  }

  function setupBookmarks() {
    const refreshBtn = document.getElementById('refreshBookmarksBtn');
    refreshBtn?.addEventListener('click', () => { loadBookmarks() });
  }

  // Message formatting — delegated to OracleMessageFormat shared component
  const { formatMessageContent, sanitizeHtml, isComplexEmailHtml, renderEmailInIframe } = window.OracleMessageFormat;
  // escapeHtml available from top-level alias
  // Render attachment in transcript slider
  // Attachment rendering — delegated to OracleMessageFormat shared component
  const { fetchGmailAttachment, renderTranscriptAttachment, showAttachmentPreview } = window.OracleMessageFormat;

  // formatDate & formatDueBy — delegated to Oracle shared component (top-level aliases)

  // Convert Slack emoji shortcodes to unicode characters
  function convertSlackEmoji(name) {
    const emojiMap = {
      // Hands & Gestures
      '+1': '👍', '-1': '👎', 'thumbsup': '👍', 'thumbsdown': '👎',
      'ok_hand': '👌', 'clap': '👏', 'wave': '👋', 'raised_hands': '🙌',
      'pray': '🙏', 'muscle': '💪', 'point_up': '☝️', 'point_right': '👉',
      'point_left': '👈', 'point_down': '👇', 'point_up_2': '👆',
      'handshake': '🤝', 'palms_up_together': '🤲', 'open_hands': '👐',
      'fist': '✊', 'facepunch': '👊', 'punch': '👊', 'v': '✌️',
      'crossed_fingers': '🤞', 'love_you_gesture': '🤟', 'metal': '🤘',
      'call_me_hand': '🤙', 'writing_hand': '✍️', 'selfie': '🤳',
      'nail_care': '💅', 'pinching_hand': '🤏', 'pinched_fingers': '🤌',
      'heart_hands': '🫶', 'index_pointing_at_the_viewer': '🫵', 'saluting_face': '🫡',
      // Faces - Smiling
      'smile': '😄', 'laughing': '😆', 'satisfied': '😆', 'joy': '😂',
      'grinning': '😀', 'smiley': '😃', 'grin': '😁', 'sweat_smile': '😅',
      'rofl': '🤣', 'slightly_smiling_face': '🙂', 'upside_down_face': '🙃',
      'wink': '😉', 'blush': '😊', 'innocent': '😇', 'smiling_face_with_three_hearts': '🥰',
      'heart_eyes': '😍', 'star_struck': '🤩', 'kissing_heart': '😘',
      'kissing': '😗', 'kissing_smiling_eyes': '😙', 'kissing_closed_eyes': '😚',
      'yum': '😋', 'stuck_out_tongue': '😛', 'stuck_out_tongue_winking_eye': '😜',
      'zany_face': '🤪', 'stuck_out_tongue_closed_eyes': '😝', 'money_mouth_face': '🤑',
      'hugging_face': '🤗', 'hugs': '🤗', 'melting_face': '🫠',
      // Faces - Neutral & Thinking
      'thinking_face': '🤔', 'thinking': '🤔', 'neutral_face': '😐',
      'expressionless': '😑', 'no_mouth': '😶', 'face_in_clouds': '😶‍🌫️',
      'dotted_line_face': '🫥', 'smirk': '😏', 'unamused': '😒',
      'face_with_rolling_eyes': '🙄', 'rolling_eyes': '🙄', 'grimacing': '😬',
      'face_exhaling': '😮‍💨', 'lying_face': '🤥', 'raised_eyebrow': '🤨',
      'face_with_monocle': '🧐', 'zipper_mouth_face': '🤐',
      // Faces - Negative
      'confused': '😕', 'worried': '😟', 'slightly_frowning_face': '🙁',
      'open_mouth': '😮', 'hushed': '😯', 'astonished': '😲', 'flushed': '😳',
      'pleading_face': '🥺', 'fearful': '😨', 'cold_sweat': '😰',
      'cry': '😢', 'sob': '😭', 'scream': '😱', 'persevere': '😣',
      'disappointed': '😞', 'sweat': '😓', 'weary': '😩', 'tired_face': '😫',
      'yawning_face': '🥱', 'angry': '😠', 'rage': '😡', 'pout': '😡',
      'face_with_symbols_on_mouth': '🤬', 'cursing_face': '🤬',
      'skull': '💀', 'skull_and_crossbones': '☠️',
      // Faces - Sick & Special
      'mask': '😷', 'face_with_thermometer': '🤒', 'face_with_head_bandage': '🤕',
      'nauseated_face': '🤢', 'face_vomiting': '🤮', 'sneezing_face': '🤧',
      'hot_face': '🥵', 'cold_face': '🥶', 'woozy_face': '🥴',
      'dizzy_face': '😵', 'exploding_head': '🤯', 'partying_face': '🥳',
      'sunglasses': '😎', 'nerd_face': '🤓', 'disguised_face': '🥸',
      'cowboy_hat_face': '🤠', 'clown_face': '🤡', 'sleeping': '😴',
      'drooling_face': '🤤', 'shushing_face': '🤫', 'face_with_peeking_eye': '🫣',
      'see_no_evil': '🙈', 'hear_no_evil': '🙉', 'speak_no_evil': '🙊',
      'poop': '💩', 'ghost': '👻', 'alien': '👽', 'robot_face': '🤖',
      'smiling_imp': '😈', 'imp': '👿',
      // Hearts & Love
      'heart': '❤️', 'red_heart': '❤️', 'orange_heart': '🧡', 'yellow_heart': '💛',
      'green_heart': '💚', 'blue_heart': '💙', 'purple_heart': '💜',
      'black_heart': '🖤', 'white_heart': '🤍', 'brown_heart': '🤎',
      'pink_heart': '🩷', 'broken_heart': '💔', 'heavy_heart_exclamation': '❣️',
      'two_hearts': '💕', 'revolving_hearts': '💞', 'heartbeat': '💓',
      'heartpulse': '💗', 'sparkling_heart': '💖', 'cupid': '💘',
      'gift_heart': '💝', 'heart_on_fire': '❤️‍🔥', 'mending_heart': '❤️‍🩹',
      // Symbols & Marks
      'white_check_mark': '✅', 'heavy_check_mark': '✔️',
      'ballot_box_with_check': '☑️', 'x': '❌', 'negative_squared_cross_mark': '❎',
      'warning': '⚠️', 'exclamation': '❗', 'question': '❓',
      'bangbang': '‼️', 'interrobang': '⁉️', 'grey_exclamation': '❕', 'grey_question': '❔',
      '100': '💯', 'anger': '💢', 'no_entry': '⛔', 'no_entry_sign': '🚫',
      'stop_sign': '🛑', 'infinity': '♾️', 'recycle': '♻️',
      // Stars, Fire & Sparkles
      'fire': '🔥', 'star': '⭐', 'star2': '🌟', 'sparkles': '✨',
      'dizzy': '💫', 'comet': '☄️',
      // Celebration
      'tada': '🎉', 'party_popper': '🎉', 'confetti_ball': '🎊',
      'trophy': '🏆', 'medal': '🏅', 'sports_medal': '🏅',
      'first_place_medal': '🥇', 'second_place_medal': '🥈', 'third_place_medal': '🥉',
      'crown': '👑', 'gem': '💎', 'balloon': '🎈', 'gift': '🎁',
      'ribbon': '🎀', 'birthday': '🎂', 'cake': '🍰', 'cupcake': '🧁',
      // Objects & Tools
      'rocket': '🚀', 'zap': '⚡', 'boom': '💥', 'collision': '💥',
      'bell': '🔔', 'no_bell': '🔕', 'mega': '📣', 'loudspeaker': '📢',
      'speech_balloon': '💬', 'thought_balloon': '💭',
      'memo': '📝', 'pencil': '✏️', 'pencil2': '✏️', 'pen': '🖊️',
      'book': '📖', 'bookmark': '🔖', 'books': '📚', 'notebook': '📓',
      'link': '🔗', 'paperclip': '📎', 'scissors': '✂️', 'pushpin': '📌',
      'lock': '🔒', 'unlock': '🔓', 'key': '🔑', 'old_key': '🗝️',
      'gear': '⚙️', 'wrench': '🔧', 'hammer': '🔨', 'tools': '🛠️',
      'nut_and_bolt': '🔩', 'chains': '⛓️',
      'hourglass': '⌛', 'hourglass_flowing_sand': '⏳', 'stopwatch': '⏱️',
      'alarm_clock': '⏰', 'watch': '⌚',
      'calendar': '📅', 'date': '📅', 'spiral_calendar': '🗓️',
      'chart_with_upwards_trend': '📈', 'chart_with_downwards_trend': '📉', 'bar_chart': '📊',
      'email': '📧', 'envelope': '✉️', 'mailbox': '📬', 'package': '📦',
      'computer': '💻', 'keyboard': '⌨️', 'iphone': '📱', 'camera': '📷',
      'tv': '📺', 'microphone': '🎤', 'headphones': '🎧',
      'bulb': '💡', 'flashlight': '🔦', 'candle': '🕯️',
      'money_with_wings': '💸', 'dollar': '💵', 'moneybag': '💰', 'credit_card': '💳',
      'shield': '🛡️',
      // Arrows
      'arrow_up': '⬆️', 'arrow_down': '⬇️', 'arrow_right': '➡️', 'arrow_left': '⬅️',
      'arrow_upper_right': '↗️', 'arrow_lower_right': '↘️',
      'arrow_right_hook': '↪️', 'leftwards_arrow_with_hook': '↩️',
      'arrows_counterclockwise': '🔄', 'fast_forward': '⏩', 'rewind': '⏪',
      // Circles & Shapes
      'red_circle': '🔴', 'orange_circle': '🟠', 'yellow_circle': '🟡',
      'green_circle': '🟢', 'blue_circle': '🔵', 'purple_circle': '🟣',
      'white_circle': '⚪', 'black_circle': '⚫',
      'large_green_circle': '🟢', 'large_red_circle': '🔴',
      'white_large_square': '⬜', 'black_large_square': '⬛',
      'small_red_triangle': '🔺', 'small_red_triangle_down': '🔻',
      // Flags
      'checkered_flag': '🏁', 'triangular_flag_on_post': '🚩',
      'waving_white_flag': '🏳️', 'rainbow_flag': '🏳️‍🌈',
      // Nature & Weather
      'sunny': '☀️', 'cloud': '☁️', 'umbrella': '☂️', 'rainbow': '🌈',
      'snowflake': '❄️', 'tornado': '🌪️', 'ocean': '🌊', 'droplet': '💧',
      'sweat_drops': '💦',
      'earth_americas': '🌎', 'earth_africa': '🌍', 'earth_asia': '🌏',
      'cherry_blossom': '🌸', 'rose': '🌹', 'sunflower': '🌻',
      'seedling': '🌱', 'four_leaf_clover': '🍀', 'herb': '🌿',
      'fallen_leaf': '🍂', 'maple_leaf': '🍁', 'mushroom': '🍄',
      // Animals
      'dog': '🐕', 'cat': '🐈', 'monkey_face': '🐵', 'monkey': '🐒',
      'bear': '🐻', 'panda_face': '🐼', 'penguin': '🐧',
      'bird': '🐦', 'eagle': '🦅', 'owl': '🦉', 'frog': '🐸',
      'snake': '🐍', 'dragon': '🐉', 'turtle': '🐢', 'whale': '🐳',
      'dolphin': '🐬', 'fish': '🐟', 'octopus': '🐙',
      'butterfly': '🦋', 'bee': '🐝', 'honeybee': '🐝', 'ant': '🐜',
      'ladybug': '🐞', 'unicorn_face': '🦄', 'unicorn': '🦄',
      'horse': '🐴', 'pig': '🐷', 'cow': '🐮', 'chicken': '🐔',
      'lion_face': '🦁', 'fox_face': '🦊', 'wolf': '🐺', 'tiger': '🐯',
      'elephant': '🐘', 'rabbit': '🐰',
      // Food & Drink
      'coffee': '☕', 'tea': '🍵', 'pizza': '🍕', 'hamburger': '🍔',
      'fries': '🍟', 'taco': '🌮', 'sushi': '🍣', 'ramen': '🍜',
      'ice_cream': '🍨', 'doughnut': '🍩', 'cookie': '🍪', 'chocolate_bar': '🍫',
      'popcorn': '🍿', 'avocado': '🥑', 'watermelon': '🍉', 'grapes': '🍇',
      'strawberry': '🍓', 'peach': '🍑', 'apple': '🍎', 'banana': '🍌',
      'pineapple': '🍍', 'lemon': '🍋', 'cherries': '🍒',
      'beer': '🍺', 'beers': '🍻', 'wine_glass': '🍷', 'cocktail': '🍸',
      'champagne': '🍾', 'bubble_tea': '🧋',
      // Thank you & Appreciation (common Slack reactions)
      'thank_you': '🙏', 'thanks': '🙏', 'thankyou': '🙏', 'ty': '🙏',
      'thank-you': '🙏', 'clapping': '👏',
      // Common Slack aliases and custom emoji fallbacks
      'plusone': '👍', 'plus1': '👍', 'thumbsup_all': '👍',
      'beer_cheers': '🍻', 'beers_cheers': '🍻', 'cheers': '🍻', 'clinking_glasses': '🥂',
      'fistbump': '🤜🤛', 'fist_bump': '🤜🤛', 'brofist': '🤜🤛',
      'this': '👆', 'point-up': '☝️', 'yes': '✅', 'no': '❌',
      'lgtm': '👍', 'shipit': '🚀', 'ship_it': '🚀', 'ship-it': '🚀',
      'parrot': '🦜', 'party_parrot': '🦜', 'party-parrot': '🦜',
      'mindblown': '🤯', 'mind_blown': '🤯', 'mind-blown': '🤯',
      'blob_wave': '👋', 'blob-wave': '👋', 'meow_wave': '👋',
      'blob_thumbsup': '👍', 'blob_clap': '👏',
      'rolling_on_the_floor_laughing': '🤣', 'face_with_tears_of_joy': '😂',
      'slightly_smiling': '🙂', 'white-check-mark': '✅',
      'heavy_plus_sign': '➕', 'heavy_minus_sign': '➖',
      'raised_hand': '✋', 'hand': '✋', 'raised_back_of_hand': '🤚',
      'the_horns': '🤘', 'sign_of_the_horns': '🤘',
      'thumbs_up': '👍', 'thumbs_down': '👎',
      'hooray': '🎉', 'celebration': '🎉', 'congrats': '🎉',
      'done': '✅', 'approved': '✅', 'merged': '✅',
      'wfh': '🏠', 'remote': '🏠', 'ack': '👍', 'noted': '📝',
      // People & Body
      'eyes': '👀', 'eye': '👁️', 'brain': '🧠',
      'speaking_head': '🗣️', 'bust_in_silhouette': '👤', 'busts_in_silhouette': '👥',
      'footprints': '👣', 'lips': '👄', 'tongue': '👅', 'ear': '👂',
      'facepalm': '🤦', 'shrug': '🤷', 'raising_hand': '🙋',
      'dancer': '💃', 'man_dancing': '🕺',
      'runner': '🏃', 'running': '🏃', 'walking': '🚶',
      'ninja': '🥷', 'astronaut': '🧑‍🚀',
      // Transport
      'car': '🚗', 'taxi': '🚕', 'bus': '🚌', 'airplane': '✈️',
      'ship': '🚢', 'bike': '🚲', 'house': '🏠', 'office': '🏢',
      'rotating_light': '🚨',
      // Misc
      'zzz': '💤', 'mag': '🔍', 'mag_right': '🔎',
      'musical_note': '🎵', 'notes': '🎶', 'guitar': '🎸',
      'soccer': '⚽', 'basketball': '🏀', 'football': '🏈',
      'dart': '🎯', 'video_game': '🎮', 'puzzle_piece': '🧩', 'game_die': '🎲',
      'microscope': '🔬', 'telescope': '🔭', 'syringe': '💉', 'pill': '💊',
      'dna': '🧬', 'test_tube': '🧪', 'adhesive_bandage': '🩹',
      'door': '🚪', 'bed': '🛏️',
      // Numbers
      'one': '1️⃣', 'two': '2️⃣', 'three': '3️⃣', 'four': '4️⃣',
      'five': '5️⃣', 'six': '6️⃣', 'seven': '7️⃣', 'eight': '8️⃣',
      'nine': '9️⃣', 'zero': '0️⃣', 'keycap_ten': '🔟',
      'a': '🅰️', 'b': '🅱️', 'ab': '🆎',
      'cool': '🆒', 'free': '🆓', 'new': '🆕', 'ok': '🆗',
      'sos': '🆘', 'up': '🆙', 'vs': '🆚', 'information_source': 'ℹ️',
    };
    if (emojiMap[name]) return emojiMap[name];

    // Handle skin-tone variants: e.g. "pray::skin-tone-2" or "thumbsup::skin-tone-3"
    const skinToneMatch = name.match(/^(.+?)::skin-tone-(\d)$/);
    if (skinToneMatch) {
      const baseEmoji = emojiMap[skinToneMatch[1]];
      if (baseEmoji) {
        const toneModifiers = { '1': '\u{1F3FB}', '2': '\u{1F3FC}', '3': '\u{1F3FD}', '4': '\u{1F3FE}', '5': '\u{1F3FF}' };
        return baseEmoji + (toneModifiers[skinToneMatch[2]] || '');
      }
    }

    return `:${name}:`;
  }

  // formatTimeAgoFresh — transcript-specific time formatter (kept locally)
  function formatTimeAgoFresh(timeString) {
    if (!timeString) return '';
    const trimmedTime = timeString.trim();
    let date;
    if (trimmedTime.match(/^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}/)) {
      if (trimmedTime.includes('Z') || trimmedTime.match(/[+-]\d{2}:\d{2}$/)) date = new Date(trimmedTime);
      else date = new Date(trimmedTime + 'Z');
    } else {
      const match = trimmedTime.match(/(\d{1,2})\s+(\w{3,})\s+(\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?/i);
      if (match) {
        const [, day, month, year, hour, min, sec, ampm] = match;
        const months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
        let h = parseInt(hour);
        if (ampm) { if (ampm.toLowerCase() === 'pm' && h !== 12) h += 12; else if (ampm.toLowerCase() === 'am' && h === 12) h = 0; }
        const mi = months[month.toLowerCase().substring(0, 3)];
        if (mi !== undefined) date = new Date(parseInt(year), mi, parseInt(day), h, parseInt(min), parseInt(sec || 0));
      }
      if (!date || isNaN(date.getTime())) date = new Date(trimmedTime);
    }
    if (!date || isNaN(date.getTime())) return trimmedTime;
    const now = new Date(), diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000), diffHours = Math.floor(diffMs / 3600000), diffDays = Math.floor(diffMs / 86400000), diffWeeks = Math.floor(diffDays / 7), diffMonths = Math.floor(diffDays / 30);
    if (diffMins < 0) return 'Just now';
    if (diffMins < 1) return 'Just now';
    if (diffMins < 60) return `${diffMins} min${diffMins > 1 ? 's' : ''} ago`;
    if (diffHours < 24) return `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
    if (diffDays < 7) return `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
    if (diffWeeks < 4) return `${diffWeeks} week${diffWeeks > 1 ? 's' : ''} ago`;
    if (diffMonths < 12) return `${diffMonths} month${diffMonths > 1 ? 's' : ''} ago`;
    return date.toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', day: 'numeric', month: 'short', year: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true });
  }

  switchMode();
  setupTodoList();
  setupBookmarks();
  setupDocuments();

  // Check if 3-column layout and load both Action and FYI
  const isThreeCol = document.querySelector('.three-column-layout') !== null;

  if (isThreeCol) {
    // Load all three columns on startup
    loadTodos('starred');
    loadFYI();
    loadNotes();
  } else {
    // Legacy: Auto-load starred todos when page loads (Action tab is active by default)
    loadTodos('starred');
  }

  window.loadTodos = loadTodos;
  window.loadFYI = loadFYI;
  window.loadTodosAnimated = loadTodosAnimated;
  window.loadFYIAnimated = loadFYIAnimated;
  window.displayTodos = displayTodos;
  window.displayTodosAnimated = displayTodosAnimated;
  window.displayFYI = displayFYI;
  window.loadBookmarks = loadBookmarks;
  window.loadNotes = loadNotes;
  window.loadDocuments = loadDocuments;
  window.showTranscriptSlider = showTranscriptSlider;
  window.updateTodoField = updateTodoField;
  window.showNoteForm = showNoteForm;
  window.hideNoteForm = hideNoteForm;
  window.showNoteViewer = window.OracleNotes.showNoteViewer;

  // Expose functions needed for Ably silent fetch
  window.updateBadge = updateBadge;
  window.updateTabCounts = updateTabCounts;
  window.buildMeetingsAccordion = buildMeetingsAccordion;
  window.setupMeetingsAccordion = setupMeetingsAccordion;
  window.buildDocumentsAccordion = buildDocumentsAccordion;
  window.setupDocumentsAccordion = setupDocumentsAccordion;

  // Setup global updates badge (unified "X new updates" next to Oracle subtitle)
  setupGlobalUpdatesBadge();

  // Dark mode is now handled by setupDarkModeGlobal() in DOMContentLoaded

  // Setup Documents functionality
  function setupDocuments() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const refreshBtn = document.getElementById('refreshDocumentsBtn');

    // Click to upload
    uploadArea?.addEventListener('click', () => fileInput?.click());

    // Drag and drop
    uploadArea?.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadArea.classList.add('dragging');
    });

    uploadArea?.addEventListener('dragleave', () => {
      uploadArea.classList.remove('dragging');
    });

    uploadArea?.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadArea.classList.remove('dragging');
      const files = e.dataTransfer.files;
      if (files.length > 0) handleFileUpload(files[0]);
    });

    // File selection
    fileInput?.addEventListener('change', (e) => {
      if (e.target.files.length > 0) handleFileUpload(e.target.files[0]);
    });

    // Refresh button
    refreshBtn?.addEventListener('click', loadDocuments);
  }

  async function handleFileUpload(file) {
    const allowedTypes = [
      'application/pdf',
      'text/plain',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    ];

    const allowedExtensions = ['.pdf', '.txt', '.doc', '.docx'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!allowedTypes.includes(file.type) && !allowedExtensions.includes(fileExtension)) {
      alert('Invalid file type. Please upload PDF, TXT, DOC, or DOCX files only.');
      return;
    }

    // Max file size: 10MB
    const maxSize = 10 * 1024 * 1024;
    if (file.size > maxSize) {
      alert('File size exceeds 10MB limit.');
      return;
    }

    const uploadProgress = document.getElementById('uploadProgress');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');

    try {
      uploadProgress.style.display = 'block';
      progressText.textContent = 'Preparing upload...';

      // Convert file to base64
      const base64 = await fileToBase64(file);

      progressFill.style.width = '30%';
      progressText.textContent = 'Uploading...';

      const payload = {
        action: 'upload_document',
        filename: file.name,
        file_data: base64,
        file_type: file.type || getMimeType(fileExtension),
        file_size: file.size,
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-newtab',
        user_id: userData?.userId,
        authenticated: true
      };

      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      progressFill.style.width = '90%';

      if (response.ok) {
        progressFill.style.width = '100%';
        progressText.textContent = 'Upload complete!';

        setTimeout(() => {
          uploadProgress.style.display = 'none';
          progressFill.style.width = '0%';
          loadDocuments();
        }, 1500);
      } else {
        throw new Error('Upload failed');
      }
    } catch (error) {
      console.error('Upload error:', error);
      progressText.textContent = 'Upload failed: ' + error.message;
      setTimeout(() => {
        uploadProgress.style.display = 'none';
        progressFill.style.width = '0%';
      }, 3000);
    }
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(',')[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  function getMimeType(extension) {
    const mimeTypes = {
      '.pdf': 'application/pdf',
      '.txt': 'text/plain',
      '.doc': 'application/msword',
      '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
    return mimeTypes[extension] || 'application/octet-stream';
  }

  async function loadDocuments() {
    const documentsList = document.getElementById('documentsList');
    if (!documentsList) return;

    documentsList.innerHTML = '<div class="loading-overlay"><div class="spinner"></div></div>';

    try {
      const payload = {
        action: 'list_documents',
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-newtab',
        user_id: userData?.userId,
        authenticated: true
      };

      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        const data = await response.json();
        let documents = Array.isArray(data) ? data : (data.documents || []);

        // Filter out empty objects that n8n sometimes returns
        documents = documents.filter(doc => doc && doc.id && (doc.filename || doc.file_name));

        allDocuments = documents;
        displayDocuments(documents);
        updateTabCounts();
      } else {
        throw new Error('Failed to load documents');
      }
    } catch (error) {
      console.error('Error loading documents:', error);
      documentsList.innerHTML = `
        <div class="empty-state">
          <h3>Failed to load documents</h3>
          <p>${error.message}</p>
        </div>
      `;
    }
  }

  function displayDocuments(documents) {
    const documentsList = document.getElementById('documentsList');
    const documentsEmpty = document.getElementById('documentsEmpty');

    if (!documents || documents.length === 0) {
      documentsList.innerHTML = '';
      if (documentsEmpty) {
        documentsEmpty.style.display = 'flex';
      }
      return;
    }

    // Hide empty state when we have documents
    if (documentsEmpty) {
      documentsEmpty.style.display = 'none';
    }

    documentsList.innerHTML = documents.map((doc, index) => {
      const icon = getFileIcon(doc.filename || doc.file_name);
      const size = formatFileSize(doc.file_size || doc.size);
      const date = formatDate(doc.created_at || doc.uploaded_at);
      const docId = doc.id || doc.document_id;
      const docUrl = doc.url || doc.file_url;

      return `
        <div class="document-item" style="animation-delay: ${index * 0.1}s">
          <div class="document-icon">${icon}</div>
          <div class="document-info">
            <div class="document-name">${escapeHtml(doc.filename || doc.file_name)}</div>
            <div class="document-meta">
              <span class="document-size">${size}</span>
              <span class="document-date">${date}</span>
            </div>
          </div>
          <div class="document-actions">
            ${docUrl ? `<div class="document-download" data-url="${docUrl}" title="Download">⬇</div>` : ''}
            <div class="document-delete" data-doc-id="${docId}" title="Delete">🗑️</div>
          </div>
        </div>
      `;
    }).join('');

    // Add event listeners for delete and download buttons
    documentsList.querySelectorAll('.document-delete').forEach(btn => {
      btn.addEventListener('click', () => {
        const docId = btn.getAttribute('data-doc-id');
        deleteDocument(docId);
      });
    });

    documentsList.querySelectorAll('.document-download').forEach(btn => {
      btn.addEventListener('click', () => {
        const url = btn.getAttribute('data-url');
        window.open(url, '_blank');
      });
    });
  }

  function getFileIcon(filename) {
    if (!filename) return '📄';
    const ext = filename.split('.').pop().toLowerCase();
    const icons = {
      'pdf': '📕',
      'txt': '📄',
      'doc': '📘',
      'docx': '📘'
    };
    return icons[ext] || '📄';
  }

  function formatFileSize(bytes) {
    if (!bytes) return '0 B';
    const kb = bytes / 1024;
    if (kb < 1024) return kb.toFixed(1) + ' KB';
    const mb = kb / 1024;
    return mb.toFixed(1) + ' MB';
  }

  // Expose deleteDocument to window for onclick access
  window.deleteDocument = deleteDocument;
  window.loadDocuments = loadDocuments;

  // Auto-load todos on page load (Action tab)
  loadTodos('starred');
}

// Global deleteDocument function (outside initializeMainInterface)
async function deleteDocument(documentId) {
  if (!confirm('Are you sure you want to delete this document?')) return;

  try {
    const payload = {
      action: 'delete_document',
      document_id: documentId,
      timestamp: new Date().toISOString(),
      source: 'oracle-chrome-extension-newtab',
      user_id: userData?.userId,
      authenticated: true
    };

    const response = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (response.ok) {
      // Call loadDocuments from window if available
      if (window.loadDocuments) {
        window.loadDocuments();
      }
    } else {
      throw new Error('Failed to delete document');
    }
  } catch (error) {
    console.error('Delete error:', error);
    alert('Failed to delete document: ' + error.message);
  }
}

// Multi-select functions
function toggleTodoSelection(todoId, element = null) {
  if (selectedTodoIds.has(todoId)) {
    selectedTodoIds.delete(todoId);
  } else {
    selectedTodoIds.add(todoId);
  }
  updateSelectionUI();
}

function clearMultiSelection() {
  selectedTodoIds.clear();
  isMultiSelectMode = false;
  updateSelectionUI();
}

function updateSelectionUI() {
  // Update todo item visual state
  document.querySelectorAll('.todo-item').forEach(item => {
    const todoId = item.dataset.todoId;
    if (selectedTodoIds.has(parseInt(todoId)) || selectedTodoIds.has(todoId)) {
      item.classList.add('multi-selected');
    } else {
      item.classList.remove('multi-selected');
    }
  });

  // Update meeting item visual state
  document.querySelectorAll('.meeting-item').forEach(item => {
    const meetingId = item.dataset.meetingId;
    if (selectedTodoIds.has(parseInt(meetingId)) || selectedTodoIds.has(meetingId)) {
      item.classList.add('multi-selected');
    } else {
      item.classList.remove('multi-selected');
    }
  });

  // Show/hide bulk action button
  const bulkBtn = document.getElementById('bulkActionBtn');
  if (bulkBtn) {
    if (selectedTodoIds.size > 0) {
      bulkBtn.style.display = 'inline-flex';
      bulkBtn.querySelector('.bulk-count').textContent = selectedTodoIds.size;
    } else {
      bulkBtn.style.display = 'none';
    }
  }
}

async function handleBulkMarkDone() {
  if (selectedTodoIds.size === 0) return;

  const bulkBtn = document.getElementById('bulkActionBtn');
  const selectedIds = new Set(selectedTodoIds); // Copy the set before clearing

  // Get full todo data for selected items (check allTodos, allFyiItems, and allCalendarItems)
  const allItems = [...allTodos, ...allFyiItems, ...allCalendarItems];
  const selectedTodos = allItems.filter(todo =>
    selectedIds.has(todo.id) || selectedIds.has(String(todo.id))
  ).map(todo => ({
    id: todo.id,
    task_title: todo.task_title || '',
    task_name: todo.task_name || '',
    status: todo.status,
    starred: todo.starred
  }));

  // Immediately animate and remove selected items from UI
  selectedIds.forEach(todoId => {
    const todoItem = document.querySelector(`.todo-item[data-todo-id="${todoId}"]`);
    if (todoItem) {
      const checkbox = todoItem.querySelector('.todo-checkbox');
      if (checkbox) {
        checkbox.classList.add('checked');
        checkbox.innerHTML = '✓';
      }
      todoItem.classList.add('completing');

      // Remove from DOM after animation completes
      setTimeout(() => {
        todoItem.remove();
      }, 400);
    }
  });

  // Update local arrays immediately (optimistic update)
  allTodos = allTodos.filter(t => !selectedIds.has(t.id) && !selectedIds.has(String(t.id)));
  allFyiItems = allFyiItems.filter(t => !selectedIds.has(t.id) && !selectedIds.has(String(t.id)));

  // Track as recently completed to prevent reappearing via pending updates
  selectedIds.forEach(todoId => {
    addRecentlyCompleted(todoId);
  });

  // Reset read state so if task comes back, it will be unread (yellow)
  selectedIds.forEach(todoId => {
    markTaskAsUnread(todoId);
  });

  // Clear selection immediately
  clearMultiSelection();

  // Show success toast
  showBulkSuccessToast(selectedTodos.length);

  // Update counts
  updateTabCounts();

  // Send to backend in background (don't await)
  try {
    const payload = {
      action: 'bulk_mark_complete',
      todos: selectedTodos,
      todo_ids: Array.from(selectedIds).map(id => parseInt(id)),
      count: selectedIds.size,
      timestamp: new Date().toISOString(),
      source: 'oracle-chrome-extension-newtab',
      user_id: userData?.userId,
      authenticated: true
    };

    fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(response => {
      if (!response.ok) {
        console.error('Failed to mark todos as done on backend');
      }
    }).catch(error => {
      console.error('Error marking todos as done:', error);
    });
  } catch (error) {
    console.error('Error preparing bulk mark done:', error);
  }
}

function showBulkSuccessToast(count) {
  const toast = document.createElement('div');
  toast.className = 'bulk-success-toast';
  toast.innerHTML = `✓ ${count} task${count > 1 ? 's' : ''} marked as done`;
  toast.style.cssText = `
    position: fixed;
    bottom: 30px;
    left: 50%;
    transform: translateX(-50%);
    background: linear-gradient(45deg, #27ae60, #2ecc71);
    color: white;
    padding: 12px 24px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 500;
    box-shadow: 0 4px 15px rgba(39, 174, 96, 0.4);
    z-index: 10000;
    animation: slideUp 0.3s ease-out;
  `;
  document.body.appendChild(toast);

  setTimeout(() => {
    toast.style.animation = 'slideDown 0.3s ease-out';
    setTimeout(() => toast.remove(), 300);
  }, 2500);
}

document.addEventListener('DOMContentLoaded', async () => {
  // Load read state from localStorage
  loadReadState();
  // Sync local variables from shared state (loadReadState populates state.*)
  const _st = window.Oracle.state;
  readTaskIds = _st.readTaskIds;
  previousTaskTimestamps = _st.previousTaskTimestamps;
  isInitialLoad = _st.isInitialLoad;

  // Attach bulk action button listeners
  const bulkMarkDoneBtn = document.getElementById('bulkMarkDoneBtn');
  const bulkCancelBtn = document.getElementById('bulkCancelBtn');

  if (bulkMarkDoneBtn) {
    bulkMarkDoneBtn.addEventListener('click', handleBulkMarkDone);
  }
  if (bulkCancelBtn) {
    bulkCancelBtn.addEventListener('click', clearMultiSelection);
  }

  // Setup dark mode toggle (outside of initializeMainInterface so it works on login screen too)
  setupDarkModeGlobal();

  // Setup fullscreen toggle
  setupFullscreenToggle();

  // Setup sidebar toggle
  setupSidebarToggle();

  // Setup chat slider
  setupChatSlider();

  // Setup new message slider
  setupNewMessageSlider();

  if (await initAuth()) {
    initializeMainInterface();
    setupProfileOverlay();
  } else {
    showLoginScreen();
  }
});

// Global dark mode setup (works even before login)
function setupDarkModeGlobal() {
  const toggle = document.getElementById('darkModeToggle');
  if (!toggle) {
    console.error('Dark mode toggle not found');
    return;
  }

  const savedMode = localStorage.getItem('oracle_dark_mode');

  // Apply saved preference
  if (savedMode === 'true') {
    document.body.classList.add('dark-mode');
    toggle.textContent = '🌙';
  } else {
    toggle.textContent = '☀️';
  }

  // Use a flag to prevent duplicate listeners
  if (toggle.dataset.listenerAttached) return;
  toggle.dataset.listenerAttached = 'true';

  toggle.addEventListener('click', function handleDarkModeClick(e) {
    e.preventDefault();
    e.stopPropagation();
    const isDark = document.body.classList.toggle('dark-mode');
    this.textContent = isDark ? '🌙' : '☀️';
    localStorage.setItem('oracle_dark_mode', isDark.toString());
    console.log('🌓 Dark mode toggled:', isDark ? 'ON' : 'OFF');
  });
}

// Quick Chat Slider
const CHAT_WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4';
const ABLY_CHAT_API_KEY = 'IROXlg.Pr4FZw:5pU-xl09axAY1_jHcnegOd6aJQBXXCiCfAjVXOAQzZI'; window.ABLY_CHAT_API_KEY = ABLY_CHAT_API_KEY; // Replace with your Ably API key

// Fullscreen toggle setup
function setupFullscreenToggle() {
  const toggle = document.getElementById('fullscreenToggle');
  if (!toggle) {
    console.error('Fullscreen toggle not found');
    return;
  }

  // Use a flag to prevent duplicate listeners
  if (toggle.dataset.listenerAttached) return;
  toggle.dataset.listenerAttached = 'true';

  // Update icon based on current fullscreen state
  function updateFullscreenIcon() {
    if (document.fullscreenElement) {
      toggle.textContent = '⛶';
      toggle.title = 'Exit Fullscreen (F)';
    } else {
      toggle.textContent = '⛶';
      toggle.title = 'Enter Fullscreen (F)';
    }
  }

  toggle.addEventListener('click', async function (e) {
    e.preventDefault();
    e.stopPropagation();

    try {
      if (!document.fullscreenElement) {
        await document.documentElement.requestFullscreen();
        console.log('🖥️ Entered fullscreen mode');
      } else {
        await document.exitFullscreen();
        console.log('🖥️ Exited fullscreen mode');
      }
    } catch (err) {
      console.error('Fullscreen error:', err);
    }
  });

  // Listen for fullscreen changes (including Escape key exit)
  document.addEventListener('fullscreenchange', updateFullscreenIcon);

  // Intercept Escape key to prevent fullscreen exit when sliders are open
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && document.fullscreenElement) {
      const hasOpenSlider = document.querySelector('.transcript-slider-overlay') ||
        document.querySelector('.note-viewer-overlay') ||
        document.querySelector('.chat-slider-overlay') ||
        document.querySelector('.new-message-slider-overlay');
      if (hasOpenSlider) {
        // Prevent the default fullscreen exit behavior
        e.preventDefault();
        e.stopPropagation();
        // The slider's own Escape handler will close it instead
      }
    }
  }, true); // Use capture phase to intercept before browser handles it
}

// Sidebar (Column 3) toggle setup
function setupSidebarToggle() {
  const toggle = document.getElementById('sidebarToggle');
  const col3 = document.getElementById('col3');
  const layout = document.querySelector('.three-column-layout');
  if (!toggle || !col3 || !layout) return;

  if (toggle.dataset.listenerAttached) return;
  toggle.dataset.listenerAttached = 'true';

  // Restore state from localStorage
  const isCollapsed = localStorage.getItem('oracle_sidebar_collapsed') === 'true';
  if (isCollapsed) {
    col3.classList.add('col3-hidden');
    layout.classList.add('col3-collapsed');
    toggle.classList.add('collapsed');
    toggle.title = 'Show Notes/Bookmarks panel (])';
  }

  toggle.addEventListener('click', function (e) {
    e.preventDefault();
    e.stopPropagation();
    const hidden = col3.classList.toggle('col3-hidden');
    layout.classList.toggle('col3-collapsed', hidden);
    toggle.classList.toggle('collapsed', hidden);
    toggle.title = hidden ? 'Show Notes/Bookmarks panel (])' : 'Hide Notes/Bookmarks panel (])';
    localStorage.setItem('oracle_sidebar_collapsed', hidden);
  });

  // Keyboard shortcut: ] to toggle
  document.addEventListener('keydown', function (e) {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable) return;
    if (e.key === ']') {
      toggle.click();
    }
  });
}

// Chat slider & New message — delegated to shared components
function setupChatSlider() {
  const chatToggle = document.getElementById('chatToggle');
  if (!chatToggle || chatToggle.dataset.listenerAttached) return;
  chatToggle.dataset.listenerAttached = 'true';
  chatToggle.addEventListener('click', () => {
    if (isChatSliderOpen) return;
    window.OracleAssistant.showChatSlider({ mode: 'fullscreen', onClose: () => { isChatSliderOpen = false; } });
    isChatSliderOpen = true;
  });
}
function formatChatResponseWithAnnotations(text) { return window.OracleAssistant.formatChatResponseWithAnnotations(text); }
function showChatSlider() { window.OracleAssistant.showChatSlider({ mode: 'fullscreen', onClose: () => { isChatSliderOpen = false; } }); }
function setupNewMessageSlider() {
  const newMsgToggle = document.getElementById('newMessageToggle');
  if (!newMsgToggle || newMsgToggle.dataset.listenerAttached) return;
  newMsgToggle.dataset.listenerAttached = 'true';
  newMsgToggle.addEventListener('click', () => { window.OracleNewMessage.showNewMessageSlider({ mode: 'fullscreen' }); });
}
function showNewMessageSlider() { window.OracleNewMessage.showNewMessageSlider({ mode: 'fullscreen' }); }

// Chat & New Message CSS
const chatStyles = document.createElement('style');
chatStyles.textContent = `
  .typing-dots span { animation: typingDot 1.4s infinite; opacity: 0; }
  .typing-dots span:nth-child(1) { animation-delay: 0s; }
  .typing-dots span:nth-child(2) { animation-delay: 0.2s; }
  .typing-dots span:nth-child(3) { animation-delay: 0.4s; }
  @keyframes typingDot { 0%, 60%, 100% { opacity: 0; } 30% { opacity: 1; } }
  .chat-input:empty:before { content: attr(placeholder); color: #95a5a6; pointer-events: none; }
  .chat-input:focus:empty:before { content: ''; }
`;
document.head.appendChild(chatStyles);

// Listen for Ably push notifications from background.js
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('📩 Newtab received message:', message);

  if (message.action === 'todoListUpdated') {
    console.log('🔄 Auto-refresh requested via Ably notification');

    // RULE 1: Check throttle - don't refresh more than once every 30 seconds
    const now = Date.now();
    const timeSinceLastRefresh = now - lastAblyRefreshTime;
    if (timeSinceLastRefresh < ABLY_REFRESH_THROTTLE_MS) {
      const waitTime = Math.ceil((ABLY_REFRESH_THROTTLE_MS - timeSinceLastRefresh) / 1000);
      console.log(`⏸️ RULE 1 BLOCKED: Throttled (last refresh was ${Math.floor(timeSinceLastRefresh / 1000)}s ago, need to wait ${waitTime}s)`);
      // Schedule a deferred refresh after the throttle window expires
      if (!window._deferredRefreshTimer) {
        const deferMs = (ABLY_REFRESH_THROTTLE_MS - timeSinceLastRefresh) + 500;
        console.log(`⏰ Scheduling deferred refresh in ${Math.ceil(deferMs / 1000)}s`);
        window._deferredRefreshTimer = setTimeout(() => {
          window._deferredRefreshTimer = null;
          console.log('⏰ Deferred refresh executing now');
          performAblyRefresh();
        }, deferMs);
      }
      sendResponse({ received: true, skipped: 'throttled', waitSeconds: waitTime });
      return true;
    }

    // With feed-like UX, we allow silent fetch even when tab has focus
    // The banner will show and user can click to load updates
    console.log('✅ Performing silent fetch for feed-like UX');
    performAblyRefresh();
    sendResponse({ received: true, refreshed: true });
  } else {
    sendResponse({ received: true });
  }

  return true; // Keep message channel open for async response
});

// Function to perform the actual Ably refresh - now with feed-like UX
function performAblyRefresh() {
  // RULE 1: Check throttle
  const now = Date.now();
  const timeSinceLastRefresh = now - lastAblyRefreshTime;
  if (timeSinceLastRefresh < ABLY_REFRESH_THROTTLE_MS) {
    console.log(`⏸️ performAblyRefresh BLOCKED by RULE 1: Throttled (last refresh was ${Math.floor(timeSinceLastRefresh / 1000)}s ago)`);
    return;
  }

  // Update last refresh timestamp BEFORE refreshing
  updateLastRefreshTime();
  console.log(`📊 Ably refresh executing at ${new Date(lastAblyRefreshTime).toLocaleTimeString()}`);

  // Fetch data silently in background, then show "N new updates" banner
  // With feed-like UX, this works even when tab has focus
  performSilentFetch();

  // Clear any pending flag
  pendingRefreshFlag = false;
}

// Silently fetch data and queue updates for feed-like UX
async function performSilentFetch() {
  console.log('🔇 Performing silent fetch for feed-like UX...');

  try {
    // Fetch Action items (starred todos)
    const actionResponse = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload({
        action: 'list_todos',
        filter: 'all',
        timestamp: new Date().toISOString()
      }))
    });

    if (actionResponse.ok) {
      const responseText = await actionResponse.text();
      let actionData = [];
      if (responseText && responseText.trim()) {
        try {
          actionData = JSON.parse(responseText);
        } catch (e) {
          console.warn('Failed to parse action response:', e);
        }
      }

      const newActionTodos = Array.isArray(actionData) ? actionData : (actionData?.todos || []);

      // Filter to only active tasks (status === 0) - exclude completed tasks
      const activeActionTodos = newActionTodos.filter(t => t.status === 0);

      // Calculate changes for Action list (non-meeting items only - includes Drive items)
      const currentActionIds = new Set(allTodos.filter(t => !isMeetingLink(t.message_link)).map(t => t.id));
      const newNonMeetingItems = activeActionTodos.filter(t => !isMeetingLink(t.message_link));

      // Debug: Log Drive items being processed
      const driveItemsInFetch = newNonMeetingItems.filter(t => isDriveLink(t.message_link));
      console.log(`🔍 Action: ${driveItemsInFetch.length} Drive items in fetch, ${newNonMeetingItems.length} total non-meeting items`);

      // Find new items (exclude recently completed tasks AND document/drive items which refresh via accordion)
      const newActionItems = newNonMeetingItems.filter(t => !currentActionIds.has(t.id) && !isRecentlyCompleted(t.id) && !isDriveLink(t.message_link));

      // Find updated items (exclude recently completed AND document/drive items)
      const updatedActionItems = newNonMeetingItems.filter(t => {
        if (isRecentlyCompleted(t.id)) return false;
        if (isDriveLink(t.message_link)) return false;
        const idStr = String(t.id);
        if (!currentActionIds.has(t.id)) return false;
        const prevTimestamp = previousTaskTimestamps.get(idStr);
        return prevTimestamp && t.updated_at && prevTimestamp !== t.updated_at;
      });

      // Debug: Log Drive items in changes
      const newDriveItems = newActionItems.filter(t => isDriveLink(t.message_link));
      const updatedDriveItems = updatedActionItems.filter(t => isDriveLink(t.message_link));
      console.log(`📊 Action changes: ${newDriveItems.length} new Drive, ${updatedDriveItems.length} updated Drive`);

      const totalActionChanges = newActionItems.length + updatedActionItems.length;

      if (totalActionChanges > 0) {
        pendingActionData = activeActionTodos.filter(t => !isRecentlyCompleted(t.id)); // Store only active, non-recently-completed tasks

        // Merge with existing pending updates using Set to deduplicate (exclude recently completed)
        const newUpdateIds = [...newActionItems.map(t => t.id), ...updatedActionItems.map(t => t.id)]
          .filter(id => !isRecentlyCompleted(id));
        const existingIds = new Set(pendingActionUpdates.map(id => String(id)));
        newUpdateIds.forEach(id => existingIds.add(String(id)));
        pendingActionUpdates = Array.from(existingIds);

        // Show banner with unique count
        showPendingUpdatesBanner('action', pendingActionUpdates.length);
        console.log(`📬 Action: ${pendingActionUpdates.length} unique pending updates`);
      }

      // Always refresh Meetings accordion immediately (not queued)
      const newMeetings = activeActionTodos.filter(t => isMeetingLink(t.message_link));
      if (newMeetings.length > 0) {
        refreshMeetingsAccordionOnly(newMeetings, activeActionTodos);
      }
    }

    // Fetch FYI items
    const fyiResponse = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload({
        action: 'list_todos',
        filter: 'all',
        starred: 0,
        status: 0,
        timestamp: new Date().toISOString()
      }))
    });

    if (fyiResponse.ok) {
      const responseText = await fyiResponse.text();
      let fyiData = [];
      if (responseText && responseText.trim()) {
        try {
          fyiData = JSON.parse(responseText);
        } catch (e) {
          console.warn('Failed to parse FYI response:', e);
        }
      }

      const allItems = Array.isArray(fyiData) ? fyiData : (fyiData?.todos || []);
      const newFyiTodos = allItems.filter(t => t.starred === 0 && t.status === 0 && !isMeetingLink(t.message_link));

      console.log(`🔍 FYI silent fetch: got ${newFyiTodos.length} items, current allFyiItems has ${allFyiItems.length}`);

      // Calculate changes for FYI list (non-document items only - documents refresh immediately)
      const currentFyiNonDocIds = new Set(allFyiItems.filter(t => !isDriveLink(t.message_link)).map(t => t.id));
      const newNonDocItems = newFyiTodos.filter(t => !isDriveLink(t.message_link));

      // Find new items (items in new fetch that aren't in current, exclude recently completed)
      const newFyiItems = newNonDocItems.filter(t => !currentFyiNonDocIds.has(t.id) && !isRecentlyCompleted(t.id));

      // Find updated items (existing items with changed timestamp, exclude recently completed)
      const updatedFyiItems = newNonDocItems.filter(t => {
        if (isRecentlyCompleted(t.id)) return false;
        const idStr = String(t.id);
        if (!currentFyiNonDocIds.has(t.id)) return false;
        const prevTimestamp = previousTaskTimestamps.get(idStr);
        return prevTimestamp && t.updated_at && prevTimestamp !== t.updated_at;
      });

      const totalFyiChanges = newFyiItems.length + updatedFyiItems.length;

      console.log(`📊 FYI changes: ${newFyiItems.length} new, ${updatedFyiItems.length} updated = ${totalFyiChanges} total`);

      if (totalFyiChanges > 0) {
        pendingFyiData = newFyiTodos.filter(t => !isRecentlyCompleted(t.id));

        // Merge with existing pending updates using Set to deduplicate (exclude recently completed)
        const newUpdateIds = [...newFyiItems.map(t => t.id), ...updatedFyiItems.map(t => t.id)]
          .filter(id => !isRecentlyCompleted(id));
        const existingIds = new Set(pendingFyiUpdates.map(id => String(id)));
        newUpdateIds.forEach(id => existingIds.add(String(id)));
        pendingFyiUpdates = Array.from(existingIds);

        // Show banner with unique count
        showPendingUpdatesBanner('fyi', pendingFyiUpdates.length);
        console.log(`📬 FYI: ${pendingFyiUpdates.length} unique pending updates`);
      }

      // Always refresh Documents accordion immediately (not queued)
      const newDocs = newFyiTodos.filter(t => isDriveLink(t.message_link));
      if (newDocs.length > 0) {
        refreshDocumentsAccordionOnly(newDocs, newFyiTodos);
      }
    }

  } catch (error) {
    console.error('Error in silent fetch:', error);
  }
}

// Refresh only the Meetings accordion without touching the list
function refreshMeetingsAccordionOnly(meetings, allNewTodos) {
  const container = document.querySelector('.todos-container');
  if (!container) return;

  // Store for accordion
  allCalendarItems = meetings;

  // Rebuild meetings accordion using window-exposed function
  const meetingsAccordionHtml = typeof window.buildMeetingsAccordion === 'function'
    ? window.buildMeetingsAccordion(meetings)
    : '';

  // Try to render in the slot above the column
  const meetingsSlot = document.getElementById('meetingsAccordionSlot');
  if (meetingsSlot) {
    meetingsSlot.innerHTML = meetingsAccordionHtml || '';
    if (meetingsAccordionHtml && typeof window.setupMeetingsAccordion === 'function') {
      window.setupMeetingsAccordion(meetingsSlot);
    }
  } else {
    // Fallback: insert inside container
    let existingAccordion = container.querySelector('.meetings-accordion');
    if (existingAccordion) {
      existingAccordion.outerHTML = meetingsAccordionHtml;
    } else if (meetingsAccordionHtml) {
      const todosList = container.querySelector('.todos-list');
      const pendingBanner = container.querySelector('.pending-updates-banner');
      if (pendingBanner) {
        pendingBanner.insertAdjacentHTML('beforebegin', meetingsAccordionHtml);
      } else if (todosList) {
        todosList.insertAdjacentHTML('beforebegin', meetingsAccordionHtml);
      }
    }
    if (typeof window.setupMeetingsAccordion === 'function') {
      window.setupMeetingsAccordion(container);
    }
  }
  console.log('✅ Meetings accordion refreshed silently');
}

// Refresh only the Documents accordion without touching the list
function refreshDocumentsAccordionOnly(docs, allNewFyiItems) {
  const container = document.querySelector('.fyi-container');
  if (!container) return;

  // Rebuild documents accordion using window-exposed function
  const documentsAccordionHtml = typeof window.buildDocumentsAccordion === 'function'
    ? window.buildDocumentsAccordion(docs, actionTabDriveFileIds)
    : '';

  // Try to render in the slot above the column
  const documentsSlot = document.getElementById('documentsAccordionSlot');
  if (documentsSlot) {
    documentsSlot.innerHTML = documentsAccordionHtml || '';
    if (documentsAccordionHtml && typeof window.setupDocumentsAccordion === 'function') {
      window.setupDocumentsAccordion(documentsSlot);
    }
  } else {
    // Fallback: insert inside container
    let existingAccordion = container.querySelector('.documents-accordion');
    if (existingAccordion) {
      existingAccordion.outerHTML = documentsAccordionHtml;
    } else if (documentsAccordionHtml) {
      const fyiList = container.querySelector('.fyi-list');
      const pendingBanner = container.querySelector('.pending-updates-banner');
      if (pendingBanner) {
        pendingBanner.insertAdjacentHTML('beforebegin', documentsAccordionHtml);
      } else if (fyiList) {
        fyiList.insertAdjacentHTML('beforebegin', documentsAccordionHtml);
      }
    }
    if (typeof window.setupDocumentsAccordion === 'function') {
      window.setupDocumentsAccordion(container);
    }
  }
  console.log('✅ Documents accordion refreshed silently');
}

// ============================================
// Global unified updates badge (next to Oracle subtitle)
// ============================================
function updateGlobalUpdatesBadge() {
  const totalCount = pendingActionUpdates.length + pendingFyiUpdates.length;
  const badge = document.getElementById('globalUpdatesBadge');
  if (!badge) {
    console.log('⚠️ globalUpdatesBadge element not found');
    return;
  }
  console.log(`📢 updateGlobalUpdatesBadge: totalCount=${totalCount} (action=${pendingActionUpdates.length}, fyi=${pendingFyiUpdates.length})`);
  if (totalCount > 0) {
    // Update only the count span text, don't replace innerHTML (preserves click handler)
    let countSpan = badge.querySelector('#globalUpdatesCount');
    if (!countSpan) {
      // If the inner structure got lost, rebuild it without losing the click handler
      badge.textContent = '';
      badge.insertAdjacentHTML('beforeend', `↑ <span id="globalUpdatesCount">${totalCount}</span> new update${totalCount > 1 ? 's' : ''}`);
    } else {
      countSpan.textContent = totalCount;
      // Update the "updates" / "update" text after the count span
      // The text structure is: "↑ {count} new update(s)"
      const textAfterCount = countSpan.nextSibling;
      if (textAfterCount && textAfterCount.nodeType === Node.TEXT_NODE) {
        textAfterCount.textContent = ` new update${totalCount > 1 ? 's' : ''}`;
      }
    }
    badge.style.display = 'inline';
    badge.style.opacity = '1';
  } else {
    badge.style.display = 'none';
  }
}

// Setup global badge click handler (called once on init)
function setupGlobalUpdatesBadge() {
  const badge = document.getElementById('globalUpdatesBadge');
  if (!badge) {
    console.log('⚠️ setupGlobalUpdatesBadge: element not found');
    return;
  }
  console.log('✅ setupGlobalUpdatesBadge: attaching click handler');
  badge.addEventListener('click', () => {
    console.log('🔄 Global badge clicked — applying all pending updates');
    // Show loading state without replacing innerHTML
    badge.textContent = '↻ Refreshing...';
    badge.style.opacity = '0.7';

    // Apply all pending updates for both tabs
    if (pendingActionData) applyPendingUpdates('action');
    if (pendingFyiData) applyPendingUpdates('fyi');

    // Cross-deduplicate: if a task moved from FYI→Action or Action→FYI, remove it from the old list
    const actionIds = new Set(allTodos.map(t => String(t.id)));
    const fyiIds = new Set(allFyiItems.map(t => String(t.id)));

    const fyiBefore = allFyiItems.length;
    const actionBefore = allTodos.length;

    // Remove from FYI anything that's now in Action
    allFyiItems = allFyiItems.filter(t => !actionIds.has(String(t.id)));
    // Remove from Action anything that's now in FYI
    allTodos = allTodos.filter(t => !fyiIds.has(String(t.id)));

    const fyiRemoved = fyiBefore - allFyiItems.length;
    const actionRemoved = actionBefore - allTodos.length;

    if (fyiRemoved > 0 || actionRemoved > 0) {
      console.log(`🔀 Cross-dedup: removed ${fyiRemoved} from FYI (moved to Action), ${actionRemoved} from Action (moved to FYI)`);
      // Re-render the affected lists
      if (fyiRemoved > 0) {
        const fyiContainer = document.querySelector('.fyi-container');
        if (fyiContainer && typeof window.displayFYI === 'function') {
          window.displayFYI(allFyiItems, fyiContainer, []);
        }
      }
      if (actionRemoved > 0 && typeof window.displayTodos === 'function') {
        window.displayTodos(allTodos, currentTodoFilter || 'starred', []);
      }
      if (typeof window.updateTabCounts === 'function') window.updateTabCounts();
    }

    // Hide the badge after a short delay and rebuild inner structure
    setTimeout(() => {
      badge.style.display = 'none';
      badge.style.opacity = '1';
      // Rebuild inner HTML for next use
      badge.innerHTML = '↑ <span id="globalUpdatesCount">0</span> new updates';
      updateGlobalUpdatesBadge(); // Re-check if anything remains
    }, 500);
  });
}

// Show "N new updates" banner
// Also updates the global unified badge next to Oracle subtitle
function showPendingUpdatesBanner(column, count) {
  // Only update the global unified badge — no per-column banners
  updateGlobalUpdatesBadge();
}

// Apply pending updates when banner is clicked
function applyPendingUpdates(column) {
  console.log(`🔄 Applying pending updates for ${column}...`);
  // Update global badge after applying
  setTimeout(() => updateGlobalUpdatesBadge(), 100);

  if (column === 'action' && pendingActionData) {
    // Remove banner from slot or container
    const actionSlot = document.getElementById('actionBannerSlot');
    const container = document.querySelector('.todos-container');
    const banner = actionSlot?.querySelector('.pending-updates-banner') || container?.querySelector('.pending-updates-banner');
    if (banner) {
      banner.classList.add('loading');
      banner.innerHTML = '<div class="pending-updates-content"><span class="pending-icon spinning">↻</span><span class="pending-text">Loading...</span></div>';
    }

    // Process and display the pending data (filter out recently completed tasks)
    const newTodos = pendingActionData.filter(t => !isRecentlyCompleted(t.id));

    // Mark updated tasks as unread
    pendingActionUpdates.forEach(id => {
      readTaskIds.delete(String(id));
    });

    // Update timestamps
    newTodos.forEach(t => {
      if (t.updated_at) {
        previousTaskTimestamps.set(String(t.id), t.updated_at);
      }
    });

    saveReadState();

    // Update global state
    allTodos = newTodos;

    // Store items to highlight before clearing
    const itemsToHighlight = [...pendingActionUpdates];

    // Clear pending state BEFORE displaying
    pendingActionData = null;
    pendingActionUpdates = [];

    // Remove banner immediately
    if (banner) banner.remove();

    // Display with highlighting using window-exposed function
    if (allTodos.length > 0 && typeof window.displayTodos === 'function') {
      window.displayTodos(allTodos, currentTodoFilter || 'starred', itemsToHighlight);
    }

    if (typeof window.updateBadge === 'function') window.updateBadge();
    if (typeof window.updateTabCounts === 'function') window.updateTabCounts();

  } else if (column === 'fyi' && pendingFyiData) {
    // Remove banner from slot or container
    const fyiSlot = document.getElementById('fyiBannerSlot');
    const container = document.querySelector('.fyi-container');
    const banner = fyiSlot?.querySelector('.pending-updates-banner') || container?.querySelector('.pending-updates-banner');
    if (banner) {
      banner.classList.add('loading');
      banner.innerHTML = '<div class="pending-updates-content"><span class="pending-icon spinning">↻</span><span class="pending-text">Loading...</span></div>';
    }

    // Process and display the pending data (filter out recently completed tasks)
    const newFyiItems = pendingFyiData.filter(t => !isRecentlyCompleted(t.id));

    // Mark updated tasks as unread
    pendingFyiUpdates.forEach(id => {
      readTaskIds.delete(String(id));
    });

    // Update timestamps
    newFyiItems.forEach(t => {
      if (t.updated_at) {
        previousTaskTimestamps.set(String(t.id), t.updated_at);
      }
    });

    saveReadState();

    // Update global state
    allFyiItems = newFyiItems;

    // Store items to highlight before clearing
    const itemsToHighlight = [...pendingFyiUpdates];

    // Clear pending state BEFORE displaying
    pendingFyiData = null;
    pendingFyiUpdates = [];

    // Remove banner immediately
    if (banner) banner.remove();

    // Display with highlighting using window-exposed function
    if (allFyiItems.length > 0 && typeof window.displayFYI === 'function') {
      window.displayFYI(allFyiItems, container, itemsToHighlight);
    }

    if (typeof window.updateTabCounts === 'function') window.updateTabCounts();
  }
}

// Clear pending updates (called on manual refresh)
function clearPendingUpdates(column) {
  if (column === 'action' || column === 'all') {
    pendingActionData = null;
    pendingActionUpdates = [];
    const actionSlot = document.getElementById('actionBannerSlot');
    if (actionSlot) actionSlot.innerHTML = '';
    const actionContainer = document.querySelector('.todos-container');
    const actionBanner = actionContainer?.querySelector('.pending-updates-banner');
    if (actionBanner) actionBanner.remove();
  }

  if (column === 'fyi' || column === 'all') {
    pendingFyiData = null;
    pendingFyiUpdates = [];
    const fyiSlot = document.getElementById('fyiBannerSlot');
    if (fyiSlot) fyiSlot.innerHTML = '';
    const fyiContainer = document.querySelector('.fyi-container');
    const fyiBanner = fyiContainer?.querySelector('.pending-updates-banner');
    if (fyiBanner) fyiBanner.remove();
  }
}

// User Profile Management
let userProfileData = null;
let originalProfileData = null; // Store original for change detection
let profileTags = []; // Array of { id: null|number, name: '', description: '' }
let originalTags = []; // Store original tags for comparison (with IDs)

// escapeHtml already available from shared component alias (line 54)

function showToastNotification(message) {
  const toast = document.createElement('div');
  toast.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: linear-gradient(45deg, #667eea, #764ba2);
    color: white;
    padding: 12px 24px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 600;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    z-index: 10001;
    animation: slideDown 0.3s ease-out;
  `;
  toast.textContent = message;
  document.body.appendChild(toast);

  setTimeout(() => {
    toast.style.animation = 'slideUp 0.3s ease-out';
    setTimeout(() => toast.remove(), 300);
  }, 2000);
}

// --- Change Detection ---
function getCurrentFormData() {
  return {
    email_ID: document.getElementById('profileEmail')?.value || '',
    password: document.getElementById('profilePassword')?.value || '',
    fr_api_key: document.getElementById('profileFrApiKey')?.value || '',
    fd_l2_api_key: document.getElementById('profileFdApiKey')?.value || '',
    slack_user_token: document.getElementById('profileSlackToken')?.value || '',
    role_description: document.getElementById('profileRoleDesc')?.value || '',
    muted_channels: userProfileData?.muted_channels || [],
    allowed_public_channels: userProfileData?.allowed_public_channels || []
  };
}

function hasProfileChanged() {
  if (!originalProfileData) return false;

  const current = getCurrentFormData();

  // Check basic fields
  if (current.email_ID !== (originalProfileData.email_ID || '')) return true;
  if (current.password !== (originalProfileData.password || '')) return true;
  if (current.fr_api_key !== (originalProfileData.fr_api_key || '')) return true;
  if (current.fd_l2_api_key !== (originalProfileData.fd_l2_api_key || '')) return true;
  if (current.slack_user_token !== (originalProfileData.slack_user_token || '')) return true;
  if (current.role_description !== (originalProfileData.role_description || '')) return true;

  // Check muted channels
  const origChannels = originalProfileData.muted_channels || [];
  const currChannels = current.muted_channels || [];
  if (JSON.stringify(origChannels) !== JSON.stringify(currChannels)) return true;

  // Check allowed public channels
  const origAllowed = originalProfileData.allowed_public_channels || [];
  const currAllowed = current.allowed_public_channels || [];
  if (JSON.stringify(origAllowed) !== JSON.stringify(currAllowed)) return true;

  // Check tags
  const cleanCurrentTags = profileTags.filter(t => t.name.trim() !== '');
  if (JSON.stringify(cleanCurrentTags) !== JSON.stringify(originalTags)) return true;

  return false;
}

function updateSaveButtonVisibility() {
  const saveBtn = document.getElementById('profileSaveBtn');
  if (!saveBtn) return;

  if (hasProfileChanged()) {
    saveBtn.classList.add('visible');
  } else {
    saveBtn.classList.remove('visible');
  }
}

// --- Tags Table (Notion-style) ---
function renderTagsTable() {
  const tbody = document.getElementById('tagsTableBody');
  if (!tbody) return;
  tbody.innerHTML = profileTags.map((tag, i) => `
    <div class="tags-table-row" data-index="${i}">
      <textarea placeholder="Tag name…" data-field="name" rows="1">${escapeHtml(tag.name)}</textarea>
      <textarea placeholder="Description…" data-field="description" rows="1">${escapeHtml(tag.description)}</textarea>
      <button class="tags-row-remove" data-index="${i}">×</button>
    </div>
  `).join('');

  // Bind textarea changes
  tbody.querySelectorAll('textarea').forEach(textarea => {
    textarea.addEventListener('input', (e) => {
      const row = e.target.closest('.tags-table-row');
      const idx = parseInt(row.dataset.index);
      const field = e.target.dataset.field;
      if (profileTags[idx]) profileTags[idx][field] = e.target.value;
      updateSaveButtonVisibility();
    });
  });

  // Bind remove buttons
  tbody.querySelectorAll('.tags-row-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt(btn.dataset.index);
      profileTags.splice(idx, 1);
      renderTagsTable();
      updateSaveButtonVisibility();
    });
  });
}

function addTagRow() {
  profileTags.push({ id: null, name: '', description: '' });
  renderTagsTable();
  updateSaveButtonVisibility();
  // Focus the new name input
  const tbody = document.getElementById('tagsTableBody');
  const lastRow = tbody.lastElementChild;
  if (lastRow) lastRow.querySelector('textarea[data-field="name"]').focus();
}

async function loadUserProfile() {
  const profileBody = document.getElementById('profileBody');

  // Show loading state
  if (profileBody) profileBody.classList.add('loading');

  try {
    const response = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload({
        action: 'get_user_profile',
        timestamp: new Date().toISOString()
      }))
    });

    const data = await response.json();
    console.log('User profile response:', data);

    // Handle both array and object responses
    if (Array.isArray(data) && data.length > 0) {
      userProfileData = data[0];
    } else if (data && typeof data === 'object') {
      userProfileData = data;
    }

    if (userProfileData) {
      console.log('Displaying profile:', userProfileData);

      // Parse allowed_public_channels from JSON string if needed
      if (typeof userProfileData.allowed_public_channels === 'string') {
        try {
          userProfileData.allowed_public_channels = JSON.parse(userProfileData.allowed_public_channels);
        } catch (e) {
          userProfileData.allowed_public_channels = [];
        }
      }
      if (!Array.isArray(userProfileData.allowed_public_channels)) {
        userProfileData.allowed_public_channels = [];
      }

      // Load tags from profile data (include IDs for tracking changes)
      if (Array.isArray(userProfileData.tags)) {
        profileTags = userProfileData.tags.map(t => ({
          id: t.id || null,
          name: t.name || '',
          description: t.description || ''
        }));
      } else {
        profileTags = [];
      }
      // Store original data for change detection
      originalProfileData = {
        email_ID: userProfileData.email_ID || '',
        password: userProfileData.password || '',
        fr_api_key: userProfileData.fr_api_key || '',
        fd_l2_api_key: userProfileData.fd_l2_api_key || '',
        slack_user_token: userProfileData.slack_user_token || '',
        role_description: userProfileData.role_description || '',
        muted_channels: [...(userProfileData.muted_channels || [])],
        allowed_public_channels: [...(userProfileData.allowed_public_channels || [])]
      };
      // Store original tags with IDs for change tracking
      originalTags = profileTags.filter(t => t.name.trim() !== '').map(t => ({ ...t }));

      displayUserProfile();
      updateSaveButtonVisibility();
    } else {
      console.error('No profile data found');
    }
  } catch (error) {
    console.error('Error loading user profile:', error);
  } finally {
    // Hide loading state
    if (profileBody) profileBody.classList.remove('loading');
  }
}

function displayUserProfile() {
  if (!userProfileData) {
    console.error('No user profile data to display');
    return;
  }

  console.log('Setting email:', userProfileData.email_ID);
  console.log('Setting password:', userProfileData.password);

  const emailField = document.getElementById('profileEmail');
  const passwordField = document.getElementById('profilePassword');

  console.log('Email field:', emailField);
  console.log('Password field:', passwordField);

  if (emailField) emailField.value = userProfileData.email_ID || '';
  if (passwordField) passwordField.value = userProfileData.password || '';

  document.getElementById('profileFrApiKey').value = userProfileData.fr_api_key || '';
  document.getElementById('profileFdApiKey').value = userProfileData.fd_l2_api_key || '';
  document.getElementById('profileSlackToken').value = userProfileData.slack_user_token || '';
  document.getElementById('profileRoleDesc').value = userProfileData.role_description || '';

  console.log('All fields set successfully');

  // Display muted channels as tags
  displayChannelTags(userProfileData.muted_channels || []);

  // Display allowed public channels as tags
  displayAllowedChannelTags(userProfileData.allowed_public_channels || []);

  // Render tags table
  renderTagsTable();
}

function displayChannelTags(channels) {
  const container = document.getElementById('profileChannelTags');
  const channelsArray = Array.isArray(channels) ? channels : (channels ? [channels] : []);

  container.innerHTML = channelsArray.map(channel => `
    <div class="profile-tag">
      ${escapeHtml(channel)}
      <span class="profile-tag-remove" data-channel="${escapeHtml(channel)}">×</span>
    </div>
  `).join('');

  // Add remove listeners
  container.querySelectorAll('.profile-tag-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const channel = btn.getAttribute('data-channel');
      removeChannel(channel);
    });
  });
}

function removeChannel(channel) {
  if (!userProfileData.muted_channels) userProfileData.muted_channels = [];
  const channels = Array.isArray(userProfileData.muted_channels)
    ? userProfileData.muted_channels
    : [userProfileData.muted_channels];

  userProfileData.muted_channels = channels.filter(ch => ch !== channel);
  displayChannelTags(userProfileData.muted_channels);
  updateSaveButtonVisibility();
}

function displayAllowedChannelTags(channels) {
  const container = document.getElementById('profileAllowedChannelTags');
  if (!container) return;
  const channelsArray = Array.isArray(channels) ? channels : (channels ? [channels] : []);

  container.innerHTML = channelsArray.map(channel => `
    <div class="profile-tag" style="background: rgba(76, 175, 80, 0.2); border-color: rgba(76, 175, 80, 0.4);">
      ${escapeHtml(channel)}
      <span class="profile-tag-remove" data-channel="${escapeHtml(channel)}">×</span>
    </div>
  `).join('');

  // Add remove listeners
  container.querySelectorAll('.profile-tag-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const channel = btn.getAttribute('data-channel');
      removeAllowedChannel(channel);
    });
  });
}

function removeAllowedChannel(channel) {
  if (!userProfileData.allowed_public_channels) userProfileData.allowed_public_channels = [];
  const channels = Array.isArray(userProfileData.allowed_public_channels)
    ? userProfileData.allowed_public_channels
    : [userProfileData.allowed_public_channels];

  userProfileData.allowed_public_channels = channels.filter(ch => ch !== channel);
  displayAllowedChannelTags(userProfileData.allowed_public_channels);
  updateSaveButtonVisibility();
}

async function saveUserProfile() {
  // Filter out empty tag rows
  const cleanTags = profileTags.filter(t => t.name.trim() !== '');

  // Calculate tag changes
  const originalTagIds = new Set(originalTags.map(t => t.id).filter(id => id !== null));
  const currentTagIds = new Set(cleanTags.map(t => t.id).filter(id => id !== null));

  // New tags: have no ID (id is null)
  const tagsAdded = cleanTags
    .filter(t => t.id === null)
    .map(t => ({ name: t.name, description: t.description }));

  // Modified tags: exist in both original and current, but content changed
  const tagsModified = cleanTags
    .filter(t => t.id !== null && originalTagIds.has(t.id))
    .filter(t => {
      const original = originalTags.find(ot => ot.id === t.id);
      return original && (original.name !== t.name || original.description !== t.description);
    })
    .map(t => ({ id: t.id, name: t.name, description: t.description }));

  // Removed tags: exist in original but not in current
  const tagsRemoved = originalTags
    .filter(t => t.id !== null && !currentTagIds.has(t.id))
    .map(t => ({ id: t.id, name: t.name }));

  // Get updated values
  const updatedData = {
    action: 'update_user_profile',
    user_id: userProfileData.id,
    email_ID: document.getElementById('profileEmail').value,
    password: document.getElementById('profilePassword').value,
    fr_api_key: document.getElementById('profileFrApiKey').value,
    fd_l2_api_key: document.getElementById('profileFdApiKey').value,
    slack_user_token: document.getElementById('profileSlackToken').value,
    role_description: document.getElementById('profileRoleDesc').value,
    muted_channels: userProfileData.muted_channels || [],
    allowed_public_channels: userProfileData.allowed_public_channels || [],
    tags_added: tagsAdded,
    tags_modified: tagsModified,
    tags_removed: tagsRemoved,
    timestamp: new Date().toISOString()
  };

  console.log('Saving profile with tag changes:', { tagsAdded, tagsModified, tagsRemoved });

  try {
    await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload(updatedData))
    });

    showToastNotification('Profile saved!');

    // Reload profile data to get updated tags with new IDs from server
    await loadUserProfile();

  } catch (error) {
    console.error('Error saving profile:', error);
    alert('Failed to save profile. Please try again.');
  }
}

function closeProfileSlider() {
  const profileOverlay = document.getElementById('profileOverlay');
  profileOverlay.classList.remove('show');
  profileOverlay.style.display = 'none';
  window.Oracle.collapseCol3AfterSlider();
}

function setupProfileOverlay() {
  const profileIcon = document.getElementById('userProfileIcon');
  const profileOverlay = document.getElementById('profileOverlay');
  const closeBtn = document.getElementById('profileCloseBtn');
  const saveBtn = document.getElementById('profileSaveBtn');
  const channelInput = document.getElementById('profileChannelInput');
  const allowedChannelInput = document.getElementById('profileAllowedChannelInput');
  const addTagBtn = document.getElementById('addTagRowBtn');

  // Ensure required elements exist
  if (!profileIcon || !profileOverlay) {
    console.error('Profile overlay elements not found, retrying in 100ms');
    setTimeout(setupProfileOverlay, 100);
    return;
  }

  // Prevent duplicate listeners
  if (profileIcon.dataset.listenerAttached) return;
  profileIcon.dataset.listenerAttached = 'true';

  // Track changes on all profile inputs
  const profileInputs = ['profileEmail', 'profilePassword', 'profileFrApiKey', 'profileFdApiKey', 'profileSlackToken', 'profileRoleDesc'];
  profileInputs.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', updateSaveButtonVisibility);
    }
  });

  // Open profile slider — overlay on col3 like transcript slider
  profileIcon.addEventListener('click', () => {
    loadUserProfile();
    // Position overlay to cover col3 only
    const col3El = document.getElementById('col3');
    const col3Rect = window.Oracle.getCol3Rect();
    const profileContainer = profileOverlay.querySelector('.profile-container');
    if (col3Rect && profileContainer) {
      profileOverlay.style.cssText = `position:fixed;top:${col3Rect.top}px;left:${col3Rect.left}px;width:${col3Rect.width}px;bottom:0;z-index:10000;pointer-events:auto;display:flex;animation:fadeIn 0.15s ease-out;`;
      profileContainer.style.cssText = `position:relative;width:100%;height:100%;background:var(--bg-primary);box-shadow:-4px 0 20px rgba(0,0,0,0.15);display:flex;flex-direction:column;overflow:hidden;border-radius:12px;border:1px solid rgba(225,232,237,0.6);transform:none;animation:slideInRight 0.3s ease-out;`;
      if (document.body.classList.contains('dark-mode')) {
        profileContainer.style.background = 'var(--bg-card)';
        profileContainer.style.boxShadow = '-4px 0 20px rgba(0,0,0,0.4)';
        profileContainer.style.borderColor = 'rgba(255,255,255,0.08)';
      }
    }
    profileOverlay.classList.add('show');
  });

  // Close profile slider
  if (closeBtn) {
    closeBtn.addEventListener('click', closeProfileSlider);
  }

  // Close on overlay click (outside panel)
  profileOverlay.addEventListener('click', (e) => {
    if (e.target === profileOverlay) {
      closeProfileSlider();
    }
  });

  // Close on Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && profileOverlay.classList.contains('show')) {
      closeProfileSlider();
    }
  });

  // Save button click
  if (saveBtn) {
    saveBtn.addEventListener('click', () => {
      if (hasProfileChanged()) {
        saveUserProfile();
      }
    });
  }

  // Add tag row button
  if (addTagBtn) {
    addTagBtn.addEventListener('click', addTagRow);
  }

  // Channel input - add on Enter
  if (channelInput) {
    channelInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const channel = channelInput.value.trim();
        if (channel) {
          if (!userProfileData.muted_channels) {
            userProfileData.muted_channels = [];
          }
          const channels = Array.isArray(userProfileData.muted_channels)
            ? userProfileData.muted_channels
            : [userProfileData.muted_channels];

          if (!channels.includes(channel)) {
            channels.push(channel);
            userProfileData.muted_channels = channels;
            displayChannelTags(channels);
            updateSaveButtonVisibility();
          }
          channelInput.value = '';
        }
      }
    });
  }

  // Allowed public channel input - add on Enter
  if (allowedChannelInput) {
    allowedChannelInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const channel = allowedChannelInput.value.trim();
        if (channel) {
          if (!userProfileData.allowed_public_channels) {
            userProfileData.allowed_public_channels = [];
          }
          const channels = Array.isArray(userProfileData.allowed_public_channels)
            ? userProfileData.allowed_public_channels
            : [userProfileData.allowed_public_channels];

          if (!channels.includes(channel)) {
            channels.push(channel);
            userProfileData.allowed_public_channels = channels;
            displayAllowedChannelTags(channels);
            updateSaveButtonVisibility();
          }
          allowedChannelInput.value = '';
        }
      }
    });
  }
}

// setupProfileOverlay is now called from DOMContentLoaded above
